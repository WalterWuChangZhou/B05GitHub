"""
NAV更新器 - 核心业务逻辑
从源文件夹(CSV)读取数据，更新目标文件(Excel)
"""
import pandas as pd
import os
import shutil
import win32com.client as win32
from datetime import datetime
from column_identifier import ColumnIdentifier


class NAVUpdater:
    """NAV更新器类"""

    def __init__(self, config):
        """
        初始化NAV更新器

        Args:
            config: ConfigManager实例
        """
        self.config = config
        self.source_folder = config.get_path('source_folder')  # CSV源文件夹
        self.target_file = config.get_path('target_file')  # Excel目标文件
        self.column_identifier = ColumnIdentifier(config)
        self.matching_config = config.get_matching_config()
        self.logging_config = config.get_logging_config()
        self.backup_config = config.get_backup_config()
        self.sheet_names = config.get_sheet_names()

        self.nav_tolerance = self.matching_config['nav_tolerance']
        self.date_match_mode = self.matching_config['date_match_mode']
        self.show_debug = self.logging_config['show_debug']
        self.max_display = self.logging_config['max_display_records']

    def run(self):
        """执行更新"""
        print("=" * 80)
        print("基金净值匹配工具 - 智能列识别版")
        print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"源文件夹(CSV): {self.source_folder}")
        print(f"目标文件(Excel): {self.target_file}")
        print("=" * 80)

        # 验证文件夹和文件
        if not os.path.exists(self.source_folder):
            print(f"❌ 错误: 源文件夹不存在: {self.source_folder}")
            return False

        if not os.path.exists(self.target_file):
            print(f"❌ 错误: 目标文件不存在: {self.target_file}")
            return False

        # 自动备份目标文件
        if self.backup_config['auto_backup']:
            self._backup_file()

        total_updates = {}

        for sheet_name in self.sheet_names:
            try:
                print("\n" + "=" * 80)
                print(f"处理【{sheet_name}】工作表")
                print("=" * 80)

                # 读取目标文件中的数据
                df = pd.read_excel(self.target_file, sheet_name=sheet_name, engine='openpyxl')
                print(f"✅ 成功读取工作表，共 {len(df)} 行")

                # 智能识别列
                col_mapping = self.column_identifier.identify_columns(df, sheet_name)

                if 'code_col' not in col_mapping or 'date_col' not in col_mapping or 'nav_col' not in col_mapping:
                    print(f"❌ 无法识别必要的列，跳过【{sheet_name}】")
                    total_updates[sheet_name] = 0
                    continue

                # 执行更新（从源文件夹读取CSV，匹配后更新df）
                df, need_update = self._update_nav(df, col_mapping, sheet_name)

                if need_update > 0:
                    # 写入目标文件
                    self._write_to_excel(sheet_name, df, col_mapping, need_update)
                    self._show_update_result(sheet_name, df, col_mapping)
                else:
                    print(f"\n✅ 【{sheet_name}】所有NAV数据都是最新的，无需更新！")

                total_updates[sheet_name] = need_update

            except Exception as e:
                print(f"❌ 处理【{sheet_name}】失败: {e}")
                total_updates[sheet_name] = 0

        # 总结
        self._show_summary(total_updates)
        return True

    def _backup_file(self):
        """备份目标文件"""
        try:
            backup_folder = self.config.get_path('backup_folder')
            os.makedirs(backup_folder, exist_ok=True)

            # 生成备份文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            base_name = os.path.basename(self.target_file)
            name, ext = os.path.splitext(base_name)
            backup_name = f"{name}_backup_{timestamp}{ext}"
            backup_path = os.path.join(backup_folder, backup_name)

            # 复制文件
            shutil.copy2(self.target_file, backup_path)
            print(f"✅ 备份文件已创建: {backup_path}")

            # 清理旧备份
            self._clean_old_backups(backup_folder)

        except Exception as e:
            print(f"⚠️ 备份失败: {e}")

    def _clean_old_backups(self, backup_folder):
        """清理旧的备份文件"""
        try:
            backup_count = self.backup_config['backup_count']
            files = [f for f in os.listdir(backup_folder) if f.endswith('.xlsm')]
            files.sort(key=lambda x: os.path.getmtime(os.path.join(backup_folder, x)))

            # 删除多余的旧备份
            while len(files) > backup_count:
                old_file = files.pop(0)
                os.remove(os.path.join(backup_folder, old_file))
                print(f"   🗑️ 删除旧备份: {old_file}")

        except Exception as e:
            print(f"   ⚠️ 清理旧备份失败: {e}")

    def _update_nav(self, df, col_mapping, sheet_name):
        """
        更新NAV数据
        从源文件夹读取CSV文件，匹配NAV值
        """
        code_col = col_mapping['code_col']
        date_col = col_mapping['date_col']
        nav_col = col_mapping['nav_col']

        print(f"\n📌 使用的列:")
        print(f"   代码列: {code_col}")
        print(f"   日期列: {date_col}")
        print(f"   NAV列: {nav_col}")

        # 转换日期
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df['date_normalized'] = df[date_col].dt.normalize()
        valid_dates = df[date_col].notna().sum()
        print(f"📊 有效日期: {valid_dates}/{len(df)}")

        print(f"\n正在从源文件夹读取CSV文件匹配NAV数据...")
        print(f"源文件夹: {self.source_folder}")

        df["CSV_NAV"] = None
        df["需要更新"] = False
        df["更新原因"] = ""

        match_count = 0
        need_update_count = 0
        debug_shown = False

        for idx, row in df.iterrows():
            code = str(row[code_col]) if pd.notna(row[code_col]) else None
            target_date = row['date_normalized'] if pd.notna(row['date_normalized']) else None
            current_nav = row[nav_col] if pd.notna(row[nav_col]) else None

            if code is None or target_date is None:
                df.at[idx, "更新原因"] = "缺少代码或日期"
                continue

            # 清理代码
            code_clean = self._clean_code(code)
            csv_path = os.path.join(self.source_folder, f"{code_clean}.csv")

            if not os.path.exists(csv_path):
                df.at[idx, "更新原因"] = "CSV文件不存在"
                continue

            try:
                csv_df = pd.read_csv(csv_path)

                # 智能识别CSV中的列
                csv_cols = self.column_identifier.identify_columns(
                    csv_df, f"{code_clean}.csv"
                )

                date_col_csv = csv_cols.get('date_col', csv_df.columns[0])
                nav_col_csv = csv_cols.get('nav_col',
                                           csv_df.columns[1] if len(csv_df.columns) > 1 else csv_df.columns[0])

                # 转换日期
                csv_df[date_col_csv] = pd.to_datetime(csv_df[date_col_csv], errors='coerce')
                csv_df['date_normalized'] = csv_df[date_col_csv].dt.normalize()

                # 调试信息
                if self.show_debug and not debug_shown:
                    print(f"\n🔍 调试信息 (基金: {code_clean}, 日期: {target_date.date()}):")
                    print(f"   CSV日期范围: {csv_df['date_normalized'].min()} 到 {csv_df['date_normalized'].max()}")
                    debug_shown = True

                # 匹配日期
                if self.date_match_mode == 'date_only':
                    matched = csv_df[csv_df['date_normalized'] == target_date]
                else:
                    matched = csv_df[csv_df[date_col_csv] == target_date]

                if matched.empty:
                    df.at[idx, "更新原因"] = f"CSV中未找到日期 {target_date.strftime('%Y/%m/%d')}"
                    continue

                match_count += 1
                csv_nav = matched.iloc[0][nav_col_csv]
                df.at[idx, "CSV_NAV"] = csv_nav

                # 判断是否需要更新
                need_update = False
                if current_nav is None or pd.isna(current_nav):
                    need_update = True
                    reason = f"NAV为空 → 更新为 {csv_nav}"
                else:
                    try:
                        if abs(float(current_nav) - float(csv_nav)) > self.nav_tolerance:
                            need_update = True
                            reason = f"NAV不一致 ({current_nav} → {csv_nav})"
                        else:
                            reason = f"NAV已是最新 ({current_nav})"
                    except:
                        need_update = True
                        reason = f"NAV格式异常 → 更新为 {csv_nav}"

                if need_update:
                    df.at[idx, "需要更新"] = True
                    df.at[idx, "更新原因"] = reason
                    need_update_count += 1
                else:
                    df.at[idx, "更新原因"] = reason

            except Exception as e:
                df.at[idx, "更新原因"] = f"读取失败: {str(e)[:30]}"

        print(f"\n📊 匹配统计:")
        print(f"   成功匹配CSV文件: {match_count} 条")
        print(f"   需要更新NAV: {need_update_count} 条")
        print(f"   NAV已是最新: {match_count - need_update_count} 条")

        # 显示未匹配的记录
        no_match_df = df[(df["CSV_NAV"].isna()) &
                         (df["更新原因"].str.contains("未找到日期", na=False))]
        if len(no_match_df) > 0:
            print(f"\n⚠️ 未匹配到CSV日期的记录（前5条）:")
            for i in range(min(5, len(no_match_df))):
                row = no_match_df.iloc[i]
                date_str = row[date_col].strftime('%Y/%m/%d') if pd.notna(row[date_col]) else "未知"
                print(f"   {row[code_col]} - {date_str} - {row['更新原因']}")

        if need_update_count > 0:
            print(f"\n📝 需要更新NAV的记录:")
            update_df = df[df["需要更新"] == True]
            for i in range(min(self.max_display, len(update_df))):
                row = update_df.iloc[i]
                date_str = row[date_col].strftime('%Y/%m/%d') if pd.notna(row[date_col]) else "未知"
                print(f"   {row[code_col]} - {date_str} - {row['更新原因']}")

        return df, need_update_count

    def _clean_code(self, code):
        """清理基金代码"""
        code_clean = code.strip()
        if '.' in code_clean:
            code_clean = code_clean.split('.')[0]
        code_clean = code_clean.replace('.', '').replace('OF', '').strip()
        return code_clean

    def _write_to_excel(self, sheet_name, df, col_mapping, need_update_count):
        """
        写入Excel目标文件
        """
        if need_update_count == 0:
            return

        nav_col = col_mapping['nav_col']
        nav_col_idx = df.columns.get_loc(nav_col)

        print(f"\n正在通过 Excel COM 写入目标文件【{self.target_file}】的【{sheet_name}】工作表...")

        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            # 打开目标文件
            wb = excel.Workbooks.Open(self.target_file)
            ws = wb.Sheets(sheet_name)

            print("✅ 成功打开目标文件")

            nav_col_excel = nav_col_idx + 1
            write_count = 0

            for idx, row in df.iterrows():
                if row["需要更新"] == True:
                    csv_nav = row["CSV_NAV"]
                    if pd.notna(csv_nav):
                        excel_row = idx + 2
                        ws.Cells(excel_row, nav_col_excel).Value = float(csv_nav)
                        write_count += 1

                        if write_count % 10 == 0:
                            print(f"✅ 已写入 {write_count} 条")

            print(f"\n📊 写入统计: 共更新 {write_count} 条NAV数据")

            if write_count > 0:
                print("\n正在保存目标文件...")
                wb.Save()
                print("✅ 保存完成")

            wb.Close()
            excel.Quit()

            print(f"\n✅ 【{sheet_name}】更新完成")

        except Exception as e:
            print(f"❌ 写入失败: {e}")
            try:
                excel.Quit()
            except:
                pass

    def _show_update_result(self, sheet_name, df, col_mapping):
        """显示更新结果"""
        update_df = df[df["需要更新"] == True]
        if len(update_df) == 0:
            return

        print(f"\n{'=' * 80}")
        print(f"【{sheet_name}】更新结果")
        print('=' * 80)

        code_col = col_mapping['code_col']
        date_col = col_mapping['date_col']
        nav_col = col_mapping['nav_col']

        for i in range(min(self.max_display, len(update_df))):
            row = update_df.iloc[i]
            excel_row = df[df.index == row.name].index[0] + 2
            code = row[code_col]
            date_str = row[date_col].strftime('%Y/%m/%d') if pd.notna(row[date_col]) else "未知"
            old_nav = row[nav_col] if pd.notna(row[nav_col]) else "空"
            new_nav = row["CSV_NAV"] if pd.notna(row["CSV_NAV"]) else "空"

            print(f"  行{excel_row}: {code} - {date_str} | NAV: {old_nav} → {new_nav}")

    def _show_summary(self, total_updates):
        """显示总结"""
        print("\n" + "=" * 80)
        print("✅ 所有操作完成！")
        print(f"   源文件夹(CSV): {self.source_folder}")
        print(f"   目标文件(Excel): {self.target_file}")
        for sheet_name, count in total_updates.items():
            print(f"   - {sheet_name}页面共更新 {count} 条NAV数据")
        print(f"   - 备份文件保存在: {self.config.get_path('backup_folder')}")
        print(f"   - 请刷新目标文件查看更新时间")
        print("=" * 80)