"""
程序入口 - 基金净值匹配工具
从源文件夹(CSV)读取数据，更新目标文件(Excel)
"""
import sys
import os
from datetime import datetime
from config_manager import ConfigManager
from nav_updater import NAVUpdater


def main():
    """主函数"""
    try:
        print("=" * 80)
        print("基金净值匹配工具 v1.0")
        print("=" * 80)

        # 加载配置
        config = ConfigManager('config.ini')

        # 确保必要目录存在
        os.makedirs(config.get_path('backup_folder'), exist_ok=True)
        os.makedirs(config.get_path('log_folder'), exist_ok=True)

        # 显示配置信息
        print(f"\n📋 当前配置:")
        print(f"   源文件夹(CSV): {config.get_path('source_folder')}")
        print(f"   目标文件(Excel): {config.get_path('target_file')}")
        print(f"   工作表: {', '.join(config.get_sheet_names())}")

        # 创建更新器并运行
        updater = NAVUpdater(config)
        success = updater.run()

        if success:
            print("\n✅ 程序执行成功！")
        else:
            print("\n⚠️ 程序执行完成，但遇到一些问题，请查看上方信息。")

    except KeyboardInterrupt:
        print("\n\n⚠️ 用户中断程序执行")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 程序执行失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()