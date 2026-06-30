import openpyxl
import statistics
import math
from collections import defaultdict

wb = openpyxl.load_workbook(r'E:\2026年LOF基金估值偏差记录.xlsm', data_only=True, keep_vba=True)
ws = wb['161124']

data = []
for row in ws.iter_rows(min_row=2, max_col=12, values_only=True):
    d = {
        'date': row[0], 'idx': row[1], 'fund': row[2], 'amt': row[3],
        'nav': row[4], 'share': row[5], 'hkd': row[6], 'add': row[7],
        'post': row[8], 'pm': row[9], 'iopv': row[10], 'br': row[11]
    }
    if d['nav'] is None:
        continue
    data.append(d)

fee_annual = 0.01  # 1% total (mgmt + custody)
fee_daily = fee_annual / 365

# ---- CURRENT MODEL (baseline) ----
current_brs = []
for i in range(1, len(data)):
    d = data[i]
    p = data[i-1]
    if p['br'] is None:
        # first data point has no IOPV
        continue
    if d['br'] is not None:
        current_brs.append(d['br'])

print('='*70)
print('CURRENT MODEL (fixed Post=0.95, 1%/365 fee, idx*FX)')
print('='*70)
print('N = %d' % len(current_brs))
print('Mean BR = %+.2f bp' % (sum(current_brs)/len(current_brs)))
print('Stdev BR = %.2f bp' % statistics.stdev(current_brs))
print('Mean |BR| = %.2f bp' % (sum(abs(b) for b in current_brs)/len(current_brs)))
print('Max |BR| = %.1f bp' % max(abs(b) for b in current_brs))
within10 = sum(1 for b in current_brs if abs(b) <= 10)
within20 = sum(1 for b in current_brs if abs(b) <= 20)
print('Days |BR|<=10bp: %d/%d (%.1f%%)' % (within10, len(current_brs), 100*within10/len(current_brs)))
print('Days |BR|<=20bp: %d/%d (%.1f%%)' % (within20, len(current_brs), 100*within20/len(current_brs)))

# ---- IMPROVED MODEL ----
# Factor 1: Dynamic equity weight with subscription cash drag
# Factor 2: FX effect (same as current for cross-border)
# Factor 3: Cash return on non-equity portion
# Factor 4: Fee accrual (same as current)
# Factor 5: Dividend adjustment (estimated from residual)
# Factor 6: Volatility-dependent tracking error
# Factor 7: Holiday gap detection

# First, compute the cash deployment effect
# When subscription happens, cash inflow = add_share * NAV_prev
# Assume cash is deployed over N_deploy days
N_deploy = 3  # cash deployed over 3 days

# Track uninvested cash ratio (as fraction of total assets)
uninvested_ratio = 0.0
cash_rate_daily = 0.015 / 365  # 1.5% annual for cash

improved_brs = []

# We need shares(t-1) for each step
# data[0] is the first date, we start computing from data[1]
uninvested_history = [0.0]  # for data[0]

for i in range(1, len(data)):
    d = data[i]
    p = data[i-1]
    
    if p['nav'] is None or d['nav'] is None:
        uninvested_history.append(0.0)
        continue
    
    # Previous day's uninvested ratio
    u_prev = uninvested_history[-1]
    
    # Today's subscription effect
    # add_share is in million shares unit? Let me check
    # From data: share = 30,971,114 and add_share values are like 24.48, -13.47
    # So add_share is in millions? Or... let me check
    # 2026-01-30: share = 40,459,721, prev share = 27,383,378
    # Change = 13,076,343 shares. add_share = 1307.63
    # 13,076,343 / 10000 = 1,307.63... so add_share is in 万份 (10,000 shares)
    
    # Actually looking more carefully:
    # 2026-01-30: share 40,459,721 - prev 27,383,378 = 13,076,343
    # add_share = 1307.63 -> 1307.63 * 10000 = 13,076,300. Close!
    # So add_share is in units of 万份 (10,000 shares)
    
    # Cash inflow from subscription (at yesterday's NAV)
    # add_share (in 万份) * 10000 * NAV_prev
    if d['add'] is not None:
        cash_inflow = d['add'] * 10000 * p['nav']
    else:
        cash_inflow = 0
    
    # Total assets at yesterday's NAV
    total_assets_prev = p['share'] * p['nav']
    
    if total_assets_prev > 0:
        # New subscription cash as ratio of total assets
        new_sub_ratio = cash_inflow / total_assets_prev if cash_inflow > 0 else 0
        new_red_ratio = -cash_inflow / total_assets_prev if cash_inflow < 0 else 0
    else:
        new_sub_ratio = 0
        new_red_ratio = 0
    
    # Update uninvested cash (subscriptions increase, deployments decrease)
    # Sub cash: fraction deployed over N_deploy days, so 1/N_deploy deployed today
    if cash_inflow > 0:
        # Subscription: add new cash, deploy 1/N of total
        u_today = (u_prev + new_sub_ratio) * (1 - 1.0/N_deploy)
    elif cash_inflow < 0:
        # Redemption: reduce cash proportionally
        # Redemption takes from cash first, then equity
        u_today = max(0, u_prev - new_red_ratio)
    else:
        # No subscription/redemption: deploy 1/N of remaining
        u_today = u_prev * (1 - 1.0/N_deploy)
    
    u_today = max(0, min(0.5, u_today))  # cap at 50%
    uninvested_history.append(u_today)
    
    # Dynamic equity weight
    W_eq = d['post'] * (1 - u_today) if d['post'] else 0.95 * (1 - u_today)
    W_eq = max(0.3, min(0.99, W_eq))  # bounds
    W_cash = 1 - W_eq
    
    # Returns
    R_idx = d['idx'] / p['idx'] if p['idx'] else 1
    R_fx = d['hkd'] / p['hkd'] if (d['hkd'] and p['hkd']) else 1
    
    # Factor 6: Volatility-dependent tracking error
    idx_move = abs(R_idx - 1)
    # Tracking error increases with |index move|
    # On normal days: ~2bp TE. On extreme days (|move|>3%): up to 50bp
    # Model: TE = alpha_TE * |idx_move|^1.5 * sign_adjustment
    alpha_TE = 5000  # calibrated below
    # The direction of TE depends on fund composition vs index
    # On big up days, fund tends to underperform (TE < 0)
    # On big down days, fund can outperform (TE > 0, less drawdown due to cash)
    # But this is already captured by dynamic equity weight
    # So TE here is residual tracking error
    # For now, set to 0 (will calibrate)
    TE_adj = 0.0
    
    # Factor 5: Dividend adjustment
    # Without ex-dividend dates, estimate from the data pattern
    # For now, use a small constant daily dividend yield
    # Typical HK index dividend yield: ~2-3% annually
    # Daily: 2.5% / 365 = 0.0000685
    DIV_yield = 0.025  # annual dividend yield
    DIV_daily = DIV_yield / 365
    
    # Improved NAV estimation
    r_portfolio = W_eq * (R_idx * R_fx - 1) + W_cash * cash_rate_daily
    nav_est = p['nav'] * (1 + r_portfolio - fee_daily + DIV_daily)
    
    # BR for improved model
    br_improved = 10000 * (nav_est - d['nav']) / d['nav']
    improved_brs.append(br_improved)

print()
print('='*70)
print('IMPROVED MODEL v1 (dynamic equity weight + cash return + dividend)')
print('  - Dynamic equity: Post * (1 - uninvested_cash_ratio)')
print('  - Cash deployment: %d-day gradual' % N_deploy)
print('  - Cash return: %.1f%% annual' % (cash_rate_daily*365*100))
print('  - Dividend yield: %.1f%% annual' % (DIV_yield*100))
print('  - Fee: %.1f%% annual' % (fee_annual*100))
print('='*70)
print('N = %d' % len(improved_brs))
print('Mean BR = %+.2f bp' % (sum(improved_brs)/len(improved_brs)))
print('Stdev BR = %.2f bp' % statistics.stdev(improved_brs))
print('Mean |BR| = %.2f bp' % (sum(abs(b) for b in improved_brs)/len(improved_brs)))
print('Max |BR| = %.1f bp' % max(abs(b) for b in improved_brs))
within10 = sum(1 for b in improved_brs if abs(b) <= 10)
within20 = sum(1 for b in improved_brs if abs(b) <= 20)
print('Days |BR|<=10bp: %d/%d (%.1f%%)' % (within10, len(improved_brs), 100*within10/len(improved_brs)))
print('Days |BR|<=20bp: %d/%d (%.1f%%)' % (within20, len(improved_brs), 100*within20/len(improved_brs)))

# ---- OPTIMIZE N_deploy and DIV_yield ----
print()
print('='*70)
print('PARAMETER OPTIMIZATION')
print('='*70)

best_stdev = 999
best_params = None
best_results = None

for N_d in [1, 2, 3, 4, 5, 7, 10]:
    for div_y in [0.0, 0.01, 0.015, 0.02, 0.025, 0.03, 0.035, 0.04]:
        for cash_r in [0.0, 0.01, 0.015, 0.02]:
            for fee_a in [0.008, 0.009, 0.01, 0.011, 0.012]:
                u_ratio = 0.0
                trial_brs = []
                u_hist = [0.0]
                
                for i in range(1, len(data)):
                    d = data[i]
                    p = data[i-1]
                    if p['nav'] is None or d['nav'] is None:
                        u_hist.append(0.0)
                        continue
                    
                    u_prev = u_hist[-1]
                    
                    if d['add'] is not None:
                        c_inflow = d['add'] * 10000 * p['nav']
                    else:
                        c_inflow = 0
                    
                    t_assets = p['share'] * p['nav']
                    
                    if t_assets > 0 and c_inflow > 0:
                        new_sub = c_inflow / t_assets
                        u_t = (u_prev + new_sub) * (1 - 1.0/N_d)
                    elif t_assets > 0 and c_inflow < 0:
                        new_red = -c_inflow / t_assets
                        u_t = max(0, u_prev - new_red)
                    else:
                        u_t = u_prev * (1 - 1.0/N_d)
                    
                    u_t = max(0, min(0.5, u_t))
                    u_hist.append(u_t)
                    
                    W_eq = 0.95 * (1 - u_t)
                    W_eq = max(0.3, min(0.99, W_eq))
                    W_cash = 1 - W_eq
                    
                    R_idx = d['idx'] / p['idx'] if p['idx'] else 1
                    R_fx = d['hkd'] / p['hkd'] if (d['hkd'] and p['hkd']) else 1
                    
                    r_port = W_eq * (R_idx * R_fx - 1) + W_cash * cash_r / 365
                    n_est = p['nav'] * (1 + r_port - fee_a/365 + div_y/365)
                    
                    br_t = 10000 * (n_est - d['nav']) / d['nav']
                    trial_brs.append(br_t)
                
                if len(trial_brs) < 10:
                    continue
                
                s = statistics.stdev(trial_brs)
                m = abs(sum(trial_brs)/len(trial_brs))
                # Objective: minimize stdev + penalty for non-zero mean
                obj = s + 0.5 * m
                
                if obj < best_stdev:
                    best_stdev = obj
                    best_params = (N_d, div_y, cash_r, fee_a)
                    best_results = trial_brs

print('Best parameters:')
print('  N_deploy = %d days' % best_params[0])
print('  Dividend yield = %.1f%%' % (best_params[1]*100))
print('  Cash return = %.1f%%' % (best_params[2]*100))
print('  Fee = %.2f%%' % (best_params[3]*100))
print()
print('OPTIMIZED MODEL RESULTS:')
print('N = %d' % len(best_results))
print('Mean BR = %+.2f bp' % (sum(best_results)/len(best_results)))
print('Stdev BR = %.2f bp' % statistics.stdev(best_results))
print('Mean |BR| = %.2f bp' % (sum(abs(b) for b in best_results)/len(best_results)))
print('Max |BR| = %.1f bp' % max(abs(b) for b in best_results))
within10 = sum(1 for b in best_results if abs(b) <= 10)
within20 = sum(1 for b in best_results if abs(b) <= 20)
print('Days |BR|<=10bp: %d/%d (%.1f%%)' % (within10, len(best_results), 100*within10/len(best_results)))
print('Days |BR|<=20bp: %d/%d (%.1f%%)' % (within20, len(best_results), 100*within20/len(best_results)))

# ---- Show improvement on extreme days ----
print()
print('='*70)
print('EXTREME DAYS COMPARISON (optimized model)')
print('='*70)
idx2 = 0
for i in range(1, len(data)):
    d = data[i]
    p = data[i-1]
    if p['nav'] is None or d['nav'] is None:
        continue
    if d['br'] is not None and abs(d['br']) > 15:
        date_str = str(d['date'])[:10]
        old_br = d['br']
        new_br = best_results[idx2]
        improvement = abs(old_br) - abs(new_br)
        print('%s: Old BR=%+.1f bp, New BR=%+.1f bp, Improvement=%+.1f bp %s' % (
            date_str, old_br, new_br, improvement, '***' if abs(old_br) > 30 else ''))
    if d['br'] is not None:
        idx2 += 1

# ---- IMPROVED MODEL v2: Add volatility tracking error ----
print()
print('='*70)
print('IMPROVED MODEL v2 (add volatility-adjusted tracking error)')
print('='*70)

N_d2, div_y2, cash_r2, fee_a2 = best_params
# Add a volatility-based TE adjustment
# On days with |idx_move| > 1.5%, fund tends to have tracking error
# Direction: on big UP days, fund slightly underperforms (negative TE)
# On big DOWN days, fund slightly outperforms (positive TE, due to cash drag)
# But this is already partly captured by dynamic equity weight
# The residual effect: |TE| ~ alpha * |idx_move|^beta

best2_stdev = 999
best2_alpha = 0
best2_beta = 1.0
best2_results = None

for alpha in [0, 100, 200, 500, 1000, 2000, 3000, 5000]:
    for beta in [1.0, 1.5, 2.0]:
        u_hist = [0.0]
        trial_brs = []
        
        for i in range(1, len(data)):
            d = data[i]
            p = data[i-1]
            if p['nav'] is None or d['nav'] is None:
                u_hist.append(0.0)
                continue
            
            u_prev = u_hist[-1]
            
            if d['add'] is not None:
                c_inflow = d['add'] * 10000 * p['nav']
            else:
                c_inflow = 0
            
            t_assets = p['share'] * p['nav']
            
            if t_assets > 0 and c_inflow > 0:
                new_sub = c_inflow / t_assets
                u_t = (u_prev + new_sub) * (1 - 1.0/N_d2)
            elif t_assets > 0 and c_inflow < 0:
                new_red = -c_inflow / t_assets
                u_t = max(0, u_prev - new_red)
            else:
                u_t = u_prev * (1 - 1.0/N_d2)
            
            u_t = max(0, min(0.5, u_t))
            u_hist.append(u_t)
            
            W_eq = 0.95 * (1 - u_t)
            W_eq = max(0.3, min(0.99, W_eq))
            W_cash = 1 - W_eq
            
            R_idx = d['idx'] / p['idx'] if p['idx'] else 1
            R_fx = d['hkd'] / p['hkd'] if (d['hkd'] and p['hkd']) else 1
            
            idx_move = R_idx - 1
            # TE: on extreme days, fund has residual tracking error
            # Negative on big up days (fund underperforms slightly)
            # Positive on big down days (fund outperforms slightly due to cash)
            te = -alpha * (abs(idx_move) ** beta) * (1 if idx_move > 0 else -1) / 10000
            
            r_port = W_eq * (R_idx * R_fx - 1 + te) + W_cash * cash_r2 / 365
            n_est = p['nav'] * (1 + r_port - fee_a2/365 + div_y2/365)
            
            br_t = 10000 * (n_est - d['nav']) / d['nav']
            trial_brs.append(br_t)
        
        if len(trial_brs) < 10:
            continue
        
        s = statistics.stdev(trial_brs)
        m = abs(sum(trial_brs)/len(trial_brs))
        obj = s + 0.5 * m
        
        if obj < best2_stdev:
            best2_stdev = obj
            best2_alpha = alpha
            best2_beta = beta
            best2_results = trial_brs

print('Best TE parameters:')
print('  alpha_TE = %d' % best2_alpha)
print('  beta_TE = %.1f' % best2_beta)
print()
print('OPTIMIZED MODEL v2 RESULTS:')
print('N = %d' % len(best2_results))
print('Mean BR = %+.2f bp' % (sum(best2_results)/len(best2_results)))
print('Stdev BR = %.2f bp' % statistics.stdev(best2_results))
print('Mean |BR| = %.2f bp' % (sum(abs(b) for b in best2_results)/len(best2_results)))
print('Max |BR| = %.1f bp' % max(abs(b) for b in best2_results))
within10 = sum(1 for b in best2_results if abs(b) <= 10)
within20 = sum(1 for b in best2_results if abs(b) <= 20)
print('Days |BR|<=10bp: %d/%d (%.1f%%)' % (within10, len(best2_results), 100*within10/len(best2_results)))
print('Days |BR|<=20bp: %d/%d (%.1f%%)' % (within20, len(best2_results), 100*within20/len(best2_results)))

# ---- IMPROVED MODEL v3: Add dividend spike detection ----
print()
print('='*70)
print('IMPROVED MODEL v3 (detect and adjust for dividend ex-dates)')
print('='*70)

# Detect potential ex-dividend dates:
# On ex-dividend date, index drops mechanically but fund NAV includes dividend
# Pattern: index drops, but NAV drops less than expected
# Or: the fund's NAV stays higher relative to index
# We look for days where:
# 1. Index dropped
# 2. The model overestimates the drop (positive BR)
# 3. No large subscription
# These are likely ex-dividend dates

potential_ex_div = []
for i in range(1, len(data)):
    d = data[i]
    p = data[i-1]
    if d['br'] is None or p['br'] is None:
        continue
    R_idx = d['idx'] / p['idx'] if p['idx'] else 1
    idx_drop = R_idx < 0.995  # index dropped more than 0.5%
    positive_br = d['br'] > 15  # model overestimates NAV
    no_large_sub = abs(d['add'] or 0) < 50 if d['add'] is not None else True
    if idx_drop and positive_br and no_large_sub:
        date_str = str(d['date'])[:10]
        potential_ex_div.append((date_str, d['br'], R_idx - 1, d['add']))

print('Potential ex-dividend dates detected:')
for p in potential_ex_div:
    print('  %s: BR=%+.1f, idx_ret=%+.2f%%, add=%s' % (p[0], p[1], p[2]*100, p[3]))

# Also check: days after large positive BR with no index explanation
# (could be accumulated dividends)
print()
print('Days with BR > +20 and |idx_move| < 2%% (likely dividend events):')
for i in range(1, len(data)):
    d = data[i]
    p = data[i-1]
    if d['br'] is None:
        continue
    R_idx = d['idx'] / p['idx'] if p['idx'] else 1
    if d['br'] > 20 and abs(R_idx - 1) < 0.02:
        date_str = str(d['date'])[:10]
        print('  %s: BR=%+.1f, idx_ret=%+.2f%%, add=%s' % (date_str, d['br'], (R_idx-1)*100, d['add']))

# ---- SUMMARY COMPARISON ----
print()
print('='*70)
print('SUMMARY COMPARISON')
print('='*70)
print('%-30s %10s %10s %10s' % ('Metric', 'Current', 'Opt_v1', 'Opt_v2'))
print('-'*60)
for label, vals in [('Mean BR (bp)', [sum(current_brs)/len(current_brs), sum(best_results)/len(best_results), sum(best2_results)/len(best2_results)]),
                     ('Stdev BR (bp)', [statistics.stdev(current_brs), statistics.stdev(best_results), statistics.stdev(best2_results)]),
                     ('Mean |BR| (bp)', [sum(abs(b) for b in current_brs)/len(current_brs), sum(abs(b) for b in best_results)/len(best_results), sum(abs(b) for b in best2_results)/len(best2_results)]),
                     ('Max |BR| (bp)', [max(abs(b) for b in current_brs), max(abs(b) for b in best_results), max(abs(b) for b in best2_results)])]:
    print('%-30s %10.2f %10.2f %10.2f' % (label, vals[0], vals[1], vals[2]))

# Days within thresholds
for t in [5, 10, 15, 20]:
    c = sum(1 for b in current_brs if abs(b) <= t)
    v1 = sum(1 for b in best_results if abs(b) <= t)
    v2 = sum(1 for b in best2_results if abs(b) <= t)
    print('Days |BR|<=%dbp       %d/%d(%4.1f%%) %d/%d(%4.1f%%) %d/%d(%4.1f%%)' % (
        t, c, len(current_brs), 100*c/len(current_brs),
        v1, len(best_results), 100*v1/len(best_results),
        v2, len(best2_results), 100*v2/len(best2_results)))

wb.close()
