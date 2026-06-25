#!/usr/bin/env python3
"""Politiques de dividendes — priorité à Nessim (100 % du capital, 45 % des parts)."""
import importlib.util, sys

# importer build_model pour récupérer les résultats nets et la cap table
spec = importlib.util.spec_from_file_location("bm", "/Users/os/Desktop/HCTech/model/build_model.py")
bm = importlib.util.module_from_spec(spec)
sys.argv = ["build_model.py"]
spec.loader.exec_module(bm)

NET = {y: bm.YA[y]['net'] for y in range(1, 8)}
SHARES = bm.SHARES
CAPITAL = bm.CAPITAL
NAMES = list(SHARES.keys())

def irr(cashflows, lo=-0.99, hi=10.0):
    def npv(r): return sum(cf/((1+r)**t) for t,cf in enumerate(cashflows))
    for _ in range(200):
        mid=(lo+hi)/2
        if npv(mid)>0: lo=mid
        else: hi=mid
    return (lo+hi)/2

def show(title, flows):
    print("="*86); print(title); print("="*86)
    print(f"{'Actionnaire':<20}"+"".join(f"{'An'+str(y):>10}" for y in range(1,8))+f"{'TOTAL':>12}")
    tot_by_name={}
    for name in NAMES:
        vals=[flows[y][name] for y in range(1,8)]; t=sum(vals); tot_by_name[name]=t
        print(f"{name:<20}"+"".join(f"{v:>10,.0f}" for v in vals)+f"{t:>12,.0f}")
    pool_tot=sum(sum(flows[y].values()) for y in range(1,8))
    print(f"{'TOTAL distribué':<20}"+"".join(f"{sum(flows[y].values()):>10,.0f}" for y in range(1,8))+f"{pool_tot:>12,.0f}")
    # Nessim payback + IRR
    ness=[flows[y]['Nessim Mami'] for y in range(1,8)]
    cum=0; payback=None
    for y in range(1,8):
        cum+=ness[y-1]
        if payback is None and cum>=CAPITAL: payback=y
    cf=[-CAPITAL]+ness
    print(f"\n>> Nessim : reçu total {tot_by_name['Nessim Mami']:,.0f} DT | "
          f"capital remboursé fin An{payback} | TRI ≈ {irr(cf)*100:,.0f}%")
    print(f">> Multiple Nessim (total reçu / capital) : {tot_by_name['Nessim Mami']/CAPITAL:.1f}×\n")

# ── Option 1 : remboursement prioritaire Nessim (capital + prime 20%) puis pro-rata ──
def opt1(net, pref_target=360000, payout=0.70):
    flows={y:{n:0 for n in NAMES} for y in range(1,8)}; rec=0
    for y in range(1,8):
        pool=payout*net[y]
        if rec<pref_target:
            pay=min(pool, pref_target-rec); flows[y]['Nessim Mami']+=pay; rec+=pay; pool-=pay
        for n,p in SHARES.items(): flows[y][n]+=pool*p
    return flows

# ── Option 2 : coupon préférentiel permanent 20% du capital + pro-rata ──
def opt2(net, coupon=0.20, payout=0.70):
    flows={y:{n:0 for n in NAMES} for y in range(1,8)}
    for y in range(1,8):
        pool=payout*net[y]; coup=min(pool, coupon*CAPITAL)
        flows[y]['Nessim Mami']+=coup; pool-=coup
        for n,p in SHARES.items(): flows[y][n]+=pool*p
    return flows

# ── Option 3 : remboursement éclair (100% à Nessim An1-2 jusqu'à capital seul), puis pro-rata ──
def opt3(net, target=300000, payout=0.85):
    flows={y:{n:0 for n in NAMES} for y in range(1,8)}; rec=0
    for y in range(1,8):
        pool=payout*net[y]
        if rec<target:
            pay=min(pool, target-rec); flows[y]['Nessim Mami']+=pay; rec+=pay; pool-=pay
        for n,p in SHARES.items(): flows[y][n]+=pool*p
    return flows

print("Résultats nets par an:", {y:round(NET[y]) for y in range(1,8)})
print("Net cumulé 7 ans:", f"{sum(NET.values()):,.0f}\n")
show("OPTION 1 — Remboursement prioritaire Nessim (capital 300k + prime 20% = 360k), payout 70%", opt1(NET))
show("OPTION 2 — Coupon préférentiel Nessim 20%/an (60k) + pro-rata, payout 70%", opt2(NET))
show("OPTION 3 — Remboursement éclair Nessim (300k prioritaire, payout 85%) puis pro-rata", opt3(NET))
