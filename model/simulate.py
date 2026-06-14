#!/usr/bin/env python3
"""
HCTECH — Modèle financier 7 ans, AUTOFINANCEMENT SEUL.
Simulation mensuelle (84 mois). Le cash pilote le déploiement des machines.
Toutes les hypothèses sont en tête de fichier (sourcées dans le BP).
"""

import math

# ────────────────────────────────────────────────────────────────────────
# HYPOTHÈSES
# ────────────────────────────────────────────────────────────────────────
MONTHS = 84
USD_TND = 2.95

# Machine (par unité)
MACHINE_USD = 18000
MACHINE_DT = MACHINE_USD * USD_TND          # 53 100 (départ usine, fournisseur Corée)
IMPORT_FEES = MACHINE_DT * 0.025            # 2,5 % frais d'import
INSTALL = 3500                              # installation locale
MACHINE_HT = MACHINE_DT + IMPORT_FEES + INSTALL   # immobilisation HT
DEPREC_YEARS = 10                           # amortissement linéaire matériel

# Modalités de règlement fournisseur (sur la part machine Corée)
DOWN_PCT = 0.20                             # 20 % à la commande
BAL_PCT = 0.80                              # solde à 120 j (~4 mois)
ORDER_LEAD = 2                              # acheminement: commande = livraison - 2 mois
BAL_LAG = 4                                 # 120 jours après la commande

# Revenu
STATION_L_DAY = 12000                       # ventes station moyenne
RECOVERY = 5/1000                           # taux récupération retenu
L_PER_MACHINE_DAY = STATION_L_DAY * RECOVERY      # 60 L/j
FUEL_P0 = 2.525                             # prix essence an 1
FUEL_GROWTH = 0.045                         # +4,5 %/an (CAGR réel 2016-2026)
HCTECH_SHARE = 0.60                         # partage 60 % HCTECH / 40 % station

# Charges
OPEX_MACHINE_MONTH = 120                    # maintenance + internet / machine / mois
EMP_CHARGES = 0.21                          # charges patronales (CNSS 17,07 + TFP 2 + FOPROLOS 1 + acc.)
SAL_FOUNDER0 = 2000                         # Lamine & Ali, mensuel an 1
SAL_FOUNDER_GROWTH = 0.50                   # +50 %/an
SAL_FOUNDER_CAP = 6000                      # plafond mensuel
SAL_TECH0 = 1200                            # technicien mensuel an 1
SAL_TECH_GROWTH = 0.06                      # +6 %/an
MACHINES_PER_TECH = 50                      # 1 technicien / 50 machines
RENT0 = 1200                                # bureau + atelier mensuel an 1
RENT_GROWTH = 0.06
OVERHEAD0 = 2000                            # frais généraux divers (compta, assur., banque, carburant…)
OVERHEAD_GROWTH = 0.06

# Leasing véhicule utilitaire 70 000 DT @ 12 %/5 ans
VEH_PRICE = 70000
VEH_RATE = 0.12
VEH_MONTHS = 60
VEH_PMT = VEH_PRICE * (VEH_RATE/12) / (1 - (1 + VEH_RATE/12)**(-VEH_MONTHS))

# Fiscalité
IS_RATE = 0.25                             # IS services/commerce
CSS_RATE = 0.03                            # contribution sociale solidarité
TAX_RATE = IS_RATE + CSS_RATE              # ~28 % sur bénéfice
MIN_TAX = 0.001                            # minimum d'impôt 0,1 % du CA

# Financement
CAPITAL = 300000                           # capital social (Nessim), injecté mois 1
CASH_BUFFER = 50000                        # trésorerie de sécurité minimale

# Déploiement cible: 14 machines à M3 puis +14 tous les 6 mois
TARGET_BATCH = 14
DEPLOY_MONTHS = list(range(3, MONTHS+1, 6))      # 3,9,15,...,81

# ────────────────────────────────────────────────────────────────────────
# HELPERS
# ────────────────────────────────────────────────────────────────────────
def year_of(m):
    return (m - 1) // 12 + 1

def fuel_price(m):
    return FUEL_P0 * (1 + FUEL_GROWTH) ** (year_of(m) - 1)

def founder_salary(m):
    y = year_of(m)
    s = SAL_FOUNDER0 * (1 + SAL_FOUNDER_GROWTH) ** (y - 1)
    return min(s, SAL_FOUNDER_CAP)

def tech_unit_salary(m):
    y = year_of(m)
    return SAL_TECH0 * (1 + SAL_TECH_GROWTH) ** (y - 1)

def rent(m):
    return RENT0 * (1 + RENT_GROWTH) ** (year_of(m) - 1)

def overhead(m):
    return OVERHEAD0 * (1 + OVERHEAD_GROWTH) ** (year_of(m) - 1)

def machine_revenue_month(m):
    # revenu HCTECH par machine en service, pour le mois m
    days = 365/12
    return L_PER_MACHINE_DAY * days * fuel_price(m) * HCTECH_SHARE

def recurring_cash_costs(m, machines_in_service):
    """Charges décaissées récurrentes du mois m (hors machines capex, hors impôt)."""
    techs = max(1, math.ceil(machines_in_service / MACHINES_PER_TECH))
    gross_payroll = 2 * founder_salary(m) + techs * tech_unit_salary(m)
    payroll = gross_payroll * (1 + EMP_CHARGES)
    veh = VEH_PMT if m <= VEH_MONTHS else 0.0
    opex_mach = machines_in_service * OPEX_MACHINE_MONTH
    return payroll + rent(m) + overhead(m) + veh + opex_mach

# ────────────────────────────────────────────────────────────────────────
# SIMULATION
# ────────────────────────────────────────────────────────────────────────
def simulate(self_financed=True, forced_batch=None):
    """
    self_financed=True : déploiement bridé par la trésorerie (>= buffer).
    forced_batch=N : déploie N machines à chaque date prévue (ignore le cash) -> mesure le besoin.
    Retourne (annual, monthly).
    """
    cash = 0.0
    deploy_plan = {}          # mois -> nb machines déployées
    payments = {m: 0.0 for m in range(1, MONTHS + 13)}  # décaissements machines planifiés
    min_cash_ever = float('inf')

    # injecter capital mois 1
    # (on l'ajoute dans la boucle)
    capital_injected = False

    def machines_in_service_at(mo, plan):
        return sum(n for dm, n in plan.items() if dm <= mo)

    def order_machine_payments(deploy_m, n):
        """programme les décaissements pour n machines déployées à deploy_m."""
        order_m = deploy_m - ORDER_LEAD
        out = {}
        # 20% à la commande
        out[order_m] = out.get(order_m, 0) + n * MACHINE_DT * DOWN_PCT
        # 80% à 120j après commande
        out[order_m + BAL_LAG] = out.get(order_m + BAL_LAG, 0) + n * MACHINE_DT * BAL_PCT
        # import + installation à la livraison
        out[deploy_m] = out.get(deploy_m, 0) + n * (IMPORT_FEES + INSTALL)
        return out

    def project_min_cash(from_m, start_cash, base_plan, base_payments, candidate_deploy_m, candidate_n, horizon=14):
        """projette la trésorerie en supposant candidate_n machines à candidate_deploy_m, sans autres batches."""
        plan = dict(base_plan)
        pays = dict(base_payments)
        if candidate_n > 0:
            plan[candidate_deploy_m] = plan.get(candidate_deploy_m, 0) + candidate_n
            for mm, amt in order_machine_payments(candidate_deploy_m, candidate_n).items():
                pays[mm] = pays.get(mm, 0) + amt
        c = start_cash
        mn = c
        for mm in range(from_m, min(MONTHS, from_m + horizon) + 1):
            mis = machines_in_service_at(mm, plan)
            c += mis * machine_revenue_month(mm)
            c -= recurring_cash_costs(mm, mis)
            c -= pays.get(mm, 0)
            mn = min(mn, c)
        return mn

    annual = {y: {} for y in range(1, 8)}
    monthly = []
    tax_payable_next = 0.0
    # accumulateurs annuels P&L
    yacc = {y: dict(rev=0, opex=0, payroll=0, rent=0, overhead=0, veh=0, deprec=0) for y in range(1,8)}

    deploy_decisions = {}  # order month handled

    for m in range(1, MONTHS + 1):
        y = year_of(m)
        # injection capital
        if m == 1:
            cash += CAPITAL

        # décision de commande: aux mois (deploy-2)
        for dm in DEPLOY_MONTHS:
            if dm - ORDER_LEAD == m:
                if forced_batch is not None:
                    n = forced_batch
                else:
                    # trouver le plus grand n<=TARGET_BATCH gardant min cash >= buffer
                    n = 0
                    for cand in range(TARGET_BATCH, 0, -1):
                        mn = project_min_cash(m, cash, deploy_plan, payments, dm, cand)
                        if mn >= CASH_BUFFER:
                            n = cand
                            break
                if n > 0:
                    deploy_plan[dm] = deploy_plan.get(dm, 0) + n
                    for mm, amt in order_machine_payments(dm, n).items():
                        payments[mm] = payments.get(mm, 0) + amt

        # flux du mois
        mis = machines_in_service_at(m, deploy_plan)
        rev = mis * machine_revenue_month(m)
        rcosts = recurring_cash_costs(m, mis)
        mach_pay = payments.get(m, 0.0)

        # impôt: payé en juin (mois 6 de chaque année) sur le bénéfice de l'année précédente
        tax_cash = 0.0
        if m % 12 == 6:
            tax_cash = tax_payable_next
            tax_payable_next = 0.0

        cash += rev - rcosts - mach_pay - tax_cash
        min_cash_ever = min(min_cash_ever, cash)

        # P&L accrual accumulators
        techs = max(1, math.ceil(mis / MACHINES_PER_TECH))
        gross_payroll = 2 * founder_salary(m) + techs * tech_unit_salary(m)
        payroll = gross_payroll * (1 + EMP_CHARGES)
        veh = VEH_PMT if m <= VEH_MONTHS else 0.0
        deprec = mis * (MACHINE_HT / DEPREC_YEARS / 12)
        yacc[y]['rev'] += rev
        yacc[y]['opex'] += mis * OPEX_MACHINE_MONTH
        yacc[y]['payroll'] += payroll
        yacc[y]['rent'] += rent(m)
        yacc[y]['overhead'] += overhead(m)
        yacc[y]['veh'] += veh
        yacc[y]['deprec'] += deprec

        monthly.append(dict(m=m, y=y, machines=mis, rev=rev, cash=cash, mach_pay=mach_pay))

        # fin d'année: calcul impôt
        if m % 12 == 0:
            a = yacc[y]
            ebitda = a['rev'] - a['opex'] - a['payroll'] - a['rent'] - a['overhead'] - a['veh']
            ebit = ebitda - a['deprec']
            if ebit > 0:
                tax = ebit * TAX_RATE
            else:
                tax = MIN_TAX * a['rev']
            tax_payable_next = tax
            annual[y] = dict(
                new_machines=sum(n for dm, n in deploy_plan.items() if year_of(dm) == y),
                machines_eoy=machines_in_service_at(m, deploy_plan),
                machine_years=a['rev'] / (machine_revenue_month(y*12) ) if False else None,
                rev=a['rev'], opex=a['opex'], payroll=a['payroll'], rent=a['rent'],
                overhead=a['overhead'], veh=a['veh'], ebitda=ebitda, deprec=a['deprec'],
                ebit=ebit, tax=tax, net=ebit - tax, cash_eoy=cash
            )

    return annual, monthly, min_cash_ever, deploy_plan


def print_annual(title, annual, min_cash, deploy_plan):
    print("=" * 78)
    print(title)
    print("=" * 78)
    hdr = f"{'':<22}" + "".join(f"{'An'+str(y):>13}" for y in range(1,8))
    print(hdr)
    def row(label, key, scale=1):
        vals = "".join(f"{annual[y][key]/scale:>13,.0f}" for y in range(1,8))
        print(f"{label:<22}{vals}")
    print(f"{'Nouvelles machines':<22}" + "".join(f"{annual[y]['new_machines']:>13}" for y in range(1,8)))
    print(f"{'Parc fin année':<22}" + "".join(f"{annual[y]['machines_eoy']:>13}" for y in range(1,8)))
    row("CA HCTECH", 'rev')
    row("OPEX machines", 'opex')
    row("Masse salariale", 'payroll')
    row("Loyer", 'rent')
    row("Frais généraux", 'overhead')
    row("Leasing véhicule", 'veh')
    row("EBITDA", 'ebitda')
    row("Amortissements", 'deprec')
    row("Résultat exploit.", 'ebit')
    row("Impôt (IS+CSS)", 'tax')
    row("Résultat net", 'net')
    row("Trésorerie fin année", 'cash_eoy')
    print(f"\nTrésorerie minimale sur 84 mois : {min_cash:,.0f} DT")
    print(f"Parc total déployé fin An7 : {sum(deploy_plan.values())} machines")
    print(f"VEH_PMT mensuel leasing : {VEH_PMT:,.1f} DT  | Machine HT : {MACHINE_HT:,.0f} DT")
    cum = 0
    print("Résultat net cumulé:", end=" ")
    for y in range(1,8):
        cum += annual[y]['net']
        print(f"An{y}={cum:,.0f}", end="  ")
    print()


# Scénario A: autofinancement bridé par le cash
annA, monA, mincashA, planA = simulate(self_financed=True)
print_annual("SCÉNARIO AUTOFINANCEMENT (déploiement bridé par la trésorerie)", annA, mincashA, planA)

# Scénario B: forcer 14/batch (28/an) pour mesurer le besoin de financement
annB, monB, mincashB, planB = simulate(self_financed=False, forced_batch=14)
print()
print_annual("SCÉNARIO CIBLE 28/AN FORCÉ (mesure du besoin de financement)", annB, mincashB, planB)
print(f"\n>>> BESOIN DE FINANCEMENT EXTERNE (trésorerie la plus basse, scénario forcé) : {mincashB:,.0f} DT")
