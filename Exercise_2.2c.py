"""
Exercise 2.2.c: Rolling Stock Scheduling - Composition Model (X_t,p formulation)
Compare runtime with Basic Model (N_u,t formulation)
"""

import pandas as pd
from gurobipy import Model, GRB, quicksum
import time

# ============================================================
# 1. Read Data
# ============================================================
xl = pd.ExcelFile('a2_part2.xlsx')
timetable_df = pd.read_excel(xl, sheet_name='Timetable')
seats_df = pd.read_excel(xl, sheet_name='Seats')

# Parse seats data
seats_df.columns = ['Line', 'Southbound', 'Northbound']
seats_df = seats_df.iloc[1:].reset_index(drop=True)
seats_df['Line'] = seats_df['Line'].astype(int)
seats_df['Southbound'] = seats_df['Southbound'].astype(int)
seats_df['Northbound'] = seats_df['Northbound'].astype(int)

seat_demand = {}
for _, row in seats_df.iterrows():
    seat_demand[(row['Line'], 'South')] = row['Southbound']
    seat_demand[(row['Line'], 'North')] = row['Northbound']

# ============================================================
# 2. Parameters
# ============================================================
# Unit parameters
cost = {'PL3': 315000, 'PL4': 385000}
capacity = {'PL3': 400, 'PL4': 600}
length = {'PL3': 80, 'PL4': 110}

# Cross-section trains (from 2.1.a)
cross_section = {
    (800, 'South'): 6, (800, 'North'): 6,
    (3000, 'South'): 6, (3000, 'North'): 5,
    (3100, 'South'): 3, (3100, 'North'): 3,
    (3500, 'South'): 4, (3500, 'North'): 4,
    (3900, 'South'): 2, (3900, 'North'): 3,
}

# Create train set
trains = []
train_info = {}
for (line, direction), num_trains in cross_section.items():
    for i in range(num_trains):
        train_id = f"{line}_{direction}_{i+1}"
        trains.append(train_id)
        train_info[train_id] = {
            'line': line,
            'direction': direction,
            'seat_demand': seat_demand[(line, direction)],
            'max_length': 200 if line == 3900 else 300
        }

print(f"Total cross-section trains: {len(trains)}")

# ============================================================
# 3. Generate Compositions
# ============================================================
# Compositions: (n_PL3, n_PL4, length, capacity, cost)
def generate_compositions(max_length):
    compositions = []
    for n_pl3 in range(5):
        for n_pl4 in range(4):
            if n_pl3 == 0 and n_pl4 == 0:
                continue
            total_length = n_pl3 * length['PL3'] + n_pl4 * length['PL4']
            total_capacity = n_pl3 * capacity['PL3'] + n_pl4 * capacity['PL4']
            total_cost = n_pl3 * cost['PL3'] + n_pl4 * cost['PL4']
            if total_length <= max_length:
                comp_id = f"{n_pl3}PL3_{n_pl4}PL4"
                compositions.append({
                    'id': comp_id,
                    'n_PL3': n_pl3,
                    'n_PL4': n_pl4,
                    'length': total_length,
                    'capacity': total_capacity,
                    'cost': total_cost
                })
    return compositions

P_general = generate_compositions(300)
P_3900 = generate_compositions(200)

print(f"Compositions for general lines (≤300m): {len(P_general)}")
print(f"Compositions for Line 3900 (≤200m): {len(P_3900)}")

# Assign valid compositions to each train (preprocess both length AND seat requirements)
train_compositions = {}
for t in trains:
    max_len = train_info[t]['max_length']
    min_seats = train_info[t]['seat_demand']
    
    # Generate compositions that satisfy BOTH length AND seat constraints
    valid_compositions = []
    for n_pl3 in range(5):
        for n_pl4 in range(4):
            if n_pl3 == 0 and n_pl4 == 0:
                continue
            total_length = n_pl3 * length['PL3'] + n_pl4 * length['PL4']
            total_capacity = n_pl3 * capacity['PL3'] + n_pl4 * capacity['PL4']
            total_cost = n_pl3 * cost['PL3'] + n_pl4 * cost['PL4']
            
            # PREPROCESSING: Filter by both length AND capacity
            if total_length <= max_len and total_capacity >= min_seats:
                comp_id = f"{n_pl3}PL3_{n_pl4}PL4"
                valid_compositions.append({
                    'id': comp_id,
                    'n_PL3': n_pl3,
                    'n_PL4': n_pl4,
                    'length': total_length,
                    'capacity': total_capacity,
                    'cost': total_cost
                })
    
    train_compositions[t] = valid_compositions

print(f"\nCompositions per train after preprocessing:")
print(f"{'Train':<20} {'Demand':<8} {'MaxLen':<8} {'Valid Compositions':<10}")
print("-" * 50)
# Show all trains grouped by line/direction
for t in trains:
    print(f"{t:<20} {train_info[t]['seat_demand']:<8} {train_info[t]['max_length']:<8} {len(train_compositions[t]):<10}")

# ============================================================
# 4. Composition Model (X_t,p formulation)
# ============================================================
print("\n" + "=" * 70)
print("COMPOSITION MODEL (X_t,p formulation)")
print("=" * 70)

start_time_comp = time.time()

model_comp = Model("RollingStock_Composition")
model_comp.setParam('OutputFlag', 0)

# Decision variables: X[t,p] = 1 if composition p is used for train t
X = {}
for t in trains:
    for p in train_compositions[t]:
        X[t, p['id']] = model_comp.addVar(vtype=GRB.BINARY, name=f"X_{t}_{p['id']}")

model_comp.update()

# Objective: Minimize total annual cost
model_comp.setObjective(
    quicksum(p['cost'] * X[t, p['id']] 
             for t in trains 
             for p in train_compositions[t]),
    GRB.MINIMIZE
)

# Constraint 1: Exactly one composition per train
for t in trains:
    model_comp.addConstr(
        quicksum(X[t, p['id']] for p in train_compositions[t]) == 1,
        name=f"one_comp_{t}"
    )

# NOTE: Seat requirement is handled implicitly by preprocessing P_t
# NOTE: Length constraint is handled implicitly by preprocessing P_t

# Constraint 2: Balance constraint (25% rule)
# Total PL3 units = sum over all trains and compositions of (n_PL3 * X[t,p])
total_PL3 = quicksum(p['n_PL3'] * X[t, p['id']] 
                     for t in trains 
                     for p in train_compositions[t])
total_PL4 = quicksum(p['n_PL4'] * X[t, p['id']] 
                     for t in trains 
                     for p in train_compositions[t])

model_comp.addConstr(total_PL3 <= 1.25 * total_PL4, name="balance_PL3")
model_comp.addConstr(total_PL4 <= 1.25 * total_PL3, name="balance_PL4")

# Solve
model_comp.optimize()
end_time_comp = time.time()
runtime_comp = end_time_comp - start_time_comp

if model_comp.status == GRB.OPTIMAL:
    print(f"\nOptimal annual cost: €{model_comp.objVal:,.0f}")
    print(f"Runtime: {runtime_comp:.4f} seconds")
    
    # Calculate totals
    total_pl3 = sum(p['n_PL3'] * X[t, p['id']].X 
                    for t in trains 
                    for p in train_compositions[t])
    total_pl4 = sum(p['n_PL4'] * X[t, p['id']].X 
                    for t in trains 
                    for p in train_compositions[t])
    
    print(f"Total PL3 units: {total_pl3:.0f}")
    print(f"Total PL4 units: {total_pl4:.0f}")
    print(f"PL3/PL4 ratio: {total_pl3/total_pl4:.3f}")
    
    # Output composition for each train
    print("\n" + "-" * 70)
    print("ASSIGNED COMPOSITIONS")
    print("-" * 70)
    print(f"{'Train ID':<20} {'Line':<6} {'Dir':<6} {'Demand':<8} {'Composition':<12} {'Capacity':<10}")
    print("-" * 70)
    
    for t in trains:
        for p in train_compositions[t]:
            if X[t, p['id']].X > 0.5:
                comp_str = f"({p['n_PL3']},{p['n_PL4']})"
                print(f"{t:<20} {train_info[t]['line']:<6} {train_info[t]['direction']:<6} "
                      f"{train_info[t]['seat_demand']:<8} {comp_str:<12} {p['capacity']:<10}")

# ============================================================
# 5. Basic Model (N_u,t formulation) for comparison (Claude)
# ============================================================
print("\n" + "=" * 70)
print("BASIC MODEL (N_u,t formulation) - for runtime comparison")
print("=" * 70)

start_time_basic = time.time()

model_basic = Model("RollingStock_Basic")
model_basic.setParam('OutputFlag', 0)

U = ['PL3', 'PL4']

# Decision variables
N = {}
for u in U:
    for t in trains:
        N[u, t] = model_basic.addVar(vtype=GRB.INTEGER, lb=0, name=f"N_{u}_{t}")

model_basic.update()

# Objective
model_basic.setObjective(
    quicksum(cost[u] * N[u, t] for u in U for t in trains),
    GRB.MINIMIZE
)

# Constraints
for t in trains:
    model_basic.addConstr(
        quicksum(capacity[u] * N[u, t] for u in U) >= train_info[t]['seat_demand'],
        name=f"seats_{t}"
    )
    model_basic.addConstr(
        quicksum(length[u] * N[u, t] for u in U) <= train_info[t]['max_length'],
        name=f"length_{t}"
    )

total_PL3_basic = quicksum(N['PL3', t] for t in trains)
total_PL4_basic = quicksum(N['PL4', t] for t in trains)
model_basic.addConstr(total_PL3_basic <= 1.25 * total_PL4_basic, name="balance_PL3")
model_basic.addConstr(total_PL4_basic <= 1.25 * total_PL3_basic, name="balance_PL4")

# Solve
model_basic.optimize()
end_time_basic = time.time()
runtime_basic = end_time_basic - start_time_basic

if model_basic.status == GRB.OPTIMAL:
    print(f"\nOptimal annual cost: €{model_basic.objVal:,.0f}")
    print(f"Runtime: {runtime_basic:.4f} seconds")

# ============================================================
# 6. Comparison (Claude)
# ============================================================
print("\n" + "=" * 70)
print("COMPARISON OF FORMULATIONS")
print("=" * 70)

print(f"\n{'Metric':<30} {'Basic (N_u,t)':<20} {'Composition (X_t,p)':<20}")
print("-" * 70)
print(f"{'Optimal cost':<30} €{model_basic.objVal:,.0f}{'':>7} €{model_comp.objVal:,.0f}")
print(f"{'Runtime (seconds)':<30} {runtime_basic:.4f}{'':>13} {runtime_comp:.4f}")
print(f"{'Number of variables':<30} {model_basic.NumVars:<20} {model_comp.NumVars:<20}")
print(f"{'Number of constraints':<30} {model_basic.NumConstrs:<20} {model_comp.NumConstrs:<20}")