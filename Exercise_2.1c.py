"""
Exercise 2.1.c: Rolling Stock Scheduling - Basic Model (N_u,t formulation)
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
seats_df = seats_df.iloc[1:].reset_index(drop=True)  # Skip header row
seats_df['Line'] = seats_df['Line'].astype(int)
seats_df['Southbound'] = seats_df['Southbound'].astype(int)
seats_df['Northbound'] = seats_df['Northbound'].astype(int)

seat_demand = {}
for _, row in seats_df.iterrows():
    seat_demand[(row['Line'], 'South')] = row['Southbound']
    seat_demand[(row['Line'], 'North')] = row['Northbound']

print("Seat demand per line/direction:")
for key, val in seat_demand.items():
    print(f"  Line {key[0]} {key[1]}: {val} seats")

# ============================================================
# 2. Calculate Cross-Section Trains
# ============================================================
# Calculate duration for each line/direction
lines_info = {
    800: {'South': ['Amr', 'Asd', 'Ut', 'Ehv', 'Std', 'Mt'],
          'North': ['Mt', 'Std', 'Ehv', 'Ut', 'Asd', 'Amr']},
    3000: {'South': ['Hdr', 'Amr', 'Asd', 'Ut', 'Nm'],
           'North': ['Nm', 'Ut', 'Asd', 'Amr', 'Hdr']},
    3100: {'South': ['Shl', 'Ut', 'Nm'],
           'North': ['Nm', 'Ut', 'Shl']},
    3500: {'South': ['Shl', 'Ut', 'Ehv', 'Vl'],
           'North': ['Vl', 'Ehv', 'Ut', 'Shl']},
    3900: {'South': ['Ehv', 'Std', 'Hrl'],
           'North': ['Hrl', 'Std', 'Ehv']}
}

T = 30  # Period time

# Calculate duration from timetable
def get_duration(line, direction, timetable_df):
    df = timetable_df[(timetable_df['Line'] == line) & (timetable_df['Direction'] == direction)]
    
    # Get first departure and last arrival
    first_dep = df[df['Type'] == 'dep'].iloc[0]['Time']
    last_arr = df[df['Type'] == 'arr'].iloc[-1]['Time']
    
    # Calculate duration (handle wrap-around)
    duration = last_arr - first_dep
    if duration <= 0:
        duration += 30  # Add one period
    
    # Count number of periods
    # Duration includes multiple periods for long trips
    route = lines_info[line][direction]
    
    # Calculate actual duration by summing segments
    total_time = 0
    for i in range(len(route) - 1):
        from_st = route[i]
        to_st = route[i + 1]
        
        # Get departure time from from_st
        dep_row = df[(df['Station'] == from_st) & (df['Type'] == 'dep')]
        arr_row = df[(df['Station'] == to_st) & (df['Type'] == 'arr')]
        
        if len(dep_row) > 0 and len(arr_row) > 0:
            dep_time = dep_row.iloc[0]['Time']
            arr_time = arr_row.iloc[0]['Time']
            segment_time = arr_time - dep_time
            if segment_time < 0:
                segment_time += 30
            total_time += segment_time
        
        # Add dwell time if not last station
        if i < len(route) - 2:
            arr_row2 = df[(df['Station'] == to_st) & (df['Type'] == 'arr')]
            dep_row2 = df[(df['Station'] == to_st) & (df['Type'] == 'dep')]
            if len(arr_row2) > 0 and len(dep_row2) > 0:
                arr_t = arr_row2.iloc[0]['Time']
                dep_t = dep_row2.iloc[0]['Time']
                dwell = dep_t - arr_t
                if dwell < 0:
                    dwell += 30
                total_time += dwell
    
    return total_time

# Use the durations from the reference (Table 7)
durations = {
    (800, 'South'): 181,
    (800, 'North'): 178,
    (3000, 'South'): 159,
    (3000, 'North'): 156,
    (3100, 'South'): 84,
    (3100, 'North'): 87,
    (3500, 'South'): 121,
    (3500, 'North'): 124,
    (3900, 'South'): 64,
    (3900, 'North'): 64,
}

# Cross-section trains from reference Table 7 (already verified in 2.1.a)
cross_section = {
    (800, 'South'): 6,
    (800, 'North'): 6,
    (3000, 'South'): 6,
    (3000, 'North'): 5,
    (3100, 'South'): 3,
    (3100, 'North'): 3,
    (3500, 'South'): 4,
    (3500, 'North'): 4,
    (3900, 'South'): 2,
    (3900, 'North'): 3,
}

print("\nCross-section trains per line/direction:")
total_cs = 0
for (line, direction), cs in cross_section.items():
    print(f"  Line {line} {direction}: {cs} trains")
    total_cs += cs
print(f"Total cross-section trains: {total_cs}")

# ============================================================
# 3. Create Cross-Section Train Set
# ============================================================
# Each cross-section train is identified by (line, direction, index)
trains = []
train_info = {}  # Store line, direction, seat demand, max length for each train

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

print(f"\nTotal trains in set T: {len(trains)}")

# ============================================================
# 4. Parameters
# ============================================================
# Rolling stock types
U = ['PL3', 'PL4']

# Unit parameters
cost = {'PL3': 315000, 'PL4': 385000}  # Annual fixed cost (€)
capacity = {'PL3': 400, 'PL4': 600}    # Seat capacity
length = {'PL3': 80, 'PL4': 110}       # Length (m)

# ============================================================
# 5. Build Gurobi Model
# ============================================================
model = Model("RollingStock_Basic")
model.setParam('OutputFlag', 0)

# Decision variables: N[u,t] = number of units of type u assigned to train t
N = {}
for u in U:
    for t in trains:
        N[u, t] = model.addVar(vtype=GRB.INTEGER, lb=0, name=f"N_{u}_{t}")

model.update()

# Objective: Minimize total annual cost
model.setObjective(
    quicksum(cost[u] * N[u, t] for u in U for t in trains),
    GRB.MINIMIZE
)

# Constraint 1: Seat requirement for each train
for t in trains:
    model.addConstr(
        quicksum(capacity[u] * N[u, t] for u in U) >= train_info[t]['seat_demand'],
        name=f"seats_{t}"
    )

# Constraint 2: Length limit for each train
for t in trains:
    model.addConstr(
        quicksum(length[u] * N[u, t] for u in U) <= train_info[t]['max_length'],
        name=f"length_{t}"
    )

# Constraint 3: Balance constraint (at most 25% difference)
total_PL3 = quicksum(N['PL3', t] for t in trains)
total_PL4 = quicksum(N['PL4', t] for t in trains)

model.addConstr(total_PL3 <= 1.25 * total_PL4, name="balance_PL3")
model.addConstr(total_PL4 <= 1.25 * total_PL3, name="balance_PL4")

# Solve
start_time = time.time()
model.optimize()
end_time = time.time()
runtime = end_time - start_time

# ============================================================
# 6. Output Results
# ============================================================
if model.status == GRB.OPTIMAL:
    print("\n" + "=" * 70)
    print("OPTIMAL SOLUTION FOUND")
    print("=" * 70)
    
    # Total costs and units
    total_cost = model.objVal
    total_PL3_units = sum(N['PL3', t].X for t in trains)
    total_PL4_units = sum(N['PL4', t].X for t in trains)
    
    print(f"\nOptimal annual cost: €{total_cost:,.0f}")
    print(f"Runtime: {runtime:.4f} seconds")
    print(f"Total PL3 units: {total_PL3_units:.0f}")
    print(f"Total PL4 units: {total_PL4_units:.0f}")
    print(f"Total units: {total_PL3_units + total_PL4_units:.0f}")
    
    # Verify balance constraint
    if total_PL4_units > 0:
        ratio = total_PL3_units / total_PL4_units
        print(f"PL3/PL4 ratio: {ratio:.3f} (must be between 0.8 and 1.25)")
    
    # Composition for each train
    print("\n" + "-" * 70)
    print("ROLLING STOCK COMPOSITION PER TRAIN")
    print("-" * 70)
    print(f"{'Train ID':<20} {'Line':<6} {'Dir':<6} {'Demand':<8} {'MaxLen':<8} {'PL3':<5} {'PL4':<5} {'Seats':<8} {'Length':<8}")
    print("-" * 70)
    
    for t in trains:
        pl3 = int(N['PL3', t].X)
        pl4 = int(N['PL4', t].X)
        seats = pl3 * capacity['PL3'] + pl4 * capacity['PL4']
        train_len = pl3 * length['PL3'] + pl4 * length['PL4']
        
        print(f"{t:<20} {train_info[t]['line']:<6} {train_info[t]['direction']:<6} "
              f"{train_info[t]['seat_demand']:<8} {train_info[t]['max_length']:<8} "
              f"{pl3:<5} {pl4:<5} {seats:<8} {train_len:<8}")
    
    # Summary by line and direction
    print("\n" + "-" * 70)
    print("SUMMARY BY LINE AND DIRECTION")
    print("-" * 70)
    print(f"{'Line':<6} {'Direction':<10} {'Trains':<8} {'PL3':<8} {'PL4':<8} {'Total Units':<12}")
    print("-" * 70)
    
    for (line, direction), num_trains in cross_section.items():
        pl3_sum = sum(N['PL3', t].X for t in trains 
                      if train_info[t]['line'] == line and train_info[t]['direction'] == direction)
        pl4_sum = sum(N['PL4', t].X for t in trains 
                      if train_info[t]['line'] == line and train_info[t]['direction'] == direction)
        print(f"{line:<6} {direction:<10} {num_trains:<8} {pl3_sum:<8.0f} {pl4_sum:<8.0f} {pl3_sum + pl4_sum:<12.0f}")

else:
    print(f"No optimal solution found. Status: {model.status}")