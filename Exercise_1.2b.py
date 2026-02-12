"""
Exercise 1.2b: PESP Timetabling Model with Extended Line 3900 (High-Frequency Service)
Synchronization is relaxed using PESP framework with appropriate bounds.
"""

import pandas as pd
from gurobipy import Model, GRB, quicksum

# ============================================================
# 1. Read Data
# ============================================================
xl = pd.ExcelFile('a2_part1.xlsx')
travel_times_df = pd.read_excel(xl, sheet_name='Travel Times')

# Create travel time dictionary (bidirectional)
travel_time = {}
for _, row in travel_times_df.iterrows():
    travel_time[(row['From'], row['To'])] = row['Travel Time']
    travel_time[(row['To'], row['From'])] = row['Travel Time']

# Define lines - Line 3900 is EXTENDED to Amsterdam
lines = {
    800: ['Amr', 'Asd', 'Ut', 'Ehv', 'Std', 'Mt'],
    3000: ['Hdr', 'Amr', 'Asd', 'Ut', 'Nm'],
    3100: ['Shl', 'Ut', 'Nm'],
    3500: ['Shl', 'Ut', 'Ehv', 'Vl'],
    3900: ['Asd', 'Ut', 'Ehv', 'Std', 'Hrl']  # EXTENDED: now starts from Asd
}

T = 30  # Period time

# ============================================================
# 2. Create Events
# ============================================================
events = []
event_idx = {}

for line, stops in lines.items():
    for direction in ['South', 'North']:
        route = stops if direction == 'South' else stops[::-1]
        
        for i, station in enumerate(route):
            if i == 0:
                event = (line, direction, station, 'dep')
                events.append(event)
                event_idx[event] = len(events) - 1
            elif i == len(route) - 1:
                event = (line, direction, station, 'arr')
                events.append(event)
                event_idx[event] = len(events) - 1
            else:
                event_arr = (line, direction, station, 'arr')
                event_dep = (line, direction, station, 'dep')
                events.append(event_arr)
                event_idx[event_arr] = len(events) - 1
                events.append(event_dep)
                event_idx[event_dep] = len(events) - 1

print(f"Total events: {len(events)}")

# ============================================================
# 3. Create Activities
# ============================================================
activities = []

# 3.1 Driving activities
for line, stops in lines.items():
    for direction in ['South', 'North']:
        route = stops if direction == 'South' else stops[::-1]
        
        for i in range(len(route) - 1):
            from_station = route[i]
            to_station = route[i + 1]
            
            from_event = (line, direction, from_station, 'dep')
            to_event = (line, direction, to_station, 'arr')
            
            tt = travel_time.get((from_station, to_station))
            if tt is None:
                print(f"Warning: No travel time for {from_station} -> {to_station}")
                continue
            
            activities.append({
                'type': 'driving',
                'from': from_event,
                'to': to_event,
                'l': tt,
                'u': tt
            })

# 3.2 Dwell activities (2-8 minutes)
for line, stops in lines.items():
    for direction in ['South', 'North']:
        route = stops if direction == 'South' else stops[::-1]
        
        for i in range(1, len(route) - 1):
            station = route[i]
            arr_event = (line, direction, station, 'arr')
            dep_event = (line, direction, station, 'dep')
            
            activities.append({
                'type': 'dwell',
                'from': arr_event,
                'to': dep_event,
                'l': 2,
                'u': 8
            })

# 3.3 Synchronization activities
# 4 trains/hour sections: exact 15-minute sync
sync_sections_4trains = [
    ('Shl', 'Ut', 3100, 3500),
    ('Ut', 'Nm', 3000, 3100),
]

for from_st, to_st, line1, line2 in sync_sections_4trains:
    for direction in ['South', 'North']:
        if direction == 'South':
            dep_station = from_st
        else:
            dep_station = to_st
        
        route1 = lines[line1] if direction == 'South' else lines[line1][::-1]
        route2 = lines[line2] if direction == 'South' else lines[line2][::-1]
        
        if dep_station in route1 and dep_station in route2:
            idx1 = route1.index(dep_station)
            idx2 = route2.index(dep_station)
            
            if idx1 < len(route1) - 1 and idx2 < len(route2) - 1:
                event1 = (line1, direction, dep_station, 'dep')
                event2 = (line2, direction, dep_station, 'dep')
                
                activities.append({
                    'type': 'sync',
                    'from': event1,
                    'to': event2,
                    'l': 15,
                    'u': 15
                })

# 6 trains/hour sections: RELAXED sync within PESP framework
# For 3 lines, ideal spacing is: two pairs ~10 min, one pair ~20 min
# We specify: first two pairs [8,12], third pair [18,22]
sync_sections_6trains = [
    # (station1, station2, [(line1, line2, l, u), ...])
    ('Asd', 'Ut', [
        (800, 3000, 8, 12),   # ~10 min
        (3000, 3900, 8, 12),  # ~10 min
        (800, 3900, 18, 22),  # ~20 min
    ]),
    ('Ut', 'Ehv', [
        (800, 3500, 8, 12),   # ~10 min
        (3500, 3900, 8, 12),  # ~10 min
        (800, 3900, 18, 22),  # ~20 min
    ]),
]

for from_st, to_st, line_pairs in sync_sections_6trains:
    for direction in ['South', 'North']:
        if direction == 'South':
            dep_station = from_st
        else:
            dep_station = to_st
        
        for line1, line2, l_bound, u_bound in line_pairs:
            route1 = lines[line1] if direction == 'South' else lines[line1][::-1]
            route2 = lines[line2] if direction == 'South' else lines[line2][::-1]
            
            if dep_station in route1 and dep_station in route2:
                idx1 = route1.index(dep_station)
                idx2 = route2.index(dep_station)
                
                if idx1 < len(route1) - 1 and idx2 < len(route2) - 1:
                    event1 = (line1, direction, dep_station, 'dep')
                    event2 = (line2, direction, dep_station, 'dep')
                    
                    activities.append({
                        'type': 'relaxed_sync',
                        'from': event1,
                        'to': event2,
                        'l': l_bound,
                        'u': u_bound
                    })

# 3.4 Headway activities at Utrecht
# Now we have MORE trains: 800, 3000, 3900 from Asd direction + 3100, 3500 from Shl direction
headway_pairs = []

# Southbound arrivals at Utrecht: Shl lines vs Asd lines
shl_lines = [3100, 3500]
asd_lines = [800, 3000, 3900]  # 3900 now comes from Asd

for shl_line in shl_lines:
    for asd_line in asd_lines:
        headway_pairs.append(((shl_line, 'South', 'Ut', 'arr'), (asd_line, 'South', 'Ut', 'arr')))

# Northbound departures from Utrecht: same pairs
for shl_line in shl_lines:
    for asd_line in asd_lines:
        headway_pairs.append(((shl_line, 'North', 'Ut', 'dep'), (asd_line, 'North', 'Ut', 'dep')))

for e1, e2 in headway_pairs:
    activities.append({
        'type': 'headway',
        'from': e1,
        'to': e2,
        'l': 3,
        'u': T - 3
    })

# NOTE: Transfer constraints at Eindhoven are DROPPED (all passengers can travel directly)

# Print activity counts
print(f"\nTotal activities: {len(activities)}")
activity_counts = {}
for a in activities:
    t = a['type']
    activity_counts[t] = activity_counts.get(t, 0) + 1
print("Activity counts:", activity_counts)

# ============================================================
# 4. Build Gurobi Model (Pure PESP - no extra variables)
# ============================================================
model = Model("PESP_HighFrequency")
model.setParam('OutputFlag', 0)

# Decision variables
pi = {}
for e in events:
    pi[e] = model.addVar(lb=0, ub=T, name=f"pi_{e}")

x = {}
p = {}
for i, a in enumerate(activities):
    x[i] = model.addVar(lb=a['l'], ub=a['u'], name=f"x_{i}")
    p[i] = model.addVar(vtype=GRB.INTEGER, lb=0, name=f"p_{i}")

model.update()

# Constraints for all activities (standard PESP constraint)
for i, a in enumerate(activities):
    e_from = a['from']
    e_to = a['to']
    model.addConstr(x[i] == pi[e_to] - pi[e_from] + T * p[i], name=f"activity_{i}")

# Fixed departure time: Line 3500 departs Schiphol at .09
fixed_event = (3500, 'South', 'Shl', 'dep')
model.addConstr(pi[fixed_event] == 9, name="fixed_3500_Shl")

# Objective: Minimize total dwell time only (no transfer constraints in this model)
dwell_terms = []
for i, a in enumerate(activities):
    if a['type'] == 'dwell':
        dwell_terms.append(x[i])

model.setObjective(quicksum(dwell_terms), GRB.MINIMIZE)

# Solve
model.optimize()

# ============================================================
# 5. Output Results
# ============================================================
if model.status == GRB.OPTIMAL:
    print("\n" + "=" * 60)
    print("OPTIMAL TIMETABLE FOUND (High-Frequency Service)")
    print("=" * 60)
    
    # Calculate objective
    total_dwell = sum(x[i].X for i, a in enumerate(activities) if a['type'] == 'dwell')
    
    print(f"Objective value (total dwell time): {model.objVal:.0f} minutes")
    
    # Show relaxed sync intervals
    print("\nRelaxed Synchronization Intervals (6 trains/hour sections):")
    print("-" * 60)
    print(f"{'Section':<12} {'Direction':<10} {'Lines':<12} {'Target':<10} {'Actual':<10}")
    print("-" * 60)
    for i, a in enumerate(activities):
        if a['type'] == 'relaxed_sync':
            e1, e2 = a['from'], a['to']
            interval = x[i].X
            line1, dir1, station1, _ = e1
            line2, _, _, _ = e2
            target = "~10 min" if a['u'] <= 12 else "~20 min"
            print(f"  {station1:<10} {dir1:<10} {line1}-{line2:<8} {target:<10} {interval:.0f} min")
    
    # Output timetable
    print("\n" + "-" * 60)
    print("TIMETABLE (times in minutes past the hour, mod 30)")
    print("-" * 60)
    
    for line in sorted(lines.keys()):
        for direction in ['South', 'North']:
            route = lines[line] if direction == 'South' else lines[line][::-1]
            
            print(f"\nLine {line} ({direction}bound): {' -> '.join(route)}")
            print(f"{'Station':<8} {'Arr':>6} {'Dep':>6}")
            print("-" * 22)
            
            for i, station in enumerate(route):
                if i == 0:
                    dep_event = (line, direction, station, 'dep')
                    dep_time = int(round(pi[dep_event].X)) % T
                    print(f"{station:<8} {'--':>6} {dep_time:>6}")
                elif i == len(route) - 1:
                    arr_event = (line, direction, station, 'arr')
                    arr_time = int(round(pi[arr_event].X)) % T
                    print(f"{station:<8} {arr_time:>6} {'--':>6}")
                else:
                    arr_event = (line, direction, station, 'arr')
                    dep_event = (line, direction, station, 'dep')
                    arr_time = int(round(pi[arr_event].X)) % T
                    dep_time = int(round(pi[dep_event].X)) % T
                    print(f"{station:<8} {arr_time:>6} {dep_time:>6}")
    
    # Compact format
    print("\n" + "=" * 60)
    print("COMPACT TIMETABLE FOR REPORT")
    print("=" * 60)
    
    for direction in ['South', 'North']:
        print(f"\n{direction}bound Timetable:")
        for line in sorted(lines.keys()):
            route = lines[line] if direction == 'South' else lines[line][::-1]
            times = []
            for i, station in enumerate(route):
                if i == 0:
                    dep_event = (line, direction, station, 'dep')
                    dep_time = int(round(pi[dep_event].X)) % T
                    times.append(f"d{dep_time:02d}")
                elif i == len(route) - 1:
                    arr_event = (line, direction, station, 'arr')
                    arr_time = int(round(pi[arr_event].X)) % T
                    times.append(f"a{arr_time:02d}")
                else:
                    arr_event = (line, direction, station, 'arr')
                    dep_event = (line, direction, station, 'dep')
                    arr_time = int(round(pi[arr_event].X)) % T
                    dep_time = int(round(pi[dep_event].X)) % T
                    times.append(f"a{arr_time:02d}/d{dep_time:02d}")
            
            route_str = " - ".join(route)
            times_str = " | ".join(times)
            print(f"  Line {line}: {route_str}")
            print(f"           {times_str}")

    # Compare with basic model
    print("\n" + "=" * 60)
    print("COMPARISON WITH BASIC MODEL (1.1.e)")
    print("=" * 60)
    print(f"Basic model objective (dwell + transfer): 60 minutes")
    print(f"High-frequency model objective (dwell only): {model.objVal:.0f} minutes")
    print(f"Note: Transfer constraints are dropped in high-frequency model")

else:
    print(f"No optimal solution found. Status: {model.status}")