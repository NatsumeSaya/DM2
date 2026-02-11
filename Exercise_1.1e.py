"""
Exercise 1.1e: PESP Timetabling Model for A2-corridor
Author: Jie Xiao
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

# Define lines
lines = {
    800: ['Amr', 'Asd', 'Ut', 'Ehv', 'Std', 'Mt'],
    3000: ['Hdr', 'Amr', 'Asd', 'Ut', 'Nm'],
    3100: ['Shl', 'Ut', 'Nm'],
    3500: ['Shl', 'Ut', 'Ehv', 'Vl'],
    3900: ['Ehv', 'Std', 'Hrl']
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
                # Origin: departure only
                event = (line, direction, station, 'dep')
                events.append(event)
                event_idx[event] = len(events) - 1
            elif i == len(route) - 1:
                # Destination: arrival only
                event = (line, direction, station, 'arr')
                events.append(event)
                event_idx[event] = len(events) - 1
            else:
                # Intermediate: arrival + departure
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

# 3.1 Driving activities (fixed running time)
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

# 3.2 Dwell activities (2-8 minutes at intermediate stations)
for line, stops in lines.items():
    for direction in ['South', 'North']:
        route = stops if direction == 'South' else stops[::-1]
        
        for i in range(1, len(route) - 1):  # Intermediate stations only
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

# 3.3 Synchronization activities (15 min)
sync_sections = [
    ('Amr', 'Asd', 800, 3000),
    ('Asd', 'Ut', 800, 3000),
    ('Shl', 'Ut', 3100, 3500),
    ('Ut', 'Nm', 3000, 3100),
    ('Ut', 'Ehv', 800, 3500),
    ('Ehv', 'Std', 800, 3900)
]

for from_st, to_st, line1, line2 in sync_sections:
    for direction in ['South', 'North']:
        # Determine departure station based on direction
        if direction == 'South':
            dep_station = from_st
        else:
            dep_station = to_st
        
        # Get routes for both lines
        route1 = lines[line1] if direction == 'South' else lines[line1][::-1]
        route2 = lines[line2] if direction == 'South' else lines[line2][::-1]
        
        # Check if departure station is not the final stop for both lines
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

# 3.4 Headway activities at Utrecht (between different directions)
# Southbound arrivals: Shl lines (3100, 3500) vs Asd lines (800, 3000)
# Northbound departures: same pairs
headway_pairs = [
    # Southbound arrivals at Utrecht
    ((3500, 'South', 'Ut', 'arr'), (800, 'South', 'Ut', 'arr')),
    ((3500, 'South', 'Ut', 'arr'), (3000, 'South', 'Ut', 'arr')),
    ((3100, 'South', 'Ut', 'arr'), (800, 'South', 'Ut', 'arr')),
    ((3100, 'South', 'Ut', 'arr'), (3000, 'South', 'Ut', 'arr')),
    # Northbound departures from Utrecht
    ((3500, 'North', 'Ut', 'dep'), (800, 'North', 'Ut', 'dep')),
    ((3500, 'North', 'Ut', 'dep'), (3000, 'North', 'Ut', 'dep')),
    ((3100, 'North', 'Ut', 'dep'), (800, 'North', 'Ut', 'dep')),
    ((3100, 'North', 'Ut', 'dep'), (3000, 'North', 'Ut', 'dep'))
]

for e1, e2 in headway_pairs:
    activities.append({
        'type': 'headway',
        'from': e1,
        'to': e2,
        'l': 3,
        'u': T - 3
    })

# 3.5 Transfer activities at Eindhoven (between 3500 and 3900 only)
# Note: 800 and 3900 are synchronized (15 min apart), so transfer time would exceed 5 min
transfer_pairs = [
    # Hrl -> Ut: arr(3900, North, Ehv) -> dep(3500, North, Ehv)
    ((3900, 'North', 'Ehv', 'arr'), (3500, 'North', 'Ehv', 'dep')),
    # Ut -> Hrl: arr(3500, South, Ehv) -> dep(3900, South, Ehv)
    ((3500, 'South', 'Ehv', 'arr'), (3900, 'South', 'Ehv', 'dep'))
]

for e1, e2 in transfer_pairs:
    activities.append({
        'type': 'transfer',
        'from': e1,
        'to': e2,
        'l': 2,
        'u': 5
    })

# Print activity counts
print(f"\nTotal activities: {len(activities)}")
activity_counts = {}
for a in activities:
    t = a['type']
    activity_counts[t] = activity_counts.get(t, 0) + 1
print("Activity counts:", activity_counts)

# =============================================================================================
# 4. Build Gurobi Model (Claude helped check if constraints is complete and correct the codes)
# =============================================================================================
model = Model("PESP")
model.setParam('OutputFlag', 0)  # Suppress solver output

# Decision variables
pi = {}  # Event times
for e in events:
    pi[e] = model.addVar(lb=0, ub=T, name=f"pi_{e}")

x = {}  # Activity durations
p = {}  # Period variables
for i, a in enumerate(activities):
    x[i] = model.addVar(lb=a['l'], ub=a['u'], name=f"x_{i}")
    p[i] = model.addVar(vtype=GRB.INTEGER, lb=0, name=f"p_{i}")

model.update()

# Constraint: Activity duration = pi_j - pi_i + T * p
for i, a in enumerate(activities):
    e_from = a['from']
    e_to = a['to']
    model.addConstr(x[i] == pi[e_to] - pi[e_from] + T * p[i], name=f"activity_{i}")

# Constraint: Fixed departure time - Line 3500 departs Schiphol at .09
fixed_event = (3500, 'South', 'Shl', 'dep')
model.addConstr(pi[fixed_event] == 9, name="fixed_3500_Shl")

# Objective: Minimize total dwell + transfer time
obj_terms = []
for i, a in enumerate(activities):
    if a['type'] in ['dwell', 'transfer']:
        obj_terms.append(x[i])

model.setObjective(quicksum(obj_terms), GRB.MINIMIZE)

# Solve
model.optimize()

# ============================================================
# 5. Output Results (with the help of Calude)
# ============================================================
if model.status == GRB.OPTIMAL:
    print("\n" + "=" * 60)
    print("OPTIMAL TIMETABLE FOUND")
    print("=" * 60)
    print(f"Objective value (total dwell + transfer time): {model.objVal:.0f} minutes")
    
    # Output timetable by line and direction
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
    
    # Compact summary table for report
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

else:
    print(f"No optimal solution found. Status: {model.status}")