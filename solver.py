from ortools.sat.python import cp_model
import datetime as dt

HOURS_IN_DAY = 24
DAYS_IN_WEEK = 7


def generate_work_patterns(shift_type: tuple[dt.time, dt.time, int]):
    start_time, end_time, num_days_off = shift_type
    start_hour = start_time.hour
    end_hour = end_time.hour
    if end_hour < start_hour:
        end_hour += HOURS_IN_DAY
    if num_days_off == 0:
        work_patterns = [[True] * DAYS_IN_WEEK]
    elif num_days_off == DAYS_IN_WEEK:
        work_patterns = [[False] * DAYS_IN_WEEK]
    else:
        import itertools
        combs = list(itertools.combinations(range(DAYS_IN_WEEK), num_days_off))
        work_patterns = [[i not in comb for i in range(DAYS_IN_WEEK)] for comb in combs]

    shift_subtypes = []
    for pattern in work_patterns:
        arr = [0] * HOURS_IN_DAY * DAYS_IN_WEEK
        for i in range(DAYS_IN_WEEK):
            for h in range(HOURS_IN_DAY):
                if (
                    pattern[i] and start_hour <= h < end_hour
                    or pattern[(i-1) % DAYS_IN_WEEK] and start_hour <= h + HOURS_IN_DAY < end_hour
                ):
                    arr[i * HOURS_IN_DAY + h] = 1
        shift_subtypes.append((arr, pattern, start_time, end_time))
    return shift_subtypes


def recommend_shifts(shift_types: list[tuple[dt.time, dt.time, int]], requirements: list[list[int]]):

    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    shift_subtypes = []
    for shift_type in shift_types:
        shift_subtypes.extend(generate_work_patterns(shift_type))
    
    model = cp_model.CpModel()
    num_needed = [
        model.NewIntVar(0, max([max(row) for row in requirements]), f"num_needed_{i}") 
        for i in range(len(shift_subtypes))
    ]

    cnt_workers = [
        model.NewIntVar(0, 10**7, f"cnt_workers_{i}") 
        for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)
    ]
    
    for h in range(HOURS_IN_DAY * DAYS_IN_WEEK):
        model.Add(
            cnt_workers[h] == 
            sum(
                num_needed[i] * shift_subtypes[i][0][h] 
                for i in range(len(shift_subtypes))
            )
        )

    requirements_flattened = [item for row in requirements for item in row]

    understaff_vars = [
        model.NewIntVar(0, 10**7, f"understaff_{i}") 
        for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)
    ]
    squared_understaff_vars = [
        model.NewIntVar(0, 10**7, f"squared_understaff_{i}") 
        for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)
    ]
    overstaff_vars = [
        model.NewIntVar(0, 10**7, f"overstaff_{i}") 
        for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)
    ]
    squared_overstaff_vars = [
        model.NewIntVar(0, 10**7, f"squared_overstaff_{i}") 
        for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)
    ]

    for h in range(HOURS_IN_DAY * DAYS_IN_WEEK):
        model.AddMaxEquality(
            understaff_vars[h], 
            [0, requirements_flattened[h] - cnt_workers[h]]
        )
        model.AddMaxEquality(
            overstaff_vars[h], 
            [0, cnt_workers[h] - requirements_flattened[h]]
        )
        model.AddMultiplicationEquality(
            squared_understaff_vars[h], 
            [understaff_vars[h], understaff_vars[h]]
        )
        model.AddMultiplicationEquality(
            squared_overstaff_vars[h], 
            [overstaff_vars[h], overstaff_vars[h]]
        )

    total_understaff = model.NewIntVar(0, 10**7, "total_understaff")
    total_squared_understaff = model.NewIntVar(0, 10**7, "total_squared_understaff")
    total_overstaff = model.NewIntVar(0, 10**7, "total_overstaff")
    total_squared_overstaff = model.NewIntVar(0, 10**7, "total_squared_overstaff")

    model.Add(total_understaff == sum(understaff_vars))
    model.Add(total_squared_understaff == sum(squared_understaff_vars))
    model.Add(total_overstaff == sum(overstaff_vars))
    model.Add(total_squared_overstaff == sum(squared_overstaff_vars))

    model.Minimize(2*total_squared_understaff + total_squared_overstaff)

    solver = cp_model.CpSolver()
    var_printer = cp_model.VarArraySolutionPrinter([total_understaff, total_overstaff])
    status = solver.Solve(model, var_printer)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        staffed = [solver.Value(cnt_workers[i]) for i in range(HOURS_IN_DAY * DAYS_IN_WEEK)]
        selected_shifts = []
        for i in range(len(shift_subtypes)):
            days_off = [
                weekday for d, weekday in enumerate(weekdays) 
                if not shift_subtypes[i][1][d]
            ]
            for _ in range(solver.Value(num_needed[i])):
                selected_shifts.append(
                    (shift_subtypes[i][2], shift_subtypes[i][3], days_off)  # start_time, end_time, days_off
                )
        return staffed, selected_shifts
