import pandas as pd
import streamlit as st
import random
from ortools.sat.python import cp_model

# Load the data
def load_dummy():
    df_Course = pd.read_excel("input_data.xlsx", sheet_name="Courses")
    D = pd.read_excel("input_data.xlsx", sheet_name="Days")
    T = pd.read_excel("input_data.xlsx", sheet_name="Time")
    I = pd.read_excel("input_data.xlsx", sheet_name="Instructors")
    R = pd.read_excel("input_data.xlsx", sheet_name="Rooms")

    instructor_df = pd.read_excel("input_data.xlsx", sheet_name="instructor_course")
    enrollment_df = pd.read_excel("input_data.xlsx", sheet_name="enrollment")
    room_capacity_df = pd.read_excel("input_data.xlsx", sheet_name="room_capacity")
    instructor_availability_class = pd.read_excel("input_data.xlsx", sheet_name="instructor_availability")

    room_availability_class = pd.read_excel("input_data.xlsx", sheet_name="room_availability")

    student_courses_class = pd.read_excel("input_data.xlsx", sheet_name="student_courses")

    # Create a dictionary for storing dataframes to be displayed in Streamlit
    data = {}
    data["Courses"] = df_Course
    data["Courses"].columns = ["Course Name", "Credit"]

    data["Days"] = D
    data["Days"].columns = ["Day"]

    data["Time"] = T
    data["Time"].columns = ["Time"]

    data["Rooms"] = R
    data["Rooms"].columns = ["Room"]

    data["Instructors"] = I
    data["Instructors"].columns = ["Instructor"]

    data["Instructor Course"] = instructor_df
    data["Instructor Course"].columns = ["Instructor", "Course"]

    data["Enrollment"] = enrollment_df
    data["Enrollment"].columns = ["Course", "Enrollment"]

    data["Room Capacity"] = room_capacity_df
    data["Room Capacity"].columns = ["Room", "Capacity"]

    data["Instructor Availability"] = instructor_availability_class
    data["Instructor Availability"].columns = ["Instructor", "Availability", "Time", "x"]
    data["Instructor Availability"] = data["Instructor Availability"].drop(columns=["x"])

    data["Room Availability"] = room_availability_class
    data["Room Availability"].columns = ["Room", "Availability", "Time", "x"]
    data["Room Availability"] = data["Room Availability"].drop(columns=["x"])

    data["Student Courses"] = student_courses_class
    data["Student Courses"].columns = ["Student", "Courses"]
    return data

def solve_university_class_scheduling(C, D, T, R, I, enrollment, instructor,
                                      room_capacity, instructor_availability, room_availability,
                                      student_courses):
    # Initialize the CP-SAT model
    model = cp_model.CpModel()

    # Decision variables
    x = {}
    for c in C:
        for d in D:
            for t in T:
                for r in R:
                    for i in I:
                        x[(c, d, t, r, i)] = model.NewBoolVar(f'x_{c}_{d}_{t}_{r}_{i}')

    # Decision variables for student attendance
    student_attendance = {}
    for student, courses in student_courses.items():
        for c in courses:
            for d in D:
                for t in T:
                    for r in R:
                        for i in I:
                            student_attendance[(student, c, d, t, r, i)] = model.NewBoolVar(f'a_{student}_{c}_{d}_{t}_{r}_{i}')

    # Slack variables for unscheduled courses
    unscheduled = {}
    for c in C:
        unscheduled[c] = model.NewBoolVar(f'unscheduled_{c}')

    # Constraints
    # 1. Each course must be scheduled exactly once
    for c in C:
        model.Add(sum(x[(c, d, t, r, i)] for d in D for t in T for r in R for i in I) + unscheduled[c] == 1)
        #model.Add(course_scheduled + unscheduled[c] == 1)

    # 2. Room capacity constraints
    for d in D:
        for t in T:
            for r in R:
              for i in I:
                model.Add(sum(enrollment[c] * x[(c, d, t, r, i)] for c in C) <= room_capacity[r])

    # 3. Instructor availability constraints
    for c in C:
        for d in D:
            for t in T:
                for r in R:
                    for i in I:
                        if (i, d, t) not in instructor_availability:
                            model.Add(x[(c, d, t, r, i)] == 0)


    # 4. Time slot availability constraints
    for d in D:
        for t in T:
            for r in R:
                model.Add(sum(x[(c, d, t, r, i)] for c in C for i in I) <= 1)

    # 5. Room availability constraints
    for r in R:
        for d in D:
            for t in T:
                if (r, d, t) not in room_availability:
                    for i in I:
                        for c in C:
                            model.Add(x[(c, d, t, r, i)] == 0)
                else:
                    for i in I:
                        for c in C:
                            model.Add(x[(c, d, t, r, i)] <= room_availability[(r, d, t)])

    # 6. Student attendance constraints using channeling (OnlyEnforceIf)
    for student, courses in student_courses.items():
        for c in courses:
            for d in D:
                for t in T:
                    for r in R:
                        for i in I:
                            # Channeling constraint for student attendance
                            model.Add(student_attendance[(student, c, d, t, r, i)] == 1).OnlyEnforceIf(x[(c, d, t, r, i)])
                            model.Add(student_attendance[(student, c, d, t, r, i)] == 0).OnlyEnforceIf(x[(c, d, t, r, i)].Not())

    # 7. Each course must be taught by the assigned professor
    for c in C:
        assigned_prof = instructor[c]
        for d in D:
            for t in T:
                for r in R:
                    for i in I:
                        if i != assigned_prof:
                            model.Add(x[(c, d, t, r, i)] == 0)

    # 8. Each professor only teach one class at the same time
    for i in I:
        for d in D:
            for t in T:
                model.Add(sum(x[(c, d, t, r, i)] for c in C for r in R) <= 1)

    # constrain for 2 credist course
    for c in credits_select :
      for d in D :
        for r in R :
          for i in I :
            for k,t in enumerate(T[:-1]) :
              curr_c = c + "-1"
              next_c = c + "-2"
              curr_slot = T[k]
              next_slot = T[k+1]
              #model.Add(x[(curr_c, d, curr_slot, r, i)] == 1).OnlyEnforceIf(x[(next_c, d, next_slot, r, i)])
              #model.Add(x[(next_c, d, next_slot, r, i)] == 1).OnlyEnforceIf(x[(curr_c, d, curr_slot, r, i)])

              # 2. Ensure both parts of the course (curr_c and next_c) are scheduled in the same room (r) and with the same instructor (i)
              model.Add(x[(curr_c, d, curr_slot, r, i)] + x[(next_c, d, next_slot, r, i)] == 2).OnlyEnforceIf(x[(curr_c, d, curr_slot, r, i)])
              model.Add(x[(curr_c, d, curr_slot, r, i)] + x[(next_c, d, next_slot, r, i)] == 2).OnlyEnforceIf(x[(next_c, d, next_slot, r, i)])
              

              # Optional: Enforce that next_c cannot be scheduled unless curr_c is scheduled first
              #model.Add(x[(next_c, d, next_slot, r, i)] == 0).OnlyEnforceIf(x[(curr_c, d, curr_slot, r, i)].Not())

    for t in [T[0], T[-1]] :
      for r in R :
        for i in I :
          for d in D :
            for c in credits_select :
              if t == T[0] :
                model.Add(x[(c+"-2", d, t, r, i)] == 0)
              else :
                model.Add(x[(c+"-1", d, t, r, i)] == 0)

    # a student only attend a class in same time slot
    for student, courses in student_courses.items():
      for d in D :
        for t in T :
          model.Add(sum(x[(c, d, t, r, i)] for c in courses for r in R for i in I) <= 1)
                                                
    # Objective - minimize total room usage (example)
    penalty_weight = 100  # Set a high weight for penalty
    model.Minimize(sum(x[(c, d, t, r, i)] for c in C for d in D for t in T for r in R for i in I) +
                   penalty_weight * sum(unscheduled[c] for c in C))

    # Create a solver and solve the model
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 600  # Increase the time limit
    solver.parameters.num_search_workers = 4   # Use 4 threads
    solver.parameters.search_branching = 3  # Try different search strategies
    solver.parameters.log_search_progress = True  # Enable logging for progress
    status = solver.Solve(model)
    print("Optimization status : ", status)

    if status == cp_model.OPTIMAL:
        # Retrieve and print the schedule
        schedule = {}
        for c in C:
            for d in D:
                for t in T:
                    for r in R:
                        for i in I:
                            if solver.Value(x[(c, d, t, r, i)]) == 1:
                                schedule[c] = (d, t, r, i)
                                break
        # print("Optimal Schedule:")
        # for c, (d, t, r, i) in schedule.items():
        #     print(f"Course {c} scheduled on Day {d}, Time Slot {t}, Room {r}, Instructor {i}")

        result_unscheduled = []
        for c in C:
          if solver.Value(unscheduled[c]) == 1:
            result_unscheduled.append(c)

        return schedule, result_unscheduled, solver

    else:
        print("No optimal solution found.")
        return None

if 'generated_data' not in st.session_state:
    st.session_state.generated_data = True
if 'view_data' not in st.session_state:
    st.session_state.view_data = False
if 'run_job' not in st.session_state:
    st.session_state.run_job = False

# Streamlit sidebar with clickable tab menu
st.sidebar.title("Menu Tab")
tabs = [
    "Courses", "Days", "Time", "Rooms", "Instructor Availability", 
    "Instructor Course", "Enrollment", "Room Capacity", 
    "Room Availability", "Student Courses"
]

generate_dummy = st.sidebar.button("Generate Dummy Data")
selected_tab = st.sidebar.radio("View Data", tabs)
run_job = st.sidebar.button("Run Optimization")

if generate_dummy :
    st.session_state.generated_data = True
if selected_tab :
    st.session_state.view_data = True
if run_job :
    st.session_state.run_job = True
    st.session_state.view_data = False

if st.session_state.generated_data :
    st.title("Generate Dummy Data")

    num_courses_default = 80
    num_days_default = 5
    num_time_slots_default = 8
    num_rooms_default = 4
    num_instructors_default = 8
    num_students_default = 4

    # Sliders with default values
    num_courses = st.slider("Number of Courses", min_value=1, max_value=100, value=num_courses_default)
    num_days = st.slider("Number of Days", min_value=1, max_value=7, value=num_days_default)
    num_time_slots = st.slider("Number of Time Slots", min_value=1, max_value=12, value=num_time_slots_default)
    num_rooms = st.slider("Number of Rooms", min_value=1, max_value=20, value=num_rooms_default)
    num_instructors = st.slider("Number of Instructors", min_value=1, max_value=20, value=num_instructors_default)
    num_students = st.slider("Number of Students", min_value=1, max_value=100, value=num_students_default)

    generate_data = st.button("Generate Data")

    if generate_data:
        # Generate course names
        C = [f'C{i+1}' for i in range(num_courses)]

        # Course credits
        credits = {}
        for course in C:
            credits[course] = random.choice([1, 2])  # Assign random credits (1 or 2)

        credits_df = pd.DataFrame(list(credits.items()), columns=['Course', 'Credits'])
        credits_df.to_excel('data/credits.xlsx', index=False)

        # Days
        D = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

        # Time slots
        T = ['Sesi-1', 'Sesi-2', 'Sesi-3', 'Sesi-4', 'Sesi-5', 'Sesi-6', 'Sesi-7', 'Sesi-8']

        # Rooms
        R = [f'Room{i+1}' for i in range(num_rooms)]

        # Instructors
        I = [f'Prof{i+1}' for i in range(num_instructors)]

        # Assign instructors to courses
        instructor = {course: random.choice(I) for course in C}

        # Assign durations to courses (1 to 3 time slots)
        duration = {course: random.randint(1, 3) for course in C}

        # Assign enrollment numbers to courses (10 to 50 students)
        enrollment = {course: random.randint(10, 50) for course in C}

        # Assign room capacities (30 to 100)
        room_capacity = {room: random.randint(30, 100) for room in R}

        # Generate instructor availability
        instructor_availability = {}
        for instruct in I:
            for day in D:
                if random.random() <= 1:
                    for time in T:
                        instructor_availability[(instruct, day, time)] = 1

        # Generate room availability
        room_availability = {}
        for room in R:
            for time in T:
                for day in D:
                    room_availability[(room, day, time)] = 1

        # Generate student registrations
        student_courses = {}
        div_course = num_courses // num_students
        k = 0
        for student in range(1, num_students + 1):
            selected_course = C[k:k + div_course]
            student_courses[f'Student{student}'] = selected_course
            k += div_course

        data = {}
        data["Courses"] = credits_df
        data["Days"] = pd.DataFrame(D)
        data["Time"] = pd.DataFrame(T)
        data["Rooms"] = pd.DataFrame(R)
        data["Instructors"] = pd.DataFrame(I)

        instructor_class = pd.DataFrame(instructor.keys())
        instructor_class["instructor"] = instructor.values()
        data["instructor_course"] = instructor_class

        enrollment_class = pd.DataFrame(enrollment.keys())
        enrollment_class["jumlah"] = enrollment.values()
        data["enrollment"] = enrollment_class

        capacity = pd.DataFrame(room_capacity.keys())
        capacity["jumlah"] = room_capacity.values()
        data["room_capacity"] = capacity

        instructor_availability_class = pd.DataFrame(instructor_availability.keys())
        instructor_availability_class["jumlah"] = instructor_availability.values()
        data["instructor_availability"] = instructor_availability_class

        room_availability_class = pd.DataFrame(room_availability.keys())
        room_availability_class["jumlah"] = room_availability.values()
        data["room_availability"] = room_availability_class

        student_courses_class = pd.DataFrame(student_courses.keys())
        student_courses_class["jumlah"] = student_courses.values()
        data["student_courses"] = student_courses_class

        with pd.ExcelWriter('input_data.xlsx') as writer:
            for elem in data.keys():
                select_data = data[elem]
                select_data.to_excel(writer, sheet_name=elem, index=False)

        st.write("Data generated successfully")
        st.session_state.generated_data = False

elif st.session_state.view_data :
    dummy_data = load_dummy()
    st.title("View Data")
    st.write(f"Selected tab: {selected_tab}")
    st.dataframe(dummy_data[selected_tab])

elif st.session_state.run_job :
    with st.spinner("Optimization is running...") :

        df_C = pd.read_excel("input_data.xlsx", sheet_name="Courses")
        C = df_C["Course"].astype(str) + "-1"

        credits_select = df_C[df_C["Credits"] > 1]
        credits_select = credits_select["Course"].to_list()

        # add prefix "-2" on credits_select variable
        credits_select_modif = [i + "-2" for i in credits_select]

        C = C.to_list() + credits_select_modif

        D = pd.read_excel("input_data.xlsx", sheet_name="Days")
        D = D[0].to_list()

        T = pd.read_excel("input_data.xlsx", sheet_name="Time")
        T = T[0].to_list()

        I = pd.read_excel("input_data.xlsx", sheet_name="Instructors")
        I = I[0].to_list()

        R = pd.read_excel("input_data.xlsx", sheet_name="Rooms")
        R = R[0].to_list()

        instructor_df = pd.read_excel("input_data.xlsx", sheet_name="instructor_course")
        instructor = {}
        for row in instructor_df.values:
            if row[0] in credits_select :
                instructor[row[0] + "-1"] = row[1]
                instructor[row[0] + "-2"] = row[1]
            else :
                instructor[row[0] + "-1"] = row[1]

        enrollment_df = pd.read_excel("input_data.xlsx", sheet_name="enrollment")
        enrollment = {}
        for elem in enrollment_df.values :
            if elem[0] in credits_select :
                enrollment[elem[0] + "-1"] = elem[1]
                enrollment[elem[0] + "-2"] = elem[1]
            else :
                enrollment[elem[0] + "-1"] = elem[1]

        room_capacity_df = pd.read_excel("input_data.xlsx", sheet_name="room_capacity")
        room_capacity = {}
        for row in room_capacity_df.values:
            room_capacity[row[0]] = row[1]

        instructor_availability_class = pd.read_excel("input_data.xlsx", sheet_name="instructor_availability")
        instructor_availability = {}
        for row in instructor_availability_class.values:
            instructor_availability[(row[0], row[1], row[2])] = row[3]

        room_availability_class = pd.read_excel("input_data.xlsx", sheet_name="room_availability")
        room_availability = {}
        for row in room_availability_class.values:
            room_availability[(row[0], row[1], row[2])] = row[3]

        student_courses_class = pd.read_excel("input_data.xlsx", sheet_name="student_courses")
        student_courses = {}
        for row in student_courses_class.values:
            row[1] = row[1].replace("[","").replace("]","").replace("'","").split(", ")
        course = []
        for elem in row[1] :
            if elem in credits_select :
                course.append(elem + "-1")
                course.append(elem + "-2")
            else :
                course.append(elem + "-1")
        student_courses[row[0]] = course

        # Print the generated dataset
        print(f"C = {C}")
        print(f"D = {D}")
        print(f"T = {T}")
        print(f"R = {R}")
        print(f"I = {I}")
        print(f"instructor = {instructor}")
        print(f"enrollment = {enrollment}")
        print(f"room_capacity = {room_capacity}")
        print(f"instructor_availability = {instructor_availability}")
        print(f"room_availability = {room_availability}")
        print(f"student_courses = {student_courses}")

        data = {}
        data["Courses"] = pd.DataFrame(C)
        data["Days"] = pd.DataFrame(D)
        data["Time"] = pd.DataFrame(T)
        data["Rooms"] = pd.DataFrame(R)
        data["Instructors"] = pd.DataFrame(I)

        instructor_class = pd.DataFrame(instructor.keys())
        instructor_class["instructor"] = instructor.values()
        data["instructor_course"] = instructor_class

        enrollment_class = pd.DataFrame(enrollment.keys())
        enrollment_class["jumlah"] = enrollment.values()
        data["enrollment"] = enrollment_class

        capacity = pd.DataFrame(room_capacity.keys())
        capacity["jumlah"] = room_capacity.values()
        data["room_capacity"] = capacity

        instructor_availability_class = pd.DataFrame(instructor_availability.keys())
        instructor_availability_class["jumlah"] = instructor_availability.values()
        data["instructor_availability"] = instructor_availability_class

        room_availability_class = pd.DataFrame(room_availability.keys())
        room_availability_class["jumlah"] = room_availability.values()
        data["room_availability"] = room_availability_class

        student_courses_class = pd.DataFrame(student_courses.keys())
        student_courses_class["jumlah"] = student_courses.values()
        data["student_courses"] = student_courses_class

        with pd.ExcelWriter('input_data_modif.xlsx') as writer:
            for elem in data.keys() :
                select_data = data[elem]
                select_data.to_excel(writer, sheet_name=elem, index=False)

        C = pd.read_excel("input_data_modif.xlsx", sheet_name="Courses")
        C = C[0].to_list()

        D = pd.read_excel("input_data_modif.xlsx", sheet_name="Days")
        D = D[0].to_list()

        T = pd.read_excel("input_data_modif.xlsx", sheet_name="Time")
        T = T[0].to_list()

        I = pd.read_excel("input_data_modif.xlsx", sheet_name="Instructors")
        I = I[0].to_list()

        R = pd.read_excel("input_data_modif.xlsx", sheet_name="Rooms")
        R = R[0].to_list()

        instructor_df = pd.read_excel("input_data_modif.xlsx", sheet_name="instructor_course")
        instructor = {}
        for row in instructor_df.values:
            instructor[row[0]] = row[1]

        enrollment_df = pd.read_excel("input_data_modif.xlsx", sheet_name="enrollment")
        enrollment = {}
        for row in enrollment_df.values:
            enrollment[row[0]] = row[1]

        capacity_df = pd.read_excel("input_data_modif.xlsx", sheet_name="room_capacity")
        room_capacity = {}
        for row in capacity_df.values:
            room_capacity[row[0]] = row[1]

        instructor_availability_class = pd.read_excel("input_data_modif.xlsx", sheet_name="instructor_availability")
        instructor_availability = {}
        for row in instructor_availability_class.values:
            instructor_availability[(row[0], row[1], row[2])] = row[3]

        room_availability_class = pd.read_excel("input_data_modif.xlsx", sheet_name="room_availability")
        room_availability = {}
        for row in room_availability_class.values:
            room_availability[(row[0], row[1], row[2])] = row[3]

        student_courses_class = pd.read_excel("input_data_modif.xlsx", sheet_name="student_courses")
        student_courses = {}
        for row in student_courses_class.values:
            row[1] = row[1].replace("[","").replace("]","").replace("'","").split(", ")
            student_courses[row[0]] = row[1]

        result, unscheduled, solve = solve_university_class_scheduling(C, D, T, R, I, enrollment, instructor,
                                        room_capacity, instructor_availability, room_availability, student_courses)


        df_result = pd.DataFrame(result, index=["Day", "Time", "Room", "Instructor"]).T
        df_result = df_result.reset_index()

        for i in range(5) :
            select_df = df_result[df_result["Day"] == D[i]].reset_index(drop=True)
            select_df["index"] = select_df["index"] + "_" + select_df["Instructor"]
            select_df = select_df.drop(["Day","Instructor"], axis=1)
            select_df = select_df.pivot(index="Room", columns="Time", values="index")
            print(D[i])
            print(select_df)
            print("-"*80)

    st.write("Optimization completed")
