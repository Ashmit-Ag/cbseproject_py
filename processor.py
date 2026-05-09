import re
import os
import glob
import pandas as pd

class Subject:
    def __init__(self, code, name, marks, grade):
        self.code = code
        self.name = name
        self.marks = marks
        self.grade = grade

class Student:
    def __init__(self, roll, name):
        self.roll = roll
        self.name = name
        self.subjects = []
        self.total = 0
        self.percentage = 0.0
        self.result = ""
        self.total_merit = 0
        self.percentage_merit = 0.0

def load_subject_codes(path):
    df = pd.read_excel(path)
    code_map = dict(zip(df['Code'], df['Subject']))
    return code_map

def parse_gazette(path, code_map):
    students = []
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for i in range(len(lines)-1):
        line = lines[i].strip()
        if re.match(r"^\d{8}", line):  # roll no starts with 8 digits
            parts = line.split()
            roll = parts[0]
            # name until subject code appears
            name_parts = []
            for p in parts[2:]:
                if re.match(r"^\d{3}$", p):
                    break
                name_parts.append(p)
            name = " ".join(name_parts)
            subject_codes = [int(p) for p in parts if re.match(r"^\d{3}$", p)]

            student = Student(roll, name)

            marks_line = lines[i+1].strip().split()
            for j, code in enumerate(subject_codes):
                idx = j*2
                if idx+1 < len(marks_line):  # safe check
                    try:
                        marks = int(marks_line[idx])
                    except ValueError:
                        marks = 0
                    grade = marks_line[idx+1]
                else:
                    marks = 0
                    grade = "NA"
                subj = Subject(code, code_map.get(code, str(code)), marks, grade)
                student.subjects.append(subj)

            students.append(student)
    return students

def calculate_results(students):
    for s in students:
        # CBSE official calculation: first five subjects (not best five)
        first_five = s.subjects[:5]
        s.total = sum(x.marks for x in first_five)
        s.percentage = round((s.total/500.0)*100, 2)

        fails = sum(1 for x in first_five if x.grade == "E")
        if fails == 0:
            s.result = "PASS"
        elif fails == 1:
            s.result = "COMPARTMENT"
        else:
            s.result = "FAIL"

def calculate_merit_list(students):
    for s in students:
        language = next((x for x in s.subjects if x.code in [1,2,3,22,301,302,303,322]), None)
        language_pass = language and language.grade != "E"
        other_subjects = [x for x in s.subjects if language is None or x.code != language.code]
        passed = [x for x in other_subjects if x.grade != "E"]

        if language_pass:
            best_four = sum(x.marks for x in sorted(passed, key=lambda x: x.marks, reverse=True)[:4])
            s.total_merit = language.marks + best_four
            s.percentage_merit = round((s.total_merit/500.0)*100, 2)
        else:
            s.total_merit = 0
            s.percentage_merit = 0.0

def export_excel(students, path):
    all_subjects = sorted({subj.name for s in students for subj in s.subjects})

    # 1. Original Result (exactly six subjects: Code, Marks, Grade)
    orig_data = []
    for s in students:
        row = {"Roll No": s.roll, "Name": s.name}
        for idx, subj in enumerate(s.subjects[:6], start=1):
            row[f"Subject{idx} Code"] = subj.code
            row[f"Subject{idx} Marks"] = subj.marks
            row[f"Subject{idx} Grade"] = subj.grade
        row["Total"] = s.total
        row["Percentage"] = s.percentage
        row["Result"] = s.result
        orig_data.append(row)
    df_orig = pd.DataFrame(orig_data).sort_values(by="Total", ascending=False).reset_index(drop=True)
    df_orig.insert(0, "Rank", df_orig.index+1)

    # 2. Merit List Sorted (English + best 4, subject-wise segregation)
    merit_data = []
    for s in students:
        row = {"Roll No": s.roll, "Name": s.name}
        for subj in s.subjects:
            row[f"{subj.name} Marks"] = subj.marks
            row[f"{subj.name} Grade"] = subj.grade
        row["Total"] = s.total_merit
        row["Percentage"] = s.percentage_merit
        row["Result"] = s.result
        merit_data.append(row)
    df_merit_sorted = pd.DataFrame(merit_data).sort_values(by="Total", ascending=False).reset_index(drop=True)
    df_merit_sorted.insert(0, "Rank", df_merit_sorted.index+1)

    # 3. Subject Toppers (include all tied toppers)
    toppers = []
    for subj in all_subjects:
        subj_scores = [(s.roll, s.name, x.marks, x.grade)
                       for s in students for x in s.subjects
                       if x.name == subj and x.grade != "E"]
        if subj_scores:
            max_marks = max(t[2] for t in subj_scores)
            for roll, name, marks, grade in subj_scores:
                if marks == max_marks:
                    toppers.append({"Subject": subj, "Roll No": roll,
                                    "Name": name, "Marks": marks, "Grade": grade})
    df_toppers = pd.DataFrame(toppers)

    # 4. Failures
    fails = []
    for s in students:
        for subj in s.subjects:
            if subj.grade == "E":
                fails.append({"Subject": subj.name, "Roll No": s.roll,
                              "Name": s.name, "Marks": subj.marks, "Grade": subj.grade})
    df_fails = pd.DataFrame(fails)

    # Handle overwrite: if file exists, create new numbered file
    base, ext = os.path.splitext(path)
    counter = 1
    new_path = path
    while os.path.exists(new_path):
        counter += 1
        new_path = f"{base}_{counter}{ext}"

    # Write all sheets
    with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
        df_orig.to_excel(writer, sheet_name="Original Result", index=False)
        df_merit_sorted.to_excel(writer, sheet_name="Merit List Sorted", index=False)
        df_toppers.to_excel(writer, sheet_name="Subject Toppers", index=False)
        df_fails.to_excel(writer, sheet_name="Failures", index=False)

    print(f"Analysis complete. File saved as {new_path}")

def process_result(txt_path, subject_path, output_path):
    code_map = load_subject_codes(subject_path)
    students = parse_gazette(txt_path, code_map)
    calculate_results(students)
    calculate_merit_list(students)
    export_excel(students, output_path)
