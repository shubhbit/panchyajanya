from openpyxl import load_workbook
import pandas as pd
import os
import time
import string
import random

#################################### CONFIGS ####################################
PRINT_LEN = 60
ROLL_BASE = 19220000
FILL_CHAR = "*"
file = "Panchyajanya_latest.xlsx" # change the file name xlsx which you want to read
output_file = "Admit_Cards"
#################################### CONFIGS ####################################

def random_string(chars):
    N = chars
    res = ''.join(random.choices(string.ascii_uppercase +
                                 string.digits, k=N))
    return res


columns_we_need = {"Regestration No.": 1, "Name": 2, "Father's Name": 3,
                   "Gender": 7, "Date of Birth": 8, "Class": 6, "School": 9, "Interest": 12}




def get_no_cols_rows(file):
    df = pd.read_excel(file)
    return len(df.columns), len(df)


NO_OF_COLS, NO_OF_ROWS = get_no_cols_rows(file)


def if_student_unique(final_data, row):
    is_unique = True
    for gender in final_data:
        for student in final_data[gender]:
            if student.lower().strip() == row[columns_we_need["Name"]].lower().strip():
                print("same name found....")
                print("old row: {0}: {1}\n".format(
                    student, final_data[gender][student]))
                print("New row: {0}\n".format(row))
                print(
                    "checking if father name is same as well, will ignore in case father is same\n")
                if final_data[gender][student]["Father's Name"].lower().strip() == row[columns_we_need["Father's Name"]].lower().strip():
                    print("Father name same, ignored...\n\n\n")
                    is_unique = False
                    break
                else:
                    print("Father name not same, not ignored...\n\n\n")
                    new_name = "{0}_{1}".format(
                        student, random_string(len(student)))
                    row.pop(columns_we_need["Name"])
                    row.insert(columns_we_need["Name"], new_name)
        if not is_unique:
            break
    return row, is_unique


def read_data(sheet_obj):
    final_data = {"Female": {}, "Male": {}}
    duplicate = 0
    for row_i in range(2, NO_OF_ROWS+2):
        rows = []
        for column_i in range(1, NO_OF_COLS+1):
            cell_obj = sheet_obj.cell(row=row_i, column=column_i)
            rows.append(cell_obj.value)
        row, is_unique = if_student_unique(final_data, rows)
        if is_unique:
            rows = row
            final_data[rows[columns_we_need['Gender']]][rows[columns_we_need['Name']]] = {
                k: rows[columns_we_need[k]] for k in columns_we_need.keys() if k != "Name"}
        else:
            duplicate += 1
    print("##################### Summary #####################")
    print("total unique students: ", len(final_data['Male'])+len(final_data['Female']))
    print("duplicates removed: ", duplicate)
    return final_data


def group_by_interest(final_data):
    data_by_interest = {"Female": {}, "Male": {}}
    for gender in final_data:
        for student in sorted(final_data[gender]):
            interest = final_data[gender][student]['Interest']
            if interest not in data_by_interest[gender]:
                data_by_interest[gender][interest] = [
                    {student: final_data[gender][student]}]
            else:
                data_by_interest[gender][interest].append(
                    {student: final_data[gender][student]})
    return data_by_interest


def assign_roll_number(student_by_interest):
    roll_base = ROLL_BASE
    for gender in student_by_interest:
        for students in student_by_interest[gender].values():
            for student in students:
                roll_base += 1
                for key in student:
                    student[key]["Roll No."] = roll_base
    return student_by_interest


def group_by_school(student_by_with_role):
    students_by_school = {}
    for gender in student_by_with_role:
        for students in student_by_with_role[gender].values():
            for student in students:
                for key in student:
                    if student[key]["School"] not in students_by_school:
                        students_by_school[student[key]["School"]] = [student]
                    else:
                        students_by_school[student[key]
                                           ["School"]].append(student)
    return students_by_school


def print_new(data_by_school):
    if os.path.exists(output_file):
        os.remove(output_file)
    st = """
{0}
{1}
{2}
{3}
{4}
{5}
{6}
"""
    for school in data_by_school:
        for student in data_by_school[school]:
            data_list = []
            for name in student:
                data_list.append("Panchyajanya".center(PRINT_LEN, FILL_CHAR))
                name_to_print = name.split("_")[0]
                data_list.append("Name: {}".format(name_to_print).center(PRINT_LEN, " "))
                for indx, k in enumerate(student[name]):
                    if indx < len(student[name])//2:
                        v = student[name][k]
                        if k == "Date of Birth":
                            v = str(v).split(" ")[0]
                        fi = "{}: {}".format(k, v)
                        new_key = list(student[name].keys())[indx+4]
                        space = PRINT_LEN//2-len(fi)
                        se = "{}{}: {}".format(
                            " "*space, new_key, student[name][new_key])
                        temp_d = "{}{}".format(fi.ljust(0), se)
                        data_list.append(temp_d)
                data_list.append("Admit Card".center(PRINT_LEN, FILL_CHAR))
            with open(output_file, "a") as f:
                f.write(st.format(*data_list))


def main():
    wb_obj = load_workbook(file)
    sheet_obj = wb_obj.active
    start_time = time.time()
    final_data = read_data(sheet_obj)
    student_by_interest = group_by_interest(final_data)
    students_with_roll = assign_roll_number(student_by_interest)
    students_by_school = group_by_school(students_with_roll)
    end_time = time.time()
    print_new(students_by_school)
    print("time took: {} seconds".format(end_time-start_time))
    print("##################### Summary #####################")

if __name__ == "__main__":
    main()
