import pandas as pd
import math
from fpdf import FPDF
import os


def main():
    def pdfgen(roll, table1, table2, table3, table4):
        pdf = FPDF()
        pdf.add_page()

        pdf.set_font("Times", size=18)
        pdf.cell(190, 10, txt=settings['School'], ln=1, align='C')

        pdf.set_font("Times", size=14)
        pdf.cell(190, 10, txt=exam + " Exam Mark Sheet", ln=1, align='C')

        pdf.set_font("Times", size=12)
        cell_width = (pdf.w - pdf.l_margin - pdf.r_margin)/len(table1[0])
        cell_height = 10
        top = pdf.y + 10
        left = pdf.x

        for i in table1:
            counter = 0
            for j in i:

                pdf.x = left + counter*cell_width

                pdf.y = top
                pdf.multi_cell(cell_width, cell_height, txt=j,  border=1)
                counter = counter + 1

        pdf.set_font("Times", size=11)
        cell_width = [10, 120, 30, 30]
        cell_height = 10
        top = pdf.y
        left = pdf.x

        for i in table2:
            counter = 0
            top = top + 10
            left = pdf.x
            for j, k in zip(i, cell_width):

                pdf.x = left

                left = left + k
                pdf.y = top
                pdf.multi_cell(k, cell_height, txt=str(j),  border=1)
                counter = counter + 1

        pdf.set_font("Times", size=12)
        cell_width = 30
        cell_height = 10
        top = pdf.y
        left = pdf.x

        for i in table4:
            counter = 0
            top = top + 10
            left = pdf.x
            for j in i:

                pdf.x = left + counter*cell_width

                pdf.y = top

                pdf.multi_cell(cell_width, cell_height,
                               txt=str(j),  border=1, align='C')
                counter = counter + 1

        pdf.set_font("Times", size=12)
        cell_width = (pdf.w - pdf.l_margin - pdf.r_margin)/len(table3[0])
        cell_height = 10
        top = pdf.y + 5
        left = pdf.x

        for i in table3:
            counter = 0
            top = top + 10
            left = pdf.x
            for j in i:

                pdf.x = left + counter*cell_width

                pdf.y = top
                pdf.multi_cell(cell_width, cell_height,
                               txt=str(j),  border=0, align='C')
                counter = counter + 1
        pdf.output(str("Mark_sheet/mark_sheet-"+str(roll)+".pdf"), dest='F')

    def gpmaker(a, b=0, c=0, d=0, total=100):
        s = a + b + c + d
        if (s >= 0.8*total):
            gp = 5.00
        elif (s >= 0.7*total):
            gp = 4.00
        elif (s >= 0.6*total):
            gp = 3.50
        elif (s >= 0.5*total):
            gp = 3.00
        elif (s >= 0.4*total):
            gp = 2.00
        elif (s >= 0.33*total):
            gp = 1.00
        else:
            gp = 0
        return gp

    def grademaker(gp):
        if (gp == 5.00):
            return 'A+'
        elif (gp >= 4.00):
            return 'A'
        elif (gp >= 3.50):
            return 'A-'
        elif (gp >= 3.00):
            return 'B'
        elif (gp >= 2.00):
            return 'C'
        elif (gp >= 1.00):
            return 'D'
        else:
            return 'F'

    if (os.path.exists("Mark_sheet")):
        pass
    else:
        os.mkdir("Mark_sheet")

    subject_by_code = {
        154: 'ICT'
    }

    settings = dict()
    setting_file = open('setting.txt', 'r').readlines()
    for setting in setting_file:
        [key, value] = setting.rstrip().split('=')
        settings[key] = value

    totals = settings['Totals'].split(',')
    exam = settings['Exam']
    subjects = settings['Subjects'].split(",")
    school = settings['School']
    class_ = settings['Class']
    number_of_students = settings['Number of Students']

    excel = settings['Excel']
    excels = []
    for row in excel.split("/"):
        column = []
        for col in row.split(","):
            column.append(col)
        excels.append(column)

    subject_file_names = settings['Subjects'].split(",")
    for subject_file_name in subject_file_names:
        filename = subject_file_name.replace(" ", "_").lower()
        exec(f"{filename} = pd.read_excel('data/{filename}.xlsx')")

    names = pd.read_excel('Data/names.xlsx')

    std_no = names.count()[0]

    subjects = settings['Subjects'].split(",")

    results = [["Roll", "Name"] + subjects + ["Total", "GPA", "Position"]]

    gpas = dict()
    tables = []
    for roll in range(0, std_no, 1):
        name = names.loc[roll].Name
        optional_code = names.loc[roll].Optional
        for subject_file_name in subject_file_names:
            filename = subject_file_name.replace(" ", "_").lower()
            exec(f"{filename}_mark = {filename}.loc[roll]")

        gp = []
        for i, subject_file_name in enumerate(subject_file_names):
            filename = subject_file_name.replace(" ", "_").lower()
            gpmaker_input = []
            for col in excels[i]:
                gpmaker_input.append(f"{filename}_mark.{col}")
            gpmaker_input.append(f"total={totals[i]}")
            mystr = f"{filename}_gp = gpmaker("+','.join(gpmaker_input)+")"

            exec(mystr)

            exec(f"gp.append({filename}_gp)")

        try:

            optional_sub = subject_by_code[optional_code]
        except:
            optional_sub = None

        if (optional_sub != None):
            number_of_subjects = len(subjects) - 1
        else:
            number_of_subjects = len(subjects)

        if (optional_sub != None):

            if 0.0 in gp:
                if (gp.index(0.0) == subjects.index(optional_sub)):
                    gpa = round(sum(gp)/number_of_subjects,2)
                else:
                    gpa = 0
            else:
                if (gp[subjects.index(optional_sub)] >= 2):
                    gpa = round(
                        (sum(gp)+gp[subjects.index(optional_sub)]-2)/number_of_subjects, 2)
        else:

            if 0.0 in gp:

                gpa = 0
            else:
                gpa = round(sum(gp)/number_of_subjects, 2)

        if (gpa > 5):
            gpa = 5.0
        gpas[roll+1] = gpa

        table1 = [["Name: "+str(name).title()+"\nRoll: " +
                   str(roll+1)+"\nSection: " + settings['Section']]]

        marks = []
        for i, subject_file_name in enumerate(subject_file_names):
            filename = subject_file_name.replace(" ", "_").lower()
            append_input = []
            for col in excels[i]:
                append_input.append(f"{filename}_mark.{col}")
            mystr = f"marks.append("+'+'.join(append_input)+")"

            exec(mystr)

        table2 = [["S.N.", "Subjects", "Marks", "Grade"]]
        for i in range(1, len(subjects)+1, 1):
            temp = []

            temp.append(i)
            if (optional_sub != None):
                if (i-1 == subjects.index(optional_sub)):
                    temp.append(subjects[i-1] + " (Optional)")
                else:
                    temp.append(subjects[i-1])
            else:
                temp.append(subjects[i-1])
            temp.append(marks[i-1])
            temp.append(grademaker(gp[i-1]))
            table2.append(temp)

        table4 = [["Total Marks", sum(marks)], ["GPA", "{:.2f}".format(gpa)], [
            "Grade", grademaker(gpa)]]

        table3 = [["", "--------------------------------"],
                  ["", "Signature of Head Teacher\n(" + settings['School'] + ")"]]

        # pdfgen(roll+1, table1, table2, table3, table4)

        tables.append([roll+1, table1, table2, table3, table4])
        gp = ["{:.2f}".format(g) for g in gp]
        result = [roll+1, name.title()] + gp + [sum(marks),
                                                "{:.2f}".format(float(gpa))]

        results.append(result)

    lst = sorted(gpas.items(), key=lambda kv: kv[1], reverse=True)
    lst = sorted(lst, key=lambda kv: results[kv[0]][-2], reverse=True)
    merit = dict()
    count = 1
    for l in lst:
        merit[l[0]] = count
        count += 1

    lst = sorted(merit.items(), key=lambda kv: kv[0])
    pos = []
    for l in lst:
        pos.append(l[1])

    for i in range(len(results)-1):

        results[i+1].append(pos[i])

    df = pd.DataFrame(results)
    df.to_excel('Result_sheet.xlsx', index=False, header=False)

    for (i, table) in enumerate(tables):
        lst = ["Merit Position", pos[i]]

        table[4].append(lst)

        pdfgen(table[0], table[1], table[2], table[3], table[4])

    gpadf = df.iloc[:, -2]
    gpa_lst = gpadf.to_list()[1:]
    fail_student = sum([1 if i == '0.00' else 0 for i in gpa_lst])
    pass_student = std_no - fail_student
    percentage_of_pass = round(pass_student*100/std_no, 2)
    first_place_index = pos.index(1)
    first_place = names.loc[first_place_index].Name
    first_place_gpa = results[first_place_index+1][-2]
    gpa5_student = sum([1 if i == '5.00' else 0 for i in gpa_lst])

    info = []
    info.append(pass_student)
    info.append(fail_student)
    info.append(percentage_of_pass)
    info.append(first_place)
    info.append(first_place_gpa)
    info.append(gpa5_student)

    return info


main()
