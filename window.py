import ttkbootstrap as ttkb
import tkinter as tk
from tkinter import *
from tkinter import ttk
import pandas as pd
import math
from fpdf import FPDF
import os


# main function
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
            if(class_ == '6' or class_ == '7' or class_ == '8'):
                optional_sub = subjects8bycode[str(optional_code)]
            elif(class_ == '9' or class_ == '10'):
                optional_sub = subjects8bycode[str(optional_code)]
        except:
            optional_sub = None

        if (optional_sub != None):
            number_of_subjects = len(subjects) - 1
        else:
            number_of_subjects = len(subjects)

        if (optional_sub != None):

            if 0.0 in gp:
                if (gp.index(0.0) == subjects.index(optional_sub)):
                    gpa = round(sum(gp)/number_of_subjects, 2)
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


# create excel function
def create_excel():

    if (os.path.exists("Data")):
        pass
    else:
        os.mkdir("Data")

    settings = dict()
    setting_file = open('setting.txt', 'r').readlines()

    for setting in setting_file:
        [key, value] = setting.rstrip().split('=')
        settings[key] = value

    std_no = settings['Number of Students']
    subjects = settings['Subjects'].split(",")
    excel = settings['Excel']
    excels = []
    for row in excel.split("/"):
        column = []
        for col in row.split(","):
            column.append(col)
        excels.append(column)

    for i, subject in enumerate(subjects):
        filename = subject.replace(" ", "_").lower() + '.xlsx'
        rolls = []

        # df = pd.DataFrame([["Roll"] + excels[i]]).append(pd.DataFrame([i] for i in range(1,int(std_no)+1,1)))
        df1 = pd.DataFrame([["Roll"] + excels[i]])
        df2 = pd.DataFrame([i] for i in range(1, int(std_no)+1, 1))
        df = pd.concat([df1, df2])
        if (os.path.exists('Data/'+filename)):
            pass
        else:
            df.to_excel('Data/'+filename, index=False, header=False)
    if (os.path.exists('Data/'+'names.xlsx')):
        pass
    else:
        df1 = pd.DataFrame([["Roll", "Name", "Optional"]])
        df2 = pd.DataFrame([i] for i in range(1, int(std_no)+1, 1))
        df = pd.concat([df1, df2])
        df.to_excel('Data/'+'names.xlsx', index=False, header=False)


# Entry status function
def status():
    ldict = dict()
    settings = dict()
    setting_file = open('setting.txt', 'r').readlines()
    for setting in setting_file:
        [key, value] = setting.rstrip().split('=')
        settings[key] = value

    response = []

    subject_file_names = settings['Subjects'].split(",")
    for subject_file_name in subject_file_names:
        filename = subject_file_name.replace(" ", "_").lower()
        exec(f"{filename} = pd.read_excel('data/{filename}.xlsx')")
        exec(f"count_nan = {filename}.isna().sum().sum()", locals(), ldict)
        count_nan = ldict['count_nan']
        if (count_nan != 0):

            response.append(str(count_nan) +
                            " data missing in " + subject_file_name)
        else:
            response.append("All data is present in " + subject_file_name)

    names = pd.read_excel('Data/names.xlsx')
    count_nan = names.isna().sum().sum()
    if (count_nan != 0):
        response.append(str(count_nan) + " data missing in " + 'Names')
    else:
        response.append("All data is present in " + 'Names')
    return response


root = Tk()
root.title("Digital Result System")
root.geometry('1000x600')
# root.resizable(width=FALSE, height=FALSE)

class_option_var = StringVar()
class_option_var.set('Select class')

subjects8bycode = dict()
subjects8byname = dict()
file8 = open('subject_list_class8.txt', 'r')
for file in file8:
    [key, value] = file.rstrip().split('=')
    subjects8bycode[key] = value
    subjects8byname[value] = key


subjects10bycode = dict()
subjects10byname = dict()
file10 = open('subject_list_class10.txt', 'r')
for file in file10:
    [key, value] = file.rstrip().split('=')
    subjects10bycode[key] = value
    subjects10byname[value] = key


def calculation():
    info = main()


def setup():
    for widget in content_frame.winfo_children():
        widget.destroy()
    setupframe = Frame(content_frame, bg='skyblue')
    setupframe.pack(fill=BOTH, expand=True)

    def save_config():
        file = open('setting.txt', 'w')
        school_name = school_entry.get()
        file.writelines(f"School={school_name}")
        class_name = class_option_var.get()
        file.writelines(f"\nClass={class_name}")
        section = section_entry.get()
        file.writelines(f"\nSection={section}")
        exam = exam_name_entry.get()
        file.writelines(f"\nExam={exam}")
        std_no = student_entry.get()
        file.writelines(f"\nNumber of Students={std_no}")
        file.writelines("\nSubjects="+",".join(subject_lst))
        file.writelines("\nExcel=")
        for i, entry in enumerate(mark_input_ref):
            if (i+1 == len(mark_input_ref)):
                file.writelines(entry.get())
            else:
                file.writelines(entry.get()+'/')
        file.writelines("\nTotals=")
        for i, entry in enumerate(total_mark_ref):
            if (i+1 == len(total_mark_ref)):
                file.writelines(entry.get())
            else:
                file.writelines(entry.get()+',')
        file.close()

        create_excel()

    exam_frame = LabelFrame(setupframe, text='Exam information')
    exam_frame.pack(padx=5, pady=5, anchor=N, side=TOP)

    subject_frame = LabelFrame(setupframe, text='Subject information')
    subject_frame.pack(padx=20, pady=5, side=LEFT, anchor=N)

    sub_detail_frame = LabelFrame(setupframe, text="Subject Details")

    school_label = Label(exam_frame, text='Name of School: ')
    school_label.grid(row=0, column=0)
    school_entry = Entry(exam_frame, width=60)
    school_entry.grid(row=0, column=1, columnspan=3)

    exam_name_label = Label(exam_frame, text='Exam Name: ')
    exam_name_label.grid(row=1, column=0)
    exam_name_entry = Entry(exam_frame)
    exam_name_entry.grid(row=1, column=1)

    class_label = Label(exam_frame, text='Class: ')
    class_label.grid(row=1, column=2)
    class_entry = Entry(exam_frame)
    options = ['6', '7','8','9','10']
    class_option_var.set('6')
    class_entry = OptionMenu(exam_frame, class_option_var, *options)
    class_entry.grid(row=1, column=3)

    section_label = Label(exam_frame, text='Section: ')
    section_label.grid(row=2, column=0)
    section_entry = Entry(exam_frame)
    section_entry.grid(row=2, column=1)

    student_label = Label(exam_frame, text='Number of students: ')
    student_label.grid(row=2, column=2)
    student_entry = Entry(exam_frame)
    student_entry.grid(row=2, column=3)

    subject_lst = []
    mark_input_ref = []
    total_mark_ref = []

    def add_sub():
        global add_sub_btn
        global subject_name
        add_sub_btn.config(state=DISABLED)
        subject_name.config(state=DISABLED)

        sub_detail_frame.pack(anchor=N, side=LEFT, padx=20, pady=5)
        Label(sub_detail_frame, text='Code').grid(row=0, column=0)
        Label(sub_detail_frame, text='Subject', width=23).grid(row=0, column=1)
        Label(sub_detail_frame, text='Column for Mark Input').grid(
            row=0, column=2)
        Label(sub_detail_frame, text='Total Marks').grid(
            row=0, column=3, padx=2)

        subjects = subject_name.selection_get()
        for i, subject in enumerate(subjects.split('\n')):
            Label(sub_detail_frame, text=subject[:3]).grid(row=i+1, column=0)
            Label(sub_detail_frame, text=subject[4:]).grid(row=i+1, column=1)
            mir = Entry(sub_detail_frame, width=24)
            mir.insert(0, "cq,mcq")
            mir.grid(row=i+1, column=2)
            tmr = Entry(sub_detail_frame, width=10)
            tmr.insert(0, 100)
            tmr.grid(row=i+1, column=3)

            subject_lst.append(subject[4:])
            mark_input_ref.append(mir)
            total_mark_ref.append(tmr)

        save_btn = Button(sub_detail_frame, text='Save Setup',
                          command=save_config)
        save_btn.grid(row=i+2, column=0, columnspan=4)

    def next_cmd():
        next_btn.config(state=DISABLED)
        global add_sub_btn
        global subject_name
        subject_name = Listbox(
            subject_frame, selectmode=MULTIPLE, width=60, height=15)
        subject_name.pack()

        if (class_option_var.get() == '6' or class_option_var.get() == '7' or class_option_var.get() == '8'):
            x = list(subjects8bycode.values())
            y = list(subjects8bycode.keys())
        elif (class_entry.get() == '9' or class_entry.get() == '10'):
            x = list(subjects10bycode.values())
            y = list(subjects10bycode.keys())

        # x = ["Bangla", "English", "Bangladesh and Global Studies",
        #    "Islam", "General Math", "Science", "Home Economics", "ICT"]

        for i in range(len(x)):

            subject_name.insert(END, y[i] + ' ' + x[i])
            subject_name.itemconfig(i, bg="lime")

        add_sub_btn = Button(
            subject_frame, text='Add Subjects', command=add_sub)
        add_sub_btn.pack(pady=5)

    next_btn = Button(exam_frame, text='Next', command=next_cmd)
    next_btn.grid(row=3, columnspan=4, pady=5)


def entry():
    for widget in content_frame.winfo_children():
        widget.destroy()
    entryframe = Frame(content_frame, bg='skyblue')
    entryframe.pack(fill=BOTH, expand=True)
    for st in status():
        Label(entryframe, text=st, font=('Times', 12),
              bg='white').pack(anchor=W, padx=10, pady=10)


def calc():
    for widget in content_frame.winfo_children():
        widget.destroy()
    calcframe = Frame(content_frame, bg='skyblue')
    calcframe.pack(fill=BOTH, expand=True)

    def calculation():
        Label(calcframe, text='Result Summary',
              font=('Times', 20)).pack(pady=10)
        info = main()
        text = 'Number of A+: ' + str(info[5])
        text = text + '\nNumber of student passed: ' + str(info[0])
        text = text + '\nNumber of student failed: ' + str(info[1])
        text = text + '\nPercentage of pass: ' + str(info[2]) + '%'

        Label(calcframe, text=text, justify=LEFT).pack()
        Label(calcframe, text='First Position',
              font=('Times', 20)).pack(pady=10)
        text = 'Name: ' + info[3].title()
        text += '\nGPA: ' + str(info[4])
        Label(calcframe, text=text, justify=LEFT).pack()

    Button(calcframe, text='Start Calculation', font=(
        'Times', 12), command=calculation).pack(pady=10)


def check():
    for widget in content_frame.winfo_children():
        widget.destroy()
    checkframe = Frame(content_frame, bg='skyblue')
    checkframe.pack(fill=BOTH, expand=True)

    settings = dict()
    try:
        setting_file = open('setting.txt', 'r').readlines()
    except:
        Label(checkframe, text='No previous setting file found',
              font=('Times', 14)).pack()
        return
    for setting in setting_file:
        [key, value] = setting.rstrip().split('=')
        settings[key] = value

    totals = settings['Totals'].split(',')
    exam = settings['Exam']
    subjects = settings['Subjects'].split(",")
    school = settings['School']
    class_ = settings['Class']
    number_of_students = settings['Number of Students']

    section = settings['Section']
    excel = settings['Excel']
    excels = []
    for row in excel.split("/"):

        excels.append(row)

    exam_frame = LabelFrame(checkframe, text='Exam information')
    exam_frame.pack(padx=5, pady=5, anchor=N, side=TOP)

    sub_detail_frame = LabelFrame(checkframe, text="Subject Details")

    school_label = Label(exam_frame, text='Name of School: ')
    school_label.grid(row=0, column=0)
    school_entry = Entry(exam_frame, width=60)
    school_entry.insert(0, school)
    school_entry.grid(row=0, column=1, columnspan=3)

    exam_name_label = Label(exam_frame, text='Exam Name: ')
    exam_name_label.grid(row=1, column=0)
    exam_name_entry = Entry(exam_frame)
    exam_name_entry.insert(0, exam)
    exam_name_entry.grid(row=1, column=1)

    class_label = Label(exam_frame, text='Class: ')
    class_label.grid(row=1, column=2)
    class_entry = Entry(exam_frame)
    class_entry.insert(0, class_)
    class_entry.grid(row=1, column=3)

    section_label = Label(exam_frame, text='Section: ')
    section_label.grid(row=2, column=0)
    section_entry = Entry(exam_frame)
    section_entry.insert(0, section)
    section_entry.grid(row=2, column=1)

    student_label = Label(exam_frame, text='Number of students: ')
    student_label.grid(row=2, column=2)
    student_entry = Entry(exam_frame)
    student_entry.insert(0, number_of_students)
    student_entry.grid(row=2, column=3)

    sub_detail_frame.pack(side=TOP, anchor=N, padx=5, pady=20)
    Label(sub_detail_frame, text='Subject Code').grid(row=0, column=0)
    Label(sub_detail_frame, text='Subject', width=23).grid(row=0, column=1)
    Label(sub_detail_frame, text='Column for Mark Input').grid(row=0, column=2)
    Label(sub_detail_frame, text='Total Marks').grid(row=0, column=3, padx=2)

    for i, subject in enumerate(subjects):
        Label(sub_detail_frame, text=subjects8byname[subject]).grid(
            row=i+1, column=0)
        Label(sub_detail_frame, text=subject).grid(row=i+1, column=1)
        mir = Entry(sub_detail_frame, width=24)
        mir.insert(0, excels[i])
        mir.grid(row=i+1, column=2)
        tmr = Entry(sub_detail_frame, width=10)
        tmr.insert(0, totals[i])
        tmr.grid(row=i+1, column=3)


navbar = Frame(root, bg="green", width=100)
navbar.pack(anchor=W, fill=Y, expand=False, side=LEFT)

content_frame = Frame(root, bg="skyblue")
content_frame.pack(anchor=N, fill=BOTH, expand=True, side=LEFT)
Label(content_frame, text="Welcome to digital Result system",
      font=('Times', 18)).pack()


check_btn = Button(navbar, text="Check", width=10, command=check)
check_btn.grid(row=0, column=0)

setup_btn = Button(navbar, text="Setup", width=10, command=setup)
setup_btn.grid(row=1, column=0)

entry_btn = Button(navbar, text="Mark Status", width=10, command=entry)
entry_btn.grid(row=2, column=0)

calc_btn = Button(navbar, text="Calculation", width=10, command=calc)
calc_btn.grid(row=3, column=0)

root.mainloop()
