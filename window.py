import ttkbootstrap as ttkb
import tkinter as tk
from tkinter import *
from tkinter import ttk
from entry_status import status
from main import main

root = Tk()
root.title("Digital Result System")
root.geometry('600x700')



def calculation():
    info = main()
    



def setup():
    for widget in content_frame.winfo_children():
        widget.destroy()
    setupframe = Frame(content_frame, bg='skyblue')
    setupframe.pack(fill=BOTH, expand=True)

    def save_config():
        file = open('set.txt','w')
        school_name = school_entry.get()
        file.writelines(f"School={school_name}")
        class_name = class_entry.get()
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
            if (i+1==len(mark_input_ref)):
                file.writelines(entry.get())
            else:
                file.writelines(entry.get()+'/')
        file.writelines("\nTotals=")
        for i, entry in enumerate(total_mark_ref):
            if (i+1==len(total_mark_ref)):
                file.writelines(entry.get())
            else:
                file.writelines(entry.get()+',')
        

    exam_frame = LabelFrame(setupframe, text='Exam information')
    exam_frame.pack(padx=5, pady=5, anchor=W)

    subject_frame = LabelFrame(setupframe, text='Subject information')
    subject_frame.pack(padx=5, pady=5, anchor=W)

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
        
        sub_detail_frame.pack(anchor=W,padx=5, pady=5)
        Label(sub_detail_frame, text='Subject Code').grid(row=0,column=0)
        Label(sub_detail_frame, text='Subject',width=23).grid(row=0, column=1)
        Label(sub_detail_frame, text='Column for Mark Input').grid(row=0, column=2)
        Label(sub_detail_frame, text='Total Marks').grid(row=0, column=3,padx=2)
        
        subjects = subject_name.selection_get()
        for i, subject in enumerate(subjects.split('\n')):
            Label(sub_detail_frame,text=154).grid(row=i+1,column=0)
            Label(sub_detail_frame,text=subject).grid(row=i+1,column=1)
            mir = Entry(sub_detail_frame,width=24)
            mir.grid(row=i+1, column=2)
            tmr = Entry(sub_detail_frame,width=10)
            tmr.grid(row=i+1, column=3)
            
            

            subject_lst.append(subject)
            mark_input_ref.append(mir)
            total_mark_ref.append(tmr)
            
            
        
        save_btn = Button(sub_detail_frame,text='Save Setup', padx=10, pady=10, command=save_config)
        save_btn.grid(row=i+2, column=0, columnspan=4)



    def next_cmd():
        global add_sub_btn
        global subject_name
        subject_name = Listbox(subject_frame, selectmode=MULTIPLE, width=76)
        subject_name.pack()

        

        x = ["Bangla", "English", "Bangladesh and Global Studies",
            "Islam", "General Math", "Science", "Home Economics", "ICT"]

        for each_item in range(len(x)):

            subject_name.insert(END, x[each_item])
            subject_name.itemconfig(each_item, bg="lime")

        add_sub_btn = Button(subject_frame, text='Add Subjects', command=add_sub)
        add_sub_btn.pack(pady=5)

    Button(exam_frame, text='Next',command=next_cmd).grid(row=3, columnspan=4, pady=5)

    
    
    
    

    

def entry():
    for widget in content_frame.winfo_children():
        widget.destroy()
    entryframe = Frame(content_frame, bg='white')
    entryframe.pack(fill=BOTH, expand=True)
    for st in status():
        Label(entryframe,text= st, font=('Times',12), bg = 'white').pack(anchor=W,padx=10,pady=10)


def calc():
    for widget in content_frame.winfo_children():
        widget.destroy()
    calcframe = Frame(content_frame, bg='blue')
    calcframe.pack(fill=BOTH, expand=True)

    def calculation():
        info = main()
        Label(calcframe,text=info).pack()
    Button(calcframe, text='Start Calculation',font=('Times',12), command=calculation).pack(pady=10)
    

def check():
    for widget in content_frame.winfo_children():
        widget.destroy()
    checkframe = Frame(content_frame, bg='blue')
    checkframe.pack(fill=BOTH, expand=True)
    
    settings = dict()
    setting_file = open('setting.config', 'r').readlines()
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
    exam_frame.pack(padx=5, pady=5, anchor=W)

    subject_frame = LabelFrame(checkframe, text='Subject information')
    subject_frame.pack(padx=5, pady=5, anchor=W)

    sub_detail_frame = LabelFrame(checkframe, text="Subject Details")

    school_label = Label(exam_frame, text='Name of School: ')
    school_label.grid(row=0, column=0)
    school_entry = Entry(exam_frame, width=60)
    school_entry.insert(0,school)
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

    sub_detail_frame.pack(anchor=W,padx=5, pady=5)
    Label(sub_detail_frame, text='Subject Code').grid(row=0,column=0)
    Label(sub_detail_frame, text='Subject',width=23).grid(row=0, column=1)
    Label(sub_detail_frame, text='Column for Mark Input').grid(row=0, column=2)
    Label(sub_detail_frame, text='Total Marks').grid(row=0, column=3,padx=2)
    

    for i,subject in enumerate(subjects):
        Label(sub_detail_frame,text=154).grid(row=i+1,column=0)
        Label(sub_detail_frame,text=subject).grid(row=i+1,column=1)
        mir = Entry(sub_detail_frame,width=24)
        mir.insert(0, excels[i])
        mir.grid(row=i+1, column=2)
        tmr = Entry(sub_detail_frame,width=10)
        tmr.insert(0, totals[i])
        tmr.grid(row=i+1, column=3)
        


navbar = Frame(root, bg="green", width=100)
navbar.pack(anchor=W, fill=Y, expand=False, side=LEFT)

content_frame = Frame(root, bg="orange")
content_frame.pack(anchor=N, fill=BOTH, expand=True, side=LEFT)


check_btn = Button(navbar, text="Check", width=10, command=check)
check_btn.grid(row=0, column=0)

setup_btn = Button(navbar, text="Setup", width=10, command=setup)
setup_btn.grid(row=1, column=0)

entry_btn = Button(navbar, text="Mark Status", width=10, command=entry)
entry_btn.grid(row=2, column=0)

calc_btn = Button(navbar, text="Calculation", width=10, command=calc)
calc_btn.grid(row=3, column=0)

root.mainloop()
