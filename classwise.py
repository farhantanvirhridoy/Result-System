import pandas as pd


class student():
    def __init__(self,name,roll,class_,school,section,exam,subjects_with_mark,optional=None):
        self.name = name
        self.roll = roll
        self.class_ = class_
        self.school = school
        self.exam = exam
        self.optional = optional
        self.subject_with_mark = subjects_with_mark

    def __str__(self):
        return self.name

    
input = pd.read_excel('input.xlsx')
print(list(input.columns))
