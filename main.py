import pandas as pd
import math
from fpdf import FPDF
import os

if(os.path.exists("Mark_sheet")): pass
else: os.mkdir("Mark_sheet")



def gpmaker(a, b=0, c=0, d=0, total=100):
    s = a + b + c + d;
    if (s>=0.8*total): gp = 5.00
    elif (s>=0.7*total): gp = 4.00
    elif (s>=0.6*total): gp = 3.50
    elif (s>=0.5*total): gp = 3.00
    elif (s>=0.4*total): gp = 2.00
    elif (s>=0.33*total): gp = 1.00
    else: gp = 0
    return gp


def grademaker(gp):
    if (gp==5.00): return 'A+'
    elif (gp>=4.00): return 'A'
    elif (gp>=3.50): return 'A-'
    elif (gp>=3.00): return 'B'
    elif (gp>=2.00): return 'C'
    elif (gp>=1.00): return 'D'
    else: return 'F'








names = pd.read_excel('data/names.xlsx')
bangla1 = pd.read_excel('data/bangla1.xlsx')
bangla2 = pd.read_excel('data/bangla2.xlsx')
english1 = pd.read_excel('data/english1.xlsx')
english2 = pd.read_excel('data/english2.xlsx')
bgs = pd.read_excel('data/bgs.xlsx')
islam = pd.read_excel('data/islam.xlsx')
gmath = pd.read_excel('data/gmath.xlsx')
science = pd.read_excel('data/science.xlsx')
homeeconomics = pd.read_excel('data/homeeconomics.xlsx')
ict = pd.read_excel('data/ict.xlsx')

std_no = names.count()[0]

subjects = ["Bangla", "English", "Bangladesh and Global Studies","Islam","General Math","Science","Home Economics", "ICT"]

results = [["Roll","Name","Bangla", "English", "BGS","Islam","General Math","Science","Home Economics", "ICT","Total","GPA","Position"]]

gpas = dict()


for roll in range(0,std_no,1):   
    name = names.loc[roll].Name
    bangla1_mark = bangla1.loc[roll]
    bangla2_mark = bangla2.loc[roll]
    english1_mark = english1.loc[roll]
    english2_mark = english2.loc[roll]
    bgs_mark = bgs.loc[roll]
    islam_mark = islam.loc[roll]
    gmath_mark = gmath.loc[roll]
    science_mark = science.loc[roll]
    homeeconomics_mark = homeeconomics.loc[roll]
    ict_mark = ict.loc[roll]

    bangla_gp = gpmaker(bangla1_mark.CQ,bangla1_mark.MCQ,bangla2_mark.CQ,total=150)
    english_gp = gpmaker(english1_mark.CQ,english2_mark.CQ,total=150)
    bgs_gp = gpmaker(bgs_mark.CQ,bgs_mark.MCQ)
    islam_gp = gpmaker(islam_mark.CQ,islam_mark.MCQ)
    gmath_gp = gpmaker(gmath_mark.CQ,gmath_mark.MCQ)
    science_gp = gpmaker(science_mark.CQ, science_mark.MCQ)
    homeeconomics_gp = gpmaker(homeeconomics_mark.CQ,homeeconomics_mark.MCQ)
    ict_gp = gpmaker(ict_mark.CQ,total=25)
    gp = [bangla_gp, english_gp, bgs_gp, islam_gp, gmath_gp, science_gp, homeeconomics_gp, ict_gp]
    
    if 0.0 in gp: gpa = 0
    else: gpa = round(sum(gp)/len(gp),2)

    gpas[roll+1] = gpa
    
    pdf = FPDF()
    pdf.add_page()

    pdf.set_font("Times", size = 18)
    pdf.cell(190, 10, txt = "Burimari Hasar Uddin High School",ln = 1, align = 'C')

    pdf.set_font("Times", size = 14) 
    pdf.cell(190, 10, txt = "Annual Exam Mark Sheet",ln = 1, align = 'C')

    table1 = [["Name: "+str(name).title()+"\nRoll: "+str(roll+1)+"\nSection: A"]]

    pdf.set_font("Times", size = 12)
    cell_width = (pdf.w - pdf.l_margin - pdf.r_margin)/len(table1[0])
    cell_height = 10
    top = pdf.y + 10
    left = pdf.x

    for i in table1:
        counter = 0
        for j in i:
            
            pdf.x = left + counter*cell_width
            
            
            pdf.y = top
            pdf.multi_cell(cell_width,cell_height,txt = j,  border=1)
            counter = counter + 1


    marks = [sum([bangla1_mark.CQ,bangla1_mark.MCQ,bangla2_mark.CQ])]
    marks.append(sum([english1_mark.CQ,english2_mark.CQ]))
    marks.append(sum([bgs_mark.CQ,bgs_mark.MCQ]))
    marks.append(sum([islam_mark.CQ,islam_mark.MCQ]))
    marks.append(sum([gmath_mark.CQ,gmath_mark.MCQ]))
    marks.append(sum([science_mark.CQ,science_mark.MCQ]))
    marks.append(sum([homeeconomics_mark.CQ,homeeconomics_mark.MCQ]))
    marks.append(sum([ict_mark.CQ]))

    gps = [bangla_gp,english_gp,bgs_gp,islam_gp,gmath_gp,science_gp,homeeconomics_gp,ict_gp]
    

    table2 = [["S.N.","Subjects","Marks","Grade"]]
    for i in range(1,9,1):
        temp = []
    
        temp.append(i)
        temp.append(subjects[i-1])
        temp.append(marks[i-1])
        temp.append(grademaker(gps[i-1]))
        table2.append(temp)
    
    

    pdf.set_font("Times", size = 11)
    cell_width = [10,120,30,30]
    cell_height = 10
    top = pdf.y
    left = pdf.x

    for i in table2:
        counter = 0
        top = top + 10
        left = pdf.x
        for j,k in zip(i,cell_width):
        
        
            pdf.x = left
        
            left = left + k
            pdf.y = top
            pdf.multi_cell(k,cell_height,txt = str(j),  border=1)
            counter = counter + 1



    table4 = [["Total Marks",sum(marks)],["GPA",gpa],["Grade",grademaker(gpa)]]

    pdf.set_font("Times", size = 12)
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
            
            pdf.multi_cell(cell_width,cell_height,txt = str(j),  border=1, align='C')
            counter = counter + 1

    table3 = [["","--------------------------------"],["","Signature of Head Teacher\n(Burimari Hasar Uddin High School)"]]

    pdf.set_font("Times", size = 12)
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
            pdf.multi_cell(cell_width,cell_height,txt = str(j),  border=0, align='C')
            counter = counter + 1
    pdf.output(str("Mark_sheet/mark_sheet-"+str(roll+1)+".pdf"),dest = 'F')

    result = [roll+1,name.title(),bangla_gp, english_gp, bgs_gp, islam_gp, gmath_gp, science_gp, homeeconomics_gp, ict_gp,sum(marks),gpa]
    results.append(result)


lst = sorted(gpas.items(), key=lambda kv:
                 kv[1],reverse=True)
merit = dict()
count = 1
for l in lst:
    merit[l[0]] = count
    count += 1

lst = sorted(merit.items(), key = lambda kv: kv[0])
pos = []
for l in lst:
    pos.append(l[1])



for i in range(len(results)-1):
    
    results[i+1].append(pos[i])
    

df = pd.DataFrame(results)
df.to_excel('Result_sheet.xlsx', index=False, header=False)

    
        

