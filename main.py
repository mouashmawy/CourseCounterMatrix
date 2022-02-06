#Some constants for your code...
first_column = 5
first_row = 2
Input_file_name = "inputExample.xlsx"
Output_file_name = "output1.xlsx"
############################################
from openpyxl import *
from openpyxl.styles import *

#importing current file that contains data
wb = load_workbook(Input_file_name)
s = wb.active
#creating new file for saving info
wb2 = Workbook()
outSheet = wb2["Sheet"]
SomeDataSheet = wb2.create_sheet("SomeDataSheet", 0)
SomeDataSheet.title = "SomeDataSheet"

#Reading coyrses names from main file
courseNamesList = []
for course in range(5,s.max_column+1):
    courseNamesList.append(s.cell(row=1,column=course).value)


#writing courses names to Matrix
outSheet['A1'].value ="MATRIX"
for rowandcolumn in range(len(courseNamesList)):
    outSheet.cell(row=rowandcolumn+2,column=1).value = courseNamesList[rowandcolumn]
    outSheet.cell(row=1, column=rowandcolumn + 2).value = courseNamesList[rowandcolumn]


#Reading data from main sheet and putting them in a List
AllStudentsList = []
for student in range(2,s.max_row+1):
    CoursesOfStudent = []
    for course in range(5,s.max_column+1):
        cellvalue =s.cell(row=student,column=course).value
        x = 1 if (cellvalue==1) else 0
        CoursesOfStudent.append(x)
    AllStudentsList.append(CoursesOfStudent)


#Calculating the matrix putting it in another list
AllCourseMatrix = []
for i in range(len(AllStudentsList[0])):
    OneCourseMatrix = []
    for j in range(len(AllStudentsList[0])):
        matchesCounter=0
        for k in range(len(AllStudentsList)):
            if(AllStudentsList[k][i] and AllStudentsList[k][j]):
                matchesCounter+=1
        OneCourseMatrix.append(matchesCounter)
    AllCourseMatrix.append(OneCourseMatrix)



#Assigning values from Matrix List to Matrix Sheet
for i in range(len(AllCourseMatrix)):
    for j in range(len(AllCourseMatrix[i])):
        outSheet.cell(row=i+2,column=j+2).value =  AllCourseMatrix[i][j]


#adding number of registered students in every course
SomeDataSheet['A1'].value ="Course"
SomeDataSheet['B1'].value ="number registered"
for i in range(2,outSheet.max_row+1):
    SomeDataSheet.cell(row=i, column=1).value = outSheet.cell(row=1, column=i).value
    SomeDataSheet.cell(row=i,column=2).value = outSheet.cell(row=i,column=i).value


#Changing styles for some cells
for i in range(2,outSheet.max_row+1):
    outSheet.cell(row=i,column=i).fill = PatternFill(start_color='FFFF00',fill_type='solid')
for i in range(1,outSheet.max_row+1):
    outSheet.cell(row=i, column=1).font = Font(bold=True)
    outSheet.cell(row=1, column=i).font = Font(bold=True)


#saving the file
wb2._sheets = [wb2._sheets[1],wb2._sheets[0]]
wb2.save(Output_file_name)
