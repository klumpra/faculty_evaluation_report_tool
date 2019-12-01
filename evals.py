'''
This program produces faculty / course evaluation reports. 
It creates section-by-section reports in a separate folder for each program.
It also produces summary reports that list course evaluation results on
a course-by-course basis, which is useful for comparing sections and 
identifying potential problems.
'''

import xlrd
import os

semester_prefix = "SP19"  ## change if future semesters.

'''enrollment.txt is a tab-delited file that Sarah Pariseau sent me somewhere
in mid to late November 2019 that shows "Course Enrollments By Subject". I
believe it originally comes from a Cognos report of the same name. It has these
fields in this order:
    CRN	SUBJ	COURSE	SECTION	COURSE TITLE	COURSE	PART OF TERM	INSTRUCTOR	CAMPUS	CAPACITY	ACTUAL	REM	BUILDING	ROOM	BEGIN_TIME	END_TIME	M	T	W	R	F	S	COLLEGE	COLLEGE_DESC	DEPARTMENT	DEPARTMENT_DESC	STATUS	SCHEDULE	SCHEDULE_DESC	INSTRUCTION_METHOD	INSTRUCTION_METHOD_DESC
with the first line stating "Course Enrollments by Subject" and the second line
containing these column headings.
'''
course_fname =  "C:\\Users\\klumpra\\Dropbox\\coast\\ray_stuff\\faculty_evaluations\\enrollment.txt"

'''scores.xlsx is the file that was downloaded from TK20 that shows the
evaluation results. This was sent to me by Sarah Pariseau who downloaded it from TK20.
The first 12 rows of the spreadsheet look like this:
    Spring 2019 CoAST course evaluations																																																																							
																																																																							
Form Instance: 	CAS New Traditional End of Course Evaluation FA18																																																																						
Sent Date: 																																																																							
Sent Term: 	Spring 2019																																																																						
Saved/Submitted Date: 																																																																							
Assessee (Last Name, First Name, PID): 																																																																							
Course(s): 	AVMT10601 | 001 | Spring 2019,AVMT11001 | 001 | Spring 2019,AVMT12001 | 001 | Spring 2019,AVMT13001 | 002 | Spring 2019,AVMT20000 | 001 | Spring 2019,AVMT20001 | 001 | Spring 2019,AVMT20001 | 002 | Spring 2019,AVMT20001 | 003 | Spring 2019,AVMT20100 | 001 | Spring 2019,AVMT20100 | 003 | Spring 2019,AVMT20200 | 001 | Spring 2019,AVMT21000 | 001 | Spring 2019,AVMT21001 | 001 | Spring 2019,AVMT21001 | 002 | Spring 2019,AVMT21001 | 003 | Spring 2019,AVMT22000 | 001 | Spring 2019,AVMT22001 | 001 | Spring 2019,AVMT22001 | 002 | Spring 2019,AVMT22001 | 004 | Spring 2019,AVMT32001 | 001 | Spring 2019,AVMT33000 | 002 | Spring 2019,AVMT33001 | 002 | Spring 2019,AVMT34000 | 001 | Spring 2019,AVMT34001 | 001 | Spring 2019,AVMT41000 | 001 | Spring 2019,AVMT41001 | 001 | Spring 2019,AVMT41001 | 003 | Spring 2019,AVMT42000 | 001 | Spring 2019,AVMT42001 | 001 | Spring 2019,AVMT42001 | 002 | Spring 2019,AVMT42500 | 001 | Spring 2019,AVMT42700 | 001 | Spring 2019,AVMT43000 | 001 | Spring 2019,AVMT43001 | 001 | Spring 2019,AVMT43001 | 002 | Spring 2019,AVMT46000 | 002 | Spring 2019,AVMT49500 | 001 | Spring 2019,AVTR10000 | 001 | Spring 2019,AVTR10000 | 002 | Spring 2019,AVTR10000 | 003 | Spring 2019,AVTR10000 | 004 | Spring 2019,AVTR10000 | 005 | Spring 2019,AVTR10000 | 006 | Spring 2019,AVTR10000 | 007 | Spring 2019,AVTR10000 | 008 | Spring 2019,AVTR10200 | 002 | Spring 2019,AVTR10200 | 003 | Spring 2019,AVTR10200 | 004 | Spring 2019,AVTR10200 | 005 | Spring 2019,AVTR12000 | 001 | Spring 2019,AVTR13000 | 001 | Spring 2019,AVTR13000 | 002 | Spring 2019,AVTR13100 | 001 | Spring 2019,AVTR13100 | 002 | Spring 2019,AVTR13100 | 003 | Spring 2019,AVTR13100 | 004 | Spring 2019,AVTR13100 | 005 | Spring 2019,AVTR21400 | 001 | Spring 2019,AVTR23100 | 001 | Spring 2019,AVTR23100 | 002 | Spring 2019,AVTR25000 | 001 | Spring 2019,AVTR25200 | 001 | Spring 2019,AVTR25200 | 002 | Spring 2019,AVTR25200 | 003 | Spring 2019,AVTR25200 | 004 | Spring 2019,AVTR25200 | 005 | Spring 2019,AVTR25200 | 006 | Spring 2019,AVTR25200 | 007 | Spring 2019,AVTR26300 | 001 | Spring 2019,AVTR26300 | 002 | Spring 2019,AVTR31300 | 001 | Spring 2019,AVTR31300 | 002 | Spring 2019,AVTR31300 | 003 | Spring 2019,AVTR32000 | 001 | Spring 2019,AVTR32100 | 001 | Spring 2019,AVTR33100 | 001 | Spring 2019,AVTR33100 | 002 | Spring 2019,AVTR34400 | 001 | Spring 2019,AVTR35300 | 001 | Spring 2019,AVTR35300 | 002 | Spring 2019,AVTR35300 | 003 | Spring 2019,AVTR35300 | 004 | Spring 2019,AVTR35600 | 001 | Spring 2019,AVTR37300 | 001 | Spring 2019,AVTR37300 | 002 | Spring 2019,AVTR37300 | 003 | Spring 2019,AVTR39001 | 001 | Spring 2019,AVTR39704 | 001 | Spring 2019,AVTR39705 | 001 | Spring 2019,AVTR39707 | 001 | Spring 2019,AVTR40700 | 001 | Spring 2019,AVTR40800 | 001 | Spring 2019,AVTR42000 | 001 | Spring 2019,AVTR45000 | 001 | Spring 2019,AVTR45000 | 002 | Spring 2019,AVTR45200 | 001 | Spring 2019,AVTR45200 | 002 | Spring 2019,AVTR45300 | 001 | Spring 2019,AVTR45700 | 001 | Spring 2019,AVTR46000 | 001 | Spring 2019,AVTR46300 | 001 | Spring 2019,AVTR46300 | 002 | Spring 2019,AVTR47500 | 001 | Spring 2019,AVTR48000 | 001 | Spring 2019,AVTR48000 | 002 | Spring 2019,AVTR48200 | 001 | Spring 2019,AVTR48400 | 001 | Spring 2019,AVTR48500 | 001 | Spring 2019,AVTR48600 | 001 | Spring 2019,AVTR49600 | 001 | Spring 2019,AVTR49900 | 003 | Spring 2019,AVTR52000 | 001 | Spring 2019,AVTR54000 | 001 | Spring 2019,AVTR58003 | 001 | Spring 2019,AVTR59600 | 001 | Spring 2019,AVTR59700 | 001 | Spring 2019,BIOL10100 | 001 | Spring 2019,BIOL10100 | 002 | Spring 2019,BIOL10200 | 001 | Spring 2019,BIOL10200 | 002 | Spring 2019,BIOL10300 | 001 | Spring 2019,BIOL10300 | 002 | Spring 2019,BIOL10300 | 003 | Spring 2019,BIOL10300 | 004 | Spring 2019,BIOL10400 | 001 | Spring 2019,BIOL10400 | 002 | Spring 2019,BIOL10400 | 003 | Spring 2019,BIOL10400 | 004 | Spring 2019,BIOL10400 | 005 | Spring 2019,BIOL10400 | 006 | Spring 2019,BIOL10600 | 001 | Spring 2019,BIOL10600 | 002 | Spring 2019,BIOL10600 | 003 | Spring 2019,BIOL10700 | 001 | Spring 2019,BIOL10700 | 002 | Spring 2019,BIOL10800 | 001 | Spring 2019,BIOL10800 | 002 | Spring 2019,BIOL11000 | 001 | Spring 2019,BIOL11100 | 001 | Spring 2019,BIOL11500 | 001 | Spring 2019,BIOL11500 | 002 | Spring 2019,BIOL11500 | 003 | Spring 2019,BIOL11600 | 001 | Spring 2019,BIOL11600 | 002 | Spring 2019,BIOL11600 | 003 | Spring 2019,BIOL11600 | 004 | Spring 2019,BIOL11600 | 005 | Spring 2019,BIOL12200 | 001 | Spring 2019,BIOL12300 | 001 | Spring 2019,BIOL19900 | 001 | Spring 2019,BIOL22000 | 001 | Spring 2019,BIOL22100 | 001 | Spring 2019,BIOL22400 | 001 | Spring 2019,BIOL22400 | 002 | Spring 2019,BIOL22500 | 001 | Spring 2019,BIOL22600 | 001 | Spring 2019,BIOL22600 | 002 | Spring 2019,BIOL22600 | 003 | Spring 2019,BIOL22600 | 004 | Spring 2019,BIOL22600 | 005 | Spring 2019,BIOL22700 | 001 | Spring 2019,BIOL27000 | 001 | Spring 2019,BIOL27000 | 002 | Spring 2019,BIOL27000 | 003 | Spring 2019,BIOL32000 | 001 | Spring 2019,BIOL32000 | 002 | Spring 2019,BIOL33500 | 001 | Spring 2019,BIOL33600 | 001 | Spring 2019,BIOL33600 | 002 | Spring 2019,BIOL35600 | 001 | Spring 2019,BIOL35700 | 001 | Spring 2019,BIOL38500 | 001 | Spring 2019,BIOL38500 | 002 | Spring 2019,BIOL38500 | 003 | Spring 2019,BIOL39000 | 001 | Spring 2019,BIOL39500 | 001 | Spring 2019,BIOL39900 | 001 | Spring 2019,BIOL40600 | 001 | Spring 2019,BIOL40600 | 002 | Spring 2019,BIOL41600 | 001 | Spring 2019,BIOL41700 | 001 | Spring 2019,BIOL42200 | 001 | Spring 2019,BIOL42200 | 002 | Spring 2019,BIOL42300 | 001 | Spring 2019,BIOL42600 | 001 | Spring 2019,BIOL43500 | 001 | Spring 2019,BIOL49000 | 002 | Spring 2019,BIOL49000 | 003 | Spring 2019,BIOL49000 | 005 | Spring 2019,BIOL49000 | 006 | Spring 2019,BIOL49600 | 001 | Spring 2019,BIOL49600 | 002 | Spring 2019,BIOL49600 | 003 | Spring 2019,BIOL49600 | 004 | Spring 2019,BIOL49600 | 005 | Spring 2019,BIOL49600 | 006 | Spring 2019,BIOL49600 | 007 | Spring 2019,BIOL49600 | 008 | Spring 2019,CHEM10500 | 001 | Spring 2019,CHEM10601 | 001 | Spring 2019,CHEM11000 | 001 | Spring 2019,CHEM11000 | 002 | Spring 2019,CHEM11100 | 001 | Spring 2019,CHEM11100 | 002 | Spring 2019,CHEM11500 | 001 | Spring 2019,CHEM11500 | 002 | Spring 2019,CHEM11500 | 003 | Spring 2019,CHEM11500 | 004 | Spring 2019,CHEM11600 | 001 | Spring 2019,CHEM11600 | 002 | Spring 2019,CHEM11600 | 003 | Spring 2019,CHEM11600 | 004 | Spring 2019,CHEM11600 | 005 | Spring 2019,CHEM11600 | 006 | Spring 2019,CHEM11600 | 008 | Spring 2019,CHEM12200 | 001 | Spring 2019,CHEM12300 | 001 | Spring 2019,CHEM22000 | 001 | Spring 2019,CHEM22100 | 001 | Spring 2019,CHEM22500 | 001 | Spring 2019,CHEM22500 | 002 | Spring 2019,CHEM22600 | 001 | Spring 2019,CHEM22600 | 002 | Spring 2019,CHEM22600 | 003 | Spring 2019,CHEM22600 | 004 | Spring 2019,CHEM23500 | 001 | Spring 2019,CHEM23600 | 001 | Spring 2019,CHEM24200 | 001 | Spring 2019,CHEM29600 | 001 | Spring 2019,CHEM30500 | 001 | Spring 2019,CHEM30600 | 001 | Spring 2019,CHEM32000 | 001 | Spring 2019,CHEM40000 | 001 | Spring 2019,CHEM40700 | 001 | Spring 2019,CHEM40800 | 001 | Spring 2019,CHEM45000 | 001 | Spring 2019,CHEM45000 | 003 | Spring 2019,CHEM46500 | 001 | Spring 2019,CHEM50100 | 001 | Spring 2019,CHEM62300 | 001 | Spring 2019,CHEM68500 | 001 | Spring 2019,CHEM69600 | 001 | Spring 2019,CHEM69800 | 001 | Spring 2019,CHEM69800 | 002 | Spring 2019,CHEM69800 | 003 | Spring 2019,CHEM69800 | 004 | Spring 2019,CHEM69800 | 005 | Spring 2019,CPEN10000 | 001 | Spring 2019,CPEN22000 | 001 | Spring 2019,CPEN31000 | 001 | Spring 2019,CPEN32000 | 001 | Spring 2019,CPEN40000 | 001 | Spring 2019,CPEN45000 | 001 | Spring 2019,CPEN49600 | 001 | Spring 2019,CPSC19403 | 001 | Spring 2019,CPSC19403 | 002 | Spring 2019,CPSC19601 | 001 | Spring 2019,CPSC19602 | 001 | Spring 2019,CPSC19603 | 001 | Spring 2019,CPSC19604 | 001 | Spring 2019,CPSC19605 | 001 | Spring 2019,CPSC19605 | 002 | Spring 2019,CPSC19606 | 001 | Spring 2019,CPSC20000 | 001 | Spring 2019,CPSC20000 | 002 | Spring 2019,CPSC21000 | 001 | Spring 2019,CPSC21000 | 002 | Spring 2019,CPSC21000 | 003 | Spring 2019,CPSC23000 | 001 | Spring 2019,CPSC23500 | 001 | Spring 2019,CPSC24500 | 001 | Spring 2019,CPSC24500 | 002 | Spring 2019,CPSC25000 | 001 | Spring 2019,CPSC28200 | 001 | Spring 2019,CPSC31500 | 001 | Spring 2019,CPSC35000 | 001 | Spring 2019,CPSC35000 | 002 | Spring 2019,CPSC35000 | 003 | Spring 2019,CPSC35000 | 004 | Spring 2019,CPSC35000 | 005 | Spring 2019,CPSC35500 | 001 | Spring 2019,CPSC36000 | 001 | Spring 2019,CPSC41500 | 002 | Spring 2019,CPSC41700 | 001 | Spring 2019,CPSC42100 | 001 | Spring 2019,CPSC42300 | 001 | Spring 2019,CPSC42500 | 001 | Spring 2019,CPSC42700 | 001 | Spring 2019,CPSC44000 | 001 | Spring 2019,CPSC44000 | 002 | Spring 2019,CPSC46000 | 001 | Spring 2019,CPSC46000 | 002 | Spring 2019,CPSC46000 | 003 | Spring 2019,CPSC46000 | 004 | Spring 2019,CPSC47000 | 001 | Spring 2019,CPSC47200 | 001 | Spring 2019,CPSC48500 | 001 | Spring 2019,CPSC49200 | 001 | Spring 2019,CPSC49200 | 002 | Spring 2019,CPSC49300 | 001 | Spring 2019,CPSC49602 | 001 | Spring 2019,CPSC49603 | 001 | Spring 2019,CPSC49800 | 001 | Spring 2019,CPSC49900 | 004 | Spring 2019,CPSC50100 | 001 | Spring 2019,CPSC50200 | 001 | Spring 2019,CPSC50300 | 001 | Spring 2019,CPSC50600 | 001 | Spring 2019,CPSC50700 | 001 | Spring 2019,CPSC51100 | 001 | Spring 2019,CPSC51500 | 001 | Spring 2019,CPSC52000 | 001 | Spring 2019,CPSC52500 | 001 | Spring 2019,CPSC54000 | 001 | Spring 2019,CPSC55500 | 001 | Spring 2019,CPSC56000 | 001 | Spring 2019,CPSC57400 | 001 | Spring 2019,CPSC59100 | 001 | Spring 2019,CPSC59601 | 001 | Spring 2019,CPSC59700 | 001 | Spring 2019,CPSC60000 | 001 | Spring 2019,CPSC62700 | 001 | Spring 2019,CPSC67000 | 001 | Spring 2019,CPSC67300 | 001 | Spring 2019,CPSC69700 | 001 | Spring 2019,CPSC69700 | 002 | Spring 2019,MATH11500 | 001 | Spring 2019,MATH11500 | 002 | Spring 2019,MATH11500 | 003 | Spring 2019,MATH11500 | 004 | Spring 2019,MATH11600 | 003 | Spring 2019,MATH11900 | 001 | Spring 2019,MATH11900 | 002 | Spring 2019,MATH12000 | 001 | Spring 2019,MATH14100 | 001 | Spring 2019,MATH14100 | 002 | Spring 2019,MATH19002 | 001 | Spring 2019,MATH20000 | 001 | Spring 2019,MATH20100 | 001 | Spring 2019,MATH20100 | 002 | Spring 2019,MATH21100 | 001 | Spring 2019,MATH21500 | 001 | Spring 2019,MATH24000 | 001 | Spring 2019,MATH25000 | 001 | Spring 2019,MATH30000 | 001 | Spring 2019,MATH30700 | 001 | Spring 2019,MATH30700 | 002 | Spring 2019,MATH30700 | 003 | Spring 2019,MATH30700 | 004 | Spring 2019,MATH31000 | 001 | Spring 2019,MATH31000 | 002 | Spring 2019,MATH31400 | 001 | Spring 2019,MATH31600 | 001 | Spring 2019,MATH32500 | 001 | Spring 2019,MATH39603 | 001 | Spring 2019,MATH44100 | 001 | Spring 2019,MATH48000 | 001 | Spring 2019,MATH48000 | 002 | Spring 2019,MATH49900 | 003 | Spring 2019,MATH49900 | 005 | Spring 2019,MATH51100 | 001 | Spring 2019,MATH51200 | 001 | Spring 2019,PHYS10000 | 001 | Spring 2019,PHYS10601 | 001 | Spring 2019,PHYS20000 | 001 | Spring 2019,PHYS20100 | 001 | Spring 2019,PHYS20500 | 001 | Spring 2019,PHYS20500 | 002 | Spring 2019,PHYS20600 | 001 | Spring 2019,PHYS20600 | 002 | Spring 2019,PHYS20600 | 003 | Spring 2019,PHYS21000 | 001 | Spring 2019,PHYS21100 | 001 | Spring 2019,PHYS21500 | 001 | Spring 2019,PHYS21500 | 002 | Spring 2019,PHYS21600 | 001 | Spring 2019,PHYS21600 | 002 | Spring 2019,PHYS21800 | 001 | Spring 2019,PHYS21900 | 001 | Spring 2019,PHYS29600 | 001 | Spring 2019,PHYS30000 | 001 | Spring 2019,PHYS31800 | 001 | Spring 2019,PHYS34100 | 001 | Spring 2019,PHYS34200 | 001 | Spring 2019,PHYS40100 | 001 | Spring 2019,PHYS44200 | 001 | Spring 2019,PHYS46500 | 001 | Spring 2019,PHYS47000 | 001 | Spring 2019,PHYS47000 | 003 | Spring 2019,PHYS47000 | 004 | Spring 2019,PHYS47000 | 005 | Spring 2019,PHYS53000 | 001 | Spring 2019,PHYS54200 | 001 | Spring 2019																																																																						
Course Number: 																																																																							
Course Section: 																																																																							
Prepared by Sarah Pariseau on Tuesday, November 19, 2019 03:32:56 PM																																																																							
Form Name	Assessee (Last Name, First Name, PID)	Functionality	CE Sent Term	CE Start Date	CE Saved/Submitted Date	Assessee Last Name	Assessee First Name	Assessee PID	Assessor/Reviewer Last Name	Assessor/Reviewer First Name	Assessor/Reviewer PID	Form Instance	Course(s)	Course Number	Course Section	Section ID	Form Title	COURSE EVALUATION FORM	Traditional evaluation questions	1. The objectives of this course were made clear_Description	1. The objectives of this course were made clear_Rating	1. The objectives of this course were made clear_Points	2. The course was well organized on a daily basis_Description	2. The course was well organized on a daily basis_Rating	2. The course was well organized on a daily basis_Points	3. It was clearly indicated what material the graded work would cover._Description	3. It was clearly indicated what material the graded work would cover._Rating	3. It was clearly indicated what material the graded work would cover._Points	4. The teaching materials required for this course were helpful_Description	4. The teaching materials required for this course were helpful_Rating	4. The teaching materials required for this course were helpful_Points	5. The methods of evaluation (examinations, papers, projects, and class  discussion) were relevant to the course objectives_Description	5. The methods of evaluation (examinations, papers, projects, and class  discussion) were relevant to the course objectives_Rating	5. The methods of evaluation (examinations, papers, projects, and class  discussion) were relevant to the course objectives_Points	6. The course aroused my curiosity and challenged me intellectually_Description	6. The course aroused my curiosity and challenged me intellectually_Rating	6. The course aroused my curiosity and challenged me intellectually_Points	7. Grading policy was clear and consistently applied_Description	7. Grading policy was clear and consistently applied_Rating	7. Grading policy was clear and consistently applied_Points	8. I would rate this course as worthwhile_Description	8. I would rate this course as worthwhile_Rating	8. I would rate this course as worthwhile_Points	9. The instructor communicated the subject matter effectively._Description	9. The instructor communicated the subject matter effectively._Rating	9. The instructor communicated the subject matter effectively._Points	10. The instructor encouraged and was responsive to student participation_Description	10. The instructor encouraged and was responsive to student participation_Rating	10. The instructor encouraged and was responsive to student participation_Points	11. The instructor was available for individual help._Description	11. The instructor was available for individual help._Rating	11. The instructor was available for individual help._Points	12. The instructor showed an interest in and respect for me as an individual_Description	12. The instructor showed an interest in and respect for me as an individual_Rating	12. The instructor showed an interest in and respect for me as an individual_Points	13. I would recommend this instructor to other students_Description	13. I would recommend this instructor to other students_Rating	13. I would recommend this instructor to other students_Points	Rubric Score_Rating	Rubric Score_Points	Rubric Mean_Rating	Rubric Mean_Points	Comments:	Essay questions	14. List the strengths of this class/instructor. Try to be specific; use examples	15. List any improvements for this class and/or instructor. Again, try to be specific; use examples	16. Are there any other recommendations or comments regarding this class and/or instructor?	Total Score_Assigned Rating	Total Score_Earned Points	Total Mean_Assigned Rating	Total Mean_Earned Points
'''
score_fname = "C:\\Users\\klumpra\\Dropbox\\coast\\ray_stuff\\faculty_evaluations\\scores.xlsx"
dest_folder = "c:\\temp\\fac_evals"


workbook = xlrd.open_workbook(score_fname, on_demand = True)
worksheet = workbook.sheet_by_index(0)
def extract_score(row,col):
    text = worksheet.cell_value(row,col).strip()
    try:
        if text != "":
            return float(text)
        else:
            return -1
    except:
        return -1
''' Returns a list of questions. There are 13 numeric questions starting at
column 20 and then occurring every three columns. Then there are additional
questions - comments, strenths, improvements, recommendations '''
def get_questions():
    questions = []
    for i in range(20,57,3):
        text = worksheet.cell_value(11,i)
        pos = text.find("_")
        space_pos = text.find(" ")
        questions.append(text[space_pos+1:pos])
    questions.append("Comments")
    questions.append("List the strengths of this class/instructor.")
    questions.append("List any improvements for this class/instructor.")
    questions.append("Are there any other recommendations or comments about this class/instructor?")
    return questions
''' the number of responses is the maximum number of responses to any of the
questions. This is used for reporting how many responded out of the toal
enrollment '''
def get_response_count(record):
    max_resp = 0
    for i in range(13):
        if record[2*i+1] > max_resp:
            max_resp = record[2*i+1]
    return max_resp
''' returns a dictionary whose key is of the format SP19-CPSC-20000-001, for
example. Each entry in the dictionary is another dictionary with fields crn,
subj,num,sec,title,level,part_of_term,instructor,campus,capacity,
enrollment,building,room,begin_time,end_time,days,college,department,
department_description,status,schedule,schedule_description,instruction_method,
instruction_method_description,course_number.
The source of this information is enrollment.txt.'''
def get_course_sections(fname,prefix):
    prev_subj = ""
    prev_num = ""
    fvar = open(fname,"r")
    course_sections = {}
    count = 0
    for line in fvar:
        if count > 2: 
            line = line.strip()
            parts = line.split("\t")
            crn            = parts[0].strip()
            subj           = parts[1].strip()
            if subj == "":
                subj = prev_subj
            prev_subj = subj
            num            = parts[2].strip()
            if num == "":
                num = prev_num
            prev_num = num
            sec            = parts[3].strip()
            key = "%s-%s-%s-%s" % (prefix,subj,num,sec)
            title          = parts[4].strip()
            level          = parts[5].strip()   # for example, UG
            part_of_term   = parts[6].strip()   # for example, full term
            instructor     = parts[7].strip()
            campus         = parts[8].strip()   # for example, ROM
            cap_str        = parts[9].strip()
            if cap_str == "":
                capacity = 0
            else:
                capacity = int(cap_str)
            enrol_str      = parts[10].strip()
            if enrol_str == "":
                enrollment = 0
            else:
                enrollment = int(enrol_str)
            building       = parts[12].strip()
            room           = parts[13].strip()
            begin_time     = parts[14].strip()
            end_time       = parts[15].strip()
            days           = parts[16].strip()+parts[17].strip()+parts[18].strip()+parts[19].strip()+parts[20].strip()+parts[21].strip()
            college        = parts[23].strip()
            dept           = parts[24].strip()
            dept_desc      = parts[25].strip()
            status         = parts[26].strip()
            schedule       = parts[27].strip()
            sched_desc     = parts[28].strip()
            instr_method   = parts[29].strip()
            inst_meth_desc = parts[30].strip()
            course_num = "%s%s" % (subj,num)
            course_info = {"crn":crn,"subj":subj,"num":num,"sec":sec,"title":title,"level":level,"part_of_term":part_of_term,"instructor":instructor,"campus":campus,
                           "capacity":capacity,"enrollment":enrollment,"building":building,"room":room,"begin_time":begin_time,"end_time":end_time,"days":days,
                           "college":college,"department":dept,"department_description":dept_desc,"status":status,"schedule":schedule,"schedule_description":sched_desc,
                           "instruction_method":instr_method,"instruction_method_description":inst_meth_desc,"course_number":course_num}
            course_sections[key] = course_info
        count = count + 1
    fvar.close()
    return course_sections

''' returns a sorted list of courses by looking at the list of course_sections
and recording the unique ones. '''
def get_courses(course_sections):
    course_list = []
    for key in course_sections:
        course_num = course_sections[key]["course_number"]
        if course_num not in course_list:
            course_list.append(course_num)
    course_list.sort()
    return course_list

''' returns a a dictionary whose key is the name of the program. The value
is a list of courses'''
def get_courses_by_program(courses):
    courses_by_program = {}
    for course in courses:
        program = course[:4]
        if program not in courses_by_program:
            courses_by_program[program] = []
        courses_by_program[program].append(course)
    for cbp in courses_by_program:
        courses_by_program[cbp].sort()
    return courses_by_program

''' scores is a dictionary indexed by a combination of fac_id and sec_id. Its
value is a list iwth many slots. I've identified the slots by number below.'''
scores = {}

'''sections is a dictionary of dictionaries. See description for
get_course_sections above '''
sections = get_course_sections(course_fname,semester_prefix)  
'''now that you have the sections, generate a list of unique courses'''
courses = get_courses(sections)
'''now that you have list of course, break them up by program '''
courses_by_program = get_courses_by_program(courses)

''' questions is a list of the text of each question. see the get_questions 
function for a description '''
questions = get_questions()
''' each row of the spreadsheet represents one student's response to the
list of questions. WE will process these now. '''
for row in range(12,worksheet.nrows):
    fac_id = worksheet.cell_value(row,8)
    sec_id = worksheet.cell_value(row,16)
    fac_lname = worksheet.cell_value(row,6)
    fac_fname = worksheet.cell_value(row,7)
    dept = worksheet.cell_value(row,13).strip()[:4]
#    print("fac_id sec_id fac_lname fac_fname",fac_id,sec_id,fac_lname,fac_fname)
#    print("col22:", worksheet.cell_value(row,22))
    ''' q is the list of scores reported by this one student '''
    q=[]
    ''' grab the scores for the 13 questions. These start in column 22 and
    appear in every third column. '''
    for i in range(13):
        q.append(extract_score(row,22+3*i))
    ''' grab the open-ended responses comment, strengths, improvments, and
    recommendations. These round out the responses from this student for this
    course section '''
    comment = worksheet.cell_value(row,63).strip().replace("n/a","").replace("-","")
    strengths = worksheet.cell_value(row,65).strip().replace("n/a","").replace("-","")
    improvements = worksheet.cell_value(row,66).strip().replace("n/a","").replace("-","")
    recommendations = worksheet.cell_value(row,67).strip().replace("n/a","").replace("-","")
    ''' we now add the scores recorded in q and the comment, strengths, 
    improvements, and recommendations strings to the list of scores. scores
    is a dictionary that is key'ed by fac_id_sec_id so that we can get a
    section by section aggregate. The value for each key is an array with
    lots of columns. The first 26 columns are the sum of scores and the number of
    scores reported for each question. For example [0] is the total score
    for the first question, and [1] is the number of people who responded to
    the first question. This will help us determine the average score for
    each question '''
    key = "%s_%s" % (fac_id,sec_id)
#    print(key)
    if key not in scores.keys(): # set up scores[key]
        scores[key] = []
        for i in range(13):
            scores[key].append(0)
            scores[key].append(0)
        ''' the next four entries are lists of comments, strengths, 
        improvements, and recommendations '''
        for i in range(4):
            scores[key].append([])
        ''' now come entires 30 through 34 with self-explanatory values '''
        scores[key].append(fac_lname) #30
        scores[key].append(fac_fname) #31
        scores[key].append(fac_id)    #32
        scores[key].append(sec_id)    #33
        scores[key].append(dept)      #34
    ''' we are now guaranteed that scores[key] has been set up. Go ahead
    and add this student's responses to the respective fields. '''
    for i in range(13):
        if q[i] >= 0:
            scores[key][2*i] = scores[key][2*i] + q[i]
            scores[key][2*i+1] = scores[key][2*i+1] + 1
    if len(comment) > 0:
        scores[key][26].append(comment)
    if len(strengths) > 0:
        scores[key][27].append(strengths)
    if len(improvements) > 0:
        scores[key][28].append(improvements)
    if len(recommendations) > 0:
        scores[key][29].append(recommendations)
    ''' the 35th slot contains a bunch of information as a dictionary. this
    stuff was returned from get_course_sections and hails from enrollment.txt.
    It is a catch-all for additional course information. If a matching
    section record was found, append it in slot 35. Otherwise, append
    an empty dictionary '''
    if sec_id in sections:   
        scores[key].append(sections[sec_id])   # course information in slot #35
    else:
        scores[key].append({})
        print("Couldn't find matching course for %s" % key)
        
''' each entry of scores represents one faculty evaluation. We have the total
and count for each numeric scores, and lists of the qualitative responses.
Let's go ahead and compute the averages for the 13 quantitative questions.
Then, store the list of averages in the dictionary in slot [35] with key
"averages". Then, compute the overall average  and append this to slot[35]'s
dictionary with key "overall" '''
for key in scores:
    averages = []
    for i in range(13):
        if scores[key][2*i+1] > 0:
            averages.append(scores[key][2*i]/scores[key][2*i+1])
        else:
            averages.append(0)
    scores[key][35]["averages"] = averages
    scores[key][35]["overall"] = sum(averages)/13
    ''' output the report for each course section '''

    lname = scores[key][30]
    fname = scores[key][31]
    sec = scores[key][33].replace("-","_")
    ''' dept is used to figure out which folder ot put this report in. This
    will create the path if it doesn't exist '''
    dept = scores[key][34]
    dest_path = "%s\\%s" % (dest_folder,dept)
    if os.path.exists(dest_path) == False:
        os.mkdir(dest_path)
    ''' the name of the file is a combination of lname, fname, and sec '''
    file_id = "%s_%s_%s" % (lname,fname,sec)
    fname = "%s\\%s.html" % (dest_path,file_id)
    fvar = open(fname,"w",encoding="utf-8")
    fvar.write("<html>")
    fvar.write("<head><title>%s</title></head>\n" % file_id)
    fvar.write('<body style="width:800px; font-family:Arial;">\n')
    fvar.write('<center><h5>Faculty Course Evaluation</h5>')
    ''' another entry to the scores[key][35] dictionary - the number of 
    respondees. This comes from enrollment.txt '''
    responded = get_response_count(scores[key])
    scores[key][35]["responded"] = responded
    enrolled = scores[key][35]["enrollment"]
    if enrolled > 0:
        resp_rate = responded/enrolled * 100
    else:
        resp_rate = 0
    course_title = scores[key][35]["title"]
    instr_method = scores[key][35]["instruction_method_description"]
    sched_desc = scores[key][35]["schedule_description"]
    dept_desc = scores[key][35]["department_description"]
    fvar.write('<h3>%s</h3><h5>%s<br />' % (file_id,course_title))
    fvar.write('%s %s in %s<br />' % (instr_method,sched_desc,dept_desc))
    fvar.write('%d students responded out of %d enrolled (rate: %.2f%%)</h5></center>' % (responded,enrolled,resp_rate))
    fvar.write('<table border="1"><tr><th> &nbsp; </th><th>Question</th><th>Mean Score</th></tr>\n')
    for i in range(13):
        fvar.write('<tr><td style="text-align:right;">%2d.</td><td>%s</td><td style="text-align:center;">%.2f</td></tr>\n'% ((i+1),questions[i],averages[i]))
    fvar.write("</table><br /><br />\n")
    fvar.write("<h3>%s</h3>" % questions[13])
    for comment in scores[key][26]:
        fvar.write("%s<br /><br />\n" % comment)
    fvar.write("<h3>%s</h3>" % questions[14])
    for strength in scores[key][27]:
        fvar.write("%s<br /><br />\n" % strength)
    fvar.write("<h3>%s</h3>" % questions[15])
    for improvement in scores[key][28]:
        fvar.write("%s<br /><br />\n" % improvement)
    fvar.write("<h3>%s</h3>" % questions[16])
    for recommendation in scores[key][29]:
        fvar.write("%s<br /><br />\n" % recommendation)
    fvar.write("</body>")
    fvar.write("</html>")
    fvar.close()
    
'''print course-by-course listing. Sections are listed horizontally so that
we can see how different sections compare. Column headings are color-coded
when the average scores is less than 4. This alerts people to a problem.
These reports are written as program summary html files '''
for program in courses_by_program:
    course_list = courses_by_program[program]
    fname = "%s\\%s_course_summary.html" % (dest_folder,program)
    fvar = open(fname,"w")
    fvar.write("<html><head><title>Course-by-Course Summary for Program %s</title></head>\n" % program)
    fvar.write('<body style="width:1024px; font-family:Arial;">\n')
    fvar.write("<h3>Course-by-Course Summary for Program %s</h3>\n" % program)
    for course in course_list:
        ''' for each course, find all the sections. The matching sections
        are stored in matches. matches is an entry in the scores directionary,
        and so it is a 35-slotted list '''
        matches = []
        for key in scores:
            if scores[key][35]["course_number"] == course:
                matches.append(scores[key])
        if len(matches) > 0:
            ''' for each match, print the average score for each of the 13
            numeric questions. Color-code the header if the overall average,
            which is in cell 35's dictionary under the key "overall", is
            less than 4.0 '''
            fvar.write("<h3>%s %s</h3>\n" % (course,matches[0][35]["title"]))
            fvar.write('<table border="1"><tr><th> &nbsp; </th><th style="width:300px;">Question</th>')
            for match in matches:
                if match[35]["overall"] < 4 and match[35]["overall"] > 0:
 #                   print(match[35]["overall"])
                    color = ' background-color:coral;'
                else:
                    color = '';
                fvar.write('<th style="text-align:center;%s"><a href="%s/%s_%s_%s.html">%s,<br/>%s</a><br/>(%s)<br />%d resp of %d</th>' % (color,match[34],match[30],match[31],match[33].replace("-","_"),match[30],match[31],match[35]['sec'],match[35]["responded"],match[35]["enrollment"]))
            fvar.write("</tr>\n")
            for i in range(13):
                fvar.write("<tr>")
                fvar.write('<td style="text-align:right;">%2d.</td>' % (i+1))
                fvar.write('<td style="width:300px;">%s</td>' % questions[i])
                for match in matches:
                    fvar.write('<td style="text-align:center;">%.2f</td>' % match[35]["averages"][i])
                fvar.write("</tr>\n")
            fvar.write("</table>\n")
    fvar.write("</body></html>")
    fvar.close()
workbook.release_resources()
del workbook
print("Done")

        











