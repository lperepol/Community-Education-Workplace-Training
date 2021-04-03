# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import datetime
from datetime import timedelta

def DescriptionTemplate(descrp, ce_url):
    Template= ""
    if r'\n' in descrp:
        print (r"'\n' in descrp:")
    descrp = descrp.replace(r'\n', '<br>')
    Template = Template + '<p>&nbsp;</p>'
    Template = Template + '<h3><strong>Course  Description:</strong></h3></p>'
    Template = Template + '<p><strong>' + descrp + '</strong></p>'
    Template = Template + '<p>&nbsp;</p>'
    Template = Template + '<p><a class="white_button_main_content" href="' + ce_url + '">Register Now</a></p>'
    Template = Template + '<p>&nbsp;</p>'
    Template = Template + ''
    return Template

def PROGAREA_CODE():
    adict = {
        "TL":"Teaching and Learning",
        "GP":"Gathering Place",
        "LE":"Applied Leadership Citation Program",
        "AC":"Arts & Culture",
        "BF":"Business, Finance and Leadership",
        "CY":"Child and Youth Programming",
        "CD":"Computers and Digital Technology",
        "ES":"Environment and Sustainability",
        "FA":"First Aid and Safety",
        "GA":"Gardening",
        "HW":"Health and Wellness",
        "HT":"Hospitality and Tourism",
        "IT":"Industry and Trades",
        "LA":"Languages",
        "RL":"Recreation and Leisure",
        "MI":"Mir Centre for Peace",
        "SKDI":"Ski Resort Operations and Management",
        "LR":"Learning in Retirement",
        "ALT":"Adult Literacy",
        "PD":"Selkirk College Professional Development",
        "BCEL":"BC Electrical Code",
        "ICE":"Refrigeration Plant Operation",
        "SPEC":"Spectrum"
    }
    return adict

def read_csv(fname):
    df = pd.read_csv(fname)
    return df

def read_excel(fname):
    df = pd.read_excel(open(fname, 'rb'), sheet_name='Sheet1')
    return df

def read_excel_2():
    file_name = 'C:/SelkirkCollege/CE_Calender/New/CE Winter 2021.Online Registration Extract.xlsx'
    xl_file = pd.ExcelFile(file_name)

    dfs = {sheet_name: xl_file.parse(sheet_name)
           for sheet_name in xl_file.sheet_names}
    return dfs

def get_Sections():
    section_set = {
        "F20G01",
        "F21C01",
        "F21G01",
        "F21K01",
        "F21K02",
        "F21K03",
        "F21K04",
        "F21P01",
        "F21R01",
        "F21R02",
        "F21T01",
        "F21T02",
        "S21A01",
        "S21C01",
        "S21C02",
        "S21C03",
        "S21C04",
        "S21C05",
        "S21C06",
        "S21G01",
        "S21G02",
        "S21G03",
        "S21K01",
        "S21K02",
        "S21N01",
        "S21N02",
        "S21N03",
        "S21R01",
        "S21R02",
        "S21R03",
        "S21R04",
        "S21R05",
        "S21R06",
        "S21R07",
        "S21R08",
        "S21R09",
        "S21R10",
        "S21R11",
        "S21R12",
        "S21R13",
        "S21R14",
        "S21R15",
        "S21R16",
        "S21R17",
        "S21R18",
        "S21R19",
        "S21T01",
        "S21T02",
        "S21T03",
        "S21T04",
        "S21T05",
        "S21T06",
        "S21T07",
        "S21T08",
        "W21A01",
        "W21A02",
        "W21C01",
        "W21C02",
        "W21C03",
        "W21C04",
        "W21C05",
        "W21C06",
        "W21C07",
        "W21C08",
        "W21C09",
        "W21C10",
        "W21C11",
        "W21C12",
        "W21C13",
        "W21C14",
        "W21C15",
        "W21G01",
        "W21G02",
        "W21G03",
        "W21G04",
        "W21G05",
        "W21G06",
        "W21G07",
        "W21G08",
        "W21K01",
        "W21K02",
        "W21N01",
        "W21N02",
        "W21N03",
        "W21N04",
        "W21N05",
        "W21N06",
        "W21P01",
        "W21P02",
        "W21R01",
        "W21R02",
        "W21R03",
        "W21R04",
        "W21R05",
        "W21R06",
        "W21R07",
        "W21R08",
        "W21R09",
        "W21R10",
        "W21R11",
        "W21R12",
        "W21R13",
        "W21R14",
        "W21R15",
        "W21R16",
        "W21R17",
        "W21R18",
        "W21R19",
        "W21R20",
        "W21R21",
        "W21T01",
        "W21T02",
        "W21T03",
        "W21T04",
        "W21T05",
        "W21T06",
        "W21T07",
        "W21T08",
        "W21T09",
        "W21T10",
        "W21T11",
        "W21T12",
        "W21T13",
        "W21T14",
        "W21T15"}
    return section_set



def process_Section():
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/CE.SECTION.xlsx'
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/cvs/CE.SECTION.csv'

    df = pd.read_csv(file_name)
    course_set = set()
    section_set = get_Sections()
    course_dict = dict()
    for index, row in df.iterrows():
        SECTION_IRN = str(row['SECTION_IRN']).strip()
        COURSE_IRN = str(row['COURSE_IRN']).strip()
        COURSE_CODE = str(row['COURSE_CODE']).strip()
        SECTION_CODE = str(row['SECTION_CODE']).strip()
        TUITION_FEE = str(row['TUITION_FEE']).strip()
        LAB_FEE = str(row['LAB_FEE']).strip()
        key = (SECTION_IRN,COURSE_IRN,COURSE_CODE,SECTION_CODE,TUITION_FEE,LAB_FEE)
        if SECTION_CODE in section_set:
            course_set.add(key)

    course_dict = dict()
    for i in course_set:
        (SECTION_IRN,COURSE_IRN,COURSE_CODE,SECTION_CODE,TUITION_FEE,LAB_FEE) = i
        if (COURSE_IRN == "37231"):
            print (COURSE_IRN)
        #if COURSE_IRN == '37139':
        #    print(COURSE_IRN)
        course_dict[SECTION_IRN] = i


    return course_dict

def process_Time_Table(course_dict):
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/CE.SECTION_TIMETABLE.xlsx'
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/cvs/CE.SECTION_TIMETABLE.csv'
    df = pd.read_csv(file_name)

    for i in course_dict:
        kkk = course_dict[i]
        (SECTION_IRN,COURSE_IRN,COURSE_CODE,SECTION_CODE,TUITION_FEE,LAB_FEE) = kkk
        if (COURSE_IRN == "37231"):
            print (COURSE_IRN)

    course_set = set()
    for index, row in df.iterrows():
        SECTION_IRN = str(row['SECTION_IRN']).strip()
        if SECTION_IRN in course_dict:
            (SECTION_IRN, COURSE_IRN, COURSE_CODE, SECTION_CODE, TUITION_FEE, LAB_FEE) = tuple(course_dict[SECTION_IRN])
            START_DATE = str(row['START_DATE']).strip()
            print(START_DATE)
            #START_DATE_dt = datetime.datetime.strptime(START_DATE, '%Y-%m-%d %H:%M:%S')
            START_DATE_dt = datetime.datetime.strptime(START_DATE, '%Y-%m-%d')

            Unpublish_date = START_DATE_dt + timedelta(days=1)
            Unpublish_date = Unpublish_date.strftime('%Y-%m-%d')

            END_DATE = str(row['END_DATE']).strip()
            #END_DATE_dt = datetime.datetime.strptime(END_DATE, '%Y-%m-%d %H:%M:%S')
            END_DATE_dt = datetime.datetime.strptime(END_DATE, '%Y-%m-%d')

            MEET_START = str(row['MEET_START']).strip()
            if len(MEET_START) > 2:
                START_DATE_dt_str = START_DATE_dt.strftime('%Y-%m-%d') + " " + MEET_START[:2] + ':' + MEET_START[2:]
            else:
                START_DATE_dt_str = START_DATE_dt.strftime('%Y-%m-%d %H:%M')

            MEET_END = str(row['MEET_END']).strip()
            if len(MEET_END) > 2:
                END_DATE_dt_str = END_DATE_dt.strftime('%Y-%m-%d') + " " + MEET_END[:2] + ':' + MEET_END[2:]
            else:
                END_DATE_dt_str = END_DATE_dt.strftime('%Y-%m-%d %H:%M')

            START_DATE = datetime.datetime.strptime(START_DATE_dt_str, '%Y-%m-%d %H:%M')
            END_DATE = datetime.datetime.strptime(END_DATE_dt_str, '%Y-%m-%d %H:%M')
            LastDayOfClass = END_DATE_dt.strftime('%Y-%m-%d %H:%M')

            dateDiff = (END_DATE - START_DATE).days
            if dateDiff > 0:
                if len(MEET_END) > 2:
                    END_DATE_dt_str = START_DATE_dt.strftime('%Y-%m-%d') + " " + MEET_END[:2] + ':' + MEET_END[2:]
                else:
                    END_DATE_dt_str = START_DATE_dt.strftime('%Y-%m-%d %H:%M')

            START_DATE = datetime.datetime.strptime(START_DATE_dt_str, '%Y-%m-%d %H:%M')
            END_DATE = datetime.datetime.strptime(END_DATE_dt_str, '%Y-%m-%d %H:%M')
            dateDiff = (END_DATE - START_DATE).seconds
            if dateDiff < 3600:
                END_DATE = END_DATE + datetime.timedelta(minutes=45)
                END_DATE_dt_str = END_DATE.strftime('%Y-%m-%d %H:%M')

            START_DATE = datetime.datetime.strptime(START_DATE_dt_str, '%Y-%m-%d %H:%M')
            END_DATE = datetime.datetime.strptime(END_DATE_dt_str, '%Y-%m-%d %H:%M')
            START_DATE = START_DATE + timedelta(seconds=1)
            START_DATE_dt_str = START_DATE.strftime('%Y-%m-%d %H:%M:%S')
            END_DATE = END_DATE + timedelta(seconds=1)
            END_DATE_dt_str = END_DATE.strftime('%Y-%m-%d %H:%M:%S')

            LOCATION = str(row['LOCATION']).strip()
            print("Location -->" + LOCATION)

            crs = (
                SECTION_IRN,
                COURSE_IRN,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass
            )

            course_set.add(crs)

    course_dict = dict()
    for i in course_set:
        (
            SECTION_IRN,
            COURSE_IRN,
            COURSE_CODE,
            SECTION_CODE,
            START_DATE_dt_str,
            END_DATE_dt_str,
            TUITION_FEE,
            LAB_FEE,
            LOCATION,
            Unpublish_date,
            LastDayOfClass
        ) = i

        if COURSE_IRN in course_dict:
            course_dict[COURSE_IRN].append(i)
        else:
            course_dict[COURSE_IRN] = list()
            course_dict[COURSE_IRN].append(i)

    for i in course_dict:
        kkk = course_dict[i]
        for jjj in kkk:
            (
                SECTION_IRN,
                COURSE_IRN,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass
            ) = jjj
            if (COURSE_IRN == "37231"):
                print (COURSE_IRN)

    return course_dict

def process_Course_Master(course_dict):

    for i in course_dict:
        kkk = course_dict[i]
        for jjj in kkk:
            (
                SECTION_IRN,
                COURSE_IRN,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass
            ) = jjj
            if (COURSE_IRN == "37231"):
                print (COURSE_IRN)

    #file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/CE.COURSE_MASTER.xlsx'
    #df = pd.read_excel(file_name)
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/cvs/CE.COURSE_MASTER.csv'
    df = pd.read_csv(file_name)
    course_set = set()
    section_set = get_Sections()
    course_set = set()
    PROGAREA_CODE_look_Up = PROGAREA_CODE()
    for index, row in df.iterrows():
        COURSE_IRN = str(row['COURSE_IRN']).strip()
        COURSE_DESCR = str(row['COURSE_DESCR']).strip()
        MAIN_PROGAREA = PROGAREA_CODE_look_Up[str(row['MAIN_PROGAREA']).strip()]
        TITLE = str(row['TITLE']).strip()
        if COURSE_IRN in course_dict:
            alsit = course_dict[COURSE_IRN]
            for i in alsit:
                (
                    SECTION_IRN,
                    COURSE_IRN,
                    COURSE_CODE,
                    SECTION_CODE,
                    START_DATE_dt_str,
                    END_DATE_dt_str,
                    TUITION_FEE,
                    LAB_FEE,
                    LOCATION,
                    Unpublish_date,
                    LastDayOfClass
                ) = i
                COURSE_CODE = str(row['COURSE_CODE']).strip()

                course_instance =                 (
                    TITLE,
                    SECTION_IRN,
                    COURSE_IRN,
                    MAIN_PROGAREA,
                    COURSE_CODE,
                    SECTION_CODE,
                    START_DATE_dt_str,
                    END_DATE_dt_str,
                    TUITION_FEE,
                    LAB_FEE,
                    LOCATION,
                    Unpublish_date,
                    LastDayOfClass,
                    COURSE_DESCR
                )
                course_set.add(course_instance)

    course_dict = dict()
    for i in course_set:
        (
            TITLE,
            SECTION_IRN,
            COURSE_IRN,
            MAIN_PROGAREA,
            COURSE_CODE,
            SECTION_CODE,
            START_DATE_dt_str,
            END_DATE_dt_str,
            TUITION_FEE,
            LAB_FEE,
            LOCATION,
            Unpublish_date,
            LastDayOfClass,
            COURSE_DESCR
        ) = i
        if TITLE == "OFA 1 Occupational First Aid level 1":
            hh = (COURSE_IRN, TITLE,START_DATE_dt_str,END_DATE_dt_str )
            #print (hh)

        if TITLE in course_dict:
            course_dict[TITLE].append(i)
        else:
            course_dict[TITLE] = list()
            course_dict[TITLE].append(i)

    for i in course_dict:
        lst = course_dict[i]
        for j in lst:
            (
                TITLE,
                SECTION_IRN,
                COURSE_IRN,
                MAIN_PROGAREA,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass,
                COURSE_DESCR
            ) = j
            if (COURSE_IRN == "37231"):
                print (COURSE_IRN)


    return course_dict


def create_upload_table(course_dict):

    for i in course_dict:
        lst = course_dict[i]
        for j in lst:
            (
                TITLE,
                SECTION_IRN,
                COURSE_IRN,
                MAIN_PROGAREA,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass,
                COURSE_DESCR
            ) = j
            if (COURSE_IRN == "37231"):
                print (COURSE_IRN)

    df = pd.DataFrame()
    df["Title"] =""
    df["Section"] =""
    df["Category"] =""
    df["Course"] =""
    df["Course Description"] =""
    df["Location"] =""
    df["LAB_FEE"] =""
    df["TUITION_FEE"] =""
    df["CE_REG_URL_TITLE"] =""
    df["CE_REG_URL"] =""
    df["START_DATE"] =""
    df["END_DATE"] =""
    df["Unpublish_date"] =""
    df["LastDayOfClass"] =""
    count = 0
    for i in course_dict:
        course = course_dict[i]
        for row in course:
            (
                Title,
                SECTION_IRN,
                COURSE_IRN,
                Category,
                COURSE_CODE,
                SECTION_CODE,
                START_DATE_dt_str,
                END_DATE_dt_str,
                TUITION_FEE,
                LAB_FEE,
                LOCATION,
                Unpublish_date,
                LastDayOfClass,
                COURSE_DESCR
            ) = row
            if "A" in SECTION_CODE:
                LOCATION = "Nelson Victoria St."
            elif "C" in SECTION_CODE:
                LOCATION = "Castlegar"
            elif "G" in SECTION_CODE:
                LOCATION = "Grand Forks"
            elif "K" in SECTION_CODE:
                LOCATION = "Kaslo"
            elif "N" in SECTION_CODE:
                LOCATION = "Nakusp"
            elif "P" in SECTION_CODE:
                LOCATION = "Nelson Tenth St."
            elif "R" in SECTION_CODE:
                LOCATION = "Nelson Silver King"
            elif "T" in SECTION_CODE:
                LOCATION = "Trail"

            if "W" in SECTION_CODE:
                SECTION_CODE =  SECTION_CODE + " (Winter)"
            elif "S" in SECTION_CODE:
                SECTION_CODE =  SECTION_CODE + " (Summer)"
            elif "F" in SECTION_CODE:
                SECTION_CODE =  SECTION_CODE + " (Fall)"
            Course = ""
            START_DATE = datetime.datetime.strptime(START_DATE_dt_str, '%Y-%m-%d %H:%M:%S')
            END_DATE = datetime.datetime.strptime(END_DATE_dt_str, '%Y-%m-%d %H:%M:%S')
            dateDiff = (END_DATE - START_DATE).days
            if dateDiff > 0:
                print ("Date diff greater than zero")

            CE_REG_URL_TITLE = "Register Now!"
            CE_REG_URL = "https://cereg.selkirk.ca/SRS/cecourses.htm#option=course&crsid="+ COURSE_CODE.replace(" ", "+") + "&allstart=N"

            if (COURSE_IRN == "37231"):
                print (COURSE_IRN)

            Course_Description = DescriptionTemplate(COURSE_DESCR, CE_REG_URL)
            df.loc[count] = [Title,SECTION_CODE,Category,Course,Course_Description,LOCATION,LAB_FEE,TUITION_FEE,CE_REG_URL_TITLE,CE_REG_URL,START_DATE_dt_str,END_DATE_dt_str,LastDayOfClass,Unpublish_date]
            count = count + 1
    df.to_csv('Test_out.csv', index=False, date_format='%Y-%m-%d %H:%M')

    return df

def populate_node_ids(df):
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/cvs/continuing_education_category.csv'
    category_df =  read_csv(file_name)

    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/cvs/continuing_education_course.csv'
    course_df =  read_csv(file_name)

    for index, row in df.iterrows():
        Category = str(row['Category']).strip()
        if Category == "Arts & Culture":
            print (Category)
        Title = str(row['Title']).strip()
        for index1, row1 in category_df.iterrows():
            category_Title = str(row1['Title']).strip()
            if Category == category_Title:
                category_nid = str(row1['Nid']).strip()
                df.loc[index, 'Category'] = category_nid
        for index1, row1 in course_df.iterrows():
            course_Title = str(row1['Title']).strip()
            if Title == course_Title:
                course_nid = str(row1['Nid']).strip()
                df.loc[index, 'Course'] = course_nid

    df.to_csv('out.csv', index=False, date_format='%Y-%m-%d %H:%M')

def main():
    course_dict = process_Section()
    course_dict = process_Time_Table(course_dict)
    course_dict = process_Course_Master(course_dict)
    #new_course_dict = process_extract()
    #merge(course_dict, new_course_dict)
    df = create_upload_table(course_dict)
    populate_node_ids(df)


if __name__ == '__main__':
    file_name = 'C:/SelkirkCollege/CE_Calender/Oracle_CE/'

    print('Start...')
    main()
    print('Stop...')
# See PyCharm help at https://www.jetbrains.com/help/pycharm/




########################################################## Junk ##########################
def process_extract():
    Extract = read_excel_2()
    new_course_set = set()
    for df in Extract:
        category = df
        for index, row in Extract[df].iterrows():
            CourseTitle = str(row['Course Title']).strip()
            Course_Description = str(row['Course Description (for website & online reg)']).strip()
            Campus = str(row['Campus']).strip()
            Dates_Times = str(row['Dates & Times (ex: 1 class: Jan 15, Mon 8:30 am-5 pm)']).strip()
            Instructor_Name = str(row["Instructor's Name"]).strip()
            row = (category,CourseTitle, Course_Description,Campus,Dates_Times,Instructor_Name)
            new_course_set.add(row)
    course_dict = dict()
    for i in new_course_set:
        (category,CourseTitle, Course_Description,Campus,Dates_Times,Instructor_Name)= i
        if CourseTitle in course_dict:
            course_dict[CourseTitle].append(i)
        else:
            course_dict[CourseTitle] = list()
            course_dict[CourseTitle].append(i)
    return course_dict



def merge(course_dict, new_course_dict):
    for course_title in course_dict:
        value = course_dict[course_title]
        for course in value:
            (TITLE, SECTION_IRN, COURSE_IRN, COURSE_CODE, SECTION_CODE, START_DATE_dt_str, END_DATE_dt_str, LOCATION) = course
            aa_course = new_course_dict[course_title]
            for row in aa_course:
                (category, CourseTitle, Course_Description, Campus, Dates_Times, Instructor_Name) = row

