import pyodbc
import pandas as pd
import numpy as np
from os import listdir, system
from pathlib import Path
import openpyxl as xl
from ctypes import *
from warnings import filterwarnings
filterwarnings("ignore", category=UserWarning)

default_removals = "inc:inc.:,inc:llc:llc.:,llc:llp:llp.:,llp:co:co.:plan:and:&:trust:ltd:,ltd:ltd.:the:401:401k:401(k):k:(k):assetmark:db"

try:
    with open('format.txt', 'r') as f:
        lines = f.readlines()
        removals = lines[0].split(':')
        SOURCE_TYPE = lines[3][9:-1]
        PLAN_NAME = [lines[5][12:-1], lines[7][7:-1], lines[8][7:-1], lines[9][7:-1]]
        PARSED_ORG_COL = lines[10][18:-1]
        PARTY_COL = lines[11][11:-1]
        PLAN_COL = lines[12][10:-1]
        INITIAL_COL = lines[14][10:-1]
        COMMENT_COL = lines[15][10:-1]
        EIN_COL = lines[18][12:-1]
    del lines
except FileNotFoundError:
    with open("format.txt", 'w') as file:
        file.write(default_removals+'\n'+"Matchtype = A"+"\n"+"Helper_ID = B"+'\n'+"Source = C"+'\n'+"Sponsor_Name = F"+'\n'+"Plan_Name = G"+'\n'+"Reg1 = S"+'\n'+"Reg2 = T"+'\n'+"Reg3 = U"+'\n'+"Reg4 = V"+'\n' +
                   "PARSED_ORG_NAME = Y"+'\n'+"PARTY_ID = BE"+'\n'+"PLAN_ID = BF"+'\n'+"Date_Column = AP"+'\n'+"Initial = AQ"+'\n'+"Comment = BD"+'\n'+"P.Checked = BA"+'\n'+"P.Search  = BB"+'\n'+"P.Results = BC"+'\n'+"EOF")
    input("Rerun the program!")
    exit()



def main():
    file_name = excel_file_name_input()
    print(f'File - {file_name}\nOpening File...')
    data_excel = xl.load_workbook(Path(file_name))
    data_sheet = data_excel.active
    print("File Opened Successfully!!")

    row_range = input("\n*Range of the records to be checked (Example: '100-200')\n or press 'Enter' key for whole file:\n")
    if row_range == "":
        r1, r2 = 2, data_sheet.max_row
    else:
        try:
            r1, r2 = [int(i) for i in row_range.split('-')]
            if r1 > r2:
                r1, r2 = r2, r1
            elif r1 == r2:
                input("Range can not be Equal!! Closing...")
                exit()
            else:
                pass
        except ValueError:
            input("Wrong input range typed!! Closing...")
            exit()
    driver = "SQL Server Native Client 11.0"
    server = input("Server Name (w000000\cgsql): ")
    database = input("Database Name(CMX_ORS_10_3): ")

    cnxn = pyodbc.connect(f"Driver={driver};Server={server};Database={database};Trusted_connection=Yes;")

    sql_querry = [f"select distinct p.rowid_object as Party_ID, pl.ROWID_OBJECT as Plan_ID, p.ORG_NAME as Org_Name, p.party_tax_id as EIN, pl.Plan_Name as Plan_Name",
                  f"from {database}..c_party p left join {database}..c_party_relationship rel on p.ROWID_OBJECT=rel.TO_PARTY left join {database}..c_party_address pa on rel.from_party_id=pa.party_id left join {database}..c_address addr on pa.ADDRESS_ID = addr.ROWID_OBJECT left join {database}..c_plan_party_role ppr on p.ROWID_OJECT=ppr.PARTY_ID left join {database}...c_plan pl on ppr.PLAN_ID = pl.ROWID_OBJECT",
                  f"and p.PARTY_TYPE_ID=2101 and p.HUB_STATE_IND=1 and pl.HUB_STATE_IND=1 and p.rowid_object in (select party_id from {database}..c_party_role where ROLE_TYPE_ID=4108) order by ORG_NAME, PLAN_ID"]
    
    type_of_rec_to_do = input("\nType of Records to be done:\n 1. All\n 2. Only EIN")

    previous_ein = count = previous_org_name = previous_plan_name = df = found_unique = found_multiple = 0
    system('cls')
    data_sheet[f"{COMMENT_COL}1"].value = "SQL.Result"
    comment_update = "not found"
    for row in range(r1, r2+1):
        print_at(0,0, f"Row No : {row}")
        if data_sheet[f"{INITIAL_COL}{row}"].value not in [None,""] or data_sheet[f"{COMMENT_COL}{row}"].value not in [None,""]:
            continue

        org_name = data_sheet[f"{PARSED_ORG_COL}{row}"].value
        system('cls')
        print_at(0,0,f"Row No : {row}\n\n")
        print(f"Given Org_Name: {org_name}\n")
        print("Source Type: ", data_sheet[f"{SOURCE_TYPE}{row}"].value)

        org_name = clean(str(org_name).lower())

        data_sheet[f"{COMMENT_COL}{row}"].value = "not found" if type_of_rec_to_do == 1 else None
        ein = data_sheet[f"{EIN_COL}{row}"].value
        if data_sheet[f"{SOURCE_TYPE}{row}"].value == "DST":
            plan_name = " ".join([str(data_sheet[f"{i}{row}"].value) for i in PLAN_NAME[1] if data_sheet[f"{i}{row}"].value not in [None, ""]]).lower()
        else:
            plan_name = str(data_sheet[f"{PLAN_NAME[0]}{row}"].value).lower()
        
        print(f"\nGiven Plan_Name: {plan_name}")

        if org_name == previous_org_name:
            if flag:
                data_sheet[f"{PARTY_COL}{row}"].value = party_id
                data_sheet[f"{PLAN_COL}{row}"].value = plan_id
            data_sheet[f"{COMMENT_COL}{row}"].value = comment_update
            continue
        elif ein in [None, "", "Not Found", "Not Searched", "Multiple Webpages", "Beacon: Found without EIN", "Mannual Check is Required"]:
            if org_name in [None, ""] or type_of_rec_to_do == '2':
                continue
            sql3 = f"where CONCAT(org_name, ' ', plan_name) like '{org_name}'"
            search_method = False
            print(f"\nSearch Org_name: {org_name}")
        elif (ein == previous_ein) and flag:
            data_sheet[f"{PARTY_COL}{row}"].value = party_id
            data_sheet[f"{PLAN_COL}{row}"].value = plan_id
            data_sheet[f"{COMMENT_COL}{row}"].value = comment_update
        else:
            search_method = True
            print(f"\nSearch EIN: {ein}")
            sql3 = f"where p.party_tax_id in ('{ein}')"
        
        flag = False
        comment_update = "not found"
        count += 1
        print(f"\nCount : {count}\nFound Unique: {found_unique}\nFound Multiple: {found_multiple}")

        code = f"{sql_querry[0]} {sql_querry[1]} {sql3} {sql_querry[2]}"
        print("\n-->Running Database Search...\n")
        print("\nPrevious Fetched Data:")
        print(f"\nSearch Org_Name: {previous_org_name}")
        print(f"Plan Name: {previous_plan_name}\n")
        print(df)
        df = pd.read_sql(code,cnxn)

        plan_types = {
            '401psp' : [['401k profit', '401(k) profit', '401 k profit'], 0],
            'psp': [['ps', 'profit', 'prof'], 0],
            'saving' : [['savings', 'saving', 'svgs'], 0],
            'sh': [['safe', 'safe harbour', 'safe harbor'], 0],
            '401': [['401'], 0],
            '403': [['403'], 0],
            'cb': [['cb', 'cash balance', 'cash'], 0],
            'ret': [['retirement', "retiremen", 'retirem', 'ret', 'rtmt'], 0]
        }

        if len(df) ==0:
            if search_method:
                comment_update = f"Beacon: Found with EIN {ein}"
                data_sheet[f"{COMMENT_COL}{row}"].value = comment_update
        elif len(df['Party_ID'].unique()) == 1:
            comment_update = "found_unique"
            data_sheet[f"{COMMENT_COL}{row}"].value = comment_update
            found_unique += 1
            party_id = int(str(df['Party_ID'][0]).strip())
            data_sheet[f"{PARTY_COL}{row}"].value = party_id

            found_any = False
            for key, value in plan_types.items():
                for v in value[0]:
                    if v in plan_name:
                        plan_types[key][1] = 1
                        found_any = True
                        break
            if found_any:
                plans_found = df[['Plan_ID', 'Plan_Name']].drop_duplicates()
                matches_with_given_plan_found = []
                for i, x in plans_found.iterrows():
                    plan_id, p = x['Plan_ID'], x['Plan_Name']

                    if p not in [None, ""]:
                        for key, value in plan_types.items():
                            matched = False
                            if value[1] == 1:
                                for v in value[0]:
                                    if v in p.lower():
                                        matched = True
                                        matches_with_given_plan_found.append(plan_id)
                                        break
                            if matched:
                                break

                if len(matches_with_given_plan_found) == 1:
                    plan_id = int(matches_with_given_plan_found[0].strip())
                    data_sheet[f"{PLAN_COL}{row}"].value = plan_id
                else:
                    plan_id = None

            else:
                plan_id = df['Plan_ID'].unique()
                if len(plan_id) == 1 and plan_id[0] != None:
                    plan_id = int(plan_id[0].strip())
                    data_sheet[f"{PLAN_COL}{row}"].value = plan_id
                else:
                    plan_id = None
            flag = True
        else:
            comment_update = "found_multiple"
            data_sheet[f"{COMMENT_COL}{row}"].value = comment_update
            found_multiple += 1

        previous_ein = ein
        previous_org_name = org_name
        previous_plan_name = plan_name

        if count % 50 == 0:
            print("\n\n*** Saving File...")
            data_excel.save(f"./{file_name[:-5]} Auto.xlsx")
            print("\n*** File Saved !!! ***")
            system('cls')

    system('cls')
    print_at(0,0,f"Row No : {row}\n\n")
    print_at(2,0,f"Count  : {count}\n\n")
    print("\n\n***Saving File...")
    data_excel.save(f"./{file_name[:-5]} Auto.xlsx")
    input("\n*** File Saved !!! ***")

def clean(plan=''):  # Remove the words & special characters to search
    temp = plan.lower().split(' ')
    plan = ''
    for word in temp:
        if (word in removals) or (len(word) < 3):
            plan += '%' + ' '
        else:
            r = ["'s", ',', '&', '-', '.', "'", '#', '(', ')']
            for i in r:
                word = word.replace(i, "%")
            plan += word + ' '
    plan = (("%" + plan + "%").replace(" %", "%")).strip()
    try:
        while "%%" in plan:
            plan = plan.replace("%%", "%")
    except:
        return None
    return plan

def excel_file_name_input():  # Select Excel File to be worked upon
    print("Excel files in same directory:")
    files = listdir('.')
    xl_files = []
    counter = 0
    for file in files:
        if file[-4:] == 'xlsx':
            counter += 1
            xl_files.append(file)
            print(f"{counter}. {file}")
    if counter == 1:
        return xl_files[0]
    elif counter == 0:
        input("\n**No Excel Files Found.\n  Closing Program !!")
        exit()
    option = int(
        input(f"\nEnter the file number corresponding to file name: ")) - 1
    if len(xl_files) > option >= 0:
        return xl_files[option]
    else:
        input("\n\n**Wrong Input !!!\n  Closing Program !!")
        exit()


class COORD(Structure):
    pass

STD_OUTPUT_HANDLE = -11
COORD._fields_ = [("X", c_short), ("Y", c_short)]

def print_at(r,c,s):
    h = windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    windll.kernel32.SetConsoleCursorPosition(h, COORD(c,r))

    c = s.encode("windows-1252")
    windll.kernel32.WriteConsoleA(h,c_char_p(c), len(c), None, None)

main()
