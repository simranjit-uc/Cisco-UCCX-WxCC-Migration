import platform, os, json, openpyxl, CCX_Sheet, WxCC


full_Path_WxCC = ""

# Creating a new Spreadsheet
wb_WxCC = openpyxl.Workbook()
WxCC_Apps_WS = wb_WxCC.create_sheet(title="Entry Points", index=0)
WxCC_Skill_WS = wb_WxCC.create_sheet(title="Skills", index=1)
WxCC_SP_WS = wb_WxCC.create_sheet(title="Skill Profile", index=2)
WxCC_SP_2_WS = wb_WxCC.create_sheet(title="Skill Profile1", index=3)
WxCC_Teams_WS = wb_WxCC.create_sheet(title="Teams", index=4)
WxCC_WrapCodes_WS = wb_WxCC.create_sheet(title="Wrap up Codes", index=5)
WxCC_IdleCodes_WS = wb_WxCC.create_sheet(title="Idle Codes", index=6)
WxCC_CSQ_WS = wb_WxCC.create_sheet(title="CSQ", index=7)
WxCC_Contacts_WS = wb_WxCC.create_sheet(title="Contacts", index=8)


# Loading the contents of the CCX-Details.xlsx sheet. Converting "Application" details into WxCC "Entry Points".
def app_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        col = 1
        try:
            WxCC_Apps_WS.cell(row=1, column=1, value="Name")
            WxCC_Apps_WS.cell(row=1, column=2, value="Description")
            WxCC_Apps_WS.cell(row=1, column=3, value="Maximum Sessions")

            CCX_Apps_WS = CCX_wb['Applications']
            for r in range(2, CCX_Apps_WS.max_row + 1):
                name = CCX_Apps_WS.cell(row=r, column=1).value
                max_Session = CCX_Apps_WS.cell(row=r, column=5).value
                WxCC_Apps_WS.cell(row=r, column=col, value=name)
                WxCC_Apps_WS.cell(row=r, column=col+1, value=name)
                WxCC_Apps_WS.cell(row=r, column=col+2, value=max_Session)
            col = 1
            print("Converting Applications : \033[92mDone\033[0m")
            skillprof_WxCC()
        except:
            print("An error was encountered while opening the file located at APPS WxCC")
    except:
        print("Program has encountered an error.")


# Loading the contents of the CCX-Details.xlsx sheet. Creating WxCC "Skill Profile" values.
def skillprof_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        result = []
        result_dict = {}

        try:
            ws_Resources = CCX_wb["Resources"]
            row_Resource = ws_Resources.max_row
            col_List = [5]

            count = 0
            heading_Check = ws_Resources[1]
            for cell in heading_Check:
                if "Skill" in str(cell.value):
                    count += 1
                    col_List.append(cell.column)

            col_List_Length = len(col_List)
            i = 1

            for col_idx in col_List:
                for row_idx in range(1, ws_Resources.max_row + 1):
                    # Getting the value from the source sheet
                    cell_value = ws_Resources.cell(row=row_idx, column=col_idx).value
                    # Setting the value in the destination sheet
                    WxCC_SP_WS.cell(row=row_idx, column=i).value = cell_value
                i += 1

            for row in range(2, WxCC_SP_WS.max_row + 1):
                team = WxCC_SP_WS.cell(row=row, column=1).value
                skill_List = []
                for col in range(2, 8):
                    cell_value = WxCC_SP_WS.cell(row=row, column=col).value
                    if cell_value is not None:
                        skill_List.append(cell_value)

                skill_Team = {"Team": team, "Skills": skill_List}

                result.append(skill_Team)

            for entry in result:
                team = entry.get('Team')
                if team:
                    values = entry.get('Skills')
                    if team in result_dict:
                        result_dict[team].extend(values)
                    else:
                        result_dict[team] = values

            result = [{'Team': key, 'Skills': values} for key, values in result_dict.items()]
            result = [item for item in result if len(item['Skills']) > 0]

            WxCC_SP_2_WS.cell(row=1, column=1, value="Skill Profile Name")
            WxCC_SP_2_WS.cell(row=1, column=2, value="Team Name")
            for a in range(0, len(result)):
                col_val = 3
                team = f"{result[a]['Team']}"
                for i in range(0, len(result[a]['Skills'])):
                    WxCC_SP_2_WS.cell(row=a + 2, column=1, value=f"{result[a]['Team']} Skill Profile")
                    WxCC_SP_2_WS.cell(row=a + 2, column=2, value=team)
                    WxCC_SP_2_WS.cell(row=1, column=col_val, value=f"Skill Name")
                    WxCC_SP_2_WS.cell(row=a + 2, column=col_val, value=result[a]['Skills'][i])
                    col_val += 1
            print("Converting Skills : \033[92mDone\033[0m")
            skills_WxCC()
        except:
            print("An error was encountered while opening the file located at SP WxCC")
    except:
        print("Program has encountered an error.")


# Loading the contents of the CCX-Details.xlsx sheet. Converting "Skill" details into WxCC "Skill Definitions".
def skills_WxCC():
    try:
        try:
            CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
            CCX_Skill_WS = CCX_wb['Skills']
            for row in CCX_Skill_WS.iter_rows():
                for cell in row:
                    WxCC_Skill_WS[cell.coordinate].value = cell.value
            print("Converting Skills : \033[92mDone\033[0m")
            teams_WxCC()
        except:
            print("An error was encountered while opening the file located at SKILLS WxCC")
    except:
        print("Program has encountered an error.")


# Loading the contents of the CCX-Details.xlsx sheet. Converting "Team" details into WxCC "Teams".
def teams_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        row = 1
        col = 1
        try:
            WxCC_Teams_WS.cell(row=1, column=1, value="Team")
            WxCC_Teams_WS.cell(row=1, column=2, value="Supervisor Name")
            WxCC_Teams_WS.cell(row=1, column=3, value="Skill Profile")
            CCX_Teams_WS = CCX_wb['Teams']
            team_Values = set()
            for r in range(1, CCX_Teams_WS.max_row+1):
                r += 1
                team = CCX_Teams_WS.cell(row=r, column=2).value
                Supervisor = CCX_Teams_WS.cell(row=r, column=3).value
                WxCC_Teams_WS.cell(row=r, column=col, value=team)
                WxCC_Teams_WS.cell(row=r, column=col+1, value=Supervisor)
                team_Values.add(CCX_Teams_WS.cell(row=r, column=2).value)

            matched_values = []

            for a in range(2, WxCC_SP_2_WS.max_row+1):
                team = WxCC_SP_2_WS.cell(row=a, column=2).value
                team_List = []
                if team in team_Values:
                    team_List.append(team)
                    team_List.append(WxCC_SP_2_WS.cell(row=a, column=1).value)
                matched_values.append(team_List)

            for item in matched_values:
                for b in range(2, WxCC_Teams_WS.max_row+1):
                    if item[0] == WxCC_Teams_WS.cell(row=b, column=1).value:
                        WxCC_Teams_WS.cell(row=b, column=3, value=item[1])
            print("Converting Teams : \033[92mDone\033[0m")
            codes_WxCC()
        except:
            print("An error was encountered while opening the file located at TEAMS WxCC")
    except:
        print("Program has encountered an error.")


# Loading the contents of the CCX-Details.xlsx sheet. Converting "Wrap-Up and Reason Codes" details into WxCC "Aux Codes".
def codes_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        col = 1
        try:
            WxCC_WrapCodes_WS.cell(row=1, column=1, value="Name")
            WxCC_WrapCodes_WS.cell(row=1, column=2, value="Description")
            WxCC_WrapCodes_WS.cell(row=1, column=3, value="Type")

            WxCC_IdleCodes_WS.cell(row=1, column=1, value="Name")
            WxCC_IdleCodes_WS.cell(row=1, column=2, value="Description")
            WxCC_IdleCodes_WS.cell(row=1, column=3, value="Type")

            CCX_WrapCodes_WS = CCX_wb['Wrap Up Codes']
            CCX_Reason_WS = CCX_wb['Reason Codes']
            for r in range(2, CCX_WrapCodes_WS.max_row + 1):
                name = CCX_WrapCodes_WS.cell(row=r, column=1).value
                WxCC_WrapCodes_WS.cell(row=r, column=col, value=name)
                WxCC_WrapCodes_WS.cell(row=r, column=col+1, value=name)
                WxCC_WrapCodes_WS.cell(row=r, column=col+2, value="WRAP_UP_CODE")

            col = 1

            for r in range(2, CCX_Reason_WS.max_row+1):
                name = CCX_Reason_WS.cell(row=r, column=1).value
                WxCC_IdleCodes_WS.cell(row=r, column=col, value=name)
                WxCC_IdleCodes_WS.cell(row=r, column=col+1, value=name)
                WxCC_IdleCodes_WS.cell(row=r, column=col+2, value="IDLE_CODE")
            print("Converting Codes : \033[92mDone\033[0m")
            csq_WxCC()
        except:
            print("An error was encountered while opening the file located at CODES WxCC")
    except:
        print("Program has encountered an error.")


# Loading the contents of the CCX-Details.xlsx sheet. Converting "CSQ" details into WxCC "Queues".
def csq_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        col = 3
        count = 1
        col_List = []
        CCX_List = []
        CCX_Main_List = []
        WxCC_Main_List = []
        WxCC_List = []
        try:
            WxCC_CSQ_WS.cell(row=1, column=1, value="Name")
            WxCC_CSQ_WS.cell(row=1, column=2, value="Description")
            WxCC_CSQ_WS.cell(row=1, column=3, value="Team")

            CCX_CSQ_WS = CCX_wb['CSQ']
            for r in range(2, CCX_CSQ_WS.max_row + 1):
                name = CCX_CSQ_WS.cell(row=r, column=2).value
            heading_Check_CCX = CCX_CSQ_WS[1]
            for cell in heading_Check_CCX:
                if "Skill Name" in str(cell.value):
                    count += 1
                    col_List.append(cell.column)

            col_List_Length = len(col_List)
            i = 1

            for row_idx in range(2, CCX_CSQ_WS.max_row + 1):
                CCX_List = []
                CCX_List.append(CCX_CSQ_WS.cell(row=row_idx, column=2).value)
                for col_idx in col_List:
                    CCX_List.append(CCX_CSQ_WS.cell(row=row_idx, column=col_idx).value)
                CCX_Main_List.append(CCX_List)

            col_SN_List = []
            col_SL_List = []

            heading_Check_WxCC = wb_WxCC['Skill Profile1'][1]
            for cell in heading_Check_WxCC:
                if "Skill Name" in str(cell.value):
                    count += 1
                    col_SN_List.append(cell.column)

            for row_idxx in range(2, wb_WxCC['Skill Profile1'].max_row + 1):
                WxCC_List = []
                WxCC_List.append(wb_WxCC['Skill Profile1'].cell(row=row_idxx, column=2).value)
                for j in range(0, len(col_SN_List),2):
                    WxCC_List.append(wb_WxCC['Skill Profile1'].cell(row=row_idxx, column=col_SN_List[j]).value)
                WxCC_Main_List.append(WxCC_List)

            r = 2

            max_len = max(len(CCX_Main_List), len(WxCC_Main_List))
            CCX_Main_List.extend([[]] * (max_len - len(CCX_Main_List)))  # Extend List to the max length
            WxCC_Main_List.extend([[]] * (max_len - len(WxCC_Main_List)))  # Extend List2 to the max length


            for h in range(0, len(CCX_Main_List)):
                WxCC_CSQ_WS.cell(row=r, column=1, value=CCX_Main_List[h][0])
                WxCC_CSQ_WS.cell(row=r, column=2, value=CCX_Main_List[h][0])
                for value in CCX_Main_List[h]:
                    # print(value)
                    if value != None:
                        for k in range(len(WxCC_Main_List)):
                            if value in WxCC_Main_List[k]:
                                WxCC_CSQ_WS.cell(row=r, column=3, value=WxCC_Main_List[k][0])
                                # print(WxCC_Main_List[k][0])
                                break
                            else:
                                continue

                r += 1
            print("Converting CSQs : \033[92mDone\033[0m")
            pb_WxCC()
        except:
            print("An error was encountered while opening the file located at CSQ WxCC")
    except:
        print("Program has encountered an error.")


# # Loading the contents of the CCX-Details.xlsx sheet. Converting "Phonebook" & its "Contact" details into WxCC "Address Books".
def pb_WxCC():
    try:
        CCX_wb = openpyxl.load_workbook(CCX_Sheet.full_Path_CCX)
        col = 1
        count = 1
        col_List = []
        Ext_List = []
        index = 1
        contact_Dict = {"Name" : "", "Extension" : ""}

        CCX_PB_WS_List = []
        pb_List = []
        contact_List = []
        try:
            sheetname = "Phonebook"
            for idx, sheet in enumerate(CCX_wb.sheetnames):
                if sheetname in sheet:
                    pb_List.append(idx)
                    CCX_PB_WS_List.append(sheet)

            for name in CCX_PB_WS_List:
                col_List = []
                Ext_List = []
                col = 1
                WxCC_PB_WS = wb_WxCC.create_sheet(title=name, index=index+8)
                WxCC_PB_WS.cell(row=1, column=1, value="Phonebook Name")
                WxCC_PB_WS.cell(row=1, column=2, value="Phonebook Description")
                heading_Check = CCX_wb[name][1]
                for cell in heading_Check:
                    if "Contact Name" in str(cell.value):
                        count += 1
                        col_List.append(cell.column)

                for cell in heading_Check:
                    if "Phone Number" in str(cell.value):
                        count += 1
                        Ext_List.append(cell.column)

                for col_idx in range(0, len(col_List)):
                    p_name = CCX_wb[name].cell(row=2, column=1).value
                    WxCC_PB_WS.cell(row=2, column=1, value=p_name)
                    WxCC_PB_WS.cell(row=2, column=2, value=p_name)
                    for r in range(2, CCX_wb[name].max_row + 1):
                        col += 2
                        WxCC_PB_WS.cell(row=1, column=col, value=f"Contact Name {col_idx + 1}")
                        WxCC_PB_WS.cell(row=r, column=col, value=CCX_wb[name].cell(row=r, column=col_List[col_idx]).value)
                        WxCC_PB_WS.cell(row=1, column=col + 1, value=f"Extension {col_idx + 1}")
                        WxCC_PB_WS.cell(row=r, column=col + 1, value=CCX_wb[name].cell(row=r, column=col_List[col_idx] + 2).value)

                index+=1

            col = 1
            print("Converting Phonebooks : \033[92mDone\033[0m")
            create_WxCC_File()
        except:
            print("An error was encountered while opening the file located at PHONEBOOK WxCC")
    except:
        print("Program has encountered an error.")


# Storing all the transformed data into a new file - "WxCC Details.xlsx"
def create_WxCC_File():
    global full_Path_WxCC
    try:
        name = "WxCC Details.xlsx"
        os_Type = platform.system()
        if os_Type == "Windows":
            current_dir = os.getcwd()
            base_Path_WxCC = os.path.join(current_dir, "Files")
        elif os_Type == "Linux":
            current_dir = os.getcwd()
            base_Path_WxCC = os.path.join(current_dir, "Files")

        if not os.path.exists(base_Path_WxCC):
            os.makedirs(base_Path_WxCC)
            print(f"Created folder: {base_Path_WxCC}")

        full_Path_WxCC = os.path.join(base_Path_WxCC, name)

        wb_WxCC.save(full_Path_WxCC)
        print(f"\033[92mUCCX configuration data has been successfully translated into WxCC format and stored "
              f"in a file {name} at {base_Path_WxCC}.\033[0m")
        print("\033[93mNow proceeding with adding the translated data into WxCC.....\033[0m")
        WxCC.create_Skill_Profile()
    except:
        print(f"\033[91mThe UCCX Config data was converted into WxCC compatible format successfully but an error "
              f"was encountered while saving the file at {base_Path_WxCC}. Please ensure that you have appropriate read/write"
              f"permissions in this directory. \033[0m")