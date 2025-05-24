import json, requests, openpyxl, re, os, platform, WxCC_Sheet
import time

from Client_OAuth import get_Access_Token


org_ID = os.getenv("ORG_ID")
instance_Name = os.getenv("WxCC_INSTANCE")

# Initializing variables with API endpoints
EP_URL= f"https://{instance_Name}/organization/{org_ID}/entry-point"
TEAM_URL = f"https://{instance_Name}/organization/{org_ID}/team"
CSQ_URL = f"https://{instance_Name}/organization/{org_ID}/contact-service-queue"
MOH_URL = f"https://{instance_Name}/organization/{org_ID}/v2/audio-file"
SKILL_URL = f"https://{instance_Name}/organization/{org_ID}/skill"
SKILL_PROFILE_URL = f"https://{instance_Name}/organization/{org_ID}/skill-profile"
SKILL_PROFILE_URL_2 = f"https://{instance_Name}/organization/{org_ID}/v2/skill-profile?search="
SITE_URL = f"https://{instance_Name}/organization/{org_ID}/v2/site"
CODE_URL = f"https://{instance_Name}/organization/{org_ID}/auxiliary-code"
A_BOOK_URL = f"https://{instance_Name}/organization/{org_ID}/v3/address-book"
A_CONTACT_URL = f"https://{instance_Name}/organization/{org_ID}/address-book/"

# Storing token and refresh token values
token, refresh = get_Access_Token()


# Reading the contents of the newly generated WxCC Details.xlsx sheet.
def read_Sheet():
    col_PB_List = []
    ext_PB_List = []
    col_Skill_List = []
    csq_Final_List = []
    idle_Final_List = []
    wrapUp_Final_List = []
    team_Final_List = []
    skill_List = []
    skillProfile_Final_List = []
    ep_Final_List = []
    WxCC_PB_WS_List = []
    pb_Final_List = []
    sheetname = "Phonebook"
    col = 1
    try:
        wb = openpyxl.load_workbook(WxCC_Sheet.full_Path_WxCC)
        ws_EP = wb["Entry Points"]
        ws_CSQ = wb["CSQ"]
        ws_Teams = wb["Teams"]
        ws_Skills = wb["Skills"]
        ws_SkillProfile = wb['Skill Profile1']
        ws_WrapUp = wb['Wrap up Codes']
        ws_Idle = wb['Idle Codes']

        row_EP = ws_EP.max_row
        row_Skills = ws_Skills.max_row
        row_SkillProfile = ws_SkillProfile.max_row
        row_Teams = ws_Teams.max_row
        row_WrapUp = ws_WrapUp.max_row
        row_Idle = ws_Idle.max_row
        row_CSQ = ws_CSQ.max_row
        # row_Address = ws_Address.max_row


        for idx, sheet in enumerate(wb.sheetnames):
            if sheetname in sheet:
                WxCC_PB_WS_List.append(sheet)

        for name in WxCC_PB_WS_List:
            pb_List = []
            heading_Check = wb[name][1]
            for cell in heading_Check:
                if "Contact Name" in str(cell.value):
                    col_PB_List.append(cell.column)

            for cell in heading_Check:
                if "Extension" in str(cell.value):
                    ext_PB_List.append(cell.column)

            for r in range(2, wb[name].max_row + 1):
                p_name = wb[name].cell(row=2, column=1).value
                pb_List.append(p_name)
                for cl in range(len(col_PB_List)):
                    col += 2
                    c_Name = wb[name].cell(row=r, column=col).value
                    c_Ext = wb[name].cell(row=r, column=col + 1).value
                    pb_List.append(c_Name)
                    pb_List.append(c_Ext)

            pb_Final_List.append(pb_List)

        # Storing Skill values
        for s in range(2, row_Skills + 1):
            skill_List.append(ws_Skills.cell(row=s, column=2).value)

        # Storing Skill Profiles
        heading_Check = ws_SkillProfile[1]
        for cell in heading_Check:
            if "Skill Name" in str(cell.value):
                col_Skill_List.append(cell.column)

        for sp in range(2, row_SkillProfile + 1):
            skillProfile_List = []
            skillProfile_List.append(ws_SkillProfile.cell(row=sp, column=1).value)
            for col_idx in col_Skill_List:
                if ws_SkillProfile.cell(row=sp, column=col_idx).value != None:
                    skillProfile_List.append(ws_SkillProfile.cell(row=sp, column=col_idx).value)
            skillProfile_Final_List.append(skillProfile_List)

        # Storing Teams
        for t in range(2, row_Teams + 1):
            team_List = []
            team_List.append(ws_Teams.cell(row=t, column=1).value)
            team_List.append(ws_Teams.cell(row=t, column=3).value)
            team_Final_List.append(team_List)

        # Storing Wrap-Up Codes
        for w in range(2, row_WrapUp + 1):
            wrapUp_List = []
            wrapUp_List.append(ws_WrapUp.cell(row=w, column=1).value)
            wrapUp_List.append(ws_WrapUp.cell(row=w, column=3).value)
            wrapUp_Final_List.append(wrapUp_List)

        # Storing Idle Codes
        for i in range(2, row_Idle + 1):
            idle_List = []
            idle_List.append(ws_Idle.cell(row=i, column=1).value)
            idle_List.append(ws_Idle.cell(row=i, column=3).value)
            idle_Final_List.append(idle_List)

        # Storing CSQ values
        for c in range(2, row_CSQ + 1):
            csq_List = []
            csq_List.append(ws_CSQ.cell(row=c, column=1).value)
            csq_List.append(ws_CSQ.cell(row=c, column=3).value)
            csq_Final_List.append(csq_List)

        # Storing Entry Point values
        for e in range(2, row_EP + 1):
            ep_List = []
            ep_Name = ws_EP.cell(row=e, column=1).value
            ep_Desc = ws_EP.cell(row=e, column=2).value
            ep_Session = ws_EP.cell(row=e, column=3).value
            ep_List.append(ep_Name)
            ep_List.append(ep_Desc)
            ep_List.append(ep_Session)
            ep_Final_List.append(ep_List)
        return pb_Final_List, skill_List, skillProfile_Final_List, team_Final_List, wrapUp_Final_List, idle_Final_List, csq_Final_List\
           , ep_Final_List
    except:
        print("\033[91mError encountered while reading WxCC Excel File. Halting the program.\033[0m")


# Creating Skills in WxCC
def create_Skill():
    global token, refresh
    _,skill,_,_,_,_,_,_ = read_Sheet()
    skill_Final = []
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }
    try:
        for sk_idx in range (0, len(skill)):
            skill_Resp = []
            data = {
            "name": skill[sk_idx],
            "description": skill[sk_idx],
            "serviceLevelThreshold": 0,
            "active": True,
            "skillType": "PROFICIENCY"
            }

            try:
                resp = requests.post(SKILL_URL, data=json.dumps(data), headers=headers)
                resp_json = json.loads(resp.content)
                skill_Resp.append(resp_json['id'])
                skill_Resp.append(resp_json['name'])
                skill_Final.append(skill_Resp)
                print(f"Adding Skills [{resp_json['name']}] : \033[92mDone\033[0m")
            except:
                print("\033[91mSkills API Request to WxCC has encountered an error.\033[0m")
        return skill_Final
    except:
        print("\033[91mFailed to run through the skills from WxCC Sheet. Halting the program.\033[0m")


# Creating Skill Profile in WxCC
def create_Skill_Profile():
    global token, refresh
    _,_,skill_Profile,_,_,_,_,_ = read_Sheet()
    skill_Final = create_Skill()
    active_List = []
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }
    print("*" * 60)

    try:
        for value in skill_Profile:
            active_List = []
            v = 2
            s = 1
            # print(len(value))
            # print(value)

            for v_idx in range(0, len(value)):
                if s < len(value) and v < len(value):
                    for sk in skill_Final:
                        s_Name = sk[1]
                        if s_Name == f"{value[s]}":
                            skillID = sk[0]
                            # print(skillID)
                            active_Dict = {
                                "booleanValue": True,
                                "skillId": skillID,
                                "proficiencyValue": value[v]
                            }
                            # print(active_Dict)
                            active_List.append(active_Dict)
                            data = {
                            "name": f"{value[0]}",
                            "description": f"{value[0]}",
                            "activeSkills": active_List
                            }
                    s += 2
                    v += 2
                else:
                    break
            try:
                resp = requests.post(SKILL_PROFILE_URL, data=json.dumps(data), headers=headers)
                # print(f"Skill Profile {v_idx} created successfully : Status Code {resp.status_code}")
                print(f"Adding Skill Profiles [{data['name']}] : \033[92mDone\033[0m")

            except:
                print("\033[91mSkill Profile API Request to WxCC has encountered an error.\033[0m")
        create_teams()
    except:
        print("\033[91mFailed to run through the skills and skill profile1 from WxCC Sheet. Halting the program\033[0m")


# Creating Teams in WxCC
def create_teams():
    skill_ID = None
    _,_,_,teams,_,_,_,_ = read_Sheet()
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    print("*" * 60)

    try:
        resp_Site = requests.get(SITE_URL, headers=headers)
        resp_Site_json = json.loads(resp_Site.text)
        site_ID = resp_Site_json["data"][0]["id"]
    except:
        print("\033[91mSites API Request to WxCC has encountered an error from within TEAMS\033[0m")

    try:
        for tm_idx in range (0, len(teams)):
            if teams[tm_idx][1] != None:
                resp_SkillID = requests.get(f"{SKILL_PROFILE_URL_2}{teams[tm_idx][1]}", headers=headers)
                skill_ID = resp_SkillID.json()["data"][0]["id"]
            data = {
            "name": teams[tm_idx][0],
            "active": True,
            "siteId": site_ID,
            "teamStatus": "IN_SERVICE",
            "teamType": "AGENT",
            "skillProfileId": skill_ID
            }
            try:
                resp = requests.post(TEAM_URL, data=json.dumps(data), headers=headers)
                print(f"Adding Teams [{data['name']}] : \033[92mDone\033[0m")
            except:
                print("\033[91mTeams API Request to WxCC has encountered an error.\033[0m")
        create_Codes()
    except:
        print("\033[91mFailed to run through the Team List from WxCC Sheet. Halting the program\033[0m")


# Creating Aux Codes in WxCC
def create_Codes():
    _,_,_,_,wrapUp,idle,_,_ = read_Sheet()
    active = True
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    print("*" * 60)

    try:
        while active:
            for wr_idx in range(0, len(wrapUp)):
                data = {
                  "name": wrapUp[wr_idx][0],
                  "description": wrapUp[wr_idx][0],
                  "defaultCode": True,
                  "active": True,
                  "isSystemCode": False,
                  "workTypeId": "bda53bf0-2e0e-46bf-a1af-9fd375db7158",
                  "workTypeCode": "WRAP_UP_CODE"
                }
                try:
                    resp = requests.post(CODE_URL, data=json.dumps(data), headers=headers)
                    print(f"Adding Wrap-Up Codes : \033[93mIn Progress.....\033[0m")
                    time.sleep(0.2)
                except:
                    print("\033[91mWrap-Up Code API Request to WxCC has encountered an error.\033[0m")
            active = False
        print(f"Adding Wrap-Up Codes : \033[92mDone\033[0m")

        print("*" * 60)
        active = True

        while active:
            for id_idx in range(0, len(idle)):
                data = {
                    "name": idle[id_idx][0],
                    "description": idle[id_idx][0],
                    "defaultCode": True,
                    "active": True,
                    "isSystemCode": False,
                    "workTypeId": "bda53bf0-2e0e-46bf-a1af-9fd375db7158",
                    "workTypeCode": "IDLE_CODE"
                }
                try:
                    resp = requests.post(CODE_URL, data=json.dumps(data), headers=headers)
                    print(f"Adding Idle Codes : \033[93mIn Progress.....\033[0m")
                    time.sleep(0.2)
                except:
                    print("\033[91mIdle Code API Request to WxCC has encountered an error.\033[0m")

            active = False
        print(f"Adding Idle Codes : \033[92mDone\033[0m")
        create_CSQ()
    except:
        print("\033[91mFailed to run through the Code List from WxCC Sheet. Halting the program\033[0m")


# Creating Queues in WxCC
def create_CSQ():
    _,_,_,_,_,_,csq_List,_ = read_Sheet()
    team_CSQ_List = []
    music_List = []
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    print("*" * 60)

    try:
        try:
            resp_Music = requests.get(MOH_URL, headers=headers)
            resp_Music_json = json.loads(resp_Music.text)
            for music in resp_Music_json["data"]:
                if "default" in music["name"]:
                    music_ID = music["id"]
                    music_List.append(music_ID)
        except:
            print("API Request to WxCC has encountered an error MoH.")

        try:
            resp_Team = requests.get(TEAM_URL, headers=headers)
            resp_Team_json = json.loads(resp_Team.text)
            for value in resp_Team_json:
                team_List = []
                team_List.append(value['id'])
                team_List.append(value['name'])
                team_CSQ_List.append(team_List)

            for csq_idx in range(0, len(csq_List)):
                for value in team_CSQ_List:
                    if csq_List[csq_idx][1] == value[1]:
                        team_ID = value[0]
                        data = {
                            "name": csq_List[csq_idx][0],
                            "description": f"{csq_List[csq_idx][0]}",
                            "queueType": "INBOUND",
                            "checkAgentAvailability": False,
                            "channelType": "TELEPHONY",
                            "serviceLevelThreshold": 1,
                            "maxActiveContacts": 0,
                            "maxTimeInQueue": 60,
                            "defaultMusicInQueueMediaFileId": music_List[0],
                            "active": True,
                            "monitoringPermitted": False,
                            "parkingPermitted": False,
                            "recordingPermitted": False,
                            "recordingAllCallsPermitted": False,
                            "pauseRecordingPermitted": False,
                            "recordingPauseDuration": 10,
                            "controlFlowScriptUrl": "http://localhost:9000",
                            "ivrRequeueUrl": "http://localhost:9000",
                            "routingType": "LONGEST_AVAILABLE_AGENT",
                            "queueRoutingType": "TEAM_BASED",
                            "callDistributionGroups": [
                                {
                                    "agentGroups": [
                                        {
                                            "teamId": team_ID
                                        }
                                    ],
                                    "order": 1,
                                    "duration": 0
                                }
                            ],
                        }
                        try:
                            resp = requests.post(CSQ_URL, data=json.dumps(data), headers=headers)
                            print(f"Adding Queues [{data['name']}] : \033[92mDone\033[0m")
                        except:
                            print("\033[91mQueues API Request to WxCC has encountered an error. Halting the program\033[0m")
            create_EP()
        except:
            print("API Request to WxCC has encountered an error TEAMS.")
    except:
        print("\033[91mFailed to run through the CSQ List from WxCC Sheet. Halting the program\033[0m")


# Creating Entry Points in WxCC
def create_EP():
    _,_,_,_,_,_,_,eps = read_Sheet()
    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    print("*" * 60)

    try:
        for ep_idx in range(0, len(eps)):
            data = {
                "name": eps[ep_idx][0],
                "description": eps[ep_idx][0],
                "entryPointType": "INBOUND",
                "channelType": "TELEPHONY",
                "active": True,
                "serviceLevelThreshold": 12,
                "maximumActiveContacts": eps[ep_idx][2],
            }
            try:
                resp = requests.post(EP_URL, data=json.dumps(data), headers=headers)
                print(f"Adding Entry Points [{data['name']}] : \033[92mDone\033[0m")
            except:
                print("\033[91mEntry Point API Request to WxCC has encountered an error. Halting the program\033[0m")
        create_Add_Book()
    except:
        print("\033[91mFailed to run through the Entry Points List from WxCC Sheet. Halting the program\033[0m")


# Creating Address Books & adding Contacts in WxCC
def create_Add_Book():
    try:
        addbook,_,_,_,_,_,_,_ = read_Sheet()
        active = True
        headers = {
            "Accept": "*/*",
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
        }

        print("*" * 60)

        pb_name = addbook[0][0]
        c_name = 1
        c_number = 2
        data_AB = {
            "name": pb_name,
            "description": pb_name,
            "parentType": "ORGANIZATION",
        }
        try:
            resp_Add_Book = requests.post(A_BOOK_URL, json.dumps(data_AB), headers=headers)
            resp_Add_Book_json = json.loads(resp_Add_Book.text)
            A_Book_ID = resp_Add_Book_json["id"]
            print(f"Adding Address Books [{data_AB['name']}] : \033[92mDone\033[0m")

            while active:
                for add_idx in range(len(addbook[0])):
                    if c_name < len(addbook[0]) and c_number < len(addbook[0]):
                        data_Contacts = {
                            "name": addbook[0][c_name],
                            "number": addbook[0][c_number],
                            "parentType": "ORGANIZATION",
                        }
                        try:
                            resp_Contacts = requests.post(f"{A_CONTACT_URL}{A_Book_ID}/entry", json.dumps(data_Contacts),
                                                          headers=headers)
                            print(f"Adding Contacts : \033[93mIn Progress.....\033[0m")
                            time.sleep(0.2)
                            c_name += 2
                            c_number += 2
                        except:
                            print(
                                "\033[91mAddress Book Contacts API Request to WxCC has encountered an error. Halting the program\033[0m")
                    else:
                        break
                active = False
            print(f"Adding Contacts : \033[92mDone\033[0m")
        except:
            print("\033[91mAddress Book API Request to WxCC has encountered an error. Halting the program\033[0m")
    except:
        print("\033[91mFailed to run through the Phonebook List from WxCC Sheet. Halting the program\033[0m")