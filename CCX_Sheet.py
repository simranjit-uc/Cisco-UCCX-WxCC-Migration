import platform, requests, os, json, openpyxl, xmltodict, WxCC_Sheet
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning


instance_Name = os.getenv("CCX_INSTANCE")
ccx_Token = os.getenv("CCX_TOKEN")
token = os.getenv("CCX_TOKEN")

app_URL = f"https://{instance_Name}/adminapi/application"
trigger_URL = f"https://{instance_Name}/adminapi/trigger"
resource_URL = f"https://{instance_Name}/adminapi/resource"
csq_URL = f"https://{instance_Name}/adminapi/csq"
team_URL = f"https://{instance_Name}/adminapi/team"
ccg_URL = f"https://{instance_Name}/adminapi/callControlGroup"
skill_URL = f"https://{instance_Name}/adminapi/skill"
wrapup_URL = f"https://{instance_Name}:8445/finesse/api/WrapUpReasons"
reason_URL = f"https://{instance_Name}:8445/finesse/api/ReasonCodes?category=ALL"
phonebook_URL = f"https://{instance_Name}:8445/finesse/api/PhoneBooks"
contact_URL = ""


full_Path_CCX = ""

# Creating an Excel workbook and initializing it with multiple sheets for different config parameters.
wb = openpyxl.Workbook()
app_WS = wb.create_sheet(title="Applications", index=0)
trigger_WS = wb.create_sheet(title="Triggers", index=1)
resource_WS = wb.create_sheet(title="Resources", index=2)
csq_WS = wb.create_sheet(title="CSQ", index=3)
team_WS = wb.create_sheet(title="Teams", index=4)
skill_WS = wb.create_sheet(title="Skills", index=5)
wrapup_WS = wb.create_sheet(title="Wrap Up Codes", index=6)
reason_WS = wb.create_sheet(title="Reason Codes", index=7)


headers = {
    "Accept": "*/*",
    "Authorization": f"Basic {token}"
}


# Getting Application details
def get_APP():
    try:
        count = 1
        app_WS.cell(row=1, column=1, value="Application Name")
        app_WS.cell(row=1, column=2, value="Description")
        app_WS.cell(row=1, column=3, value="Script Name")
        app_WS.cell(row=1, column=4, value="Status")
        app_WS.cell(row=1, column=5, value="Maximum Sessions")

        try:
            disable_warnings(InsecureRequestWarning)
            resp = requests.get(app_URL,headers=headers, verify=False)
            data_dict = xmltodict.parse(resp.content)

            # Checking if there are any Applications in CCX
            if data_dict['applications'] is not None:
                resp_Type = type(data_dict['applications']['application'])
                # Checking if the Application API response from CCX contains one or more applications.
                if resp_Type == dict:
                    app_Name = data_dict["applications"]["application"]["applicationName"]
                    description = data_dict["applications"]["application"]["description"]
                    script_Name = data_dict["applications"]["application"]["ScriptApplication"]["script"]
                    status = data_dict["applications"]["application"]["enabled"]
                    max_Session = data_dict["applications"]["application"]["maxsession"]

                    app_WS.cell(row=2, column=1, value=app_Name)
                    app_WS.cell(row=2, column=2, value=description)
                    app_WS.cell(row=2, column=3, value=script_Name)
                    app_WS.cell(row=2, column=4, value=status)
                    app_WS.cell(row=2, column=5, value=max_Session)
                else:
                    for i in range(0, len(data_dict["applications"]["application"])):
                        app_Name = data_dict["applications"]["application"][i]["applicationName"]
                        description = data_dict["applications"]["application"][i]["description"]
                        script_Name = data_dict["applications"]["application"][i]["ScriptApplication"]["script"]
                        status = data_dict["applications"]["application"][i]["enabled"]
                        max_Session = data_dict["applications"]["application"][i]["maxsession"]
                        count += 1

                        app_WS.cell(row=count, column=1, value=app_Name)
                        app_WS.cell(row=count, column=2, value=description)
                        app_WS.cell(row=count, column=3, value=script_Name)
                        app_WS.cell(row=count, column=4, value=status)
                        app_WS.cell(row=count, column=5, value=max_Session)

                    count = 1
            else:
                print("No Applications")
            print(f"Application Details : \033[92mDone\033[0m")
            get_Trigger()
        except:
            print("\033[91mApplication API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



# Getting Call Control Group details from CCX
def get_CCG():
    ccg_List = [{"ID": "", "Type": ""}]
    count = 1
    try:
        disable_warnings(InsecureRequestWarning)
        resp = requests.get(ccg_URL,headers=headers, verify=False)
        data_dict = xmltodict.parse(resp.content)
        json_data = json.dumps(data_dict)
        json_loads = json.loads(json_data)

        # Checking if there are any Call Control Groups in CCX
        if data_dict['callControlGroups'] is not None:
            resp_Type = type(data_dict['callControlGroups']['callControlGroup'])
            # Checking if the CCG API response from CCX contains one or more CCGs.
            if resp_Type == dict:
                id = str(data_dict["callControlGroups"]["callControlGroup"]["id"])
                ccg_Type = data_dict["callControlGroups"]["callControlGroup"]["outboundGroup"]
                description = data_dict["callControlGroups"]["callControlGroup"]["description"]
                if ccg_Type == "false":
                    ccg_Type = "Inbound"
                else:
                    ccg_Type = "Outbound"
                    ccg_List.append({"ID": id, "Type": ccg_Type})

                ccg_List.pop(0)
                return ccg_List
            else:
                for i in range(0,len(data_dict["callControlGroups"]["callControlGroup"])):
                    id = str(data_dict["callControlGroups"]["callControlGroup"][i]["id"])
                    ccg_Type = data_dict["callControlGroups"]["callControlGroup"][i]["outboundGroup"]
                    description = data_dict["callControlGroups"]["callControlGroup"][i]["description"]
                    count += 1
                    if ccg_Type == "false":
                        ccg_Type = "Inbound"
                    else:
                        ccg_Type = "Outbound"

                    ccg_List.append({"ID" : id, "Type" : ccg_Type})
                ccg_List.pop(0)
                return ccg_List
        else:
            print("No CCG")
    except:
        print("\033[91mCall Control Group API Request to UCCX has encountered an error.\033[0m")



# Getting Trigger details from CCX
def get_Trigger():
    try:
        trigger_WS.cell(row=1, column=1, value="Application Name")
        trigger_WS.cell(row=1, column=2, value="Trigger Number")
        trigger_WS.cell(row=1, column=3, value="Description")
        trigger_WS.cell(row=1, column=4, value="Type")
        trigger_WS.cell(row=1, column=5, value="Call Control Group")
        trigger_WS.cell(row=1, column=6, value="Status")

        get_CCG()
        count = 1

        try:
            disable_warnings(InsecureRequestWarning)
            resp = requests.get(trigger_URL,headers=headers, verify=False)
            data_dict = xmltodict.parse(resp.content)

            # Checking if there are any Triggers in CCX
            if data_dict['triggers'] is not None:
                resp_Type = type(data_dict['triggers']['trigger'])
                # Checking if the Trigger API response from CCX contains one or more Triggers.
                if resp_Type == dict:
                    app_Name = data_dict['triggers']['trigger']['application']['@name']
                    trigger_Number = data_dict['triggers']['trigger']['directoryNumber']
                    description = data_dict['triggers']['trigger']['description']
                    # trigger_Type = data_dict['triggers']['trigger']['deviceName']
                    trigger_Type = data_dict['triggers']['trigger']['callControlGroup']['@name']
                    trigger_CCG = data_dict['triggers']['trigger']['callControlGroup']['@name']
                    status = data_dict['triggers']['trigger']['triggerEnabled']

                    trigger_WS.cell(row=2, column=1, value=app_Name)
                    trigger_WS.cell(row=2, column=2, value=trigger_Number)
                    trigger_WS.cell(row=2, column=3, value=description)
                    trigger_WS.cell(row=2, column=4, value=trigger_Type)
                    trigger_WS.cell(row=2, column=6, value=status)
                    for ccg in get_CCG():
                        if trigger_CCG == ccg['ID']:
                            trigger_WS.cell(row=2, column=5, value=ccg['Type'])
                            break
                else:
                    for i in range(0, len(data_dict['triggers']['trigger'])):
                        app_Name = data_dict['triggers']['trigger'][i]['application']['@name']
                        trigger_Number = data_dict['triggers']['trigger'][i]['directoryNumber']
                        description = data_dict['triggers']['trigger'][i]['description']
                        trigger_Type = data_dict['triggers']['trigger'][i]['callControlGroup']['@name']
                        trigger_CCG = data_dict['triggers']['trigger'][i]['callControlGroup']['@name']
                        status = data_dict['triggers']['trigger'][i]['triggerEnabled']

                        count += 1

                        trigger_WS.cell(row=count, column=1, value=app_Name)
                        trigger_WS.cell(row=count, column=2, value=trigger_Number)
                        trigger_WS.cell(row=count, column=3, value=description)
                        trigger_WS.cell(row=count, column=4, value=trigger_Type)
                        trigger_WS.cell(row=count, column=6, value=status)
                        for ccg in get_CCG():
                            if trigger_CCG == ccg['ID']:
                                trigger_WS.cell(row=count, column=5, value=ccg['Type'])
                                break
                    count = 1
            else:
                print("No Trigger")
            print(f"Trigger Details : \033[92mDone\033[0m")
            get_Resource()
        except:
            print("\033[91mTrigger API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



# Getting Resource details from CCX
def get_Resource():
    try:
        count = 1
        num = 8
        skill_List = []
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(resource_URL, headers=headers, verify=False)
            resource_WS.cell(row=1, column=1, value="Resource ID")
            resource_WS.cell(row=1, column=2, value="First Name")
            resource_WS.cell(row=1, column=3, value="Last Name")
            resource_WS.cell(row=1, column=4, value="Extension")
            resource_WS.cell(row=1, column=5, value="Team Name")
            resource_WS.cell(row=1, column=6, value="Is Primary Supervisor")
            resource_WS.cell(row=1, column=7, value="Is Secondary Supervisor")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)

            # Checking if there are any Resources in CCX
            if data_dict['resources'] is not None:
                resp_Type = type(data_dict['resources']['resource'])
                # Checking if the Resource API response from CCX contains one or more Resources.
                if resp_Type == dict:
                    num = 8
                    res_ID = data_dict['resources']['resource']['userID']
                    res_First_Name = data_dict['resources']['resource']['firstName']
                    res_Last_Name = data_dict['resources']['resource']['lastName']
                    res_Ext = data_dict['resources']['resource']['extension']
                    res_Team = data_dict['resources']['resource']['team']['@name']
                    resource_WS.cell(row=2, column=1, value=res_ID)
                    resource_WS.cell(row=2, column=2, value=res_First_Name)
                    resource_WS.cell(row=2, column=3, value=res_Last_Name)
                    resource_WS.cell(row=2, column=4, value=res_Ext)
                    resource_WS.cell(row=2, column=5, value=res_Team)

                    # Checking if the Resource API response from CCX contains one or more skills associated with resources.
                    if type(data_dict['resources']['resource']['skillMap']['skillCompetency']) == dict:
                        res_Skill = \
                        data_dict['resources']['resource']['skillMap']['skillCompetency']['skillNameUriPair']['@name']
                        res_Skill_Level = data_dict['resources']['resource']['skillMap']['skillCompetency'][
                            'competencelevel']
                    else:
                        skill_Count = len(data_dict['resources']['resource']['skillMap']['skillCompetency'])
                        for b in range(0, skill_Count):
                            res_Skill = \
                            data_dict['resources']['resource']['skillMap']['skillCompetency'][b]['skillNameUriPair'][
                                '@name']
                            res_Skill_Level = data_dict['resources']['resource']['skillMap']['skillCompetency'][b][
                                'competencelevel']
                            resource_WS.cell(row=1, column=num, value=f"Skill Name  {b + 1}")
                            resource_WS.cell(row=count, column=num, value=f"{res_Skill}")
                            num += 1
                            resource_WS.cell(row=1, column=num, value=f"Skill Level  {b + 1}")
                            resource_WS.cell(row=count, column=num, value=res_Skill_Level)
                            num += 1
                else:
                    for i in range(0, len(data_dict['resources']['resource'])):
                        res_ID = data_dict['resources']['resource'][i]['userID']
                        res_First_Name = data_dict['resources']['resource'][i]['firstName']
                        res_Last_Name = data_dict['resources']['resource'][i]['lastName']
                        res_Ext = data_dict['resources']['resource'][i]['extension']
                        res_Team = data_dict['resources']['resource'][i]['team']['@name']
                        # print(f"{res_ID}, {res_First_Name}")
                        if data_dict['resources']['resource'][i]['skillMap'] == None:
                            res_Skill = "No Skills are assigned"
                            res_Skill_Level = "0"
                        else:
                            num = 8
                            if type(data_dict['resources']['resource'][i]['skillMap']['skillCompetency']) == dict:
                                res_Skill = data_dict['resources']['resource'][i]['skillMap']['skillCompetency'][
                                    'skillNameUriPair']['@name']
                                res_Skill_Level = data_dict['resources']['resource'][i]['skillMap']['skillCompetency'][
                                    'competencelevel']
                                resource_WS.cell(row=1, column=num, value=f"Skill Name 1")
                                resource_WS.cell(row=count + 1, column=num, value=f"{res_Skill}")
                                num += 1
                                resource_WS.cell(row=1, column=num, value=f"Skill Level 1")
                                resource_WS.cell(row=count + 1, column=num, value=res_Skill_Level)
                                num += 1
                            else:
                                skill_Count = len(data_dict['resources']['resource'][i]['skillMap']['skillCompetency'])
                                # num = 11
                                for b in range(0, skill_Count):
                                    res_Skill = data_dict['resources']['resource'][i]['skillMap']['skillCompetency'][b][
                                        'skillNameUriPair']['@name']
                                    res_Skill_Level = \
                                    data_dict['resources']['resource'][i]['skillMap']['skillCompetency'][b][
                                        'competencelevel']

                                    resource_WS.cell(row=1, column=num, value=f"Skill Name  {b + 1}")
                                    resource_WS.cell(row=count + 1, column=num, value=f"{res_Skill}")
                                    num += 1
                                    resource_WS.cell(row=1, column=num, value=f"Skill Level  {b + 1}")
                                    resource_WS.cell(row=count + 1, column=num, value=res_Skill_Level)
                                    num += 1

                        count += 1

                        resource_WS.cell(row=count, column=1, value=res_ID)
                        resource_WS.cell(row=count, column=2, value=res_First_Name)
                        resource_WS.cell(row=count, column=3, value=res_Last_Name)
                        resource_WS.cell(row=count, column=4, value=res_Ext)
                        resource_WS.cell(row=count, column=5, value=res_Team)
                    count = 1
            else:
                print("No Resources")
            print(f"Resource Details : \033[92mDone\033[0m")
            get_CSQ()
        except:
            print("\033[91mResource API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



# Getting CSQ details from CCX
def get_CSQ():
    try:
        count = 1
        csq_Skill = ""
        csq_Skill_Level = ""
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(csq_URL, headers=headers, verify=False)
            csq_WS.cell(row=1, column=1, value="CSQ ID ID")
            csq_WS.cell(row=1, column=2, value="CSQ Name")
            csq_WS.cell(row=1, column=3, value="CSQ Type")
            csq_WS.cell(row=1, column=4, value="Routing Type")
            csq_WS.cell(row=1, column=5, value="Algorithm")
            csq_WS.cell(row=1, column=6, value="Auto Work")
            csq_WS.cell(row=1, column=7, value="Wrap up Time")
            csq_WS.cell(row=1, column=8, value="Resource Pool")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT

            # Checking if there are any CSQs in CCX
            if data_dict['csqs'] is not None:
                resp_Type = type(data_dict['csqs']['csq'])
                # Checking if the CSQ API response from CCX contains one or more CSQs.
                if resp_Type == dict:
                    csq_ID = data_dict['csqs']['csq']['id']
                    csq_Name = data_dict['csqs']['csq']['name']
                    csq_Type = data_dict['csqs']['csq']['queueType']
                    csq_Routing = data_dict['csqs']['csq']['routingType']
                    csq_Algo = data_dict['csqs']['csq']['queueAlgorithm']
                    csq_Pool_Type = data_dict['csqs']['csq']['resourcePoolType']

                    # Checking if the CSQ API response from CCX contains one or more skills associated with it.
                    if type(data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency']) == dict:
                        csq_Skill = data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency'][
                            'skillNameUriPair']['@name']
                        csq_Skill_Level = data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency'][
                            'competencelevel']
                    else:
                        skill_Count = len(data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency'])
                        for b in range(0, skill_Count):
                            csq_Skill = \
                            data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency'][b][
                                'skillNameUriPair']['@name']
                            csq_Skill_Level = \
                            data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['skillCompetency'][b][
                                'competencelevel']
                            csq_Skill += ", "
                            csq_Skill_Level += ", "

                    csq_Selection = data_dict['csqs']['csq']['poolSpecificInfo']['skillGroup']['selectionCriteria']

                    csq_WS.cell(row=2, column=1, value=csq_ID)
                    csq_WS.cell(row=2, column=2, value=csq_Name)
                    csq_WS.cell(row=2, column=3, value=csq_Type)
                    csq_WS.cell(row=2, column=4, value=csq_Routing)
                    csq_WS.cell(row=2, column=5, value=csq_Algo)
                    csq_WS.cell(row=2, column=8, value=csq_Pool_Type)
                    csq_WS.cell(row=2, column=9, value=csq_Skill)
                    csq_WS.cell(row=2, column=10, value=csq_Skill_Level)
                    csq_WS.cell(row=2, column=11, value=csq_Selection)
                else:
                    for i in range(0, len(data_dict['csqs']['csq'])):
                        csq_ID = data_dict['csqs']['csq'][i]['id']
                        csq_Name = data_dict['csqs']['csq'][i]['name']
                        csq_Type = data_dict['csqs']['csq'][i]['queueType']
                        csq_Routing = data_dict['csqs']['csq'][i]['routingType']
                        csq_Algo = data_dict['csqs']['csq'][i]['queueAlgorithm']
                        csq_Pool_Type = data_dict['csqs']['csq'][i]['resourcePoolType']

                        if "skillCompetency" not in data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']:
                            csq_Skill = "No Skills are assigned"
                            csq_Skill_Level = "0"
                        else:
                            num = 9
                            if type(data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup'][
                                        'skillCompetency']) == dict:
                                csq_Skill = \
                                data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']['skillCompetency'][
                                    'skillNameUriPair']['@name']
                                csq_Skill_Level = \
                                data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']['skillCompetency'][
                                    'competencelevel']
                                csq_WS.cell(row=1, column=num, value=f"Skill Name 1")
                                csq_WS.cell(row=count + 1, column=num, value=f"{csq_Skill}")
                                num += 1
                                csq_WS.cell(row=1, column=num, value=f"Skill Level 1")
                                csq_WS.cell(row=count + 1, column=num, value=csq_Skill_Level)
                                num += 1
                            else:
                                skill_Count = len(
                                    data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']['skillCompetency'])
                                for b in range(0, skill_Count):
                                    csq_Skill = \
                                    data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']['skillCompetency'][b][
                                        'skillNameUriPair']['@name']
                                    csq_Skill_Level = \
                                    data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup']['skillCompetency'][b][
                                        'competencelevel']

                                    csq_WS.cell(row=1, column=num, value=f"Skill Name  {b + 1}")
                                    csq_WS.cell(row=count + 1, column=num, value=f"{csq_Skill}")
                                    num += 1
                                    csq_WS.cell(row=1, column=num, value=f"Skill Level  {b + 1}")
                                    csq_WS.cell(row=count + 1, column=num, value=csq_Skill_Level)
                                    num += 1

                        csq_Selection = data_dict['csqs']['csq'][i]['poolSpecificInfo']['skillGroup'][
                            'selectionCriteria']

                        count += 1

                        csq_WS.cell(row=count, column=1, value=csq_ID)
                        csq_WS.cell(row=count, column=2, value=csq_Name)
                        csq_WS.cell(row=count, column=3, value=csq_Type)
                        csq_WS.cell(row=count, column=4, value=csq_Routing)
                        csq_WS.cell(row=count, column=5, value=csq_Algo)
                        csq_WS.cell(row=count, column=8, value=csq_Pool_Type)

                        csq_Skill = ""
                        csq_Skill_Level = ""

                    count = 1
            else:
                print("No CSQ")
            print(f"CSQ Details : \033[92mDone\033[0m")
            get_Team()
        except:
            print("\033[91mCSQ API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



# Getting Team details from CCX
def get_Team():
    try:
        count = 1
        team_Sec_Sup = ""
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(team_URL, headers=headers, verify=False)
            team_WS.cell(row=1, column=1, value="Team ID")
            team_WS.cell(row=1, column=2, value="Team Name")
            team_WS.cell(row=1, column=3, value="Primary Supervisor")
            team_WS.cell(row=1, column=4, value="Secondary Supervisor")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)

            # Checking if there are any Teams in CCX
            if data_dict['teams'] is not None:
                resp_Type = type(data_dict['teams']['team'])
                # Checking if the Team API response from CCX contains one or more Teams.
                if resp_Type == dict:
                    team_ID = data_dict['teams']['team']['teamId']
                    team_Name = data_dict['teams']['team']['teamname']
                    team_WS.cell(row=2, column=1, value=team_ID)
                    team_WS.cell(row=2, column=2, value=team_Name)
                else:
                    for i in range(0, len(data_dict['teams']['team'])):
                        team_ID = data_dict['teams']['team'][i]['teamId']
                        team_Name = data_dict['teams']['team'][i]['teamname']
                        if "primarySupervisor" in data_dict['teams']['team'][i]:
                            team_Pri_Sup = data_dict['teams']['team'][i]['primarySupervisor']['@name']
                        else:
                            team_Pri_Sup = "NA"

                        team_Sec_Sup = "NA"
                        count += 1

                        team_WS.cell(row=count, column=1, value=team_ID)
                        team_WS.cell(row=count, column=2, value=team_Name)
                        team_WS.cell(row=count, column=3, value=team_Pri_Sup)
                        team_WS.cell(row=count, column=4, value=team_Sec_Sup)

                    count = 1
            else:
                print("No Team")
            print(f"Team Details : \033[92mDone\033[0m")
            get_Skills()
        except:
            print("\033[91mTeam API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



# Getting Skill details from CCX
def get_Skills():
    try:
        count = 1
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(skill_URL, headers=headers, verify=False)
            skill_WS.cell(row=1, column=1, value="ID")
            skill_WS.cell(row=1, column=2, value="Skill Name")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)

            # Checking if there are any Skills in CCX
            if data_dict['skills'] is not None:
                resp_Type = type(data_dict['skills']['skill'])
                # Checking if the Skill API response from CCX contains one or more Skills.
                if resp_Type == dict:
                    id = data_dict['skills']['skill']['skillId']
                    name = data_dict['skills']['skill']['skillName']
                    skill_WS.cell(row=2, column=1, value=id)
                    skill_WS.cell(row=2, column=2, value=name)
                else:
                    for i in range(0, len(data_dict['skills']['skill'])):
                        id = data_dict['skills']['skill'][i]['skillId']
                        name = data_dict['skills']['skill'][i]['skillName']
                        count += 1
                        skill_WS.cell(row=count, column=1, value=id)
                        skill_WS.cell(row=count, column=2, value=name)

                count = 1
            print(f"Skill Details : \033[92mDone\033[0m")
            get_Wrapup()
        except:
            print("\033[91mSkill API Request to UCCX has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")



####### FINESSE REQUESTS #######

# Getting Wrap Up codes from Finesse
def get_Wrapup():
    try:
        count = 1
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(wrapup_URL, headers=headers, verify=False)
            wrapup_WS.cell(row=1, column=1, value="Label")
            wrapup_WS.cell(row=1, column=2, value="Is Global ?")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)

            # Checking if there are any Wrap Up Codes in CCX
            if data_dict['WrapUpReasons'] is not None:
                resp_Type = type(data_dict['WrapUpReasons']['WrapUpReason'])
                # Checking if the Wrap-Up API response from CCX contains one or more Wrap Up codes.
                if resp_Type == dict:
                    name = data_dict['WrapUpReasons']['WrapUpReason']['label']
                    isGlobal = data_dict['WrapUpReasons']['WrapUpReason']['forAll']
                    wrapup_WS.cell(row=2, column=1, value=name)
                    wrapup_WS.cell(row=2, column=2, value=isGlobal)
                else:
                    for i in range(0, len(data_dict['WrapUpReasons']['WrapUpReason'])):
                        name = data_dict['WrapUpReasons']['WrapUpReason'][i]['label']
                        isGlobal = data_dict['WrapUpReasons']['WrapUpReason'][i]['forAll']

                        count += 1

                        wrapup_WS.cell(row=count, column=1, value=name)
                        wrapup_WS.cell(row=count, column=2, value=isGlobal)

                    count = 1
            else:
                print("No Wrap-up Codes")
            print(f"Wrap-Up Code Details : \033[92mDone\033[0m")
            get_Reason()
        except:
            print("\033[91mWrap-Up API Request to Finesse has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encpountered an error. Halting the program.\033[0m")


# Getting Reason codes from Finesse
def get_Reason():
    try:
        count = 1
        disable_warnings(InsecureRequestWarning)
        try:
            resp = requests.get(reason_URL, headers=headers, verify=False)
            reason_WS.cell(row=1, column=1, value="Label")
            reason_WS.cell(row=1, column=2, value="Category")
            reason_WS.cell(row=1, column=3, value="Is Global ?")
            reason_WS.cell(row=1, column=4, value="Code")
            reason_WS.cell(row=1, column=5, value="systemCode")
            data_dict = xmltodict.parse(resp.content)
            # data_dict = xmltodict.parse(resp)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)

            # Checking if there are any Reason Codes in CCX
            if data_dict['ReasonCodes'] is not None:
                resp_Type = type(data_dict['ReasonCodes']['ReasonCode'])
                # Checking if the Reason API response from CCX contains one or more Reason codes.
                if resp_Type == dict:
                    label = data_dict['ReasonCodes']['ReasonCode']['label']
                    category = data_dict['ReasonCodes']['ReasonCode']['category']
                    isGlobal = data_dict['ReasonCodes']['ReasonCode']['forAll']
                    code = data_dict['ReasonCodes']['ReasonCode']['code']
                    sysCode = data_dict['ReasonCodes']['ReasonCode']['systemCode']

                    reason_WS.cell(row=2, column=1, value=label)
                    reason_WS.cell(row=2, column=2, value=category)
                    reason_WS.cell(row=2, column=3, value=isGlobal)
                    reason_WS.cell(row=2, column=4, value=code)
                    reason_WS.cell(row=2, column=5, value=sysCode)
                else:
                    for i in range(0, len(data_dict['ReasonCodes']['ReasonCode'])):
                        label = data_dict['ReasonCodes']['ReasonCode'][i]['label']
                        category = data_dict['ReasonCodes']['ReasonCode'][i]['category']
                        isGlobal = data_dict['ReasonCodes']['ReasonCode'][i]['forAll']
                        code = data_dict['ReasonCodes']['ReasonCode'][i]['code']
                        sysCode = data_dict['ReasonCodes']['ReasonCode'][i]['systemCode']

                        count += 1

                        reason_WS.cell(row=count, column=1, value=label)
                        reason_WS.cell(row=count, column=2, value=category)
                        reason_WS.cell(row=count, column=3, value=isGlobal)
                        reason_WS.cell(row=count, column=4, value=code)
                        reason_WS.cell(row=count, column=5, value=sysCode)

                    count = 1
            else:
                print("No Reason Codes")
            print(f"Reason Code Details : \033[92mDone\033[0m")
            get_Phonebooks()
        except:
            print("\033[91mReason Code API Request to Finesse has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")


# Getting Phonebooks and Contact details from CCX
def get_Phonebooks():
    try:
        p_count = 1
        index = 7
        col = 4
        disable_warnings(InsecureRequestWarning)
        try:
            resp_PB = requests.get(phonebook_URL, headers=headers, verify=False)
            data_dict_PB = xmltodict.parse(resp_PB.content)
            # data_dict_PB = xmltodict.parse(resp_PB)
            # print(type(data_dict))  # DICT
            json_data = json.dumps(data_dict_PB)
            # print(type(json_data))  # STR
            json_loads = json.loads(json_data)
            # print(type(json_loads))  # DICT
            # print(json_loads)
            resp_Type = type(data_dict_PB['PhoneBooks']['PhoneBook'])

            # Checking if there are any Phonebooks in CCX
            if data_dict_PB['PhoneBooks'] is not None:
                resp_Type = type(data_dict_PB['PhoneBooks']['PhoneBook'])
                # Checking if the Phonebook API response from CCX contains one or more Phonebooks.
                if resp_Type == dict:
                    phonebook_WS = wb.create_sheet(title=f"Phonebook", index=8)
                    name = data_dict_PB['PhoneBooks']['PhoneBook']['name']
                    uri = data_dict_PB['PhoneBooks']['PhoneBook']['uri']
                    isGlobal = data_dict_PB['PhoneBooks']['PhoneBook']['type']
                    phonebook_WS.cell(row=1, column=1, value="Name")
                    phonebook_WS.cell(row=1, column=2, value="ID")
                    phonebook_WS.cell(row=1, column=3, value="Is Global ?")

                    p_count += 1

                    uri_Split = uri.split('/')
                    pb_ID = uri_Split[4]

                    # Getting contact details associated with each phonebook
                    contact_URL = f"https://uccx1.dcloud.cisco.com:8445/finesse/api/PhoneBook/{pb_ID}/Contacts"
                    resp_Contacts = requests.get(contact_URL, headers=headers, verify=False)
                    data_dict_Contact = xmltodict.parse(resp_Contacts.content)

                    for a in range(0, len(data_dict_Contact['Contacts']['Contact'])):
                        f_Name = data_dict_Contact['Contacts']['Contact'][a]['firstName']
                        l_Name = data_dict_Contact['Contacts']['Contact'][a]['lastName']
                        c_Name = f"{f_Name} {l_Name}"
                        desc = data_dict_Contact['Contacts']['Contact'][a]['description']
                        ph_Number = data_dict_Contact['Contacts']['Contact'][a]['phoneNumber']

                        phonebook_WS.cell(row=p_count, column=1, value=name)
                        phonebook_WS.cell(row=p_count, column=2, value=pb_ID)
                        phonebook_WS.cell(row=p_count, column=3, value=isGlobal)

                        phonebook_WS.cell(row=1, column=col, value=f"Contact Name {a + 1}")
                        phonebook_WS.cell(row=p_count, column=col, value=c_Name)
                        col += 1
                        phonebook_WS.cell(row=1, column=col, value=f"Description {a + 1}")
                        phonebook_WS.cell(row=p_count, column=col, value=desc)
                        col += 1
                        phonebook_WS.cell(row=1, column=col, value=f"Phone Number {a + 1}")
                        phonebook_WS.cell(row=p_count, column=col, value=ph_Number)
                        col += 1

                    p_count = 1
                else:
                    for i in range(0, len(data_dict_PB['PhoneBooks']['PhoneBook'])):
                        col = 4
                        phonebook_WS = wb.create_sheet(title=f"Phonebook {i + 1}", index=index + p_count)
                        name = data_dict_PB['PhoneBooks']['PhoneBook'][i]['name']
                        uri = data_dict_PB['PhoneBooks']['PhoneBook'][i]['uri']
                        isGlobal = data_dict_PB['PhoneBooks']['PhoneBook'][i]['type']
                        phonebook_WS.cell(row=1, column=1, value="Name")
                        phonebook_WS.cell(row=1, column=2, value="ID")
                        phonebook_WS.cell(row=1, column=3, value="Is Global ?")

                        p_count += 1

                        uri_Split = uri.split('/')
                        pb_ID = uri_Split[4]
                        contact_URL = f"https://uccx1.dcloud.cisco.com:8445/finesse/api/PhoneBook/{pb_ID}/Contacts"
                        resp_Contacts = requests.get(contact_URL, headers=headers, verify=False)
                        data_dict_Contact = xmltodict.parse(resp_Contacts.content)

                        for a in range(0, len(data_dict_Contact['Contacts']['Contact'])):
                            phonebook_WS.cell(row=p_count, column=1, value=name)
                            phonebook_WS.cell(row=p_count, column=2, value=pb_ID)
                            phonebook_WS.cell(row=p_count, column=3, value=isGlobal)

                            f_Name = data_dict_Contact['Contacts']['Contact'][a]['firstName']
                            l_Name = data_dict_Contact['Contacts']['Contact'][a]['lastName']
                            c_Name = f"{f_Name} {l_Name}"
                            desc = data_dict_Contact['Contacts']['Contact'][a]['description']
                            ph_Number = data_dict_Contact['Contacts']['Contact'][a]['phoneNumber']

                            phonebook_WS.cell(row=1, column=col, value=f"Contact Name {a + 1}")
                            phonebook_WS.cell(row=p_count, column=col, value=c_Name)
                            col += 1
                            phonebook_WS.cell(row=1, column=col, value=f"Description {a + 1}")
                            phonebook_WS.cell(row=p_count, column=col, value=desc)
                            col += 1
                            phonebook_WS.cell(row=1, column=col, value=f"Phone Number {a + 1}")
                            phonebook_WS.cell(row=p_count, column=col, value=ph_Number)
                            col += 1

                        p_count = 1
            else:
                print("No Phonebooks")
            print(f"Phonebook Details : \033[92mDone\033[0m")
            create_CCX_File()
        except:
            print("\033[91mPhonebook API Request to Finesse has encountered an error.\033[0m")
    except:
        print("\033[91mProgram has encountered an error. Halting the program.\033[0m")


# Loading all the captured details in an Excel file
def create_CCX_File():
    global full_Path_CCX
    try:
        name = "CCX-Details.xlsx"
        os_Type = platform.system()
        if os_Type == "Windows":
            current_dir = os.getcwd()
            base_Path_CCX = os.path.join(current_dir, "Files")
        elif os_Type == "Linux":
            current_dir = os.getcwd()
            base_Path_CCX = os.path.join(current_dir, "Files")

        if not os.path.exists(base_Path_CCX):
            os.makedirs(base_Path_CCX)
            print(f"Created folder: {base_Path_CCX}")

        full_Path_CCX = os.path.join(base_Path_CCX, name)

        wb.save(full_Path_CCX)
        print("*" * 60)
        print(f"\033[92mUCCX Data has been captured successfully.\033[0m\n")
        print("\033[94mDo you want to proceed with the second stage of importing this data into WxCC ?\033[0m")
        ques = input('\033[94mEnter "Y" to proceed or "N" to stop this program and review the CCX data first : \033[0m')
        if ques == "Y" or ques == "y":
            print("\033[93mConverting UCCX Data into WxCC compatible format.....\033[0m")
            WxCC_Sheet.app_WxCC()
        else:
            print("*" * 60)
            print(f"\033[92mClosing the app.\nA file with a name {name} containing UCCX details has been created at {base_Path_CCX}\033[0m")
            print("*" * 60)
    except:
        print(f"\033[91mThe UCCX Config data was captured successfully but an error was encountered while saving "
              f"the file at {base_Path_CCX}. Please ensure that you have appropriate read/write"
              f"permissions in this directory. \033[0m")

