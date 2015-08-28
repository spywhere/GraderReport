import getpass
import os
import re
import sys
import time
import gspread
import mechanize
from bs4 import BeautifulSoup
from oauth2client.client import OAuth2WebServerFlow


OAUTH_SCOPES = ["https://spreadsheets.google.com/feeds"]
GRADER_BASE = "http://grader.eng.src.ku.ac.th"
UTILS_BASE = "http://digitalparticle.com/graderta"


def parse_file(data):
    datalist = data.replace("\r\n", "\n").split("\n")
    return {
        "grader_user": datalist[0],
        "grader_password": datalist[1],
        "client_id": datalist[2],
        "client_secret": datalist[3],
        "spreadsheet": datalist[4],
        "worksheet": datalist[5],
        "merger_no_columns": datalist[6],
        "merger_no_rows": datalist[7],
        "merger_keep_columns": datalist[8],
        "merger_keep_rows": datalist[9],
        "normalizer_include_contests": datalist[10],
        "normalizer_calculate_columns": datalist[11],
        "normalizer_withdrawn_users": datalist[12],
        "normalizer_exclude_users": datalist[13],
        "normalizer_include_columns": datalist[14]
    }


def startsin(arr, st):
    for item in arr:
        if st.startswith(item):
            return True
    return False


def trim(st):
    if st.startswith(" ") or st.startswith("\n") or st.startswith("\t"):
        return trim(st[1:])
    elif st.endswith(" ") or st.endswith("\n") or st.endswith("\t"):
        return trim(st[:-1])
    return st


def run(args):
    if len(args) < 2:
        print("usage: " + args[0] + " [option] <project_file> [flags]")
        print("Options:")
        print("  new\t: Create a new grader project")
        print("  edit\t: Change grader project settings")
        print("Flags:")
        print("  -score <file name>\t: Read grader's score from file")
        print("  -autoretry\t: Repeat update on uncritical errors")
        print("  -infinite\t: Repeat update until interrupted")
        print("  -forever\t: Alias of -infinite")
        print("  -n <num>\t: Repeat updating n times")
        print("  -times <num>\t: Alias of -n")
        print("  -l <num>\t: Delay between each update (default is 1 minute)")
        print("  -delay <num>\t: Alias of -l")
        print("  -break [google,grader,merger,normalizer,update]\t: Pause on breakpoints (can be separated by ,)")
        print("  -final\t: Alias of -break with all possible breakpoints")
        print("  -force\t: Force update on warning (if any)")
        print("  -store\t: Save fetched result page to file")
        print("  -silent\t: Update result without printing anything")
        return
    flags = {
        # Repeat times
        "n": 1,
        # Delay between each times
        "l": 60,
        "store_source": False,
        "force": False,
        "silent": False,
        "autoretry": False,
        "pause_google": False,
        "pause_grader": False,
        "pause_merger": False,
        "pause_normalizer": False,
        "pause_update": False,
        "first": True,
        "score": None,
        "credentials": None
    }
    for aid in range(2, len(args)):
        if args[aid] == "-autoretry":
            flags["autoretry"] = True
        elif args[aid] == "-infinite" or args[aid] == "-forever":
            flags["n"] = -1
        elif args[aid] == "-n" or args[aid] == "-times":
            if int(args[aid+1]) > 0:
                flags["n"] = int(args[aid+1])
        elif args[aid] == "-l" or args[aid] == "-delay":
            if int(args[aid+1]) > 0:
                flags["l"] = int(args[aid+1])
        elif args[aid] == "-score":
            flags["score"] = args[aid+1]
        elif args[aid] == "-force":
            flags["force"] = True
        elif args[aid] == "-store":
            flags["store_source"] = True
        elif args[aid] == "-silent":
            flags["silent"] = True
        elif args[aid] == "-break":
            breakpoints = args[aid+1].split(",")
            for breakpoint in breakpoints:
                if breakpoint == "google":
                    flags["pause_google"] = True
                elif breakpoint == "grader":
                    flags["pause_grader"] = True
                elif breakpoint == "merger":
                    flags["pause_merger"] = True
                elif breakpoint == "normalizer":
                    flags["pause_normalizer"] = True
                elif breakpoint == "update":
                    flags["pause_update"] = True
        elif args[aid] == "-final":
            flags["pause_google"] = True
            flags["pause_grader"] = True
            flags["pause_merger"] = True
            flags["pause_normalizer"] = True
            flags["pause_update"] = True
    input_file = args[1]
    mode = ""
    if args[1] == "new":
        mode = "new"
        input_file = args[2]
    elif args[1] == "edit":
        mode = "edit"
        input_file = args[2]
    update(mode, input_file, flags)


def update(mode="", input_file="", flags={}):
    if mode == "new":
        if os.path.exists(input_file):
            print("File already exists")
            return
        client = None
        while True:
            print("==== Grader ====")
            graderu = raw_input("Grader user: ")
            if graderu == "":
                return
            graderp = getpass.getpass("Grader password: ")
            if graderp == "":
                return
            print("Connecting to grader...")
            br = mechanize.Browser()
            br.set_handle_robots(False)
            try:
                br.open(GRADER_BASE)
            except Exception as msg:
                print("Grader Error! "+str(msg))
                return
            if len(list(br.forms())) < 1:
                print("No login form in grader... Please check the grader...")
                return
            br.form = list(br.forms())[0]
            br.form["login"] = graderu
            br.form["password"] = graderp
            print("Logging into grader...")
            if re.search("Wrong password", br.submit().read()) is not None:
                print("Invalid user or password.")
                continue
            break
        while True:
            print("==== GMail ====")
            client_id = raw_input("Client ID: ")
            if client_id == "":
                return
            client_secret = getpass.getpass("Client Secret: ")
            if client_secret == "":
                return
            print("Getting user credential...")
            flow = OAuth2WebServerFlow(client_id, client_secret, " ".join(OAUTH_SCOPES))
            flow_info = flow.step1_get_device_and_user_codes()
            print("Enter the following code at {0}: {1}".format(flow_info.verification_url,
                                                                flow_info.user_code))
            raw_input("Then press Enter/Return.")
            try:
                credentials = flow.step2_exchange(device_flow_info=flow_info)
            except:
                print("Failed to get user credentials.")
                continue
            print("Logging in...")
            try:
                client = gspread.authorize(credentials)
                break
            except gspread.AuthenticationError:
                print("Invalid email or password.")
                continue
        while True:
            sheetkeyword = raw_input("Spreeadsheets: ")
            if sheetkeyword == "":
                sheetkeyword = None
            print("Listing spreadsheets...")
            sheets = client.openall()
            filteredsheets = []
            for sheet in sheets:
                if sheetkeyword is None or sheet.title.lower().startswith(sheetkeyword.lower()):
                    filteredsheets.append(sheet)
            if len(filteredsheets) <= 0:
                print("No spreadsheet found")
                continue
            for sheet in filteredsheets:
                print("  Name: "+sheet.title)
                print("  Key: "+sheet.id)
                print("--------")
            back = False
            while True:
                currentsheet = None
                sheetkey = raw_input("Select spreadsheet key: ")
                if sheetkey == ".":
                    currentsheet = filteredsheets[0]
                    break
                if sheetkey == "<":
                    back = True
                    break
                if sheetkey == "":
                    return
                for sheet in filteredsheets:
                    if sheet.id == sheetkey:
                        currentsheet = sheet
                        break
                if currentsheet is None:
                    print("Invalid spreadsheet key")
                    continue
                else:
                    break
            if back:
                continue
            print("Spreadsheet: "+currentsheet.title)
            worksheetstr = ""
            print("Listing worksheets...")
            worksheets = currentsheet.worksheets()
            for worksheet in worksheets:
                if worksheetstr == "":
                    worksheetstr = worksheetstr+worksheet.title
                else:
                    worksheetstr = worksheetstr+", "+worksheet.title
            print("Worksheets: "+worksheetstr)
            while True:
                workingsheet = None
                sheetname = raw_input("Select worksheet: ")
                if sheetname == ".":
                    workingsheet = worksheets[0]
                    break
                if sheetname == "<":
                    back = True
                    break
                if sheetname == "":
                    return
                for sheet in worksheets:
                    if sheet.title.lower() == sheetname.lower():
                        workingsheet = sheet
                        break
                if workingsheet is None:
                    print("Invalid worksheet name")
                    continue
                else:
                    break
            if back:
                continue
            print("==== Merger ====")
            no_col = raw_input("No Columns: ")
            no_row = raw_input("No Rows: ")
            keep_col = raw_input("Keep Columns: ")
            keep_row = raw_input("Keep Rows: ")
            print("==== Normalizer ====")
            inc_contest = raw_input("Include Contest: ")
            cal_col = raw_input("Calculate Columns: ")
            w_user = raw_input("Withdrawn Users: ")
            ex_user = raw_input("Exclude Users: ")
            ex_col = raw_input("Include Columns: ")
            print("==== Grader ====")
            print("Grader User: "+graderu)
            print("Grader Password: "+("*"*len(graderp)))
            print("==== Google Spreadsheet ====")
            print("Client ID: "+client_id)
            print("Client Secret: "+("*"*len(client_secret)))
            print("Spreadsheet: "+currentsheet.title)
            print("Spreadsheet Key: "+currentsheet.id)
            print("Worksheet: " + workingsheet.title)
            print("==== Merger ====")
            print("No Columns: " + no_col)
            print("No Rows: " + no_row)
            print("Keep Columns: " + keep_col)
            print("Keep Rows: " + keep_row)
            print("==== Normalizer ====")
            print("Include Contests: " + inc_contest)
            print("Calculate Columns: " + cal_col)
            print("Withdrawn Users: " + w_user)
            print("Exclude Users: " + ex_user)
            print("Include Columns: " + ex_col)
            print("================")
            confirm = raw_input("Confirm create? (y/n): ")
            if confirm.lower() == "y":
                # Save to file
                f = open(input_file, "w")
                data = []
                data.append(graderu)
                data.append(graderp)
                data.append(client_id)
                data.append(client_secret)
                data.append(currentsheet.id)
                data.append(workingsheet.title)
                data.append(no_col)
                data.append(no_row)
                data.append(keep_col)
                data.append(keep_row)
                data.append(inc_contest)
                data.append(cal_col)
                data.append(w_user)
                data.append(ex_user)
                data.append(ex_col)
                f.write(str.join("\n", data))
                f.close()
                print("Project saved to " + input_file)
            return
        return
    if mode == "edit":
        if not os.path.exists(input_file):
            print("File is not exists")
            return
        f = open(input_file, "r")
        if f is None:
            print("Error occured while reading a file")
            return
        file_info = parse_file(f.read())
        f.close()
        print("==== Grader ====")
        print("Grader User: "+file_info["grader_user"])
        print("Grader Password: "+("*"*len(file_info["grader_password"])))
        print("==== Google Spreadsheet ====")
        print("Client ID: "+file_info["client_id"])
        print("Client Secret: "+("*"*len(file_info["client_secret"])))
        print("Spreadsheet Key: "+file_info["spreadsheet"])
        print("Worksheet: " + file_info["worksheet"])
        print("==== Merger ====")
        print("No Columns: " + file_info["merger_no_columns"])
        print("No Rows: " + file_info["merger_no_rows"])
        print("Keep Columns: " + file_info["merger_keep_columns"])
        print("Keep Rows: " + file_info["merger_keep_rows"])
        print("==== Normalizer ====")
        print("Include Contests: " + file_info["normalizer_include_contests"])
        print("Calculate Columns: " + file_info["normalizer_calculate_columns"])
        print("Withdrawn Users: " + file_info["normalizer_withdrawn_users"])
        print("Exclude Users: " + file_info["normalizer_exclude_users"])
        print("Include Columns: " + file_info["normalizer_include_columns"])
        while True:
            edit = False
            print("==== Grader ====")
            confirm = raw_input("Change grader user? (y/n): ")
            if confirm.lower() == "y":
                file_info["grader_user"] = raw_input("New grader user: ")
                edit = True
            confirm = raw_input("Change grader password? (y/n): ")
            if confirm.lower() == "y":
                file_info["grader_password"] = getpass.getpass("New grader password: ")
                edit = True
            print("==== GMail ====")
            confirm = raw_input("Change client ID? (y/n): ")
            if confirm.lower() == "y":
                file_info["client_id"] = raw_input("New client ID: ")
                edit = True
            confirm = raw_input("Change client secret? (y/n): ")
            if confirm.lower() == "y":
                file_info["client_secret"] = getpass.getpass("New client secret: ")
                edit = True
            editspreadsheet = False
            confirm = raw_input("Change spreadsheet? This will also change worksheet (y/n): ")
            if confirm.lower() == "y":
                client_id = raw_input("Client ID: ")
                if client_id == "":
                    return
                client_secret = getpass.getpass("Client Secret: ")
                if client_secret == "":
                    return
                print("Getting user credential...")
                flow = OAuth2WebServerFlow(file_info["client_id"], file_info["client_secret"], " ".join(OAUTH_SCOPES))
                flow_info = flow.step1_get_device_and_user_codes()
                print("Enter the following code at {0}: {1}".format(flow_info.verification_url,
                                                                    flow_info.user_code))
                raw_input("Then press Enter/Return.")
                try:
                    credentials = flow.step2_exchange(device_flow_info=flow_info)
                except:
                    print("Failed to get user credentials.")
                    return

                client = None
                print("Logging in...")
                try:
                    client = gspread.authorize(credentials)
                    while True:
                        sheetkeyword = raw_input("Spreadsheets filter: ")
                        if sheetkeyword == "":
                            sheetkeyword = None
                        print("Listing spreadsheets...")
                        sheets = client.openall()
                        filteredsheets = []
                        for sheet in sheets:
                            if sheetkeyword is None or sheet.title.lower().startswith(sheetkeyword.lower()):
                                filteredsheets.append(sheet)
                        if len(filteredsheets) <= 0:
                            print("No spreadsheet found")
                            continue
                        for sheet in filteredsheets:
                            print("  Name: "+sheet.title)
                            print("  Key: "+sheet.id)
                            print("--------")
                        back = False
                        stop = False
                        while True:
                            currentsheet = None
                            sheetkey = raw_input("Select spreadsheet key: ")
                            if sheetkey == ".":
                                currentsheet = filteredsheets[0]
                                break
                            if sheetkey == "<":
                                back = True
                                break
                            if sheetkey == "":
                                stop = True
                                break
                            for sheet in filteredsheets:
                                if sheet.id == sheetkey:
                                    currentsheet = sheet
                                    break
                            if currentsheet is None:
                                print("Invalid spreadsheet key")
                                continue
                            else:
                                break
                        if back:
                            continue
                        if stop:
                            break

                        if back:
                            continue
                        else:
                            edit = True
                            break
                    if edit:
                        file_info["spreadsheet"] = currentsheet.id
                        editspreadsheet = True
                except gspread.AuthenticationError:
                    print("Invalid credentials.")

            if not editspreadsheet:
                confirm = raw_input("Change worksheet? (y/n): ")
            if editspreadsheet or confirm.lower() == "y":
                client = None

                print("Getting user credential...")
                flow = OAuth2WebServerFlow(file_info["client_id"], file_info["client_secret"], " ".join(OAUTH_SCOPES))
                flow_info = flow.step1_get_device_and_user_codes()
                print("Enter the following code at {0}: {1}".format(flow_info.verification_url,
                                                                    flow_info.user_code))
                raw_input("Then press Enter/Return.")
                try:
                    credentials = flow.step2_exchange(device_flow_info=flow_info)
                except:
                    print("Failed to get user credentials.")
                    break

                if not editspreadsheet:
                    print("Logging in...")
                try:
                    client = gspread.authorize(credentials)
                    print("Opening spreadsheets...")
                    try:
                        currentsheet = client.open_by_key(file_info["spreadsheet"])
                    except gspread.SpreadsheetNotFound:
                        print("Spreadsheet is not found")
                        break
                    print("Spreadsheet: "+currentsheet.title)
                    worksheetstr = ""
                    print("Listing worksheets...")
                    worksheets = currentsheet.worksheets()
                    for worksheet in worksheets:
                        if worksheetstr == "":
                            worksheetstr = worksheetstr+worksheet.title
                        else:
                            worksheetstr = worksheetstr+", "+worksheet.title
                    print("Worksheets: "+worksheetstr)
                    while True:
                        workingsheet = None
                        sheetname = raw_input("Select worksheet: ")
                        if sheetname == ".":
                            workingsheet = worksheets[0]
                            break
                        if sheetname == "":
                            return
                        for sheet in worksheets:
                            if sheet.title.lower() == sheetname.lower():
                                workingsheet = sheet
                                break
                        if workingsheet is None:
                            print("Invalid worksheet name")
                            continue
                        else:
                            edit = True
                            break
                    if edit:
                        file_info["worksheet"] = workingsheet.title
                except gspread.AuthenticationError:
                    print("Invalid credentials.")
            print("==== Merger ====")
            confirm = raw_input("Change no columns? (y/n): ")
            if confirm.lower() == "y":
                file_info["merger_no_columns"] = raw_input("New no columns: ")
                edit = True
            confirm = raw_input("Change no rows? (y/n): ")
            if confirm.lower() == "y":
                file_info["merger_no_rows"] = raw_input("New no rows: ")
                edit = True
            confirm = raw_input("Change keep columns? (y/n): ")
            if confirm.lower() == "y":
                file_info["merger_keep_columns"] = raw_input("New keep columns: ")
                edit = True
            confirm = raw_input("Change keep rows? (y/n): ")
            if confirm.lower() == "y":
                file_info["merger_keep_rows"] = raw_input("New keep rows: ")
                edit = True
            print("==== Normalizer ====")
            confirm = raw_input("Change include contests? (y/n): ")
            if confirm.lower() == "y":
                file_info["normalizer_include_contests"] = raw_input("New include contests: ")
                edit = True
            confirm = raw_input("Change calculate columns? (y/n): ")
            if confirm.lower() == "y":
                file_info["normalizer_calculate_columns"] = raw_input("New calculate columns: ")
                edit = True
            confirm = raw_input("Change withdrawn users? (y/n): ")
            if confirm.lower() == "y":
                file_info["normalizer_withdrawn_users"] = raw_input("New withdrawn users: ")
                edit = True
            confirm = raw_input("Change exclude users? (y/n): ")
            if confirm.lower() == "y":
                file_info["normalizer_exclude_users"] = raw_input("New exclude users: ")
                edit = True
            confirm = raw_input("Change include columns? (y/n): ")
            if confirm.lower() == "y":
                file_info["normalizer_include_columns"] = raw_input("New include columns: ")
                edit = True
            if edit:
                print("==== Grader ====")
                print("Grader User: "+file_info["grader_user"])
                print("Grader Password: "+("*"*len(file_info["grader_password"])))
                print("==== Google Spreadsheet ====")
                print("Client ID: "+file_info["client_id"])
                print("Client Secret: "+("*"*len(file_info["client_secret"])))
                print("Spreadsheet Key: "+file_info["spreadsheet"])
                print("Worksheet: " + file_info["worksheet"])
                print("==== Merger ====")
                print("No Columns: " + file_info["merger_no_columns"])
                print("No Rows: " + file_info["merger_no_rows"])
                print("Keep Columns: " + file_info["merger_keep_columns"])
                print("Keep Rows: " + file_info["merger_keep_rows"])
                print("==== Normalizer ====")
                print("Include Contests: " + file_info["normalizer_include_contests"])
                print("Calculate Columns: " + file_info["normalizer_calculate_columns"])
                print("Withdrawn Users: " + file_info["normalizer_withdrawn_users"])
                print("Exclude Users: " + file_info["normalizer_exclude_users"])
                print("Include Columns: " + file_info["normalizer_include_columns"])
                print("================")
                confirm = raw_input("Confirm changes? (y/n): ")
                if confirm.lower() == "y":
                    f = open(input_file, "w")
                    data = []
                    data.append(file_info["grader_user"])
                    data.append(file_info["grader_password"])
                    data.append(file_info["client_id"])
                    data.append(file_info["client_secret"])
                    data.append(file_info["spreadsheet"])
                    data.append(file_info["worksheet"])
                    data.append(file_info["merger_no_columns"])
                    data.append(file_info["merger_no_rows"])
                    data.append(file_info["merger_keep_columns"])
                    data.append(file_info["merger_keep_rows"])
                    data.append(file_info["normalizer_include_contests"])
                    data.append(file_info["normalizer_calculate_columns"])
                    data.append(file_info["normalizer_withdrawn_users"])
                    data.append(file_info["normalizer_exclude_users"])
                    data.append(file_info["normalizer_include_columns"])
                    f.write(str.join("\n", data))
                    f.close()
                    print("Project saved to " + input_file)
                    return
                continue
            break
        return
    if not os.path.exists(input_file):
        print("File is not exists")
        return
    f = open(input_file, "r")
    if f is None:
        print("Error occured while reading a file")
        return
    file_info = parse_file(f.read())
    f.close()

    if flags["first"] and not flags["silent"]:
        print("==== Grader ====")
        print("Grader User: "+file_info["grader_user"])
        print("Grader Password: "+("*"*len(file_info["grader_password"])))
        print("==== Google Spreadsheet ====")
        print("Client ID: "+file_info["client_id"])
        print("Client Secret: "+("*"*len(file_info["client_secret"])))
        print("Spreadsheet Key: "+file_info["spreadsheet"])
        print("Worksheet: " + file_info["worksheet"])
        print("==== Merger ====")
        print("No Columns: " + file_info["merger_no_columns"])
        print("No Rows: " + file_info["merger_no_rows"])
        print("Keep Columns: " + file_info["merger_keep_columns"])
        print("Keep Rows: " + file_info["merger_keep_rows"])
        print("==== Normalizer ====")
        print("Include Contests: " + file_info["normalizer_include_contests"])
        print("Calculate Columns: " + file_info["normalizer_calculate_columns"])
        print("Withdrawn Users: " + file_info["normalizer_withdrawn_users"])
        print("Exclude Users: " + file_info["normalizer_exclude_users"])
        print("Include Columns: " + file_info["normalizer_include_columns"])
        print("================")
        if not flags["force"]:
            confirm = raw_input("Continue? (type 'n' or 'cancel' to cancel): ")
            if confirm.lower() == "n" or confirm.lower() == "cancel":
                return

    # ======
    # Retry loop
    # ======
    retry = 0
    started = time.mktime(time.localtime())

    if not flags["credentials"]:
        if not flags["silent"]:
            print("Getting user credential...")
        flow = OAuth2WebServerFlow(file_info["client_id"], file_info["client_secret"], " ".join(OAUTH_SCOPES))
        flow_info = flow.step1_get_device_and_user_codes()
        print("Enter the following code at {0}: {1}".format(flow_info.verification_url,
                                                            flow_info.user_code))
        raw_input("Then press Enter/Return.")
        try:
            flags["credentials"] = flow.step2_exchange(device_flow_info=flow_info)
        except:
            print("Failed to get user credentials.")
            return

    while retry >= 0:
        if retry > 0:
            print("================")
            print("Failed to update results on " + time.strftime("%d/%m/%Y %H:%M", time.localtime()) + "... (in " + str(time.mktime(time.localtime()) - started) + "s)")
            print("Retrying...")
            print("================")
        retry = -1
        started = time.mktime(time.localtime())
        # ======
        # Google Spreadsheet
        # ======
        client = None

        if not flags["silent"]:
            print("Logging in...")
        try:
            client = gspread.authorize(flags["credentials"])
        except gspread.AuthenticationError:
            print("Invalid credentials.")
            return
        if not flags["silent"]:
            print("Opening spreadsheets...")
        try:
            currentsheet = client.open_by_key(file_info["spreadsheet"])
        except gspread.SpreadsheetNotFound:
            print("Spreadsheet is not found")
            return
        if not flags["silent"]:
            print("Opening worksheets...")
        worksheets = currentsheet.worksheets()
        workingsheet = None
        for worksheet in worksheets:
            if worksheet.title == file_info["worksheet"]:
                workingsheet = worksheet
                break
        if workingsheet is None:
            print("Worksheet \""+file_info["worksheet"]+"\" is not found")
            if flags["autoretry"]:
                retry = 1
            continue
        if flags["pause_google"]:
            raw_input("[Google] Press 'enter' or 'return' to continue...")
        if not flags["silent"]:
            print("Gathering data from worksheet \""+workingsheet.title+"\"...")
        # totalrow = 0
        exclude_cols = ["Total", "Passed", "Withdrawn"]
        exclude_rows = ["Total Score", "Mean", "Min", "Max", "Next Update"]
        rows = workingsheet.get_all_values()
        selected_col = []
        selected_row = []
        if not flags["silent"]:
            print("Filtering columns...")
        for col in rows[0]:
            if col is not None and col != "" and not startsin(exclude_cols, col):
                selected_col.append(len(selected_col))
        if not flags["silent"]:
            print("Filtering rows...")
        for row in rows:
            if row is not None and row[0] != "" and not startsin(exclude_rows, row[0]):
                selected_row.append(len(selected_row))
        if not flags["silent"]:
            print("> Selected columns: " + str(len(selected_col)))
            print("> Selected rows: " + str(len(selected_row)))
            print("Tabularize data...")
        old_sheet = []
        for rownum in selected_row:
            orow = []
            row = rows[rownum]
            for colnum in selected_col:
                orow.append(row[colnum])
            old_sheet.append(str.join("\t", orow))
        # ======
        # Grader
        # ======
        br = mechanize.Browser()
        br.set_handle_robots(False)
        if flags["score"] is not None:
            if not flags["silent"]:
                print("Gathering result from file...")
            result_file = open(flags["score"], "r")
            result_html = BeautifulSoup(result_file.read())
            result_file.close()
        else:
            if not flags["silent"]:
                print("Gathering result from grader...")
                print("Connecting to grader...")
            try:
                br.open(GRADER_BASE)
            except Exception as msg:
                print("Grader Error! "+str(msg))
                if flags["autoretry"]:
                    retry = 1
                continue
            if len(list(br.forms())) < 1:
                print("No login form in grader... Please check the grader...")
                if flags["autoretry"]:
                    retry = 1
                continue
            br.form = list(br.forms())[0]
            br.form["login"] = file_info["grader_user"]
            br.form["password"] = file_info["grader_password"]
            if not flags["silent"]:
                print("Logging into grader...")
            if re.search("Wrong password", br.submit().read()) is not None:
                print("Invalid user or password.")
                return
            if flags["pause_grader"]:
                raw_input("[Grader] Press 'enter' or 'return' to continue...")
            if not flags["silent"]:
                print("Browsing grader results...")
            found = False
            for link in br.links():
                if link.url == "/user_admin/user_stat":
                    found = True
                    break
            if not found:
                print("Grader has no link to results page... Please check the grader...")
                if flags["autoretry"]:
                    retry = 1
                continue
            try:
                br.open(GRADER_BASE + "/user_admin/user_stat")
            except Exception as msg:
                print("Grader Results Error! "+str(msg))
                if flags["autoretry"]:
                    retry = 1
                continue
            if not flags["silent"]:
                print("Collecting grader results...")
            result_html = BeautifulSoup(br.response().read())
        result_table_html = result_html.find("table", class_="info")
        if result_table_html is None:
            print("Grader results page has no table... Please check the grader...")
            if flags["autoretry"]:
                retry = 1
            continue
        if flags["store_source"]:
            result_file = open(time.strftime("Result_" + os.path.splitext(os.path.basename(input_file))[0] + "_%Y%m%d%H%M%S.html", time.localtime()), "w")
            result_file.write(str(result_html))
            result_file.close()
        result_table = BeautifulSoup(str(result_table_html))
        result_rows = result_table("tr")
        grader_result = []
        if not flags["silent"]:
            print("Tabularize grader results...")
        for result_row in result_rows:
            result_cols = BeautifulSoup(str(result_row))
            result_col = result_cols("td")
            if len(result_col) <= 0:
                result_col = result_cols("th")
            orow = []
            for col in result_col:
                orow.append(trim(BeautifulSoup(str(col)).get_text()))
            grader_result.append(str.join("\t", orow))
        # ======
        # Merger
        # ======
        if not flags["silent"]:
            print("Connecting to merger...")

        try:
            br.open(UTILS_BASE+"/merger.php?api")
        except Exception as msg:
            print("Merger Error! "+str(msg))
            if flags["autoretry"]:
                retry = 1
            continue

        if len(list(br.forms())) < 1:
            print("No submit form in merger...")
            if flags["autoretry"]:
                retry = 1
            continue
        br.form = list(br.forms())[0]
        br.form["source1"] = str.join("\n", old_sheet).encode("utf8")
        br.form["source2"] = str.join("\n", grader_result).encode("utf8")
        br.form["no_columns"] = file_info["merger_no_columns"]
        br.form["no_rows"] = file_info["merger_no_rows"]
        br.form["keep_columns"] = file_info["merger_keep_columns"]
        br.form["keep_rows"] = file_info["merger_keep_rows"]
        if flags["pause_merger"]:
            raw_input("[Merger] Press 'enter' or 'return' to continue...")
        if not flags["silent"]:
            print("Merging...")
        merged_result = re.sub("<script.*$", "", br.submit().read())
        # ======
        # Normalizer
        # ======
        if not flags["silent"]:
            print("Connecting to normalizer...")
        try:
            br.open(UTILS_BASE+"/normalizer.php?api")
        except Exception as msg:
            print("Normalizer Error! "+str(msg))
            if flags["autoretry"]:
                retry = 1
            continue
        if len(list(br.forms())) < 1:
            print("No submit form in normalizer...")
            if flags["autoretry"]:
                retry = 1
            continue
        br.form = list(br.forms())[0]
        br.form["source"] = merged_result
        br.form["include_contests"] = file_info["normalizer_include_contests"]
        br.form["calculate_columns"] = file_info["normalizer_calculate_columns"]
        br.form["withdrawn_users"] = file_info["normalizer_withdrawn_users"]
        br.form["exclude_users"] = file_info["normalizer_exclude_users"]
        br.form["include_columns"] = file_info["normalizer_include_columns"]
        br.form["datetime"] = time.strftime("%d/%m/%Y %H:%M", time.localtime())
        if flags["pause_normalizer"]:
            raw_input("[Normalizer] Press 'enter' or 'return' to continue...")
        if not flags["silent"]:
            print("Normalizing...")
        normalized_result = trim(re.sub("<script.*$", "", br.submit().read()))
        if normalized_result.startswith("%error%"):
            print("Normalization Alert! "+normalized_result[7:])
            return
        # ======
        # Update
        # ======
        if not flags["silent"]:
            print("Converting results into list of lists...")
        normalized_result_list = normalized_result.split("\n")
        if len(normalized_result_list) < 1:
            print("Normalized results appear to have an invalid data...")
            if flags["autoretry"]:
                retry = 1
            continue
        if flags["n"] > 0:
            flags["n"] = flags["n"]-1
        result_list = []
        for row in normalized_result_list:
            result_list.append(row.split("\t"))
        if flags["n"] < 0 or flags["n"] > 0:
            result_list.append(["Next Update", time.strftime("%d/%m/%Y %H:%M", time.localtime(time.mktime(time.localtime()) + flags["l"]))])
        else:
            result_list.append([None, None])
        if not flags["silent"]:
            print("Gathering cells in worksheet...")
        range_from = workingsheet.get_addr_int(1, 1)
        range_to = workingsheet.get_addr_int(len(result_list), len(result_list[0]))
        if not flags["silent"]:
            print("From "+range_from+" to "+range_to+"...")
        result_cells = workingsheet.range(range_from+":"+range_to)
        for cell in result_cells:
            colnum = cell.col-1
            rownum = cell.row-1
            if rownum < len(result_list) and colnum < len(result_list[rownum]):
                if result_list[rownum][colnum] is None:
                    cell.value = ""
                elif trim(result_list[rownum][colnum]) != "":
                    cell.value = result_list[rownum][colnum].decode("utf8")
        if flags["pause_update"]:
            raw_input("[Update] Press 'enter' or 'return' to continue...")
        if not flags["silent"]:
            print("Updating results...")
        workingsheet.update_cells(result_cells)
        if not flags["silent"]:
            print("================")
            print("Results has been updated successfully on "+time.strftime("%d/%m/%Y %H:%M", time.localtime())+"... (in "+str(time.mktime(time.localtime())-started)+"s)")

        if flags["n"] < 0 or flags["n"] > 0:
            flags["first"] = False
            if flags["n"] > 0:
                if flags["n"] < 2:
                    print(str(flags["n"])+" update remaining...")
                else:
                    print(str(flags["n"])+" updates remaining...")

            timeremain = time.mktime(time.localtime())-started
            if timeremain < flags["l"]:
                print("Next update: "+time.strftime("%d/%m/%Y %H:%M", time.localtime(time.mktime(time.localtime()) + flags["l"])))

            while time.mktime(time.localtime())-started < flags["l"]:
                time.sleep(1)
            print("================")
            update(mode, input_file, flags)
            return
        # print("Results update successfully...")
    return

if __name__ == "__main__":
    run(sys.argv)
