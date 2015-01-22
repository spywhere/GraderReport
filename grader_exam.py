import getpass
import os
import re
import sys
import time
import threading
import traceback
import gspread
import mechanize
import argparse
try:
    import readline
    readline
except ImportError:
    pass
from bs4 import BeautifulSoup

GRADER_BASE = "http://grader.eng.src.ku.ac.th"

GOOGLE_EMAIL = ""
GOOGLE_PASSWORD = ""
SPREADSHEET_KEY = ""
CONTEST_WORKSHEET = "Contest Type"
LOG_WORKSHEET = "Logs"

SUBMISSION_PROBLEM_ID = "submission_problem_id"


class LogThread(threading.Thread):
    def __init__(self, logsheet, adminOptions, username, browser):
        self.logsheet = logsheet
        self.adminOptions = adminOptions
        self.username = username
        self.browser = browser
        self.problems = None
        self.acceptPattern = None
        self.delay = 5
        threading.Thread.__init__(self)
        self.running = False
        self.quit = False
        self.start()

    def updateInfo(self, problems, acceptPattern):
        self.problems = problems
        self.acceptPattern = acceptPattern

    def stop(self):
        self.logout()

    def isQuit(self):
        return self.quit

    def remark(self, remark):
        found = False
        selectedRow = 1
        for value in self.logsheet.col_values(1):
            if value == self.username:
                found = True
                break
            selectedRow += 1

        if not found:
            return

        cell = self.logsheet.cell(selectedRow, 4)
        cell.value = remark
        self.logsheet.update_cells([cell])

    def adminCommand(self):
        if self.adminOptions is None:
            return
        if self.adminOptions.remark is not None:
            self.remark(self.adminOptions.remark)

    def logout(self):
        found = False
        selectedRow = 1
        for value in self.logsheet.col_values(1):
            if value == self.username:
                found = True
                break
            selectedRow += 1

        if not found:
            return

        cell = self.logsheet.cell(selectedRow, 4)
        cell.value = "Logout on " + time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
        self.logsheet.update_cells([cell])
        self.running = False

    def command(self):
        started = time.mktime(time.localtime())
        while self.running:
            try:
                found = False
                selectedRow = 1
                for value in self.logsheet.col_values(1):
                    if value == self.username:
                        found = True
                        break
                    selectedRow += 1

                if not found:
                    continue

                range_from = self.logsheet.get_addr_int(selectedRow, 1)
                range_to = self.logsheet.get_addr_int(selectedRow, 6)
                cells = self.logsheet.range(range_from+":"+range_to)

                infoText = ""
                commandResponse = ""
                for cell in cells:
                    if cell.col == 5:
                        command = re.sub(";.*", "", cell.value)
                        if command == "":
                            break
                        if self.problems is None or self.acceptPattern is None:
                            commandResponse = "Command is not ready"
                            continue
                        if command.startswith("os "):
                            commandResponse = trim(os.popen(command[3:]).read())
                        elif command.startswith("bg "):
                            commandResponse = run_command(command[3:], self.browser, self.problems, self.acceptPattern, background=True)
                        else:
                            run_command(command, self.browser, self.problems, self.acceptPattern, background=False)
                            printProblems(self.problems, self.username, background=True)
                            commandResponse = "Run as foreground"
                        if commandResponse is None:
                            self.running = False
                            self.quit = True
                            commandResponse = "Logout"
                        elif type(commandResponse) is bool and not commandResponse:
                            commandResponse = "Invalid command."
                        cell.value = ";Run on " + time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
                    elif cell.col == 6:
                        cell.value = commandResponse.replace("=", "-")
                self.logsheet.update_cells(cells)
                while time.mktime(time.localtime())-started < self.delay:
                    time.sleep(1)
                started = time.mktime(time.localtime())
            except (KeyboardInterrupt, SystemExit):
                self.running = False
                return
            except Exception:
                continue

    def run(self):
        self.running = True
        while self.running:
            try:
                found = False
                selectedRow = 1
                for value in self.logsheet.col_values(1):
                    if selectedRow == 1:
                        selectedRow += 1
                        continue
                    if value == "" or value == self.username:
                        break
                    selectedRow += 1

                if selectedRow == 1:
                    selectedRow = 2
                range_from = self.logsheet.get_addr_int(1, 1)
                range_to = self.logsheet.get_addr_int(selectedRow, 6)
                cells = self.logsheet.range(range_from+":"+range_to)

                infoText = ""
                for cell in cells:
                    if cell.row == 1:
                        if cell.col == 1:
                            cell.value = "User"
                        elif cell.col == 2:
                            cell.value = "Number of Logins"
                        elif cell.col == 3:
                            cell.value = "Last Login"
                        elif cell.col == 4:
                            cell.value = "Remark"
                        elif cell.col == 5:
                            cell.value = "Command"
                        elif cell.col == 6:
                            cell.value = "Command Response"
                    elif cell.row == selectedRow:
                        if cell.col == 1 and cell.value == "":
                            cell.value = self.username
                        elif cell.col == 2:
                            oldValue = cell.value
                            if oldValue == "":
                                cell.value = "1"
                            elif parsable(oldValue):
                                cell.value = str(int(oldValue)+1)
                            else:
                                infoText = "Invalid: " + cell.value
                                cell.value = "1"
                        elif cell.col == 3:
                            cell.value = time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())
                        elif cell.col == 4:
                            cell.value = infoText
                    elif cell.row > selectedRow:
                        break
                self.logsheet.update_cells(cells)
            except (KeyboardInterrupt, SystemExit):
                self.running = False
                return
            except Exception:
                continue
            self.adminCommand()
            self.command()
        clear()


def clear():
    os.system("cls" if os.name == "nt" else "clear")

def parsable(string):
    try:
        int(string)
        return True
    except Exception:
        return False
    return False


def trim(st):
    if st.startswith(" ") or st.startswith("\n") or st.startswith("\t"):
        return trim(st[1:])
    elif st.endswith(" ") or st.endswith("\n") or st.endswith("\t"):
        return trim(st[:-1])
    return st

def run():
    parser = argparse.ArgumentParser(description="Cafe Grader Exam System.")
    parser.add_argument("-a", "--admin", dest="admin", action="store_true", default=False, help="enable administrator special commands")
    parser.add_argument("-v", "--verbose", dest="verbose", action="store_true", default=False, help="show all messages in details")
    parser.add_argument("username", nargs="?", type=str, help="username to pre-entered")
    # parser = optparse.OptionParser(usage="usage: %prog [options] [username]")
    # parser.set_defaults(admin=False, verbose=False)
    # parser.add_option("-a", "--admin", action="store_true", dest="admin", help="Enable administrator special commands")
    # parser.add_option("-v", "--verbose", action="store_true", dest="verbose", help="Show all messages in details")
    options = parser.parse_args()
    while True:
        ret = subrun(options)
        if "options" in ret:
            options = ret["options"]
        if not ret["return"]:
            break

def run_command(command, browser, problems, acceptPattern, background=False):
    if not background:
        clear()
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("-p", "--problem", dest="submit_problem", nargs="?", type=int)
    parser.add_argument("-f", "--file", dest="submit_file", nargs="?" , type=str)
    parser.add_argument("arg1", nargs="?", type=str)
    parser.add_argument("arg2", nargs="?" , type=str)
    cmdoptions = parser.parse_args(command.split())

    if cmdoptions.submit_problem is None:
        if cmdoptions.arg1 is not None and parsable(cmdoptions.arg1):
            cmdoptions.submit_problem = int(cmdoptions.arg1)
            cmdoptions.arg1 = None
    if cmdoptions.submit_file is None:
        if cmdoptions.arg1 is not None and not parsable(cmdoptions.arg1) and os.path.exists(cmdoptions.arg1):
            cmdoptions.submit_file = cmdoptions.arg1
            cmdoptions.arg1 = None
        elif cmdoptions.arg2 is not None and not parsable(cmdoptions.arg2) and os.path.exists(cmdoptions.arg2):
            cmdoptions.submit_file = cmdoptions.arg2
            cmdoptions.arg2 = None
    if cmdoptions.arg1 is not None:
        if cmdoptions.arg1.lower() == "help":
            if background:
                return "Help is not support in admin command"
            else:
                print("==== HELP ====")
                print("Type \"[File Path]\" to submit file to grader.")
                print("Type \"[Problem Number]\" to view compiler message.")
                print("Type \"[Problem Number] [File Path]\" to submit file to correspond problem.")
                print("Type \"logout\" to logout")
                print("Type nothing to refresh")
        elif cmdoptions.arg1.lower() == "logout":
            if background:
                return None
            else:
                return False
        else:
            if background:
                return False
            else:
                print("Invalid command. Type \"help\" for help.")
    submitmsg = ""
    if cmdoptions.submit_file is not None:
        if re.match(acceptPattern, cmdoptions.submit_file) is None:
            if background:
                return "File \"" + cmdoptions.submit_file + "\" is not allowed to submit"
            else:
                print("File \"" + cmdoptions.submit_file + "\" is not allowed to submit")
                return True
        if cmdoptions.submit_problem is None:
            cmdoptions.submit_problem = -1
            if background:
                submitmsg = "Submitting file \"" + cmdoptions.submit_file + "\"..."
            else:
                print("Submitting file \"" + cmdoptions.submit_file + "\"...")
        else:
            if cmdoptions.submit_problem < 1:
                if background:
                    return "Problem number must greater or equal to 1"
                else:
                    print("Problem number must greater or equal to 1")
                    return True
            elif cmdoptions.submit_problem > len(problems):
                if background:
                    return "Problem number is exceed total problems"
                else:
                    print("Problem number is exceed total problems")
                    return True
            if background:
                submitmsg = "Submitting file \"" + cmdoptions.submit_file + "\" to problem \"" + problems[cmdoptions.submit_problem-1]["name"] + "\"..."
            else:
                print("Submitting file \"" + cmdoptions.submit_file + "\" to problem \"" + problems[cmdoptions.submit_problem-1]["name"] + "\"...")
    elif cmdoptions.submit_problem is not None:
        if cmdoptions.submit_problem < 1:
            if background:
                return "Problem number must greater or equal to 1"
            else:
                print("Problem number must greater or equal to 1")
                return True
        elif cmdoptions.submit_problem > len(problems):
            if background:
                return "Problem number is exceed total problems"
            else:
                print("Problem number is exceed total problems")
                return True
        if not background:
            print("Gathering compiler message...")
        compiler_msg_link = problems[cmdoptions.submit_problem-1]["compiler_msg"]
        if compiler_msg_link is None:
            if background:
                return "No compiler message yet"
            else:
                clear()
                print("No compiler message yet")
        else:
            try:
                response = browser.open(GRADER_BASE+compiler_msg_link)
            except Exception as msg:
                if background:
                    return "Grader Error! "+str(msg)
                else:
                    print("Grader Error! "+str(msg))
                    traceback.print_exc()
                    return True
            if not background:
                clear()
            compiler_page = BeautifulSoup(browser.response().read())
            titlelength = 0
            bgmsg = ""
            for child in compiler_page.find("body").children:
                if child.name == "h2":
                    titlelength = len(child.string)
                    if background:
                        bgmsg += "==== %s ====" % (trim(child.string))
                    else:
                        print("==== %s ====" % (trim(child.string)))
                elif child.name == "p":
                    # print(child)
                    for p in child.children:
                        if p.name == "p":
                            for line in p.children:
                                if line.name != "br":
                                    if background:
                                        bgmsg += "  "+trim(line)
                                    else:
                                        print("  "+trim(line))
                        elif p.name != "br" and trim(str(p)) != "":
                            print("  "+trim(str(p)))
            if background:
                return bgmsg
            else:
                getpass.getpass("")
                clear()
        return True
    elif cmdoptions.submit_problem is None and cmdoptions.submit_file is None:
        # Refresh
        if background:
            return ""
        else:
            return True

    browser.form = list(browser.forms())[0]
    browser.form["submission[problem_id]"] = [str(problems[cmdoptions.submit_problem-1]["id"])]
    f = open(cmdoptions.submit_file)
    browser.form.add_file(open(cmdoptions.submit_file), 'text/plain', cmdoptions.submit_file)
    browser.submit()
    if background:
        return submitmsg
    else:
        return True

def printProblems(problems, username, background=False):
    print("==== [%s] %s ====" % (username, time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())))
    print("Total %s problems" % (len(problems)))
    problemNumber = 0
    for problem in problems:
        problemNumber += 1
        print("%2d> %s" % (problemNumber, problem["name"]))
        if problem["description"] is not None:
            print("    Description: %s" % (problem["description"]))
        print("    Results: %s" % (problem["status"]))
    if background:
        sys.stdout.write("> ")
        sys.stdout.flush()

def subrun(options):
    logThread = None
    try:
        adminOptions = None
        if options.admin:
            parser = argparse.ArgumentParser()
            parser.add_argument("-r", "--remark", dest="remark")
            adminOptions = parser.parse_args(getpass.getpass("> ").split())
            # return {"return": False}
        if options.username is None:
            options.username = raw_input("User: ")
        else:
            print("User: %s" % (options.username))
        if options.username == "":
            return {"return": False}

        client = None
        if options.verbose:
            print("Logging into GMail...")
        else:
            print("Please wait...")
        try:
            client = gspread.login(GOOGLE_EMAIL, GOOGLE_PASSWORD)
        except gspread.AuthenticationError, e:
            if options.verbose:
                print("Invalid email or password.")
            else:
                print("Please contact grader administrator [gcredential]")
            return {"return": False}
        if options.verbose:
            print("Opening spreadsheet...")
        try:
            currentsheet = client.open_by_key(SPREADSHEET_KEY)
        except gspread.SpreadsheetNotFound, e:
            if options.verbose:
                print("Spreadsheet is not found")
            else:
                print("Please contact grader administrator [infosheet]")
            return {"return": False}
        except Exception as e:
            if options.verbose:
                print("Google Error! "+str(e))
                traceback.print_exc()
            else:
                print("Unexpected error occurred. Maybe internet connection failed?")
            print("Retry in 5 seconds")
            time.sleep(5)
            return {"return": True, "options": options}
        if options.verbose:
            print("Opening worksheet in \""+currentsheet.title+"\"...")
        worksheets = currentsheet.worksheets()
        contestsheet = None
        logsheet = None
        for worksheet in worksheets:
            if contestsheet is not None and logsheet is not None:
                break
            if worksheet.title == CONTEST_WORKSHEET:
                contestsheet = worksheet
            elif worksheet.title == LOG_WORKSHEET:
                logsheet = worksheet
        if contestsheet is None:
            if options.verbose:
                print("Worksheet \""+CONTEST_WORKSHEET+"\" is not found")
            else:
                print("Please contact grader administrator [contest]")
            return {"return": False}
        if logsheet is None:
            if options.verbose:
                print("Worksheet \""+LOG_WORKSHEET+"\" is not found")
            else:
                print("Please contact grader administrator [log]")
            return {"return": False}
        if options.verbose:
            print("Gathering data from worksheet \""+contestsheet.title+"\"...")
        rows = contestsheet.get_all_values()
        acceptPattern = ".*"
        loginMode = "close"
        serializeRows = []
        row_size = 0
        for row in rows:
            row_size = max(row_size, len(row))
            for cell in row:
                if cell.startswith(";"):
                    continue
                serializeRows.append(cell)
        currentCell = 0
        overriding = False
        endOverride = row_size
        while currentCell < len(serializeRows):
            cell = serializeRows[currentCell]
            if overriding and cell == "":
                endOverride -= 1
                if endOverride == 0:
                    break
            else:
                endOverride = row_size
            if overriding and cell == options.username:
                loginMode = serializeRows[currentCell+1].lower()
                break
            elif not overriding and cell == "Default Mode":
                loginMode = serializeRows[currentCell+1].lower()
                currentCell += 1
            elif not overriding and cell == "Accept File Pattern":
                acceptPattern = serializeRows[currentCell+1]
                currentCell += 1
            elif not overriding and cell == "Override Mode":
                overriding = True

            currentCell += 1
        userPassword = None
        availableMode = ["hw", "exam"]
        if loginMode.lower() not in availableMode:
            print("GraderExam is disabled right now. Please contact grader administrator.")
            return {"return": False}
        if loginMode == "exam":
            currentCell = 0
            userListing = False
            endListing = row_size
            while currentCell < len(serializeRows):
                cell = serializeRows[currentCell]
                if userListing and cell == "":
                    endListing -= 1
                    if endListing == 0:
                        break
                else:
                    endListing = row_size
                if userListing and cell == options.username:
                    userPassword = serializeRows[currentCell+1]
                    break
                elif not userListing and cell == "User List":
                    userListing = True
                currentCell += 1
        grader_result = ""
        while True:
            if loginMode == "hw":
                userPassword = getpass.getpass("Password: ")
            elif userPassword is None:
                if options.verbose:
                    print("User is not found in user list")
                else:
                    print("User not found. Press 'Enter' to exit (without typing anything).")
                return {"return": True}

            br = mechanize.Browser()
            br.set_handle_robots(False)
            try:
                br.open(GRADER_BASE)
            except Exception as msg:
                print("Grader Error! "+str(msg))
                traceback.print_exc()
                return {"return": False}
            if len(list(br.forms())) < 1:
                print("No login form in grader... Please check the grader...")
                return {"return": False}
            br.form = list(br.forms())[0]
            br.form["login"] = options.username
            br.form["password"] = userPassword
            print("Logging into grader...")
            if re.search("Wrong password", br.submit().read()) is not None:
                print("Wrong password")
                continue
            break
        logThread = LogThread(logsheet, adminOptions, options.username, br)
        clear()
        while True:
            try:
                response = br.open(GRADER_BASE+"/main/list")
            except Exception as msg:
                print("Grader Error! "+str(msg))
                traceback.print_exc()
                continue
            grader_page = BeautifulSoup(br.response().read())
            submission = grader_page.find("select", id=SUBMISSION_PROBLEM_ID)
            problems = submission.find_all("option")
            problems_list = []
            for problem in problems:
                if int(problem["value"]) < 0:
                    continue
                problems_list.append({"id": int(problem["value"]), "name": problem.string})

            problems_dict = {}
            infoTables = grader_page.find_all("table")
            for table in infoTables:
                if "class" not in table.attrs:
                    continue
                elif "info" not in table.attrs["class"]:
                    continue
                for row in table.find_all("tr"):
                    if "info-head" in row.attrs["class"]:
                        continue
                    problem = {"name": "", "description": None, "status": "", "compiler_msg": None}
                    cellNumber = 0
                    for cell in row.find_all("td"):
                        if cellNumber == 1:
                            problemName = ""
                            for child in cell.children:
                                if child.name == "a":
                                    problem["description"] = child["href"]
                                else:
                                    problemName += trim(child.string).replace("\n", " ")
                            problemName = re.sub("\\s{2,}", " ", problemName)
                            problemName = re.sub("\\s*\\|.*", "", problemName)

                            problem["name"] = problemName
                        elif cellNumber == 3:
                            status = ""
                            for child in cell.children:
                                if child.name == "a":
                                    if child.string == "[compiler msg]":
                                        problem["compiler_msg"] = child["href"]
                                    continue
                                status += trim(child.string).replace("\n", " ")
                            status = re.sub("\\s{2,}", " ", status)
                            status = re.sub("\\s*\\|.*", "", status)
                            problem["status"] = status
                        cellNumber += 1
                    matched_problem = None
                    for p in problems_list:
                        if p["name"] in problem["name"]:
                            matched_problem = p
                            break
                    if matched_problem is not None:
                        problem["id"] = matched_problem["id"]
                        problems_dict[matched_problem["id"]] = problem

            problems = []
            for problem in problems_dict:
                problems.append(problems_dict[problem])

            printProblems(problems, options.username)

            logThread.updateInfo(problems, acceptPattern)

            if (logThread is not None and logThread.isQuit()) or not run_command(raw_input("> "), br, problems, acceptPattern):
                print("Logging out...")
                if logThread is not None:
                    logThread.stop()
                break
        return {"return": False}
    except (KeyboardInterrupt, SystemExit):
        if logThread is not None:
            logThread.stop()
        return {"return": False}
    except Exception as e:
        print("Fetal error: " + str(e))
        traceback.print_exc()
        if logThread is not None:
            logThread.stop()
        return {"return": False}


if __name__ == "__main__":
    run()
