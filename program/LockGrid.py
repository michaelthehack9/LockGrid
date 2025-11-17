import os
import pyodbc
import pandas as pd
import warnings

base_dir = os.path.dirname(os.path.abspath(__file__))

db_file = os.path.join(base_dir, "lockgrid.accdb")

conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={db_file};"
)

session = None

def main():
    global session
    conn = pyodbc.connect(conn_str)
    running = True
    while (running):
        os.system("cls")
        option = input("Welcome to LockGrid!\n\n1. Exit\n2. AM\n3. PM\n4. ALL\n\nEnter an option: ").strip()
        match(option):
            case "1":
                running = False
            case "2":
                session = "AM"
            case "3":
                session = "PM"
            case "4":
                session = "ALL"

        if (running and session is not None):
            menu(conn)
            session = None
    conn.close()

def menu(conn):
    running = True
    while (running):
        os.system("cls")
        print(f"Welcome to LockGrid! ({session})\n\nWhat would you like to do:\n\n1. Exit\n2. Check Combo\n3. Assign\n4. Check Student\n5. Check Locker\n6. Get Report\n7. Unassign Locker\n8. Unassign All\n")
        option = input("Enter an option: ").strip()
        os.system("cls")
        match option:
            case "1":
                running = False
            case "2":
                checkCombo(conn)
            case "3":
                assign(conn)
            case "4":
                checkStudent(conn)
            case "5":
                checkLocker(conn)
            case "6":
                getReport(conn)
            case "7":
                unassignLocker(conn)
            case "8":
                unassignAll(conn)

def checkCombo(conn):
    serial = ""

    running = True

    while (running):
        serial = input("Serial Number (-1 to cancel): ").strip()
        if (len(serial) == 8 or serial == "-1"):
            running = False
        else:
            os.system("cls")
            print("Please enter a valid serial number!\n")

    if serial != "-1":
        cur = conn.cursor()

        cur.execute(f"SELECT combo FROM locks WHERE serial = '{serial}'")
        row = cur.fetchone()
        
        if (row is None):
            print("\nLock or Combo not found!\n")
        else:
            os.system("cls")
            print(f"Serial: {serial}\n\nCombo: {row[0]}\n")

        cur.close()

        os.system("pause")

def assign(conn):
    running = True

    while (running):
        student = getStudentID(conn)
        
        if (student == -1):
            running = False
        else:

            cur = conn.cursor()

            cur.execute(f"SELECT fname, lname FROM students WHERE ID = {student}")
            row = cur.fetchone()

            if (row is None):
                os.system("cls")
                print("Student does not exist!\n")
            else:
                fname = row[0]
                lname = row[1]
                cur.execute(f"SELECT student_ID FROM lockers WHERE student_ID = {student}")
                row = cur.fetchone()

                if (row is not None):
                    os.system("cls")
                    print("Student already has a locker!\n")
                else:
                    os.system("cls")
                    
                    serial = ""

                    while (serial == ""):
                        tempSerial = input("Serial Number: ").strip()
                        if (tempSerial.isnumeric() and len(tempSerial) == 8):
                            serial = tempSerial

                            cur.execute(f"SELECT serial FROM locks WHERE serial = '{serial}'")
                            row = cur.fetchone()

                            if (row is None):
                                os.system("cls")
                                print("Lock Serial does not exist!\n")
                                serial = ""
                            else:
                                cur.execute(f"SELECT lock_ID FROM lockers WHERE lock_ID = '{serial}'")
                                row = cur.fetchone()

                                if (row is not None):
                                    os.system("cls")
                                    print("Lock Serial already assigned to someone!\n")
                                    serial = ""
                                else:
                                    os.system("cls")
                                    locker = -1

                                    while (locker == -1):
                                        strLocker = input("Locker Number: ").strip()
                                        if (strLocker.isnumeric()):
                                            locker = int(strLocker)

                                            cur.execute(f"SELECT ID FROM lockers WHERE ID = {locker}")
                                            row = cur.fetchone()

                                            if (row is None):
                                                os.system("cls")
                                                print("Locker does not exist!\n")
                                                locker = -1
                                            else:
                                                cur.execute(f"SELECT ID FROM lockers WHERE ID = {locker} AND student_ID IS NULL")
                                                row = cur.fetchone()

                                                if (row is None):
                                                    os.system("cls")
                                                    print("Locker already assigned to someone!\n")
                                                    locker = -1
                                                else:
                                                    os.system("cls")
                                                    cur.execute(f"UPDATE lockers SET student_ID = {student}, lock_ID = {serial} WHERE ID = {locker}")
                                                    rows = cur.rowcount
                                                    conn.commit()

                                                    if (rows > 0):
                                                        print(f"Student: {formatName(fname, lname)}\n\nSerial Number: {serial}\nLocker Number: {locker}\n\nSuccessfully Added!\n")
                                                        os.system("pause")
                                                    else:
                                                        print("Unknown Failure!\n")
                                                        os.system("pause")

                                                    os.system("cls")
                                        else:
                                            os.system("cls")
                                            print("Please enter a valid Locker Number!\n")
                        else:
                            os.system("cls")
                            print("Please enter a valid Lock Serial!\n")
            cur.close()

def checkStudent(conn):
    running = True

    while (running):
        student = getStudentID(conn)

        if (student == -1):
            running = False
        else:
            os.system("cls")
            cur = conn.cursor()

            cur.execute(f"SELECT fname, lname FROM students WHERE ID = {student}")
            row = cur.fetchone()
            
            if (row is None):
                print("Student does not exist!")
            else:
                print(f"Name: {formatName(row[0], row[1])}\n")

                cur.execute(f"SELECT ID, lock_ID FROM lockers WHERE student_ID = {student}")
                row = cur.fetchone()

                if (row is None):
                    print("Student does not have a locker!\n")
                else:
                    print(f"Locker: {row[0]}\nLock Serial: {row[1]}")

                    if (row[1] is None):
                        print("Student does not have a lock!\n")
                    else:
                        cur.execute(f"SELECT combo FROM locks WHERE serial = '{row[1]}'")
                        row = cur.fetchone()

                        if (row is None):
                            print("Combo not found for lock!\n")
                        else:
                            print(f"Combo: {row[0]}\n")
                        os.system("pause")
                        os.system("cls")

            cur.close()

def checkLocker(conn):
    locker = 0
    running = True

    while (running):
        strLocker = input("Locker Number (-1 to cancel): ").strip()
        if (strLocker.isnumeric() or strLocker == "-1"):
            locker = int(strLocker)
            running = False
        else:
            os.system("cls")
            print("Please enter a valid locker number!\n")

    if locker != -1:
        cur = conn.cursor()

        cur.execute(f"SELECT student_ID, lock_ID FROM lockers WHERE ID = {locker}")
        row = cur.fetchone()
        
        if (row is None):
            print("\nLocker not found!\n")
        else:
            lockSerial = row[1]

            if (row[0] is None):
                print("\nNo Student Assigned!\n")
            else:
                cur.execute(f"SELECT fname, lname FROM students WHERE ID = {row[0]}")
                row = cur.fetchone()

                fname = row[0]
                lname = row[1]

                if (lockSerial is None):
                    print("\nNo Lock Assigned!\n")
                else:
                    cur.execute(f"SELECT combo FROM locks WHERE serial = '{lockSerial}'")
                    row = cur.fetchone()

                    if (row is None):
                        print("\nCombo not found!")

                print(f"\nStudent: {formatName(fname, lname)}\n\nLocker: {locker}\nLock Serial: {lockSerial}\nCombo: {row[0]}\n")

        cur.close()

        os.system("pause")

def getReport(conn):
    commandText = "SELECT lockers.ID, students.fname, students.lname, lockers.lock_ID, locks.combo"
    whereText = f"WHERE students.session = '{session.lower()}'"

    if (session == "ALL"):
        commandText += ",students.session"
        whereText = ""
    
    commandText += f" FROM (lockers INNER JOIN students ON lockers.student_ID = students.ID) LEFT JOIN locks ON lockers.lock_ID = locks.serial {whereText} ORDER BY lockers.ID ASC"

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=UserWarning)
        df = pd.read_sql(commandText, conn)
        if (session == "ALL"):
            df.columns = ["Locker", "First Name", "Last Name", "Lock Serial Number", "Lock Combination", "Session"]
        else:
            df.columns = ["Locker", "First Name", "Last Name", "Lock Serial Number", "Lock Combination"]
        
        df.to_excel(os.path.join(os.path.dirname(__file__), f"../locker_report_{session}.xlsx"), index=False)

    print("Exported Successfully!\n")

    os.system("pause")

def unassignLocker(conn):
    locker = -2

    while (locker == -2):
        strLocker = input("What locker would you like to unassign (-1 to cancel): ")

        if (strLocker.isnumeric() or strLocker == "-1"):
            locker = int(strLocker)
            
            os.system("cls")

            if (locker != -1):
                cur = conn.cursor()
                cur.execute(f"UPDATE lockers SET student_ID = NULL, lock_ID = NULL WHERE ID = {locker}")
                rows = cur.rowcount
                conn.commit()
                cur.close()

                if (rows > 0):
                    print(f"Successfully unassigned locker {locker}!\n")
                else:
                    print(f"Failed to unassign locker {locker}!\n")
                    locker = -2
            
                os.system("pause")  
            
            os.system("cls")
        else:
            os.system("cls")
            print("Locker is not valid!\n")

def unassignAll(conn):
    print("WARNING: This will UNASSIGN ALL LOCKERS in the ENTIRE database!\n")
    confirmation = input("Please type 'CONFIRM' to unassign all: ").strip()

    if (confirmation == "CONFIRM"):
        cur = conn.cursor()

        cur.execute("UPDATE lockers SET student_ID = NULL, lock_ID = NULL")
        rows = cur.rowcount
        conn.commit()
        cur.close()

        if (rows > 0):
            print("\nSuccessfully unassigned all!\n")
        else:
            print("\nFailed to unassign all!\n")
    else:
        print("\nCancelling Operation!\n")

    os.system("pause")

def getStudentID(conn):
    studentID = -1

    running = True

    while (running):
        fname = input("Enter First Name (-1 to stop): ").strip()
        
        if (fname == "-1"):
            running = False
        else:
            if (len(fname) > 0):
                cur = conn.cursor()
                commandText = f"SELECT ID, fname, lname FROM students WHERE fname LIKE '{fname}%'"
                if (session != "ALL"):
                    commandText += f" AND session = '{session.lower()}'"
                cur.execute(commandText)
                results = cur.fetchall()

                if (len(results) == 1):
                    os.system("cls")

                    check = True

                    while (check):
                        confirmation = input(f"Name: {formatName(results[0].fname, results[0].lname)} (y or n): ").strip()
                        if (confirmation.lower() == "y"):
                            studentID = int(results[0].ID)
                            check = False
                            running = False
                        elif (confirmation.lower() == "n"):
                            os.system("cls")
                            check = False
                        else:
                            os.system("cls")
                            print("Invalid option, Try again!\n")

                elif (len(results) > 1):
                    chooseStudent = True

                    while (chooseStudent):
                        os.system("cls")
                        print("Multiple students found:\n")

                        for i, row in enumerate(results, start=1):
                            print(f"{i}. {row.fname} {row.lname} (ID: {row.ID})")

                        strChoice = input("\nPlease choose a student (-1 to cancel): ").strip()

                        if (strChoice.isnumeric() or strChoice == "-1"):
                            os.system("cls")

                            choice = int(strChoice)
                            if (choice == -1):
                                chooseStudent = False
                            elif (choice >= 1 and choice <= len(results)):
                                check = True

                                while (check):
                                    confirmation = input(f"Name: {formatName(results[choice - 1].fname, results[choice - 1].lname)} (y or n): ").strip()
                                    if (confirmation.lower() == "y"):
                                        studentID = int(results[choice - 1].ID)
                                        check = False
                                        chooseStudent = False
                                        running = False
                                    elif (confirmation.lower() == "n"):
                                        os.system("cls")
                                        check = False
                                    else:
                                        os.system("cls")
                                        print("Invalid option, Try again!\n")
                            else:
                                os.system("cls")
                                print("Invalid option, Try Again!")
                        else:
                            os.system("cls")
                            print("Choice is not a number!\n")
                else:
                    os.system("cls")
                    print("No Students Found!\n")
            else:
                os.system("cls")
                print("Please enter a name!\n")

    return studentID

def formatName(fname, lname):
    return f"{fname.capitalize()} {lname.capitalize()}"

main()
