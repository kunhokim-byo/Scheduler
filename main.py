import os
import xlsxwriter
from datetime import datetime

from Moduleoffering import Moduleoffering
from Module import Module
from Schedule import Schedule
from Location import Location
from Program import Program

from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm, inch
from reportlab.lib.utils import ImageReader

moduleOfferings = []  # list of Moduleoffering objects
modules = []  # list of Module objects
schedules = []  # list of Schedule objects
locations = []  # List of Locaton objects
programs = []  # List of Program objects


def repeatedFilter(filteredSchedules):
    print("1 List By Lecturer")
    print("2 List By Location")
    print("3 List By Date")
    print("4 List By Date Range")
    print("5 List By Time Range")
    print("6 List BY Day")
    print("7 List By Module")
    print("8 List By Room")
    selectedListOption = int(input("Select the list option: "))
    while selectedListOption not in [i for i in range(1, 9)]:  # Used generator expression
        selectedListOption = int(input("Invalid input. Select the list option again: "))
    if selectedListOption == 1:
        inputValue = input("Type the lecturer: ").strip()
        return listByLecturer(inputValue, filteredSchedules)
    elif selectedListOption == 2:
        inputValue = input("Type the location: ").strip()
        return listByLocation(inputValue, filteredSchedules)
    elif selectedListOption == 3:
        inputValue = input("Type the Date(dd/mm/year): ").strip()
        return listByDate(inputValue, filteredSchedules)
    elif selectedListOption == 4:
        startDate = input("Type the start date(dd/mm/year): ").strip()
        endDate = input("Type the end date(dd/mm/year): ").strip()
        return listByDateRange(startDate, endDate, filteredSchedules)
    elif selectedListOption == 5:
        startTime = input("Type the start time(e.g. 8am = 0800, 6pm = 1800): ").strip()
        endTime = input("Type the end time(e.g. 8am = 0800, 6pm = 1800): ").strip()
        return listByTimeRange(startTime, endTime, filteredSchedules)
    elif selectedListOption == 6:
        inputValue = input("Type the day(Mon/Tue/Wed/Thu/Fri/Sat/Sun): ").strip()
        return listByDay(inputValue, filteredSchedules)
    elif selectedListOption == 7:
        inputValue = input("Type the module code: ").strip()
        return listByModule(inputValue, filteredSchedules)
    elif selectedListOption == 8:
        inputValue = input("Type the room: ").strip()
        return listByRoom(inputValue, filteredSchedules)


# { }
# 1. Moo1 S1
# { Moo1: [S1] }
# 2. Moo1 S2
# { Moo1: [S1, S2] }


def listAllSchedule():
    out = {}
    for moo in moduleOfferings:  # moo = object of moduleOffering
        for s in moo.getSchedules():  # moo.getSchedules() = get list of Schedule objects inside of moo object
            print(moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                               1:-1])  # remove curly bracket
            if moo in out:
                out[moo].append(s)  # if moo already exist in out, append only schedule.
            else:
                out[moo] = [s]  # if moo not exist, append both.
    return out


def listAllModuleandlecturer():  # to show the list of psb modules and psb lecturers to user
    i = 0
    k = 0
    print("[PSB Modules]")
    for m in modules:
        i = i + 1
        print(i, m)
    not_duplicate = []
    for schedule in schedules:
        if schedule.getLecturerName() not in not_duplicate:
            not_duplicate.append(schedule.getLecturerName())
    print("[PSB Lecturers]")
    for d in not_duplicate:
        k = k + 1
        print(k, " " + d)


def listByModule(name, str_schedules=None):
    if str_schedules == None:  # Apply only one filter/ first filter
        out = {}
        for moo in moduleOfferings:
            if moo.getModule().getModuleCode() == name:
                for s in moo.getSchedules():
                    # If module name == name, then print/append to output list
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])  # remove bracket
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]

        return out

    else:  # if str_schedules is not None, Apply second filter
        out = {}
        for moo, schedule_list in str_schedules.items():  # str_schedules = {moo1 : S1, moo2 : S2, moo3 : S3 ...}
            if moo.getModule().getModuleCode() == name:
                for s in schedule_list:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByLecturer(name, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                if s.getLecturerName() == name:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                if s.getLecturerName() == name:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByLocation(location, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                if s.getLocation().getZone() == location:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                if s.getLocation().getZone() == location:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByRoom(room, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                if s.getLocation().getRoom() == room:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                if s.getLocation().getRoom() == room:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByDate(date, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                if s.getDate() == date:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                if s.getDate() == date:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
        return out


def listByDateRange(start, end, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                format = '%d/%m/%Y'
                start_dt = datetime.strptime(start, format)
                end_dt = datetime.strptime(end, format)
                if datetime.strptime(s.getDate(), format) >= start_dt and datetime.strptime(s.getDate(),
                                                                                            format) <= end_dt:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                format = '%d/%m/%Y'
                start_dt = datetime.strptime(start, format)
                end_dt = datetime.strptime(end, format)
                if datetime.strptime(s.getDate(), format) >= start_dt and datetime.strptime(s.getDate(),
                                                                                            format) <= end_dt:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByTimeRange(start, end, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                a = s.getStartTime()
                c = s.getEndTime()
                if a >= start and c <= end:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                a = s.getStartTime()
                c = s.getEndTime()
                if a >= start and c <= end:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def listByDay(day, str_schedules=None):
    if str_schedules == None:
        out = {}
        for moo in moduleOfferings:
            for s in moo.getSchedules():
                day_short = s.getDay()[0:3]  # To allow user to type abbreviation of days(e.g. Monday = Mon)
                if day_short == day:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out
    else:
        out = {}
        for moo, schedule_list in str_schedules.items():
            for s in schedule_list:
                day_short = s.getDay()[0:3]
                if day_short == day:
                    print(
                        moo.toString() + moo.getModule().toString() + s.toString() + str(moo.getModule().getPrograms())[
                                                                                     1:-1])
                    if moo in out:
                        out[moo].append(s)
                    else:
                        out[moo] = [s]
        return out


def deleteSchedule(fields):
    tokens = fields[2].split("/")
    to = []
    for token in tokens:
        if token[0] == "0":
            to.append(token[1:])
        else:
            to.append(token)
    fields[2] = "/".join(to)

    for moo in moduleOfferings:
        if moo.getIntakeCode() == fields[1] and moo.getModule().getModuleCode() == fields[0]:
            for schedule in moo.getSchedules():
                if str(schedule.getDate()) == str(fields[2]) and schedule.getStartTime() == fields[
                    3] and schedule.getLecturerName() == fields[4]:
                    moo.delteSchedule(schedule)
                    print("Schedule is deleted successfully")
                    print("Deleted Schedule Detail: ", schedule)
                    print("*")
                    for i in moo.getSchedules():
                        print(i)
                    print("*")
                    return
    print("Invalid schedule")


def findlocation(zone, room):
    for loca in locations:
        if loca.getZone() == zone and loca.getRoom() == room:
            return loca
    return None


def createSchedule(attributes):
    tokens = attributes[4].split("/")
    to = []
    for token in tokens:
        if token[0] == "0":
            to.append(token[1:])
        else:
            to.append(token)
    attributes[4] = "/".join(to)

    for moo in moduleOfferings:
        if moo.getIntakeCode() == attributes[0] and moo.getModule().getModuleCode() == attributes[
            1] and moo.getStudyMode() == attributes[2] and str(moo.getModule().getPrograms()[0].getProgramCode()) == \
                attributes[3]:
            loca = findlocation(attributes[11], attributes[12])
            if not loca:  # if Location object is not exists, make new Location object and append it to the locations list. This is for preventing from duplication.
                loca = Location(attributes[11], attributes[12])
                locations.append(loca)
            new = Schedule(str(attributes[4]), attributes[5], attributes[6], attributes[7], str(attributes[8]),
                           str(attributes[9]), attributes[10], loca)
            moo.addSchedule(new)
            print("Scheduled is added successfully")
            print(moo.getSchedules())
            return
    print("Invalid")


def updateSchedule(update):
    for moo in moduleOfferings:
        # for update the schedule, program has to find one exact schedule to update the schedule. To find the exact schedule, it has to check if the schedule is exist in ModuleOffering.
        if moo.getIntakeCode() == update[0] and moo.getModule().getModuleCode() == update[1] and moo.getStudyMode() == \
                update[2] and moo.getModule().getPrograms()[0].getProgramCode() == update[3]:
            # find the date and start time of the schedule that will be updated. Start time need to be checked because there are schedules that same module in the same day.
            user2 = input("Type the date for finding specific schedule(dd/mm/year): ").strip()
            user3 = input("Type the start time for finding specific schedule(hhmm): ").strip()
            for s in moo.getSchedules():
                if s.getDate() == user2 and s.getStartTime() == user3:
                    print("1 Date")
                    print("2 Day")
                    print("3 Start Time")
                    print("4 End Time")
                    print("5 Duration")
                    print("6 Room")
                    print("7 Zone")
                    print("8 Lecturer")
                    print("9 Planned Size")
                    user = int(input("Select the field to update: "))
                    while user not in [i for i in range(1, 10)]:
                        user = int(input("Invalid input. Select the field again: "))
                    if user == 1:
                        user3 = input("Type the date to replace: ").strip()
                        s.setDate(user3)
                        # for checking whether the schedule is really changed or not.For checking from this, I have to change first schedule of csv file.
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 2:
                        user3 = input("Type the Day to replace: ").strip()
                        s.setDay(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 3:
                        user3 = input("Type the start time to replace: ").strip()
                        s.setStartTime(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 4:
                        user3 = input("Type the end time to replace: ").strip()
                        s.setEndTime(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 5:
                        user3 = input("Type the duration to replace: ").strip()
                        s.setDuration(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 6:
                        user3 = input("Type the room to replace: ").strip()
                        s.getLocation().setRoom(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 7:
                        user3 = input("Type the zone to replace: ").strip()
                        s.getLocation().setZone(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 8:
                        user3 = input("Type the lecturer name to replace: ").strip()
                        s.setLecturerName(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    elif user == 9:
                        user3 = input("Type the plan size to replace: ").strip()
                        s.setPlanSize(user3)
                        print("Changed Successfully: {}".format(moo.getSchedules()[0]))
                    return
    print('No matching schedule is identified!')


def locationExists(location, lecture_room):
    for loca in locations:
        if loca.getZone() == location and loca.getRoom() == lecture_room:
            return True
    return False


def programExists(program_code):
    for program in programs:
        if program.getProgramCode() == program_code:
            return True
    return False


def moduleExists(module_code):
    for module in modules:
        if module.getModuleCode() == module_code:
            return True
    return False


def getModule(module_code):
    for module in modules:
        if module.getModuleCode() == module_code:
            return module
    return None


def getProgram(program_code):
    for program in programs:
        if program.getProgramCode() == program_code:
            return program
    return None


def getLocation(location, lecture_room):
    for loca in locations:
        if loca.getZone() == location and loca.getRoom() == lecture_room:
            return loca


def promptSort(schedules):
    user = int(input("Do you want to sort the schedule(0: NO / 1: YES)? "))
    while user not in [i for i in range(0, 2)]:
        user = int(input("Invalid input, Type again: "))
    if user == 0:
        out = []  # use list for making excel cuz excel needs list, not dict.
        for moo, ss in schedules.items():
            for s in ss:
                out.append([moo, s])
        return out
    while True:
        print("1 Intake Code")
        print("2 Study Mode")
        print("3 Module Name")
        print("4 Module Code")
        print("5 Date")
        print("6 Day")
        print("7 Start Time")
        print("8 End Time")
        print("9 Plan Size")
        print("10 Duration")
        print("11 Lecturer")
        print("12 Zone")
        print("13 Room")
        print("14 Program Code")
        print("15 Program Name")
        selectsort = int(input("Select the field to sort: "))
        schedules = sort(schedules, selectsort)  # return sort(schedules)
        printSortResults(schedules)
        user = int(input("Do you want to sort again(0: NO / 1: YES)? "))
        if user == 0:
            return schedules


def printSortResults(l):
    for items in l:
        print(items[0], items[0].getModule(), items[0].getModule().getPrograms(), items[1])


# Moo1: [S1, S2, S3]
# [[Moo0, S1], [Moo1, S2], ...]

def sort(schedules, field):
    out = []
    if isinstance(schedules, dict):
        for moo, ss in schedules.items():
            for s in ss:
                out.append([moo, s])
    else:
        out = schedules
    mergeSort(out, field)
    return out


# Convert dictionary to List of List of Moo and S
# for moo, ss in schedules.items():
#     for s in ss:
#         out.append([moo, s])
# Perform Merge Sort

def isLessThan(e1, e2, field):  # Comparison operator for each field.

    if field == 1:
        return e1[0].getIntakeCode() <= e2[0].getIntakeCode()
    elif field == 2:
        return e1[0].getStudyMode() <= e2[0].getStudyMode()
    elif field == 3:
        return e1[0].getIntakeCode() <= e2[0].getIntakeCode()
    elif field == 5:
        date1 = formatDateForSorting(e1[1].getDate())
        date2 = formatDateForSorting(e2[1].getDate())

        return date1 < date2

    elif field == 6:  # use day_mapper to sort the days in Monday to Friday order, not sort by length of string.
        day_mapper = {
            'Monday': 0,
            'Tuesday': 1,
            'Wednesday': 2,
            'Thursday': 3,
            'Friday': 4,
            'Saturday': 5,
            'Sunday': 6
        }
        return day_mapper[e1[1].getDay()] <= day_mapper[e2[1].getDay()]
    elif field == 7:
        return e1[1].getStartTime() <= e2[1].getStartTime()
    elif field == 8:
        return e1[1].getEndTime() <= e2[1].getEndTime()
    elif field == 9:
        return e1[1].getPlanSize() <= e2[1].getPlanSize()
    elif field == 10:
        return e1[1].getDuration() <= e2[1].getDuration()
    elif field == 11:
        return e1[1].getLecturerName() <= e2[1].getLecturerName()
    elif field == 12:
        return e1[1].getLocation().getZone() <= e2[1].getLocation().getZone()
    elif field == 13:
        return e1[1].getLocation().getRoom() <= e2[1].getLocation().getRoom()
    elif field == 14:
        return e1[0].getModule().getPrograms() <= e2[0].getModule().getPrograms()
    elif field == 15:
        pass


def formatDateForSorting(date):
    tokens = date.split('/')
    l = []
    for token in tokens[:-1]:
        if len(token) < 2:
            l.append('0' + token)
        else:
            l.append(token)
    l.append(tokens[2])
    return '/'.join([l[2], l[1], l[0]])


def mergeSort(arr, field):
    if len(arr) > 1:

        mid = len(arr) // 2

        L = arr[:mid]

        R = arr[mid:]

        mergeSort(L, field)

        mergeSort(R, field)

        i = j = k = 0

        while i < len(L) and j < len(R):

            if isLessThan(L[i], R[j], field):  # L[i] < R[j]:
                arr[k] = L[i]
                i += 1
            else:
                arr[k] = R[j]
                j += 1
            k += 1

        while i < len(L):
            arr[k] = L[i]
            i += 1
            k += 1

        while j < len(R):
            arr[k] = R[j]
            j += 1
            k += 1


def promptExportData(data):
    # Asks for user input
    choice = int(input("Do you want to make Time Table(0: NO / 1: YES)? "))
    while choice not in range(2):
        choice = int(input("Invalid input. Input again: "))
    if choice == 0:
        return
    output = data
    if type(data) == 'dict':
        output = []
        for moo, ss in data.items():
            for s in ss:
                output.append([moo, s])
    etype = int(input("Input export format for data(0: Excel, 1: PDF, 2: Excel and PDF): "))
    if etype == 0:
        print("Excel is generated")
        exportExcel(output)
    elif etype == 1:
        print("PDF is generated")
        exportPDF(output)
    elif etype == 2:
        print("PDF, Exel are generated")
        exportExcel(output)
        exportPDF(output)


def exportPDF(data):
    output = data

    if type(data) == 'dict':
        output = []
        for moo, ss in data.items():
            for s in ss:
                output.append([moo, s])

    doc = SimpleDocTemplate("simple_table.pdf", pagesize=(18 * inch, 11 * inch))
    # set title
    styles = getSampleStyleSheet()
    # set logo
    logo = ImageReader('logo.png')

    modCodeandName = set()
    for m in output:  # output = [[moo,s], [], []...] m = [moo,s]
        a = (m[0].getModule().getModuleCode(), m[0].getModule().getModuleName())
        modCodeandName.add(a)
    ProCode = set()
    for m in output:
        a = (m[0].getModule().getPrograms()[0].getProgramCode(), m[0].getModule().getPrograms()[0].getProgramName())
        ProCode.add(a)

    def onFirstPage(canvas, document):
        canvas.drawImage(logo, 1155, 750, mask='auto')
        canvas.setFillColorRGB(1, 0, 0)
        canvas.setFont('Helvetica-Bold', 15)
        canvas.drawString(63, 750, '[PSB Academy]')
        canvas.setFillColorRGB(0, 0, 0)
        canvas.setFont('Helvetica', 15)
        canvas.drawString(63, 730, 'Jackson Square 11 Blk A, Lor 3 Toa Payoh, #01-01, 319579')
        canvas.setFillColorRGB(1, 0, 0)
        canvas.setFont('Helvetica-Bold', 15)
        canvas.drawString(63, 710, '[Tel]')
        canvas.setFillColorRGB(0, 0, 0)
        canvas.setFont('Helvetica', 15)
        canvas.drawString(63, 690, '(+65) 6390 9000')
        canvas.setFillColorRGB(1, 0, 0)
        canvas.setFont('Helvetica-Bold', 15)
        canvas.drawString(63, 670, '[Program Information]')
        canvas.setFillColorRGB(0, 0, 0)
        canvas.setFont('Helvetica', 15)
        canvas.drawString(63, 650, str(ProCode).replace("{", '').replace("}", '').replace("'", '').replace(",", ' = '))
        canvas.setFillColorRGB(1, 0, 0)
        canvas.setFont('Helvetica-Bold', 15)
        canvas.drawString(63, 630, '[Module Information]')
        canvas.setFillColorRGB(0, 0, 0)
        canvas.setFont('Helvetica', 15)
        col = 610
        for k in modCodeandName:
            canvas.drawString(63, col, str(k).replace("'", ''))
            col -= 20

    parsedData = []
    headers = ["Intake Code", "Study Mode", "Module Code", "Date", "Day", "Start Time", "End Time",
               "Plan Size", "Lecturer", "Room", "Zone", "Program Code"]
    parsedData.append(headers)

    for moo, s in output:
        line = [moo.getIntakeCode(), moo.getStudyMode(),
                moo.getModule().getModuleCode(), s.getDate(), s.getDay(), s.getStartTime(), s.getEndTime(),
                s.getPlanSize(), s.getLecturerName(),
                s.getLocation().getRoom(), s.getLocation().getZone(), moo.getModule().getPrograms()[0].getProgramCode()
                ]
        parsedData.append(line)

    table = Table(parsedData, rowHeights=0.3 * inch)
    table.setStyle(TableStyle([('BACKGROUND', (0, 1), (-1, -1), colors.aqua),
                               ('BACKGROUND', (0, 0), (-1, 0), colors.green),
                               ('BACKGROUND', (0, 1), (0, -1), colors.rosybrown),
                               ('BACKGROUND', (1, 1), (1, -1), colors.skyblue),
                               ('BACKGROUND', (3, 1), (3, -1), colors.wheat),
                               ('BACKGROUND', (4, 1), (4, -1), colors.lightgrey),
                               ('BACKGROUND', (5, 1), (5, -1), colors.peachpuff),
                               ('BACKGROUND', (6, 1), (6, -1), colors.teal),
                               ('BACKGROUND', (7, 1), (7, -1), colors.gold),
                               ('BACKGROUND', (8, 1), (8, -1), colors.royalblue),
                               ('BACKGROUND', (9, 1), (9, -1), colors.darkcyan),
                               ('BACKGROUND', (10, 1), (10, -1), colors.tomato),
                               ('BACKGROUND', (11, 1), (11, -1), colors.chocolate),
                               ('BACKGROUND', (12, 1), (12, -1), colors.peru),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                               ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                               ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                               ('FONTSIZE', (0, 0), (-1, -1), 14)]))

    elements = [
        Spacer(10 * cm, 10 * cm),
        Paragraph('PSB Academy Time Table', styles['Title']),
        table,
        Spacer(10 * cm, 10 * cm),
    ]
    # write the document to disk
    doc.build(elements, onFirstPage=onFirstPage)


# [ [moo1, s1], [moo1, s2], [moo2, s3] ]
def exportExcel(data):
    workbook = xlsxwriter.Workbook('Timetable.xlsx')
    worksheet = workbook.add_worksheet()

    parsedData = []
    headers = ["", "Intake Code", "Study Mode", "Module Name", "Module Code", "Date", "Day", "Start Time", "End Time",
               "Plan Size", "Duration", "Lecturer", "Room", "Zone", "Program Code", "Program Name"]
    parsedData.append(headers)

    linecount = 1
    for moo, s in data:
        line = [linecount, moo.getIntakeCode(), moo.getStudyMode(), moo.getModule().getModuleName(),
                moo.getModule().getModuleCode(), s.getDate(), s.getDay(), s.getStartTime(), s.getEndTime(),
                s.getPlanSize(), s.getDuration(), s.getLecturerName(),
                s.getLocation().getRoom(), s.getLocation().getZone(), moo.getModule().getPrograms()[0].getProgramCode(),
                moo.getModule().getPrograms()[0].getProgramName()]
        parsedData.append(line)
        linecount += 1

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    for line in parsedData:
        col = 0
        for cell in line:
            worksheet.write(row, col, cell)
            col += 1
        row += 1

    workbook.close()

def selectUser():
    print("1 Student/Lecturer")
    print("2 Scheduler Admin")
    usertype = int(input("Select user type:"))
    while usertype != 1 and usertype != 2:
        usertype = int(input("Invalid input. Select user type again: "))
    return usertype

def findAndProcessData():
    program_mapper = {
        'DICT-DNDFC': 'Diploma in InfoComm Technology & Diploma in Network Defense and Forensic Countermeasures'
    }
    cwd = os.getcwd()
    while True:
        folder = input("Type the folder:").strip()
        fullPath = os.path.join(cwd, folder)
        if os.path.exists(fullPath):
            break
        print("Wrong folder. Type again:")
    for file in os.listdir(fullPath):
        f = open(os.path.join(os.getcwd(), folder, file))
        isFirstLine = True
        for line in f.readlines()[1:]:
            colum = line.split(",")[1:]

            name = colum[0]
            name = name.split('_')
            # DICT-DNDFC  _  120  _  FT  _  DDMG  _  Lec01
            program_code = name[0].strip() # DICT-DNDFC
            intake_code = name[1].strip() # 120
            study_mode = name[2].strip() # FT
            module_code = name[3] # DDMG

            module_name = colum[1] #

            date = colum[2]

            day = colum[3]

            start_time = colum[4]

            end_time = colum[5]

            duration = colum[6]

            lecture_room = colum[7].strip()

            size = colum[8]

            lecturer = colum[9]

            location = colum[10].strip()

            # Create ModuleOffering object only when reading first line
            if isFirstLine:  # elements which are okay to get info from only from first line of csv file = program name/code, module name/code
                # Check if Program object exists with program_code
                if not programExists(program_code):
                    p = Program(program_code, program_mapper[program_code])
                    programs.append(p)
                # Check if a Module object exists with module_code
                if not moduleExists(module_code):
                    m = Module(module_name, module_code)
                    modules.append(m)  # m1(module_code) .programs
                found = getModule(module_code)  # module object is found here
                # Create new moduleoffering object
                mo = Moduleoffering(intake_code, study_mode, found)
                moduleOfferings.append(mo)
                isFirstLine = False  # Q

                found = getModule(module_code)  # module object is found here
                # check if program object exists in the program list of the Module object with module_code
                programAlreadyExists = False
                for p in found.getPrograms():
                    if p.getProgramCode() == program_code:
                        programAlreadyExists = True
                        break
                if not programAlreadyExists:
                    p = getProgram(program_code)
                    found.addProgram(p)

                # check if module object exists in the module list of the program object with program_code
                found2 = getProgram(program_code)  # program object is found
                moduleAlreadyExists = False
                for m in found2.getModules():
                    if m.getModuleCode() == module_code:
                        moduleAlreadyExists = True
                        break
                if not moduleAlreadyExists:
                    m = getModule(module_code)
                    found2.addModules(m)

            # Checks if location exists
            l = findlocation(location, lecture_room)
            if not l:
                l = Location(location, lecture_room)
                locations.append(l)

            s = Schedule(date, day, start_time, end_time, size, duration, lecturer,
                         l)  # schedule objects cannot be duplicate, don't have to check if it's exist.
            schedules.append(s)
            for moo in moduleOfferings:
                if moo.getIntakeCode() == intake_code and moo.getModule().getModuleCode() == module_code and moo.getStudyMode() == study_mode:  # don't have to check if it's exist with same reason in schedule object
                    moo.addSchedule(s)
        f.close()
def promptListingOption():
    print("1 List All Schedules")
    print("2 List Schedule By Criteria")
    print("3 List All Modules & Lecturers")
    select = int(input("Select the task: "))
    while select not in [i for i in range(1, 4)]:
        select = int(input("Invalid input. Select the task again: "))
    return select

def performListAllSchedule():
    moduleOfferingAndSchedules = listAllSchedule()
    result = []  # [moo1,s1] , [moo1, s2], [moo1, s3] ...[moo2, s1], [moo2, s2] ...
    for moduleOffering, schedules in moduleOfferingAndSchedules.items():
        for s in schedules:
            result.append([moduleOffering, s])
    return result

def promptSearchCriteria():
    print("1 List By Module")
    print("2 List By Lecturer")
    print("3 List By Location")
    print("4 List By Date")
    print("5 List By Date Range")
    print("6 List By Time Range")
    print("7 List By Day")
    print("8 List By Room")
    criteria = int(input("Select the criteria: "))
    while criteria not in [i for i in range(1, 9)]:
        criteria = int(input("Invalid input. Select the criteria again: "))
    return criteria

def performSearchByModule():
    type_module = input("Type the module code: ").strip()
    result = listByModule(type_module)
    while result == {}:  # This is the case that user type the invalid input so that nothing inside of dictionary
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter module code, 0 to exit: "))
        if redo == 1:
            type_module = input("Type the module: ").strip()
            result = listByModule(type_module)
        else:  # if user does not type 1
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByLecturer():
    type_lecturer = input("Type the lecturer:").strip()
    result = listByLecturer(type_lecturer)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter lecturer, 0 to exit: "))
        if redo == 1:
            type_lecturer = input("Type the lecturer: ").strip()
            result = listByLecturer(type_lecturer)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByZone():
    type_zone = input("Type the zone: ").strip()
    result = listByLocation(type_zone)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter zone, 0 to exit: "))
        if redo == 1:
            type_zone = input("Type the zone: ").strip()
            result = listByLocation(type_zone)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByDate():
    type_date = input("type the date(dd/m/year): ").strip()
    result = listByDate(type_date)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter date, 0 to exit: "))
        if redo == 1:
            type_date = input("Type the date: ").strip()
            result = listByDate(type_date)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByDateRange():
    startDate = input("Type the start date(dd/m/year): ").strip()
    endDate = input("Type the end date(dd/m/year): ").strip()
    result = listByDateRange(startDate, endDate)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter date range, 0 to exit: "))
        if redo == 1:
            startDate = input("Type the start date(dd/m/year): ").strip()
            endDate = input("Type the end date(dd/m/year): ").strip()
            result = listByDateRange(startDate, endDate)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByTimeRange():
    startTime = input("Type the start time(type in 4 digits / e.g. 2pm = 1400 8am = 0800): ").strip()
    endTime = input("Type the end time(type in 4 digits / e.g. 2pm = 1400 8am = 0800): ").strip()
    result = listByTimeRange(startTime, endTime)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter time range, 0 to exit: "))
        if redo == 1:
            startTime = input(
                "Type the start time(type in 4 digits / e.g. 2pm = 1400 8am = 0800): ").strip()
            endTime = input(
                "Type the end time(type in 4 digits / e.g. 2pm = 1400 8am = 0800): ").strip()
            result = listByTimeRange(startTime, endTime)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByDay():
    type_day = input("Type the day(Mon/Tue/Wed/Thu/Fri/Sat/Sun): ")
    result = listByDay(type_day)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter day, 0 to exit: "))
        if redo == 1:
            type_day = input("Type the day: ").strip()
            result = listByDay(type_day)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def performSearchByRoom():
    type_room = input("Type the room: ").strip()
    result = listByRoom(type_room)
    while result == {}:
        print("No Schedule Matched")
        redo = int(input("Input 1 to re-enter room, 0 to exit: "))
        if redo == 1:
            type_room = input("Type the room: ").strip()
            result = listByRoom(type_room)
        else:
            exit()
    result = promptRepeatingSearch(result)
    return result

def promptRepeatingSearch(result):
    while True:
        ask_repeat_filter = int(input("Do you want to filter another time(0: NO / 1: YES)? "))
        while ask_repeat_filter not in [i for i in range(0, 2)]:
            ask_repeat_filter = int(input("Invalid Input. Type again: "))
        if ask_repeat_filter == 1:
            result = repeatedFilter(result)
        elif ask_repeat_filter == 0:
            break
    return result

def promptAdminTasksOption():
    print("1 Change Existing Schedule")
    print("2 Delete Schedule Data")
    print("3 Add new Schedule Data")
    selectedTask = int(input("select the task: "))
    while selectedTask not in [i for i in range(1, 4)]:
        selectedTask = int(input("Invalid input. Select the task again: "))
    return selectedTask

def performUpdateSchedule():
    updates = ['Intake Code', 'Module Code', 'Study Mode', 'Program Code']
    userInput = []
    for update in updates:
        inputValue = input("Type " + update + ":").strip()
        userInput.append(inputValue)  # user input from updates list will be append to userinput list
    updateSchedule(userInput)

def performDeleteSchedule():
    fields = ['Module Code', 'Intake Code', 'Date', 'StartTime(hhmm)', 'Lecturer']
    userInput = []
    for field in fields:
        inputValue = input("Type the " + field + ":").strip()
        userInput.append(inputValue)
    # userinput example = ['DCNG', '221', '29/04/2021', '1200', 'Dr Liau Vui Kien']
    deleteSchedule(userInput)

def performInsertSchedule():
    attributes = ['Intake Code', 'Module Code', 'Study Mode', 'Program Code', 'Date', 'Day', 'Start Time',
                  'End Time', 'Plan Size', 'Duration', 'Lecturer', 'zone', 'room']
    userInput = []
    for attribute in attributes:
        inputValue = input("Type " + attribute + ":").strip()
        userInput.append(inputValue)
    createSchedule(userInput)
def main():

    usertype = selectUser()

    findAndProcessData()

    if usertype == 1:

        selectedOption = promptListingOption()

        if selectedOption == 1:

            result = performListAllSchedule()

        elif selectedOption == 2:

            criteria = promptSearchCriteria()

            if criteria == 1:

                result = performSearchByModule()

            elif criteria == 2:

                result = performSearchByLecturer()

            elif criteria == 3:

                result = performSearchByZone()

            elif criteria == 4:

                result = performSearchByDate()

            elif criteria == 5:

                result = performSearchByDateRange()

            elif criteria == 6:

                result = performSearchByTimeRange()

            elif criteria == 7:

                result = performSearchByDay()

            elif criteria == 8:

                result = performSearchByRoom()

            if result == {}:
                print("Sorting unavailable")
            else:
                result = promptSort(result)  # only apply to the after listing task

        elif selectedOption == 3:

            listAllModuleandlecturer()

        if selectedOption == 1 or selectedOption == 2:

            promptExportData(result)

    # admin functions start.
    else:
        selectedTask = promptAdminTasksOption()

        if selectedTask == 1:

            performUpdateSchedule()

        elif selectedTask == 2:

            performDeleteSchedule()

        elif selectedTask == 3:

            performInsertSchedule()
if __name__ == "__main__":
    main()
