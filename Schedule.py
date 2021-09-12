from Location import Location
class Schedule:
    def __init__(self, date, day, startTime, endTime, planSize, duration, lecturerName, location):
        self.__date = date
        self.__day = day

        self.__startTime = self.processTime(startTime)
        self.__endTime = self.processTime(endTime)
        self.__planSize = planSize
        self.__duration = duration
        self.__lecturerName = lecturerName
        self.__location = location

    def getDate(self):
        return self.__date
    def setDate(self, new):
        self.__date = new
    def getDay(self):
        return self.__day
    def setDay(self, new):
        self.__day = new
    def getStartTime(self):
        return self.__startTime
    def setStartTime(self, new):
        self.__startTime = new
    def getEndTime(self):
        return self.__endTime
    def setEndTime(self, new):
        self.__endTime = new
    def getPlanSize(self):
        return self.__planSize
    def setPlanSize(self, new):
        self.__planSize = new
    def getDuration(self):
        return self.__duration
    def setDuration(self, new):
        self.__duration = new
    def getLecturerName(self):
        return self.__lecturerName
    def setLecturerName(self, new):
        self.__lecturerName = new
    def getLocation(self):
        return self.__location
    def setLocation(self, new):
        self.__location = new

    def processTime(self, time): # '8:30:00' -> ['8', '30', '0'] -> ['08', '30', '0'] -> '0830'
        time_tokens = time.split(":")
        outlist = []
        for token in time_tokens[0:2]:
            if len(token) == 1:
                outlist.append("0" + token)
            else:
                outlist.append(token)
        return ''.join(outlist)


    def __repr__(self):
        return "[Date]: " + self.__date + " " + "[Day]: " + self.__day + " " + "[Start Time]: " + str(
            self.__startTime) + " " + "[End Time]: " + str(self.__endTime) + " " + "[Plan Size]: " + str(
            self.__planSize) + " " + "[Duration]: " + str(
            self.__duration) + " " + "[Lecturer]: " + self.__lecturerName + " " + self.__location.toString() + " "

    def toString(self):
        return "[Date]: " + self.__date + " " + "[Day]: " + self.__day + " " + "[Start Time]: " + str(self.__startTime) + " " + "[End Time]: " + str(self.__endTime) + " " + "[Plan Size]: " + str(
            self.__planSize) + " " + "[Duration]: " + str(self.__duration) + " " + "[Lecturer]: " + self.__lecturerName + " " + self.__location.toString() + " "

    def print(self):
        print("location:", self.__location)
