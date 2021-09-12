class Moduleoffering:
    def __init__(self,intakeCode, studyMode, module):
        self.__intakeCode = intakeCode
        self.__studyMode = studyMode
        self.__schedules = []
        self.__module = module
    def getIntakeCode(self):
        return self.__intakeCode
    def getStudyMode(self):
        return self.__studyMode
    def getModule(self):
        return self.__module
    def setIntakeCode(self, newCode):
        self.__intakeCode = newCode
    def setStudyMode(self, newMode):
        self.__studyMode = newMode
    def getSchedules(self):
        return self.__schedules
    def addSchedule(self, add):
        self.__schedules.append(add)
    def delteSchedule(self, delete):
        self.__schedules.remove(delete)
    def toString(self):
        return "[Intake Code]: " + self.__intakeCode + ", [Study Mode]: " + self.__studyMode
    def __str__(self):
        return "[Intake Code]: " + self.__intakeCode + ", [Study Mode]: " + self.__studyMode + ", [Module Code]: " + self.getModule().getModuleCode()