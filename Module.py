class Module:
    def __init__(self, modName, modCode):
        self.__modName = modName
        self.__modCode = modCode
        self.__programs = []

    def setModuleName(self, newModName):
        self.__modName = newModName
    def setModuleCode(self, newModCode):
        self.__modCode = newModCode
    def getModuleName(self):
        return self.__modName
    def getModuleCode(self):
        return self.__modCode
    def getPrograms(self):
        return self.__programs
    def addProgram(self, add):
        self.__programs.append(add)
    def deleteProgram(self, delete):
        self.__programs.remove(delete)

    def print(self):
        print("Module Name:", self.__modName, "Module Code:", self.__modCode)

    def toString(self):
        return " [Module Name]: " + self.__modName + ", [Module Code]: " + self.__modCode

    def __str__(self):
        return " Module Name: " + self.__modName + ", Module Code: " + self.__modCode

    def __repr__(self):
        return " Module Name: " + self.__modName + ", Module Code: " + self.__modCode


