class Program:
    def __init__(self, programCode, programName):
        self.__programCode = programCode
        self.__programName = programName
        self.__modules = []

    def getProgramCode(self):
        return self.__programCode
    def getProgramName(self):
        return self.__programName
    def setProgramCode(self, new):
        self.__programCode = new
    def setProgramName(self, new):
        self.__programName = new
    def getModules(self):
        return self.__modules
    def addModules(self, add):
        self.__modules.append(add)
    def deleteModules(self, delete):
        self.__modules.remove(delete)

    def print(self):
        print("program code:", self.__programCode, ", program name:", self.__programName)

    def toString(self):
        return "[Program Code]: " + self.__programCode + "[Program Name]: " + self.__programName
    def __repr__(self): # allows to print the object in a list // for cases when the object is in the collection a = [obj 1, obj2]  str = if object not in the list.
        return "[Program Code]: " + self.__programCode + " [Program Name]: " + self.__programName