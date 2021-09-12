class Location:

    def __init__(self, zone, room):
        self.__zone = zone
        self.__room = room

    def getZone(self):
        return self.__zone
    def getRoom(self):
        return self.__room
    def setZone(self, newZone):
        self.__zone = newZone
    def setRoom(self, newRoom):
        self.__room = newRoom
    def toString(self):
        return "[Room]: " + self.__room + " [Zone]: " + self.__zone + "\n"
    def __str__(self):
        return "[Room]: " + self.__room + " [Zone]: " + self.__zone + "\n"
    def __repr__(self):
        return  "[Room]: " + self.__room + " [Zone]: " + self.__zone + "\n"


