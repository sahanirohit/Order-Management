class Base:
    def __init__(self):
        self.name = "Rohit"
        self.__username = "admin"
        self.__password = "admin@123"

    def test(self):
        print("Parent",self.__username)
        print("Parent",self.__password)
        # print(self.name)

class Child(Base):
    def __init__(self):
        super().__init__()
        print("Child",self.name)
        # try:
        #     self.test()
        # except Exception as e:
        #     print(e)
        
class Person(Child):
    def __init__(self) -> None:
        super().__init__()
        print("Person",self.name)
        self.test()

user = Person()