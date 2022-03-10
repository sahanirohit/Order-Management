class Human:
    def __init__(self,name, o) -> None:
        self.name = name
        self.occupation = o
        self.test("demo")

    def do_work(self):
        if self.occupation == "Actor":
            print(self.name + " shoot a Film.")
        elif self.occupation == "Cricketer":
            print(self.name + " Plays Cricket.")

    def test(self, t):
        self.t = t
        print("Hello")
        if t == "demo":
            print("Tested")

person1 = Human("Rohit","Actor")
person1.do_work()
person2 = Human("Vyas", "Cricketer")
person2.do_work()
# user = input("Enter something")
# person1.test(user)