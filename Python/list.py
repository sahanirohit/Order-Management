class Fruits:
    def __init__(self, fruit1, fruit2, fruit3) -> None:
        self.fruit1 = fruit1
        self.fruit2 = fruit2
        self.fruit3 = fruit3
        self.myList = [self.fruit1, self.fruit2, self.fruit3]
        print(self.myList)
    def update(self):
        self.myList[0] = "Watermelon"
        self.myList.append("Chiku")
        self.myList.insert(1, "Mango") # we are using .insert method to add item in a list at a specific location
        
        print(self.myList)
obj = Fruits("Apple","Banana","Grap")
obj.update()