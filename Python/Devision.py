n = int(input("Enter a number: "))
d = int(input("Enter devider: "))

n2 = d
while n2>1:
    if n%d==0:
        n2 = n/d
        print(n2)
    if n2==0:
        break