t = int(input())
for i in range(t):
    n = int(input())
    n = n%4
    if n==0:
        print("Good")
    else:
        print("Not Good")