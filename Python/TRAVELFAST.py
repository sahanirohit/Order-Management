t = int(input())
for i in range(t):
    x, y = map(int, input().split())
    if x > y:
        print("BIKE")
    elif y > x:
        print("CAR")
    else:
        print("SAME")