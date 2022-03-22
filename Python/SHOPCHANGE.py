t = int(input())
for n in range(t):
    a, b, x, y = map(int, input().split())
    if a*b <= x*y:
        print("Yes")
    else:
        print("No")
