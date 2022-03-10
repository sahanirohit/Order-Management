import math
import random

otp = random.randint(100000,999999)
print(otp)
print(type(otp))
user = int(input("Enter your otp here "))
print(user)
print(type(user))
if user == otp:
    print("Login Success")
else:
    print("OTP Invalid")
