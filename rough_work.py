import time


TIME = 0
while True:
    if time.time() > TIME:
        TIME = time.time() + 5
        print("Hello")