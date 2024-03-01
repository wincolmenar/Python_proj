import pyautogui as pag
import time
import random

while True:
    x = random.randint(50, 100)
    y = random.randint(20, 80)
    pag.moveTo(x, y, 0.5)
    time.sleep(1)

