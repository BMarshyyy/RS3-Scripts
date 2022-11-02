import pyautogui as auto
import time 


time.sleep(5)
currentMouseX, currentMouseY = auto.position()

print("Backpack")
print(str(currentMouseX) + ", " + str(currentMouseY))