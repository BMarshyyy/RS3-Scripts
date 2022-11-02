
import pyautogui as auto
import pyautogui 
import time 
import random
import bezier
import numpy as np


# Any duration less than this is rounded to 0.0 to instantly move the mouse.
auto.MINIMUM_DURATION = 0  # Default: 0.1
# Minimal number of seconds to sleep between mouse moves.
auto.MINIMUM_SLEEP = 0.005  # Default: 0.05
# The number of seconds to pause after EVERY public function call.
auto.PAUSE = 0  # Default: 0.1


time.sleep(5)

movement_types = [auto.easeOutQuad, auto.easeInOutQuad, auto.easeInQuad]

def click_in_box(x_coord_min, x_coord_max, y_coord_min, y_coord_max):
    #random_movement_int = random.uniform(.02, 2)

    end_coord_x = random.randint(x_coord_min, x_coord_max)
    end_coord_y = random.randint(y_coord_min, y_coord_max)

    # We'll wait 5 seconds to prepare the starting position
    start_delay = 3 
    print("Drawing curve from mouse in {} seconds.".format(start_delay))
    pyautogui.sleep(start_delay)

    # For this example we'll use four control points, including start and end coordinates
    start = pyautogui.position()
    end = end_coord_x, end_coord_y

    # Two intermediate control points that may be adjusted to modify the curve.
    #control1 = start[0]+125, start[1]+100
    control1 = start[0]+random.randint(50, 400), start[1]+random.randint(50, 400)
    control2 = start[0]+random.randint(50, 800), start[1]+random.randint(50, 800)

    # Format points to use with bezier
    control_points = np.array([start, control1, control2, end])
    points = np.array([control_points[:,0], control_points[:,1]]) # Split x and y coordinates

    # You can set the degree of the curve here, should be less than # of control points
    degree = 3
    # Create the bezier curve
    curve = bezier.Curve(points, degree)
    curve_steps = 50  # How many points the curve should be split into. Each is a separate pyautogui.moveTo() execution
    delay = 1/curve_steps  # Time between movements. 1/curve_steps = 1 second for entire curve

    # Move the mouse
    for i in range(1, curve_steps + 1):
        # The evaluate method takes a float from [0.0, 1.0] and returns the coordinates at that point in the curve
        # Another way of thinking about it is that i/steps gets the coordinates at (100*i/steps) percent into the curve
        x, y = curve.evaluate(i/curve_steps)
        print(str(x) + " - " + str(y))

        pyautogui.moveTo(x, y)  # Move to point in curve
        pyautogui.sleep(delay)  # Wait delay

def wait_next_step():
    time = random.uniform(.5, 2.5)
    print("Pausing for " + str(time) + " seconds.")

def key_press_wait(key):
    key_wait = random.uniform(0.8, 1.1)
    pyautogui.keyDown(key)
    time.sleep(key_wait)
    pyautogui.keyDown(key)

def click_press_wait():
    key_wait = random.uniform(0.8, 1.1)
    pyautogui.mouseDown()
    time.sleep(key_wait)
    pyautogui.mouseUp()

def press_one_wait(key):
    key_wait = random.uniform(.1, .3)
    pyautogui.keyDown(key)
    time.sleep(key_wait)
    pyautogui.keyDown(key)

###################################################################################################

while 0 == 0:
    # Moving mouse to GE
    click_in_box(850, 1050, 700, 900)

    # Pause
    wait_next_step()

    # Click and wait 
    click_press_wait()

    # Pause
    wait_next_step()
    time.sleep(3)

    # Click 1 preset 
    click_in_box(1310, 1330, 531, 555)
 
    # Pausess
    wait_next_step()

    # Click and wait 
    click_press_wait()

    # Pausess
    wait_next_step()

    # Move to item in inventory
    click_in_box(1188, 1213, 680, 700)

    # Pause
    wait_next_step()

    # Click and wait 
    click_press_wait()

    # Pause
    wait_next_step()
    time.sleep(3)

    # press space
    key_press_wait("space")

    # wait atleast 16 seconds
    fletching_wait = random.uniform(16.5, 18.0)
    time.sleep(fletching_wait)

    # Pause
    wait_next_step()