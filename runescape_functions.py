import pyautogui 
import time 
import random
import bezier
import numpy as np


def move_small_curve(x_coord_min, x_coord_max, y_coord_min, y_coord_max):
    #random_movement_int = random.uniform(.02, 2)

    end_coord_x = random.randint(x_coord_min, x_coord_max)
    end_coord_y = random.randint(y_coord_min, y_coord_max)

    # We'll wait 5 seconds to prepare the starting position
    start_delay = 3 
    pyautogui.sleep(start_delay)

    # For this example we'll use four control points, including start and end coordinates
    start = pyautogui.position()
    end = end_coord_x, end_coord_y

    # Two intermediate control points that may be adjusted to modify the curve.
    control1 = start[0]+random.randint(5, 15), start[1]+random.randint(5, 15)
    control2 = start[0]+random.randint(5, 15), start[1]+random.randint(5, 15)

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
        pyautogui.moveTo(x, y)  # Move to point in curve
        pyautogui.sleep(delay)  # Wait delay

def move_to_box(x_coord_min, x_coord_max, y_coord_min, y_coord_max):
    #random_movement_int = random.uniform(.02, 2)

    end_coord_x = random.randint(x_coord_min, x_coord_max)
    end_coord_y = random.randint(y_coord_min, y_coord_max)

    # We'll wait 5 seconds to prepare the starting position
    start_delay = 3 
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
        pyautogui.moveTo(x, y)  # Move to point in curve
        pyautogui.sleep(delay)  # Wait delay

def wait_next_step():
    sl_time = random.uniform(.3, .8)
    time.sleep(sl_time)

def key_press_wait(key):
    key_wait = random.uniform(0.2, 0.5)
    pyautogui.keyDown(key)
    time.sleep(key_wait)
    pyautogui.keyDown(key)

def click_press_wait():
    key_wait = random.uniform(0.7, 1.0)
    pyautogui.mouseDown()
    time.sleep(key_wait)
    pyautogui.mouseUp()

def right_click_press_wait():
    key_wait = random.uniform(0.8, 1.1)
    pyautogui.mouseDown(button='right')
    time.sleep(key_wait)
    pyautogui.mouseUp(button='right')

def press_one_wait(key):
    key_wait = random.uniform(.1, .3)
    pyautogui.keyDown(key)
    time.sleep(key_wait)
    pyautogui.keyDown(key)

def wait_smelting():
    sl_time = random.uniform(53, 60)
    time.sleep(sl_time)

def wait_logging():
    sl_time = random.uniform(10.0, 15.0)
    time.sleep(sl_time)

def random_wait(range1, range2):
    sl_time = random.uniform(range1, range2)
    time.sleep(sl_time)

def wait_five_to_seven():
    sl_time = random.uniform(5.0, 7.0)
    time.sleep(sl_time)

def click_press_wait():
    key_wait = random.uniform(0.8, 1.1)
    pyautogui.mouseDown()
    time.sleep(key_wait)
    pyautogui.mouseUp()

def delete_inv_trash(png_name):
    png_name = pyautogui.locateOnScreen('C:\\Users\\annam\\Documents\\Python\\RS Fletching\\img_select\\' + png_name + '.png', confidence=.95)

    if png_name is not None:
            move_small_curve(png_name[0], png_name[0] + png_name[2], png_name[1], png_name[1] + png_name[3])
            wait_next_step()
            pyautogui.mouseDown()
            wait_next_step()
            move_to_box(1228, 1313, 916, 989)
            wait_next_step()
            pyautogui.mouseUp()
            wait_next_step()
            move_to_box(1071, 1130, 924, 936)
            wait_next_step()
            click_press_wait()
            wait_next_step()