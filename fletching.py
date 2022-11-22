
import pyautogui as auto
import time 
import random
import runescape_functions as rsf

# Any duration less than this is rounded to 0.0 to instantly move the mouse.
auto.MINIMUM_DURATION = 0  # Default: 0.1
# Minimal number of seconds to sleep between mouse moves.
auto.MINIMUM_SLEEP = 0.005  # Default: 0.05
# The number of seconds to pause after EVERY public function call.
auto.PAUSE = 0  # Default: 0.1


time.sleep(5)

movement_types = [auto.easeOutQuad, auto.easeInOutQuad, auto.easeInQuad]

###################################################################################################
batches = 0
while True:
    # Moving mouse to GE
    rsf.move_to_box(850, 1050, 700, 900)

    # Pause
    rsf.wait_next_step()

    # Click and wait 
    rsf.click_press_wait()

    # Pause
    rsf.wait_next_step()
    time.sleep(1)

    # Click 1 preset 
    rsf.move_to_box(1310, 1330, 531, 555)
 
    # Pausess
    rsf.wait_next_step()

    # Click and wait 
    rsf.click_press_wait()

    # Pausess
    rsf.wait_next_step()
  
    # Move to item in inventory
    rsf.move_to_box(1351, 1368, 975, 991)

    # Pause
    rsf.wait_next_step()

    # Click and wait 
    rsf.click_press_wait()

    # Pause
    rsf.wait_next_step()
    time.sleep(1.5)
    # press space
    rsf.key_press_wait("space")

    # wait atleast 16 seconds
    fletching_wait = random.uniform(16.5, 18.0)
    time.sleep(fletching_wait)

    batches += 1
    print("Completed batch #" + str(batches)) 
