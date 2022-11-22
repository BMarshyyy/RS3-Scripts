
import pyautogui 
import time 
import runescape_functions as rsf

# Any duration less than this is rounded to 0.0 to instantly move the mouse.
pyautogui.MINIMUM_DURATION = 0  # Default: 0.1
# Minimal number of seconds to sleep between mouse moves.
pyautogui.MINIMUM_SLEEP = 0.005  # Default: 0.05
# The number of seconds to pause after EVERY public function call.
pyautogui.PAUSE = 0  # Default: 0.1

pyautogui.FAILSAFE = False


time.sleep(5)

movement_types = [pyautogui.easeOutQuad, pyautogui.easeInOutQuad, pyautogui.easeInQuad]

##########################################

def bait_exist():
    time.sleep(.5)
    bait = pyautogui.locateOnScreen('C:\\Users\\annam\\Documents\\Python\\RS Fletching\\img_select\\fishing\\bait.png', confidence=.90)
    if bait is not None:
        return True

def check_inventory_full():
    time.sleep(.5)

    full_backpack = pyautogui.locateOnScreen('C:\\Users\\annam\\Documents\\Python\\RS Fletching\\img_select\\fishing\\fish_full.png', confidence=.90)
    logic = False
    if full_backpack is not None:
        logic = True

    return logic
        
def find_bubbles():
    while True:
        fish = pyautogui.locateOnScreen('C:\\Users\\annam\\Documents\\Python\\RS Fletching\\img_select\\fishing\\fish_bubbles.png', confidence=.90)
        try:
            rsf.move_to_box(fish[0], fish[0] + fish[2], fish[1], fish[1] + fish[3])
            rsf.wait_next_step()
            break
        except:
            time.sleep(.25)

def locate_chest():
    logic = True
    while logic == True:
        for x in range(1, 5):
            chest = pyautogui.locateOnScreen('C:\\Users\\annam\\Documents\\Python\\RS Fletching\\img_select\\fishing\\chest' + str(x) + '.png', confidence=.90)

            if chest is not None:
                rsf.move_to_box(chest[0], chest[0] + chest[2], chest[1], chest[1] + chest[3])
                rsf.wait_next_step()
                rsf.click_press_wait()
                rsf.wait_next_step()
                # Walking to chest
                rsf.random_wait(5, 7)
                logic = False
                break


##########################################

while True:
    # Locate fishing spot
    find_bubbles() # Wait next step baked in
        
    # Ensuring fish are still present 
    if bait_exist() == True:
        rsf.click_press_wait()
        rsf.random_wait(5, 7)
        rsf.move_to_box(951, 966, 576, 584) # Centering Mouse
        time.sleep(1)
    while bait_exist() == True:
        rsf.random_wait(10, 15)
        
        if check_inventory_full() == True:
            locate_chest() # Now at the chest

            rsf.move_to_box(1345, 1364, 535, 554)
            rsf.wait_next_step()
            rsf.click_press_wait()
            rsf.wait_next_step()
        else:
            rsf.move_to_box(951, 966, 576, 584) # Centering Mouse