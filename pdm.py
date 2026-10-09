import pyautogui
import random
import time

# Disable the corner failsafe so ONLY Ctrl+C in the terminal can stop it
pyautogui.FAILSAFE = False

print("Starting Anti-Away simulation. Press Ctrl+C to abort.")
try:
    while True:
        # Wait a long, random time before the next "accidental" touch
        time.sleep(random.uniform(10.0, 45.0))

        # Simulate a tiny, erratic brush against the touchpad
        x_offset = random.randint(-15, 15)
        y_offset = random.randint(-15, 15)
        duration = random.uniform(0.05, 0.2)
        
        pyautogui.moveRel(x_offset, y_offset, duration, pyautogui.easeOutQuad)
        
        # Tap F15: The gold standard for bypassing Teams/Slack "Away" statuses
        pyautogui.press('f15')
        
except KeyboardInterrupt:
    print("\nSimulation terminated.")