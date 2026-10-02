# get_pos.py
import time
from pynput.mouse import Controller

mouse = Controller()
print("Hover your mouse over the target in VS Code. Logging position every 2 seconds (Ctrl+C to stop):")
try:
    while True:
        print(f"Current Position: {mouse.position}")
        time.sleep(2)
except KeyboardInterrupt:
    print("\nDone.")