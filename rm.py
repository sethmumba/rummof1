import subprocess
import sys
import time

print("Starting pde.py and pdm.py...")
print("Press Ctrl+C in this terminal to stop both scripts.")

try:
    # sys.executable ensures they both use your active Python 3.10 virtual environment
    pde_process = subprocess.Popen([sys.executable, "pde.py"])
    pdm_process = subprocess.Popen([sys.executable, "pdm.py"])

    # Keep the master script alive to monitor the child processes
    while True:
        time.sleep(1)
        
        # Exit if both processes happen to finish on their own
        if pde_process.poll() is not None and pdm_process.poll() is not None:
            print("Both scripts have finished execution.")
            break

except KeyboardInterrupt:
    print("\nCtrl+C detected. Shutting down both scripts...")
    
    # Send termination signals to both child processes
    pde_process.terminate()
    pdm_process.terminate()
    
    # Wait for them to cleanly exit
    pde_process.wait()
    pdm_process.wait()
    
    print("Both scripts terminated successfully.")