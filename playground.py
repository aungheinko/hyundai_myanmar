import threading
import time
import random

# Initialize the timer and a lock for thread safety
timer = 60
lock = threading.Lock()
game_active = True

def countdown():
    """Countdown function to reduce the timer."""
    global timer, game_active
    while timer > 0 and game_active:
        time.sleep(1)
        with lock:
            timer -= 1
        print(f"Time left: {timer} seconds")

    if timer == 0:
        print("Time's up! The game is over.")

def player(name):
    """Simulate a player pressing the button."""
    global timer, game_active
    while game_active:
        time.sleep(random.randint(1, 5))  # Random delay to simulate human behavior
        with lock:
            if timer > 0:
                print(f"{name} pressed the button! Resetting timer to 60 seconds.")
                timer = 60
            else:
                print(f"{name} tried to press the button but the game is over.")
                game_active = False

# Create and start threads for countdown and players
countdown_thread = threading.Thread(target=countdown)
player_threads = [threading.Thread(target=player, args=(f"Player {i+1}",)) for i in range(3)]

countdown_thread.start()
for thread in player_threads:
    thread.start()

# Wait for all threads to finish
countdown_thread.join()
for thread in player_threads:
    thread.join()
