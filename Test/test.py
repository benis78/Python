import pyautogui
import time

# Henter skærmopløsningen
screen_width, screen_height = pyautogui.size()

try:
    while True:
        # Henter musens position i pixels
        x, y = pyautogui.position()

        # Beregner positionen i procent
        x_percent = (x / screen_width) * 100
        y_percent = (y / screen_height) * 100

        print(f"Musens position: X: {x} px ({x_percent:.2f}%), Y: {y} px ({y_percent:.2f}%)")
        time.sleep(0.1)  # Opdaterer hver 0,1 sekund
except KeyboardInterrupt:
    print("Afbrudt af bruger")
