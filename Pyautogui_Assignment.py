import pyautogui
import time
from datetime import datetime
import pyperclip

# 1. Open Microsoft Edge
pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("msedge")
pyautogui.press("enter")
time.sleep(5)

# 2. Open weather website
pyautogui.hotkey("ctrl", "l")
pyautogui.write("https://www.weather.com/")
pyautogui.press("enter")
time.sleep(7)

# 3. Copy information from webpage
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")
time.sleep(2)

# Get copied data
fetched_data = pyperclip.paste()

print("Fetched information:")
print(fetched_data[:500])

# 4. Get today's date and time
date_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")

# 5. Open Excel
pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("excel")
pyautogui.press("enter")
time.sleep(7)

# 6. Create a new workbook
pyautogui.hotkey("ctrl", "n")
time.sleep(3)

# 7. Enter today's date and time
pyautogui.write(date_time)
pyautogui.press("tab")

# 8. Enter fetched data
pyautogui.write(fetched_data[:300])
pyautogui.press("tab")

# 9. Enter comment
pyautogui.write("Good for outdoor activities")
pyautogui.press("enter")

time.sleep(2)

# 10. Save Excel file with today's date
today = datetime.now().strftime("%Y-%m-%d")

pyautogui.hotkey("ctrl", "shift", "s")
time.sleep(3)

# Type filename in Save As window
pyautogui.hotkey("ctrl", "a")
pyautogui.write(f"Pyautogui_Assignment_{today}.xlsx")
pyautogui.press("enter")

time.sleep(5)

# 11. Take screenshot of final Excel sheet
screenshot_path = rf"B:\Pyauto Folder\Pyautogui_Assignment_{today}.png"

pyautogui.screenshot(screenshot_path)

print("Excel file saved successfully.")
print(f"Filename: Pyautogui_Assignment_{today}.xlsx")
print(f"Screenshot saved: {screenshot_path}")