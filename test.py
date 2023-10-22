import pyautogui
import time
import openpyxl
import keyboard

XlsxFileName = '1.xlsx'
workbook = openpyxl.load_workbook(XlsxFileName)
worksheet = workbook.active

def automatic_double_click(x, y, duration=0):
    pyautogui.doubleClick(x, y, duration=duration)
    time.sleep(0.5)

def check_termination_key():
    if keyboard.is_pressed('esc'):
        print("Stopping script...")
        workbook.close()
        exit()

if __name__ == "__main__":
    time.sleep(5)

    for row in worksheet.iter_rows():
        time.sleep(1)
        check_termination_key()

        automatic_double_click(716,148,duration=1)
        time.sleep(0.4)

        check_termination_key()



        automatic_click(821,965,duration=1)
        with open('message.txt', 'r') as file:
            pyautogui.hotkey('ctrl','a')
            time.sleep(0.4)
            pyautogui.hotkey('delete')

            for line in file:
                check_termination_key()
                pyautogui.write(line.strip())
                time.sleep(0.4)
                pyautogui.hotkey('shift', 'enter')

        check_termination_key()
        time.sleep(0.5)
        automatic_double_click(1865,1009,duration=1)
        automatic_double_click(1842,1007,duration=0.5)
    workbook.close()