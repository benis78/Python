from pyautogui import press, hotkey
from time import sleep
from threading import Thread

esc=1

def normal():
    global esc
    while esc==1:
        press('f15')
        sleep(5)
        if esc==False:
            break

def get_input():
    global esc
    keystrk=input('Press a key \n')
    # print('You pressed: ', keystrk)
    esc=False


n=Thread(target=normal)
i=Thread(target=get_input)
n.start()
i.start()



# try:
#     while True:
#         time.sleep(5)
#         press('f15')
#         print('hej')
        
# except KeyboardInterrupt:
#     pass
