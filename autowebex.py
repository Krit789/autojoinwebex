from ctypes.wintypes import BOOL
from datetime import date
from sys import exc_info,exit
import pyautogui
import os
import time
import win32gui, win32con
import schedule
from os.path import join, dirname
from dotenv import load_dotenv
import time_passed
import ctypes
ctypes.windll.kernel32.SetConsoleTitleW("Autojoinwebex v0.3 by Krit789")
current_dir = os.getcwd()

dotenv_path = join(current_dir, '.env')
load_dotenv(dotenv_path)

elapsed_time = 0
#pip install pyautogui pywin32 schedule pillow python-dotenv DateTime


meet_address = os.environ.get("MEET_URL")
morning_start = os.environ.get("MORNING_START")
noon_start = os.environ.get("NOON_START")
test_mode = os.environ.get("TEST_MODE").lower()
set_music_mode = os.environ.get("SET_MUSIC_MODE").lower()



if meet_address == None:
    raise NameError('MEET_URL in .env is blank.')

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def getKeyboardLayout():
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    kb_layout_id = user32.GetKeyboardLayout(user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), 0))
    return hex(kb_layout_id & (2**16 - 1))

def getScreenScaling():
    os.system("taskkill /T /F /IM ptoneclk*")
    os.system("taskkill /T /F /IM atmgr*")
    pyautogui.moveTo(100, 100, 0)
    time.sleep(5)
    print('       >Launching Cisco Webex Meetings...', end=" ")
    try:
        os.startfile('C:\\Program Files (x86)\\Webex\\Webex\\Applications\\ptoneclk.exe')
    except FileNotFoundError as not_found:
        print(bcolors.FAIL +'ERROR: '+ bcolors.ENDC + not_found.filename +' is missing.\n Are you using the correct version of Cisco Webex Meeting?\n Which can be download from ' + bcolors.UNDERLINE +'https://akamaicdn.webex.com/client/webexapp.msi' + bcolors.ENDC)
        time.sleep(15)
        exit()
    print('Done!')
    time.sleep(2)
    try:
        hwnd = win32gui.FindWindow(None,"Cisco Webex Meetings")
        win32gui.ShowWindow(hwnd,5)
        win32gui.SetForegroundWindow(hwnd)
    except:
        time.sleep(5)
        hwnd = win32gui.FindWindow(None,"Cisco Webex Meetings")
        win32gui.ShowWindow(hwnd,5)
        win32gui.SetForegroundWindow(hwnd)
    time.sleep(2)
    print('       >Detecting keyboard layout...')
    while True:
        if getKeyboardLayout() != '0x409':
            pyautogui.hotkey('win','space')
            time.sleep(1)
        elif getKeyboardLayout() == '0x409':
            break
    print('       >Detecting display scaling...', end=" ")
    global joinameeting_pos
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/100%/JoinAMeeting.PNG', grayscale=True)
    if joinameeting_pos is not None:
        print('100%')
        return 100
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/125%/JoinAMeeting.png', grayscale=True)
    if joinameeting_pos is not None:
        print('125%')
        return 125
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/150%/JoinAMeeting.png', grayscale=True)
    if joinameeting_pos is not None:
        print('150%')
        return 150
    if joinameeting_pos is None:
        print(bcolors.FAIL + '\nERROR: ' + bcolors.ENDC + 'Detection Failed. Quitting...')
        time.sleep(10)
        exit()




def joinMeetings(wscale):
    print('       >Attempting to join meetings')
    if wscale == 100:
        print('       >Entering text field...', end=" ")
        pyautogui.leftClick(joinameeting_pos.x, joinameeting_pos.y+40)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('del')
        pyautogui.typewrite(meet_address, interval=0)
        pyautogui.press('enter')
        print('Done!')
        wscale = str(wscale)
        while True:
            video_blur = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/video_blur.PNG', grayscale=False)
            video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
            join_as_guest = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/joinasguest.PNG', grayscale=True)

            if join_as_guest is not None:
                pyautogui.leftClick(join_as_guest.x,join_as_guest.y)
                print("Can Join As guest")

            if video_blur is not None and join_as_guest is None:
                print("Blur detected but join as guest button is not")
                time.sleep(1)
            if video_blur is None and join_as_guest is None:
                if video_system is not None:
                    print("Join as guest didn't appear or got clicked")
                    win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
                    time.sleep(1)
                    video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
                    pyautogui.leftClick(video_system.x+520,video_system.y+60)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
        if set_music_mode == 'true':
            while True:
                unmute_more = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/unmute.PNG', grayscale=False)
                print("Looking for audio settings")
                if unmute_more is not None:
                    pyautogui.leftClick(unmute_more.x+60, unmute_more.y)
                    time.sleep(1)
                    pyautogui.leftClick(pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/settings.PNG', grayscale=True))
                    time.sleep(2)
                    if pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/musicmodeenabled.PNG', grayscale=False) is None:
                        music_mode = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/MusicMode.PNG', grayscale=True)
                        if music_mode is not None:
                            pyautogui.leftClick(music_mode.x - 45, music_mode.y)
                    exit_setting = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/exit_sound_settings.PNG', grayscale=True)
                    pyautogui.leftClick(exit_setting.x, exit_setting.y)
                    break
        pyautogui.moveTo(250, 200, 0)
        print('Done!')



    if wscale == 125:
        print('       >Entering text field...', end=" ")
        pyautogui.leftClick(joinameeting_pos.x-20, joinameeting_pos.y+50)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('del')
        pyautogui.typewrite(meet_address, interval=0)
        pyautogui.press('enter')
        print('Done!')
        wscale = str(wscale)
        while True:

            video_blur = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/video_blur.PNG', grayscale=False)
            video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
            join_as_guest = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/joinasguest.PNG', grayscale=True)

            if join_as_guest is not None:
                pyautogui.leftClick(join_as_guest.x,join_as_guest.y)

                print("Can Join As guest")

            if video_blur is not None and join_as_guest is None:
                print("Blur detected but join as guest button is not")
                time.sleep(1)
            if video_blur is None and join_as_guest is None:
                if video_system is not None:
                    print("Join as guest didn't appear or got clicked")
                    win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
                    time.sleep(1)
                    video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
                    pyautogui.leftClick(video_system.x+610,video_system.y+70)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
        if set_music_mode == 'true':
            while True:
                unmute_more = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/unmute.PNG', grayscale=False)
                print("Looking for audio settings")
                if unmute_more is not None:
                    pyautogui.leftClick(unmute_more.x+69, unmute_more.y)
                    time.sleep(1)
                    audio_settings = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/settings.PNG', grayscale=True)
                    pyautogui.leftClick(audio_settings.x,audio_settings.y)
                    time.sleep(2)
                    if pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/musicmodeenabled.PNG', grayscale=False) is None:
                        music_mode = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/MusicMode.PNG', grayscale=True)
                        if music_mode is not None:
                            pyautogui.leftClick(music_mode.x - 55, music_mode.y)
                    exit_setting = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/exit_sound_settings.PNG', grayscale=True)
                    pyautogui.leftClick(exit_setting.x, exit_setting.y)
                    break
        pyautogui.moveTo(250, 200, 0)
        print('Done!')


    if wscale == 150:
        print('       >Entering text field...', end=" ")
        pyautogui.leftClick(joinameeting_pos.x-20, joinameeting_pos.y+60)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('del')
        pyautogui.typewrite(meet_address, interval=0)
        pyautogui.press('enter')
        print('Done!')
        wscale = str(wscale)
        while True:
            video_blur = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/video_blur.PNG', grayscale=False)
            video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
            join_as_guest = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/joinasguest.PNG', grayscale=True)

            if join_as_guest is not None:
                pyautogui.leftClick(join_as_guest.x,join_as_guest.y)
                print("Can Join As guest")

            if video_blur is not None and join_as_guest is None:
                print("Blur detected but join as guest button is not")
                time.sleep(1)
            if video_blur is None and join_as_guest is None:
                if video_system is not None:
                    print("Join as guest didn't appear or got clicked")
                    win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
                    time.sleep(1)
                    video_system = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/system.PNG', grayscale=True)
                    pyautogui.leftClick(video_system.x+600,video_system.y+80)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
        if set_music_mode == 'true':
            while True:
                unmute_more = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/unmute.PNG', grayscale=False)
                print("Looking for audio settings")
                if unmute_more is not None:
                    pyautogui.leftClick(unmute_more.x+80, unmute_more.y)
                    time.sleep(1)
                    pyautogui.leftClick(pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/settings.PNG', grayscale=True))
                    time.sleep(2)
                    if pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/musicmodeenabled.PNG', grayscale=False) is None:
                        music_mode = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/MusicMode.PNG', grayscale=True)
                        if music_mode is not None:
                            pyautogui.leftClick(music_mode.x - 55, music_mode.y)
                    exit_setting = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/exit_sound_settings.PNG', grayscale=True)
                    pyautogui.leftClick(exit_setting.x, exit_setting.y)
                    break
        pyautogui.moveTo(250, 200, 0)
        print('Done!')


def cls():
    os.system('cls')

        
def start_meet():
    joinMeetings(getScreenScaling())


def main():
    if test_mode == "true":
        print("Test Mode: True", morning_start , noon_start)
        joinMeetings(getScreenScaling())
        print(bcolors.OKGREEN + 'END OF TEST. QUITTING.' + bcolors.ENDC)
        time.sleep(5)
        exit()

    elif test_mode == "false":
        #print("Test Mode: False" , morning_start , noon_start)
        if time_passed.passed_time(8,0,0).convert() < time_passed.now() < time_passed.passed_time(8,30,0).convert():
            print('Out of order startup is about to commence which will'+ bcolors.FAIL +' TERMINATE current session of Webex meetings'+ bcolors.ENDC +'\nPress Ctrl + C to cancel.\nIgnore this message if you want to continue to join meetings.')
            try:
                time.sleep(5)
            except KeyboardInterrupt:
                cls()
                print('Operation Cancelled. Quitting in 3 econds')
                time.sleep(3)
                exit()
            start_meet()

        elif time_passed.passed_time(12,25,0).convert() < time_passed.now() < time_passed.passed_time(1,40,0).convert():
            print('Out of order startup is about to commence which will'+ bcolors.FAIL +' TERMINATE current session of Webex meetings'+ bcolors.ENDC +'\nPress Ctrl + C to cancel.\nIgnore this message if you want to continue to join meetings.')
            try:
                time.sleep(5)
            except KeyboardInterrupt:
                cls()
                print('Operation Cancelled. Quitting in 3 seconds')
                time.sleep(3)
                exit()
            start_meet()


        schedule.every().day.at(morning_start).do(start_meet)
        schedule.every().day.at(noon_start).do(start_meet)



if __name__ == '__main__':
    if date.today().weekday() < 5 or test_mode == 'true':
        main()
    else:
        print('It\'s weekend so there isn\'t anything today. Quitting')
        exit()



while True:
    try:
        cls()
        schedule.run_pending()
        idle_time = int(float("{:.2f}".format(schedule.idle_seconds())))
        time_unit = "Second(s)"
        if 60 <= idle_time <= 3600:
            idle_time = float("{:.1f}".format(idle_time/60))
            time_unit = "Minute(s)"
        if idle_time > 3600:
            idle_time = float("{:.2f}".format(idle_time/3600))
            time_unit = "Hour(s)"
        elapsed_time = elapsed_time+1

        count_display = str(bcolors.BOLD +'│'+ bcolors.ENDC + (f'Next Schedule task will start in {idle_time} {time_unit}.').center(78) + bcolors.BOLD + '│\n│' + bcolors.ENDC +(f'{elapsed_time} Seconds elapsed since startup').center(78)) + bcolors.BOLD +'│' + bcolors.ENDC
        blank = str(bcolors.BOLD +'│' +(' ').center(78) + '│' + bcolors.ENDC)



        print(bcolors.BOLD + '╒═══════════════════════════════╡ Auto Join Webex ╞════════════════════════════╕\n'+ bcolors.ENDC + blank)
        print(count_display)
        for i in range(14):
            print(blank)
        print(bcolors.BOLD +'│'+ bcolors.ENDC + '  Press Ctrl + C to terminate this program                                    ' + bcolors.BOLD +'│'+ bcolors.BOLD)
        print(bcolors.BOLD +f'╘═══╡{date.today().weekday()}╞══════════════════════════════════╡ '+ bcolors.OKGREEN +'build 30-11-2021_01'+ bcolors.ENDC +' ╞════╡ '+ bcolors.OKGREEN + 'VER 0.3'+ bcolors.ENDC +' ╞╛' + bcolors.ENDC)
        time.sleep(1)
    except KeyboardInterrupt:
        print('\nQuitting...')
        time.sleep(1)
        exit()