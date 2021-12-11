import pyautogui
import os
import time
import win32gui, win32con
import schedule
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

elapsed_time = 0

#pip install pyautogui pywin32 schedule pillow

meet_address = os.environ.get("MEET_URL")
morning_start = os.environ.get("MORNING_START")
noon_start = os.environ.get("NOON_START")
test_mode = os.environ.get("TEST_MODE")



if meet_address == None:
    raise NameError('MEET_URL in .env is blank.')


def getScreenScaling():
    os.system("taskkill /T /F /IM ptoneclk*")
    os.system("taskkill /T /F /IM atmgr*")
    time.sleep(5)
    print('       >Launching Cisco Webex Meetings...', end=" ")
    os.startfile('C:\\Program Files (x86)\\Webex\\Webex\\Applications\\ptoneclk.exe')
    print('Done!')
    time.sleep(2)
    hwnd = win32gui.FindWindow(None,"Cisco Webex Meetings")
    win32gui.ShowWindow(hwnd,5)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(2)
    global joinameeting_pos
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/100%/JoinAMeeting.PNG', grayscale=True)
    if joinameeting_pos is not None:
        return 100
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/125%/JoinAMeeting.png', grayscale=True)
    if joinameeting_pos is not None:
        return 125
    joinameeting_pos = pyautogui.locateCenterOnScreen('images/150%/JoinAMeeting.png', grayscale=True)
    if joinameeting_pos is not None:
        return 150
    



def joinMeetings(wscale):
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
            join_as_guest = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/joinasguest.PNG', grayscale=True, confidence=0.9)

            if join_as_guest is not None:
                pyautogui.leftClick(join_as_guest.x,join_as_guest.y)
                print("Can Join As guest")

            if video_blur is not None and join_as_guest is None:
                print("Blur detected but join as guest button is not")
                time.sleep(1)
            if video_blur is None and join_as_guest is None:
                if video_system is not None:
                    print("Join as guest didn't appear or got clicked")
                    pyautogui.leftClick(video_system.x+520,video_system.y+60)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
        while True:
            unmute_more = pyautogui.locateCenterOnScreen('images/' + wscale + '%/' + "dark" + '/unmute.PNG', grayscale=False)
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
                time.sleep(1)
                pyautogui.moveTo(unmute_more.x-100, unmute_more.y-200)
                break

        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
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
                    pyautogui.leftClick(video_system.x+610,video_system.y+70)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
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
                time.sleep(1)
                pyautogui.moveTo(unmute_more.x-100, unmute_more.y-200)
                break

        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
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
                    pyautogui.leftClick(video_system.x+500,video_system.y+80)
                    time.sleep(3)
                    break
                print("Nothing is being detected")
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
                time.sleep(1)
                pyautogui.moveTo(unmute_more.x-100, unmute_more.y-200)
                break

        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MAXIMIZE)
        print('Done!')


def cls():
    os.system('cls')

def main():
    if test_mode == "True":
        print("Test Mode: True")
        joinMeetings(getScreenScaling())

    elif test_mode == "False":
        print("Test Mode: False" , morning_start , noon_start)
        schedule.every().day.at(morning_start).do(start_meet)
        schedule.every().day.at(noon_start).do(start_meet)
        #schedule.every().day.at("08:00:00").do(joinMeetings,getScreenScaling())
        #schedule.every().day.at("12:25:00").do(joinMeetings,getScreenScaling())
        
def start_meet():
    joinMeetings(getScreenScaling)


if __name__ == '__main__':
    main()



while True:
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
    print('╒═══════════════════════════════ Auto Webex by Krit789════════════════════════════════╕\n')
    print(f'  Next Schedule task will start in {idle_time} {time_unit}. {elapsed_time} Seconds elapsed since startup')
    time.sleep(1)