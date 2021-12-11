from datetime import time,datetime as dt


class passed_time():
    def __init__(self,hours,minutes,seconds):
        self.hours = int(hours)
        self.minutes = int(minutes)
        self.seconds = int(seconds)

    def convert(self):
        return self.hours*3600+self.minutes*60+self.seconds


def now():
    time = dt.now()
    time_now = (int(time.strftime("%H")), int(time.strftime("%M")), int(time.strftime("%S")))
    return (time_now[0]*60+time_now[1])*60+time_now[2]

def convert_back(secs):
    hours = secs//3600
    minutes =  secs//60-hours*60
    seconds = secs - ( hours*3600 + minutes*60)
    minutes = f"{minutes:02d}"
    seconds = f"{seconds:02d}"
    return hours,minutes,seconds


