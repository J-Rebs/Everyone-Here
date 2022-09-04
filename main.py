from check import Check
from datetime import datetime, timedelta
from time import sleep

if __name__ == '__main__':
    checker = Check()

    begin = datetime.now()
    end = begin + timedelta(hours=24)

    checker.get_meeting_info(begin,end)
    sleep(60)
