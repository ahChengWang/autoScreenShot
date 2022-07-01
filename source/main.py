import requests
import base64
import urllib.parse
import aspose.words as aw
import glob
import json
import schedule
import time
from BaseService import BaseService 
import os


def main():

    _fileName = ""
    _folderPath = ""
    _url = ""
    _scheduleTime = []
    

    schedule.every().days.at('08:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('10:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('12:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('14:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('16:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('18:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('20:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('22:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('00:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('02:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('04:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))
    schedule.every().days.at('06:30').do(BaseService.do_screenShot(_fileName, _folderPath, _url))

    while True:
        # every daya at specific time_
        # schedule.every(30).seconds.do(do_screenShot)
        # print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()