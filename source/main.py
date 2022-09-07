import json
import schedule
import time
import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from BondingShot import BondingShot
from BaseService import BaseService
import rotatescreen


def main():

    _scheduleJob = BackgroundScheduler()

    # 讀取設定檔
    with open('.\\scheduleConfig.json', encoding="utf-8") as f:
        _scheduleCfig = json.load(f)
        # 依序處理
        for config in _scheduleCfig['setting']:
            # new 要執行的 py
            _className = eval('BondingShot')
            _service = _className()
            # _bondingShot = BondingShot()

            # 依據時間設定排程
            for trigTime in config['time']:
                _timeArray = str(trigTime).split(':')
                trigger = CronTrigger(
                    year="*", month="*", day="*", hour=_timeArray[0], minute=_timeArray[1], second="0")

                _scheduleJob.add_job(_service.do_screenShot,
                                   trigger=trigger,
                                   args=[config['fileName'], config['folderPath'], config['url'], config['elements'], config['zoom'], config['rotationZoom'], config['scrollTop']],
                                   id=F"Job_{_timeArray[0]}{_timeArray[1]}")
                # _job.remove()

        _scheduleJob.start()


if __name__ == '__main__':
    
    main()

    # # 測試段落
    # _bondingShot = BondingShot()
    # _bondingShot.do_screenShot('Bonding_Report', '.\\Report', 'https://docs.python.org/zh-tw/3.7/tutorial/venv.html', 
    # ["introduction","creating-virtual13543-environments","managing-packages-with-pip"],"0.88","0.56", 80)

    while True:
        # every daya at specific time_
        # schedule.every(30).seconds.do(do_screenShot)
        # print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        schedule.run_pending()
        time.sleep(1)
