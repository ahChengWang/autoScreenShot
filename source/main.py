import json
import schedule
import time
import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from BondingShot import BondingShot
from BaseService import BaseService


def main():

    _scheduler = BackgroundScheduler()

    _fileName = ""
    _folderPath = ""
    _url = ""
    _scheduleTime = []

    with open('D:\\PyProj\\autoScreenShot\\source\\scheduleConfig.json', encoding="utf-8") as f:
        _scheduleCfig = json.load(f)
        for config in _scheduleCfig['setting']:
            _className = eval('BondingShot')
            _service = _className()
            # _bondingShot = BondingShot()

            for trigTime in config['time']:
                _timeArray = str(trigTime).split(':')
                trigger = CronTrigger(
                    year="*", month="*", day="*", hour=_timeArray[0], minute=_timeArray[1], second="0")

                _scheduler.add_job(_service.do_screenShot,
                                   trigger=trigger,
                                   args=[config['fileName'], config['folderPath'], config['url'], config['elements']])
                # schedule.every().days.at(time).do(
                #     _bondingShot.do_screenShot(config['fileName'], config['folderPath'], config['url'], config['elements']))
                # schedule.every().days.at('10:30').do(BaseService.do_screenShot(schedule['fileName'], _folderPath, _url))

            _scheduler.start()


if __name__ == '__main__':

    main()

    while True:
        # every daya at specific time_
        # schedule.every(30).seconds.do(do_screenShot)
        # print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        schedule.run_pending()
        time.sleep(1)
