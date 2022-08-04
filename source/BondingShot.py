from genericpath import isfile
from multiprocessing.dummy import Array
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
import datetime
import time
import os
import shutil
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT
from BaseService import BaseService  # 處理文件的直向/橫向
from PIL import Image
from selenium.common.exceptions import NoSuchElementException
import rotatescreen


class BondingShot(BaseService):

    def do_action(self):

        print(self._nowTime.strftime('%Y-%m-%d %H:%M:%S'))

        # 刪除三天前資料夾
        if self._nowTime.strftime('%H') == '00':
            _removeFolder = (
                self._nowTime + datetime.timedelta(days=-3)).strftime('%Y%m%d')
            _dirList = os.listdir(self._shareFolderPath)

            for dirName in _dirList:
                if int(dirName) <= int(_removeFolder):
                    shutil.rmtree(f'{self._shareFolderPath}\{dirName}')

        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(
            executable_path=".\\Driver\\chromedriver.exe", chrome_options=options)
        driver.get(self._url)

        # 計算分時看板站點圖表數量
        _chartCnt = driver.find_elements_by_xpath(
            "//img[@style='height:230px;width:480px;border-width:0px;Position:absolute;top:10px;left:5px;']")

        if(not os.path.isdir(f'{self._shareFolderPath}\{self._strDate}')):
            os.mkdir(f'{self._shareFolderPath}\{self._strDate}')

        # 小於7個站點, 顯示調整為直向, 截一張圖
        if len(_chartCnt) <= 7:
            driver.get('chrome://settings/')
            driver.execute_script(
                F"chrome.settingsPrivate.setDefaultZoom({self._rotationZoom});")
            driver.get(self._url)
            driver.fullscreen_window()

            screen = rotatescreen.get_primary_display()
            start_pos = screen.current_orientation

            pos = abs((start_pos - 90) % 360)
            screen.rotate_to(pos)

            _picName = f"{self._docFileName}_{self._pngTime}.png"
            driver.get_screenshot_as_file(_picName)

            if(os.path.exists(f'.\{_picName}')):
                shutil.move(f'.\{_picName}',
                            f'{self._shareFolderPath}\{self._strDate}\{_picName}')

            driver.quit()

            pos = abs((start_pos - 360) % 360)
            screen.rotate_to(pos)

        # 大於7個站點, 橫向顯示, 分區截圖再合併
        else:
            driver.get('chrome://settings/')
            driver.execute_script(
                F"chrome.settingsPrivate.setDefaultZoom({self._zoom});")
            driver.get(self._url)
            driver.fullscreen_window()

            elementArray = ["Chart1", "Chart5", "Chart8",
                            "Chart" + str(len(_chartCnt) - 1)]
            for idx, ele in enumerate(elementArray):
                _picName = f"{self._docFileName}_{self._strTime}_t{idx + 1}.png"
                if((idx + 1) == 1):
                    driver.get_screenshot_as_file(_picName)
                else:
                    try:
                        # Handle element not found
                        charts = driver.find_element_by_id(ele)
                        action = ActionChains(driver)
                        action.move_to_element(charts).perform()
                        time.sleep(1)
                        driver.get_screenshot_as_file(_picName)
                    except NoSuchElementException:
                        continue

                self._picNameArray.append(_picName)
                
                time.sleep(1)
                if(os.path.exists(f'.\{_picName}')):
                    shutil.move(f'.\{_picName}',
                                f'{self._shareFolderPath}\{self._strDate}\{_picName}')

            driver.quit()
            
            time.sleep(1)

            self.do_spell()

            # region local test (截圖插入 word 轉成 PDF)

            '''
            if(not os.path.isdir(f'.\Report\{_strDate}')):
                os.mkdir(f'.\Report\{_strDate}')
             
            _picName = f"Bonding_Output_{_strTime}_1.png"
            _picNameArray.append(_picName)
            driver.get_screenshot_as_file(_picName)
            shutil.move(f'.\{_picName}',f'.\Report\{_strDate}\{_picName}')
            # charts = driver.find_element_by_tag_name("details")
            charts = driver.find_element_by_id("Chart8")
            action = ActionChains(driver)
            action.move_to_element(charts).perform()
            time.sleep(2)
            _picName2 = f"Bonding_Output_{_strTime}_2.png"
            _picNameArray.append(_picName2)
            driver.get_screenshot_as_file(_picName2)
            shutil.move(f'.\{_picName2}',f'.\Report\{_strDate}\{_picName2}')
            
            driver.quit()
            
            document = Document()
            section = document.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            new_width, new_height = Cm(29.7), Cm(21)
            
            section.page_width = new_width
            
            section.page_height= new_height
            section.left_margin=Cm(1.27)
            section.right_margin=Cm(1.27)
            section.top_margin=Cm(1.27)
            section.bottom_margin=Cm(1.27)
            
            p = document.add_paragraph()
            r = p.add_run()
            for picName in _picNameArray:
                r.add_picture(f'.\Report\{_strDate}\{picName}',width=Cm(27))
                
            document.save(f'.\Report\\{_strDate}\\{_docFileName}.docx')
            
            source_doc = aw.Document(f'.\Report\\{_strDate}\\{_docFileName}.docx')
            # Save as PDF
            source_doc.save(f'.\Report\\{_strDate}\\{_docFileName}.pdf')
            '''
            # endregion

    def do_spell(self):

        if len(self._picNameArray) > 2 and len(self._picNameArray) % 2 == 0:
            for i in range(0, len(self._picNameArray), 2):
                self.do_process(
                    [self._picNameArray[i], self._picNameArray[i+1]], i+1)

    def do_process(self, picArray: list, cnt: int):

        im = Image.open(
            F'{self._shareFolderPath}\{self._strDate}\{picArray[0]}')  # 開啟圖片
        xsize, ysize = im.size
        # 產生一張全黑圖片, 大小 x:長與截圖相同 y:長乘上截圖數量
        bg = Image.new('RGB', (xsize, ysize*len(picArray)), '#000000')
        im.close()

        for i, ele in enumerate(picArray):
            # 開啟圖片
            img = Image.open(F'{self._shareFolderPath}\{self._strDate}\{ele}')
            # 貼上圖片 左上座標(0, 0)
            bg.paste(img, (0, i*ysize))
            img.close()

        bg.save(F'{self._shareFolderPath}\{self._strDate}\{self._docFileName}_{self._pngTime}_{str(cnt)}.png')

        for tmp in picArray:
            if os.path.isfile(F'{self._shareFolderPath}\{self._strDate}\{tmp}'):
                os.remove(F'{self._shareFolderPath}\{self._strDate}\{tmp}')


