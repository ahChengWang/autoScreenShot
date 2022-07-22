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
        driver = webdriver.Chrome(executable_path=".\\Driver\\chromedriver.exe", chrome_options=options)
        driver.get('chrome://settings/')
        driver.execute_script(F"chrome.settingsPrivate.setDefaultZoom({self._zoom});")
        driver.get(self._url)
        # driver.get("http://10.132.133.164/Function/ENG1/Bonding_Output.aspx") # 3F Bonding
        # driver.get("http://10.132.23.123:81/Function/ENG1/Bonding_Output.aspx") # 3F Bonding from 中控

        driver.fullscreen_window()

        time.sleep(1)

        if(not os.path.isdir(f'{self._shareFolderPath}\{self._strDate}')):
            os.mkdir(f'{self._shareFolderPath}\{self._strDate}')

        for idx, ele in enumerate(self._elementArray):
            _picName = f"{self._docFileName}_{self._strTime}_{idx + 1}.png"            
            if((idx + 1) == 1):
                driver.get_screenshot_as_file(_picName)
            else:
                try:
                    # Handle element not found
                    charts = driver.find_element_by_id(ele)
                    action = ActionChains(driver)
                    action.move_to_element(charts).perform()
                    time.sleep(2)
                    driver.get_screenshot_as_file(_picName)
                except NoSuchElementException:                    
                    continue

            self._picNameArray.append(_picName)
            if(os.path.exists(f'.\{_picName}')):
                shutil.move(f'.\{_picName}',
                            f'{self._shareFolderPath}\{self._strDate}\{_picName}')

        driver.quit()

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

        im = Image.open(
            F'{self._shareFolderPath}\{self._strDate}\{self._picNameArray[0]}')  # 開啟圖片
        xsize, ysize = im.size
        # 產生一張全黑圖片, 大小 x:長與截圖相同 y:長乘上截圖數量
        bg = Image.new('RGB', (xsize, ysize*len(self._picNameArray)), '#000000')

        for i, ele in enumerate(self._picNameArray):
            # 開啟圖片
            img = Image.open(F'{self._shareFolderPath}\{self._strDate}\{ele}') 
            # 貼上圖片 左上座標(0, 0)
            bg.paste(img, (0, i*ysize))

        bg.save(F'{self._shareFolderPath}\{self._strDate}\{self._docFileName}_{self._pngTime}.png')
