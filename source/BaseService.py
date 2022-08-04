import datetime


class BaseService():

    def __init__(self) -> None:
        self._nowTime = datetime
        self._strDate = ''
        self._strTime = ''
        self._pngTime = ''
        self._docFileName = ''  # f'Bonding_Report_{_strTime}'
        self._picNameArray = []
        # f'Z:\\02-共用資料區(3G)\\03-戰情\\3F_Bonding\\Daily'
        self._shareFolderPath = ''
        self._url = ''
        self._elementArray = []
        self._zoom = ''
        self._rotationZoom = ''

    def do_screenShot(self, fileName: str, folderPath: str, url: str, elements: list, zoom: str, rotationZoom: str):
        self.__init__()
        self._nowTime = datetime.datetime.now()
        self._strDate = self._nowTime.strftime('%Y%m%d')
        self._strTime = self._nowTime.strftime('%y%m%d%H%M%S')
        self._pngTime = self._nowTime.strftime('%y%m%d%H%M')
        self._docFileName = fileName  # f'Bonding_Report_{_strTime}'
        # f'Z:\\02-共用資料區(3G)\\03-戰情\\3F_Bonding\\Daily'
        self._shareFolderPath = folderPath
        self._url = url
        self._elementArray = elements
        self._zoom = zoom
        self._rotationZoom = rotationZoom

        self.do_action()

    def do_action(self):
        pass
