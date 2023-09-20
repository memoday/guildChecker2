import os, sys
import requests
from bs4 import BeautifulSoup
import datetime
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import webbrowser
import csv

__version__ = "v2.0.1"

header = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Whale/3.18.154.13 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
}

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

icon = resource_path('assets/memo.ico')
form = resource_path('ui/main.ui')

form_class = uic.loadUiType(form)[0]

def selectedWorld(worldName):
    world = [None,'리부트','리부트2','오로라','레드','이노시스','유니온','스카니아','루나','제니스','크로아','베라','엘리시움','아케인','노바']
    worldIndex = world.index(worldName)
    return worldIndex

class compare(QThread):

    updateChangesListSignal = pyqtSignal(str)
    updateStatusBarSignal = pyqtSignal(str)

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        self.parent.btn_start.setDisabled(True)
        self.parent.btn_check.setDisabled(True)
        self.updateStatusBarSignal.emit('최신 길드원 정보 불러오는 중...')

        worldName = str(self.parent.combo_serverName.currentText())
        worldNumber = selectedWorld(worldName)

        guildName = self.parent.input_guildName.text()
        if guildName == "" or len(guildName) < 2:
            self.updateStatusBarSignal.emit('변동사항 확인: 정확한 길드명을 입력해주세요')
            return

        url = f"https://maplestory.nexon.com/Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
        recentMemberList = []
        try:
            raw = requests.get(url,headers=header)
            html = BeautifulSoup(raw.text,"html.parser")

            checkRankTag = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr')['class']
            if bool(checkRankTag) == False:
                href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > a')['href']
            else:
                href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > dl > dt > a')['href']

            guildUrl = 'https://maplestory.nexon.com'+href
            pageNumber = 1

            while True:
                self.updateStatusBarSignal.emit(f'{guildName} 최신 길드원 정보 추출 중: {pageNumber} /10')
                guildUrlPage = guildUrl+f'&page={pageNumber}'
                try:
                    r = requests.get(guildUrlPage,headers=header)
                    html = BeautifulSoup(r.text,"html.parser") 
                    members = html.select('#container > div > div > table > tbody > tr')
                    for member in members:
                        nick = member.select_one('td.left > span > img')['alt']
                        # job = member.select_one('td.left > dl > dd').text #일부 직업들은 기사단/마법사/전사/해적 등으로 표시되어있음
                        recentMemberList.append([nick][0])
                    print('크롤링한 페이지: ',pageNumber)
                except Exception as e:
                    print(e)
                    break

                if pageNumber == 10:
                    print(f'최신 길드원 목록 - ({len(recentMemberList)}) 명')
                    print(recentMemberList)

                    self.updateStatusBarSignal.emit(f'변동사항 확인 중... {guildName}')

                    newcomerList, leaveList = self.checkChanges(recentMemberList)
                    for i in range(len(newcomerList)):
                        self.updateChangesListSignal.emit('[신규] ' + newcomerList[i])
                    for i in range(len(leaveList)):
                        self.updateChangesListSignal.emit('[탈퇴] ' + leaveList[i])

                    changeCount = len(newcomerList) + len(leaveList)
                    self.parent.changeCount.setText(str(changeCount)+' 명')
                    
                    self.updateStatusBarSignal.emit('변동사항 확인 완료 '+guildName)
                    break
                else:
                    pageNumber += 1

        except Exception as e:
            self.updateStatusBarSignal.emit(f'[ERROR] {e}')
        
    def checkChanges(self,recentMemberList):
        # self.parent.guildMembers_changed.setText('')
        set1 = set(recentMemberList)
        set2 = set(oldGuildList)
        changeCount = 0
        
        guildIn = list(set1 - set2)
        guildOut = list(set2 - set1)

        newcomerList = []
        trackList = [] #닉변 추적 기능 임시 비활성화, 기존 길드 = 변경 길드 일 경우 찾을 수 없음으로 표기
        leaveList = []

        for i in range(len(guildIn)):
            print('[신규]',guildIn[i])
            newcomerList.append(guildIn[i])
            changeCount += 1
        
        for i in range(len(guildOut)):
            print('[탈퇴]',guildOut[i])
            leaveList.append(guildOut[i])
            changeCount += 1

        return newcomerList, leaveList

class execute(QThread):

    updateStatusBarSignal = pyqtSignal(str)

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
    def run(self):
        self.parent.btn_start.setDisabled(True)
        self.parent.btn_check.setDisabled(True)
        self.updateStatusBarSignal.emit('길드원 추출 준비 중..')

        worldName = str(self.parent.combo_serverName.currentText())
        worldNumber = selectedWorld(worldName)

        guildName = self.parent.input_guildName.text()
        if guildName == "" or len(guildName) < 2:
            self.updateStatusBarSignal.emit('추출하기: 정확한 길드명을 입력해주세요')
            return
        
        now = datetime.datetime.now()
        now = now.strftime('%Y-%m-%d')
        folder_path = "GuildData"
        csv_file_name = f'{worldName}_{guildName}_{now}.csv'

        if not os.path.isdir(folder_path): #폴더가 존재하지 않는 경우 폴더 생성
            os.mkdir(folder_path)

        csv_file_path = os.path.join(folder_path, csv_file_name)

        if os.path.exists(csv_file_path): #파일이 존재하는 경우 작업 취소
            self.updateStatusBarSignal.emit('이미 존재하는 파일입니다')
            return

        url = f"https://maplestory.nexon.com/Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
        membersList = []

        try:
            raw = requests.get(url,headers=header)
            html = BeautifulSoup(raw.text,"html.parser")

            checkRankTag = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr')['class']
            if bool(checkRankTag) == False:
                href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > a')['href']
            else:
                href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > dl > dt > a')['href']

            guildUrl = 'https://maplestory.nexon.com'+href
            pageNumber = 1

            while True:
                self.updateStatusBarSignal.emit(f'{guildName} 길드원 추출 중: {pageNumber} /10')
                guildUrlPage = guildUrl+f'&page={pageNumber}'
                try:
                    self.crawlMembers(guildUrlPage,membersList)
                    print('크롤링한 페이지: ',pageNumber)
                except Exception as e:
                    print(e)
                    break

                if pageNumber == 10:

                    print(membersList)
                    print(len(membersList))

                    with open(csv_file_path, 'w', newline='') as csvfile:
                        writer = csv.writer(csvfile)
                        for i in range(len(membersList)):
                            writer.writerow(membersList[i])
                    
                    self.updateStatusBarSignal.emit('추출하기 완료 '+guildName)

                    break
                else:
                    pageNumber += 1

        except Exception as e:
            self.updateStatusBarSignal.emit(f'[ERROR] {e}')
        
    def crawlMembers(self,guildUrl,membersList):
        r = requests.get(guildUrl,headers=header)
        html = BeautifulSoup(r.text,"html.parser") 
        members = html.select('#container > div > div > table > tbody > tr')
        for member in members:
            nick = member.select_one('td.left > span > img')['alt']
            # job = member.select_one('td.left > dl > dd').text #일부 직업들은 기사단/마법사/전사/해적 등으로 표시되어있음
            membersList.append([nick])
      
class WindowClass(QMainWindow, form_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        #프로그램 기본설정
        self.setWindowIcon(QIcon(icon))
        self.setWindowTitle('Guild Checker '+__version__)
        self.statusBar().showMessage('프로그램 정상 구동 중')

        #실행 후 기본값 설정
        self.input_guildName.returnPressed.connect(self.main)

        #버튼 기능
        self.btn_start.clicked.connect(self.main)
        self.btn_load.clicked.connect(self.fileLoad)
        self.btn_check.clicked.connect(self.checkInfo)
        self.btn_github.clicked.connect(self.github)

    @pyqtSlot(str)
    def updateChangesList(self, msg):
        self.guildMembers_changed.append(msg)
    
    @pyqtSlot(str)
    def updateStatusBar(self, msg):
        self.statusBar().showMessage(msg)
    
    def main(self):
        x = execute(self)
        x.updateStatusBarSignal.connect(self.updateStatusBar)
        x.start()
        x.finished.connect(self.on_finished)

    def on_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_check.setEnabled(True)

    def fileLoad(self): #파일 불러오기
        global oldGuildList
        fname = QFileDialog.getOpenFileName(self,'','','Excel(*.csv);;All File(*)')
        
        loadedFile = QFileInfo(fname[0]).fileName()
        if loadedFile != "":
            self.guildMembers.setText('')
            self.guildMembers_changed.setText('')
            self.count.setText('- 명')
            self.changeCount.setText('- 명')
            self.statusBar().showMessage('파일을 불러왔습니다. '+loadedFile)
        else:
            return

        try:
            loadedFileServer, loadedFileGuild, loadedFileDate = loadedFile.split('_')
            self.input_guildName.setText(loadedFileGuild)
            self.combo_serverName.setCurrentText(loadedFileServer)
        except ValueError:
            pass

        count = 0
        oldGuildList = []
        
        try:
            if fname[0]:
                with open(fname[0], newline='') as csvfile:
                    reader = csv.reader(csvfile, delimiter=',')
                    
                    count = 0
                    oldGuildList = []

                    for row in reader:
                        count += 1
                        oldGuildList.append(row[0])
                        self.guildMembers.append(row[0])
        except Exception as e:
            print(e)
            self.statusBar().showMessage(f'지원하지 않거나 손상된 파일입니다.')
            return

        self.count.setText(str(count)+' 명')

    def checkInfo(self): #변동사항 확인
        y = compare(self)
        y.updateChangesListSignal.connect(self.updateChangesList)
        y.updateStatusBarSignal.connect(self.updateStatusBar)
        y.start()
        y.finished.connect(self.on_finished)

    def github(self):
        webbrowser.open_new_tab('https://github.com/memoday/guildChecker2')

    def closeEvent(self, event):
        sys.exit(0)

if __name__ == "__main__":
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
