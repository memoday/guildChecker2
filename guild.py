import os, sys
import requests
from bs4 import BeautifulSoup
import datetime
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import checkInfo as ci
import tracker as track
import webbrowser
import csv

__version__ = "DEV"

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
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
    
    def run(self):
        self.parent.btn_start.setDisabled(True)
        self.parent.statusBar().showMessage('최신 길드원 정보 불러오는 중...')

        worldName = str(self.parent.combo_serverName.currentText())
        worldNumber = selectedWorld(worldName)

        guildName = self.parent.input_guildName.text()
        if guildName == "":
            self.parent.statusBar().showMessage('변동사항 확인: 길드 이름을 입력해주세요')
            return

        url = f"https://maplestory.nexon.com/Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
        recentMemberList = []

        try:
            raw = requests.get(url,headers=header)
            html = BeautifulSoup(raw.text,"html.parser")
            href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > a')['href']

            guildUrl = 'https://maplestory.nexon.com'+href
            pageNumber = 1

            while True:
                self.parent.statusBar().showMessage(f'{guildName} 최신 길드원 정보 추출 중: {pageNumber} /10')
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
                    print('10번째 페이지로 크롤링을 종료합니다\n')
                    print(f'최신 길드원 목록{len(recentMemberList)}')
                    print(recentMemberList)

                    self.parent.statusBar().showMessage(f'변동사항 확인 중... {guildName}')

                    newcomerList, leaveList = self.checkChanges(recentMemberList)
                    for i in range(len(newcomerList)):
                        self.parent.guildMembers_changed.append('[신규] '+newcomerList[i])
                    for i in range(len(leaveList)):
                        self.parent.guildMembers_changed.append('[탈퇴] '+leaveList[i])

                    changeCount = len(newcomerList) + len(leaveList)
                    self.parent.changeCount.setText(str(changeCount)+' 명')
                    
                    self.parent.statusBar().showMessage('변동사항 확인 완료 '+guildName)
                    self.parent.btn_start.setEnabled(True)

                    break
                else:
                    pageNumber += 1

        except Exception as e:
            self.parent.statusBar().showMessage(f'[ERROR] {e}')
            self.parent.btn_start.setEnabled(True)
        
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
            changed = ('[신규] '+guildIn[i])
            newcomerList.append(guildIn[i])
            changeCount += 1
        
        for i in range(len(guildOut)):
            print('[탈퇴]',guildOut[i])
            changed = ('[탈퇴] '+guildOut[i])
            leaveList.append(guildOut[i])
            changeCount += 1

        return newcomerList, leaveList



class execute(QThread):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
    def run(self):

        self.parent.btn_start.setDisabled(True)
        self.parent.statusBar().showMessage('길드원 추출 준비 중..')

        worldName = str(self.parent.combo_serverName.currentText())
        worldNumber = selectedWorld(worldName)

        guildName = self.parent.input_guildName.text()
        if guildName == "":
            self.parent.statusBar().showMessage('추출하기: 길드 이름을 입력해주세요')
            return

        url = f"https://maplestory.nexon.com/Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
        membersList = []

        try:
            raw = requests.get(url,headers=header)
            html = BeautifulSoup(raw.text,"html.parser")
            href = html.select_one('#container > div > div > div:nth-child(4) > div.rank_table_wrap > table > tbody > tr > td:nth-child(2) > span > a')['href']

            guildUrl = 'https://maplestory.nexon.com'+href
            pageNumber = 1

            while True:
                self.parent.statusBar().showMessage(f'{guildName} 길드원 추출 중: {pageNumber} /10')
                guildUrlPage = guildUrl+f'&page={pageNumber}'
                try:
                    self.crawlMembers(guildUrlPage,membersList)
                    print('크롤링한 페이지: ',pageNumber)
                except Exception as e:
                    print(e)
                    break

                if pageNumber == 10:

                    print('10번째 페이지로 크롤링을 종료합니다\n')
                    print(membersList)
                    print(len(membersList))

                    now = datetime.datetime.now()
                    now = now.strftime('%Y-%m-%d')

                    with open(f'{worldName}_{guildName}_{now}.csv', 'w', newline='') as csvfile:
                        writer = csv.writer(csvfile)
                        for i in range(len(membersList)):
                            writer.writerow(membersList[i])
                    
                    self.parent.statusBar().showMessage('추출하기 완료. '+guildName)
                    self.parent.btn_start.setEnabled(True)

                    break
                else:
                    pageNumber += 1

        except Exception as e:
            self.parent.statusBar().showMessage(f'[ERROR] {e}')
            self.parent.btn_start.setEnabled(True)
        
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
    
    def main(self):
        self.guildMembers_changed.setText('')
        x = execute(self)
        x.start()

    def fileLoad(self): #파일 불러오기
        global oldGuildList
        fname = QFileDialog.getOpenFileName(self,'','','Excel(*.xlsx, *.csv);; ;;All File(*)')
        
        self.guildMembers.setText('')
        self.guildMembers_changed.setText('')
        self.changeCount.setText('- 명')

        loadedFile = QFileInfo(fname[0]).fileName()
        if loadedFile != "":
            self.statusBar().showMessage('파일을 불러왔습니다. '+loadedFile)

        try:
            loadedFileServer, loadedFileGuild, loadedFileDate = loadedFile.split('_')
            self.input_guildName.setText(loadedFileGuild)
            self.combo_serverName.setCurrentText(loadedFileServer)
        except ValueError:
            pass

        count = 0
        oldGuildList = []
        
        if fname[0]:
            with open(fname[0], newline='') as csvfile:
                reader = csv.reader(csvfile, delimiter=',')
                
                count = 0
                oldGuildList = []

                for row in reader:
                    count += 1
                    oldGuildList.append(row[0])
                    self.guildMembers.append(row[0])
                    print(row[0])

        self.count.setText(str(count)+' 명')

    def checkInfo(self): #변동사항 확인
        y = compare(self)
        y.start()

    def github(self):
        webbrowser.open_new_tab('https://github.com/memoday/guildMemberChecker/releases')

    def closeEvent(self, event):
        sys.exit(0)

if __name__ == "__main__":
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
