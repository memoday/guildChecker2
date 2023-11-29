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

__version__ = "v2.1.1"

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

class compareCSV(QThread):
    updateChangesListSignal = pyqtSignal(str)
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
            self.updateStatusBarSignal.emit('추출하기: 정확한 길드명을 입력해주세요.')
            return
        
        now = datetime.datetime.now()
        now = now.strftime('%Y-%m-%d')
        folder_path = "GuildData"
        csv_file_name = f'{worldName}_{guildName}_{now}.csv'

        if not os.path.isdir(folder_path): #폴더가 존재하지 않는 경우 폴더 생성
            os.mkdir(folder_path)

        try:
            old_csv_file_path = os.path.join(folder_path,fname)
        except NameError:
            self.updateStatusBarSignal.emit('길드 파일을 불러와주세요.')
            return
        csv_file_path = os.path.join(folder_path, csv_file_name)

        if os.path.exists(old_csv_file_path) == False: #불러온 길드 파일이 존재하지 않는 경우
            self.updateStatusBarSignal.emit('길드 파일이 존재하지 않습니다.')
            return

        if os.path.exists(csv_file_path) == False: #파일이 존재하는 않는 경우 새로 크롤링
            print('최신 길드 정보 파일이 없어 최신 길드 정보를 추출합니다.')
            try:
                self.startCrawl(guildName,worldNumber,csv_file_path)
            except requests.exceptions.RequestException:
                self.updateStatusBarSignal.emit('서버에 접속하는 중 오류가 발생했습니다. 나중에 다시 시도 해주세요.')
                return
            except TypeError:
                self.updateStatusBarSignal.emit('길드 랭킹에서 해당 길드를 찾을 수 없습니다.')
                return
            except:
                self.updateStatusBarSignal.emit('알 수 없는 오류가 발생했습니다.')
                return

        try:
            nickChangedList,newMembersList,leavedMembersList = self.compare(old_csv_file_path,csv_file_path)
        except UnicodeDecodeError:
            self.updateStatusBarSignal.emit('지원하지 않거나 손상된 파일입니다.')
            return

        for i in range(len(newMembersList)):
            self.updateChangesListSignal.emit('[신규] ' + newMembersList[i])
        for i in range(len(leavedMembersList)):
            self.updateChangesListSignal.emit('[탈퇴] ' + leavedMembersList[i])
        for i in range(len(nickChangedList)):
            self.updateChangesListSignal.emit('[닉변] ' + nickChangedList[i][0] + ' -> ' + nickChangedList[i][1])

        changeCount = len(newMembersList) + len(leavedMembersList) + len(nickChangedList)

        self.parent.newCounts.setText(str(len(newMembersList)))
        self.parent.leavedCounts.setText(str(len(leavedMembersList)))
        self.parent.changedCounts.setText(str(len(nickChangedList)))

        
        self.updateStatusBarSignal.emit('변동사항 확인 완료. '+guildName)
    
    def startCrawl(self,guildName,worldNumber,csv_file_path):
        url = f"https://maplestory.nexon.com/N23Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
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
                    print("execute:"+e)
                    break

                if pageNumber == 10:

                    with open(csv_file_path, 'w', newline='',encoding='utf-8-sig') as csvfile:
                        writer = csv.writer(csvfile)
                        for i in range(len(membersList)):
                            writer.writerow(membersList[i])
                    
                    self.updateStatusBarSignal.emit('추출하기 완료. '+guildName)

                    break
                else:
                    pageNumber += 1

        except requests.exceptions.RequestException as e:
            print(e)
            raise e

        except TypeError:
            raise TypeError

        except Exception as e:
            self.updateStatusBarSignal.emit(f'[ERROR] {e}')
            raise e

    def crawlMembers(self,guildUrl,membersList):
        r = requests.get(guildUrl,headers=header)
        html = BeautifulSoup(r.text,"html.parser") 
        members = html.select('#container > div > div > table > tbody > tr')
        for member in members:
            nick = member.select_one('td.left > span > img')['alt']
            # job = member.select_one('td.left > dl > dd').text #직업군으로 표시되는 부분을 개선하기 위해 getRankingInfo function 활용해야함
            level = member.select_one('td:nth-child(3)').text
            exp = member.select_one('td:nth-child(4)').text
            fame = member.select_one('td:nth-child(5)').text
            character_link = member.select_one('td.left > dl > dt > a')['href']
            jobAndRankData = self.getRankingInfo(nick,character_link)

            while len(jobAndRankData) < 9:
                jobAndRankData.append('')  # Add an empty string

            membersList.append([nick,jobAndRankData[0],level,exp,fame,jobAndRankData[1],jobAndRankData[2],jobAndRankData[3],jobAndRankData[4],jobAndRankData[5],jobAndRankData[6],jobAndRankData[7],jobAndRankData[8]])

    def getRankingInfo(self,nickname,characterHref):
        jobAndRankData = []
        characterHref = characterHref.replace("?p=","/Ranking?p=")
        character_link = f"https://maplestory.nexon.com{characterHref}"

        try:
            r = requests.get(character_link,headers=header)
            html = BeautifulSoup(r.text,"html.parser")

            job_details = html.select_one('#wrap > div.center_wrap > div.char_info_top > div.char_info > dl:nth-child(2) > dd').text
            jobType, job = job_details.split("/")
            jobAndRankData.append(job)
            rankDates = html.select('#container > div.con_wrap > div.contents_wrap > div > table > tbody > tr')
            for rankDate in rankDates:
                date = rankDate.select_one('td.date').text
                comprehensiveRanking = rankDate.select_one('td:nth-child(2)').text

                rankingData = date+'R'+comprehensiveRanking
                jobAndRankData.append(rankingData)
            return jobAndRankData
        except:
            print(f'{nickname} 랭킹정보를 불러올 수 없습니다.')
            return jobAndRankData

    def read_csv_into_dict(self, file_path):
        data_dict = {}
        with open(file_path, 'r', newline='', encoding='utf-8-sig') as csv_file:
            csv_reader = csv.reader(csv_file)
            # Skip the header row if needed
            header = next(csv_reader)  # Read and discard the header row
            for row in csv_reader:
                nickname = row[0]
                rank_data = []
                for i in range(5, 13):  # Columns for 'rank1' to 'rank8'
                    rank_data.append(row[i] if len(row) > i else '')  # Check if rank column exists
                data_dict[nickname] = {'rankdata': rank_data}
        return data_dict

    def compare(self,old_csv_file_path,csv_file_path):
        old_csv_dict = self.read_csv_into_dict(old_csv_file_path)
        new_csv_dict = self.read_csv_into_dict(csv_file_path)

        new_nicknames_dict = {nickname: new_csv_dict[nickname]['rankdata'] for nickname in new_csv_dict if nickname not in old_csv_dict}
        removed_nicknames_dict = {nickname: old_csv_dict[nickname]['rankdata'] for nickname in old_csv_dict if nickname not in new_csv_dict}

        dict2_reverse = {}
        for nickname, rankdata_list in new_nicknames_dict.items():
            for rankdata in rankdata_list:
                dict2_reverse[rankdata] = nickname

        # Find matches
        matches = []
        for nickname1, rankdata_list1 in removed_nicknames_dict.items():
            for rankdata1 in rankdata_list1:
                if rankdata1 in dict2_reverse:
                    if rankdata1 == '':
                        continue
                    nickname2 = dict2_reverse[rankdata1]
                    matches.append([nickname1, nickname2])
                    break

        for old_nick, new_nick in matches:
            if old_nick in removed_nicknames_dict:
                del removed_nicknames_dict[old_nick]
            if new_nick in new_nicknames_dict:
                del new_nicknames_dict[new_nick]

        removed_nicknames_key = list(removed_nicknames_dict.keys())
        new_nicknames_key = list(new_nicknames_dict.keys())

        print(f'닉네임 변경: {matches}')
        print(f'신규 길드원: {new_nicknames_key}')
        print(f'탈퇴 길드원: {removed_nicknames_key}')

        return matches, new_nicknames_key, removed_nicknames_key
   
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
            self.updateStatusBarSignal.emit('추출하기: 정확한 길드명을 입력해주세요.')
            return
        
        now = datetime.datetime.now()
        now = now.strftime('%Y-%m-%d')
        folder_path = "GuildData"
        csv_file_name = f'{worldName}_{guildName}_{now}.csv'

        if not os.path.isdir(folder_path): #폴더가 존재하지 않는 경우 폴더 생성
            os.mkdir(folder_path)

        csv_file_path = os.path.join(folder_path, csv_file_name)

        if os.path.exists(csv_file_path): #파일이 존재하는 경우 작업 취소
            self.updateStatusBarSignal.emit('최신 길드 데이터 파일이 이미 존재합니다.')
            return

        url = f"https://maplestory.nexon.com/N23Ranking/World/Guild?w={worldNumber}&t=1&n={guildName}"
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
                    print("execute:"+e)
                    break

                if pageNumber == 10:

                    print(membersList)
                    print(len(membersList))

                    with open(csv_file_path, 'w', newline='',encoding='utf-8-sig') as csvfile:
                        writer = csv.writer(csvfile)
                        for i in range(len(membersList)):
                            writer.writerow(membersList[i])
                    
                    self.updateStatusBarSignal.emit('추출하기 완료. '+guildName)

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
            # job = member.select_one('td.left > dl > dd').text #직업군으로 표시되는 부분을 개선하기 위해 getRankingInfo function 활용해야함
            level = member.select_one('td:nth-child(3)').text
            exp = member.select_one('td:nth-child(4)').text
            fame = member.select_one('td:nth-child(5)').text
            character_link = member.select_one('td.left > dl > dt > a')['href']
            jobAndRankData = self.getRankingInfo(nick,character_link)

            while len(jobAndRankData) < 9:
                jobAndRankData.append('')  # Add an empty string

            membersList.append([nick,jobAndRankData[0],level,exp,fame,jobAndRankData[1],jobAndRankData[2],jobAndRankData[3],jobAndRankData[4],jobAndRankData[5],jobAndRankData[6],jobAndRankData[7],jobAndRankData[8]])

    def getRankingInfo(self,nickname,characterHref):
        jobAndRankData = []
        characterHref = characterHref.replace("?p=","/Ranking?p=")
        character_link = f"https://maplestory.nexon.com{characterHref}"

        try:
            r = requests.get(character_link,headers=header)
            html = BeautifulSoup(r.text,"html.parser")

            job_details = html.select_one('#wrap > div.center_wrap > div.char_info_top > div.char_info > dl:nth-child(2) > dd').text
            jobType, job = job_details.split("/")
            jobAndRankData.append(job)
            rankDates = html.select('#container > div.con_wrap > div.contents_wrap > div > table > tbody > tr')
            for rankDate in rankDates:
                date = rankDate.select_one('td.date').text
                comprehensiveRanking = rankDate.select_one('td:nth-child(2)').text

                rankingData = date+'R'+comprehensiveRanking
                jobAndRankData.append(rankingData)
            return jobAndRankData
        except:
            print(f'{nickname} 랭킹정보를 불러올 수 없습니다.')
            return jobAndRankData
      
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
        self.btn_discord.clicked.connect(self.discord)

    @pyqtSlot(str)
    def updateChangesList(self, msg):
        self.guildMembers_changed.append(msg)
    
    @pyqtSlot(str)
    def updateStatusBar(self, msg):
        self.statusBar().showMessage(msg)
        
    def main(self):
        self.input_guildName.blockSignals(True)
        x = execute(self)
        x.updateStatusBarSignal.connect(self.updateStatusBar)
        x.start()
        x.finished.connect(self.on_finished)

    def on_finished(self):
        self.btn_start.setEnabled(True)
        try:
            fname
            self.btn_check.setEnabled(True)
        except NameError:
            self.btn_check.setEnabled(False)
        self.input_guildName.blockSignals(False)

    def fileLoad(self): #파일 불러오기
        global fname
        
        script_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
        initial_dir = os.path.join(script_dir, "GuildData")

        if not os.path.exists(initial_dir):
            initial_dir = script_dir

        fname, _ = QFileDialog.getOpenFileName(self,'',initial_dir,'Excel(*.csv);;All File(*)')
        loadedFile = QFileInfo(fname).fileName()

        try:
            with open(fname, 'r', newline='', encoding='utf-8-sig') as csv_file:
                csv_reader = csv.reader(csv_file)
                # Skip the header row if needed
                header = next(csv_reader)
        except:
            self.statusBar().showMessage('지원하지 않거나 손상된 파일입니다.')
            return

        if loadedFile != "":
            print(fname)
            self.guildMembers_changed.setText('')
            self.newCounts.setText('-')
            self.leavedCounts.setText('-')
            self.changedCounts.setText('-')
            self.csvFilepath.setText(fname)
            self.statusBar().showMessage('파일을 불러왔습니다. '+loadedFile)
            self.btn_check.setEnabled(True)
        else:
            return

        try:
            loadedFileServer, loadedFileGuild, loadedFileDate = loadedFile.split('_')
            self.input_guildName.setText(loadedFileGuild)
            self.combo_serverName.setCurrentText(loadedFileServer)
        except ValueError:
            pass

    def checkInfo(self): #변동사항 확인
        self.input_guildName.blockSignals(True)
        self.guildMembers_changed.setText('')
        y = compareCSV(self)
        y.updateChangesListSignal.connect(self.updateChangesList)
        y.updateStatusBarSignal.connect(self.updateStatusBar)
        y.start()
        y.finished.connect(self.on_finished)

    def discord(self):
        webbrowser.open_new_tab('https://discord.gg/GTXVQqqTT2')

    def closeEvent(self, event):
        sys.exit(0)

if __name__ == "__main__":
    app = QApplication(sys.argv) 
    myWindow = WindowClass()
    #open qss file
    # file = open("ui/MaterialDark.qss",'r')

    # with file:
    #     qss = file.read()
    #     app.setStyleSheet(qss)

    myWindow.show()
    app.exec_()
