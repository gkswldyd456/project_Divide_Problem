import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
import os, shutil

import win32com.client as win
import time
from PIL import Image
import olefile
import re





hwp=win.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") #한글 권한 설정 보안 레지스트리
hwp.XHwpWindows.Item(0).Visible=True #백그라운드에서 동작(False) / 보이면서 동작(True)
hwp.SetMessageBoxMode(0x00020000) # 예/아니오 -> 아니오
hwp.SetMessageBoxMode(0x00000001) # 확인 박스 -> OK
# hwp.SetMessageBoxMode(0x00000010) # 확인/취소 박스 -> OK


root = Tk()
root.title("Postmath_Divide_HWP")

# 한글 문제 해설 정답 쪼개기 관련 함수 모음

def find_mizu(): # 미주 찾아가기
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 31)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)


def find_mizunum(): # 미주 번호 찾아가기
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 32)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)


def count_mizu(): # 미주 개수 세기 -> 변수 cnt_mizu 에 저장
    hwp.MovePos(2) # 문서 제일 앞으로
    find_mizunum() # 미주번호 찾아가 : hwp.MovePos(2) 이거 때매 맨 첫 미주번호로
    global cnt_mizu # 미주개수를 전역 변수로
    cnt_mizu = 1
    while hwp.HAction.Run("NoteToNext") == True:
        cnt_mizu = cnt_mizu + 1
    hwp.MovePos(2) # 문서 제일 앞으로
    


def onepageoneproblem(): # 한쪽에 한문제 (첫페이지는 표지)
    hwp.MovePos(2)
    i=0
    while i < cnt_mizu :
        find_mizu()
        hwp.HAction.Run("BreakPage"); # ctrl+enter
        hwp.HAction.Run("MoveRight");
        i = i+1
    hwp.HAction.Run("MoveColumnEnd"); #단의 끝점으로 이동
    hwp.HAction.Run("BreakPage");
    hwp.MovePos(3) # 문서 제일 뒤로
    hwp.HAction.Run("MoveLeft");
    hwp.HAction.Run("BreakPage");
    hwp.HAction.Run("MoveRight");
    hwp.HAction.Run("BreakPage");
    hwp.MovePos(2) # 문서 제일 앞으로
    

def count_common_problem(file_fullname): # common_problems -> 공통지문 리스트 / cnt_common_problems -> 공통지문 개수 / 한글 열기 전에 써야함
    f = olefile.OleFileIO(file_fullname) # HWP 파일 열기
    encoded_text = f.openstream('PrvText').read() # PrvText 스트림의 내용 꺼내기
    decoded_text = encoded_text.decode('UTF-16') # 유니코드를 UTF-16으로 디코딩
    global common_problems, cnt_common_problems
    common_problems = re.findall('\$\$(\d+)#(\d+)\$\$', decoded_text)
    cnt_common_problems = len(common_problems)
    f.close()

def find_commonproble(): # 공통지문 찾아가기 (공통지문 맨 앞으로)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.ReplaceString = "";
    hwp.HParameterSet.HFindReplace.FindString = "\\$\\$\\d+#\\d+\\$\\$";
    hwp.HParameterSet.HFindReplace.IgnoreReplaceString = 0;
    hwp.HParameterSet.HFindReplace.IgnoreFindString = 0;
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward");
    hwp.HParameterSet.HFindReplace.WholeWordOnly = 0;
    hwp.HParameterSet.HFindReplace.UseWildCards = 0;
    hwp.HParameterSet.HFindReplace.SeveralWords = 0;
    hwp.HParameterSet.HFindReplace.AllWordForms = 0;
    hwp.HParameterSet.HFindReplace.MatchCase = 0;
    hwp.HParameterSet.HFindReplace.ReplaceMode = 0;
    hwp.HParameterSet.HFindReplace.ReplaceStyle = "";
    hwp.HParameterSet.HFindReplace.FindStyle = "";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HParameterSet.HFindReplace.FindJaso = 0;
    hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0;
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
    hwp.HParameterSet.HFindReplace.FindType = 1;
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.Run("Cancel");
    hwp.HAction.Run("MoveLineBegin");
    
def find_sonproble(): # 새끼문제 @ 찾아가기 (@앞으로)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.ReplaceString = "";
    hwp.HParameterSet.HFindReplace.FindString = "@";
    hwp.HParameterSet.HFindReplace.IgnoreReplaceString = 0;
    hwp.HParameterSet.HFindReplace.IgnoreFindString = 0;
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward");
    hwp.HParameterSet.HFindReplace.WholeWordOnly = 0;
    hwp.HParameterSet.HFindReplace.UseWildCards = 0;
    hwp.HParameterSet.HFindReplace.SeveralWords = 0;
    hwp.HParameterSet.HFindReplace.AllWordForms = 0;
    hwp.HParameterSet.HFindReplace.MatchCase = 0;
    hwp.HParameterSet.HFindReplace.ReplaceMode = 0;
    hwp.HParameterSet.HFindReplace.ReplaceStyle = "";
    hwp.HParameterSet.HFindReplace.FindStyle = "";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HParameterSet.HFindReplace.FindJaso = 0;
    hwp.HParameterSet.HFindReplace.HanjaFromHangul = 0;
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1;
    hwp.HParameterSet.HFindReplace.FindType = 1;
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.Run("Cancel");
    hwp.HAction.Run("MoveLeft");

    
def onepage_commonproble(): # 공통지문 형식 찾아가서 ctrl+enter 
    hwp.MovePos(2)
    if cnt_common_problems == 0 :
        pass
    else:
        for i in range(1, cnt_common_problems+1):
            find_commonproble()
            hwp.HAction.Run("BreakPage");
            hwp.HAction.Run("MoveNextParaBegin");
    hwp.MovePos(2)


def page_size_set(): # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
    hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)
    hwp.HParameterSet.HSecDef.PageDef.PaperWidth = hwp.MiliToHwpUnit(106.5)
    hwp.HParameterSet.HSecDef.PageDef.PaperHeight = hwp.MiliToHwpUnit(1110.0)
    hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(0.0)
    hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)


def multicolumn_1(): # 페이지 1단으로 
    hwp.HAction.GetDefault("MultiColumn", hwp.HParameterSet.HColDef.HSet)
    hwp.HParameterSet.HColDef.Count = 1
    hwp.HParameterSet.HColDef.SameGap = hwp.MiliToHwpUnit(0.0)
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyClass", 832)
    hwp.HParameterSet.HColDef.HSet.SetItem("ApplyTo", 6)
    hwp.HAction.Execute("MultiColumn", hwp.HParameterSet.HColDef.HSet)


def Allreplace_circ(): # 원숫자 다 괄호숫자로 바꾸기
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "①"; 
    hwp.HParameterSet.HFindReplace.ReplaceString = "(1)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "②";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(2)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "③";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(3)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "④";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(4)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "⑤";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(5)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "⑥";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(6)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "⑦";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(7)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "⑧";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(8)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "⑨";
    hwp.HParameterSet.HFindReplace.ReplaceString = "(9)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);


def Allreplace_rhfqoddl(): # 괄호숫자 앞에 골뱅이 붙여 바꾸기
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(1)"; 
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(1)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(2)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(2)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(3)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(3)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(4)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(4)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(5)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(5)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(6)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(6)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(7)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(7)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(8)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(8)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "(9)";
    hwp.HParameterSet.HFindReplace.ReplaceString = "@(9)";
    hwp.HParameterSet.HFindReplace.FindRegExp = 0;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);


def tabdiv_pro_sol(): # 1탭에 문제만 2탭에 해설만 / 작동 후 해설 맨 처음
    # 새 탭 만들고 다시 돌아와
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("WindowNextTab")


    # 미주 찾아가서 그 미주있는 페이지 복사 후 새탭에 넣어 (새탭1)
    # hwp.MovePos(2) # 문서 제일 앞으로
    find_mizu()
    hwp.HAction.Run("MoveSelPageDown")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    time.sleep(0.1)
    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("PasteOriginal")
    time.sleep(0.1)
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)


    # 미주 번호 뒤에 있는 내용(해설내용) 복사 후 새탭에 넣어 (새탭2)
    hwp.MovePos(2)
    find_mizunum()
    time.sleep(0.1)
    hwp.HAction.Run("SelectAll")
    time.sleep(0.1)
    hwp.HAction.Run("Copy")
    time.sleep(0.1)
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("PasteOriginal")
    time.sleep(0.1)
    hwp.MovePos(2)
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("DeleteBack")
    
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)


    # 문제 부분 미주 없애고 해설 페이지로 돌아가
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("WindowNextTab")
    hwp.MovePos(2)
    find_mizu()
    hwp.HAction.Run("MoveSelRight")
    hwp.HAction.Run("DeleteBack")
    hwp.HAction.Run("MoveViewEnd");
    hwp.HAction.Run("MoveSelTopLevelEnd");
    hwp.HAction.Run("Delete");
    hwp.HAction.Run("WindowNextTab")
    time.sleep(0.1)
    hwp.MovePos(3) # 해설 맨 마지막 끝으로가서
    
    global solution_page
    solution_page = hwp.KeyIndicator()[3] # 현재 커서의 페이지 번호 저장
    
    hwp.MovePos(2) # 해설 맨 처음으로
    
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet); # 해설 내용 중 정답이 있을 수 있으니 일단 하나 써주고
    hwp.HParameterSet.HInsertText.Text = "[정답] ";
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
    
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); 
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "\\[정답\\]\\b*\\[*\\(*정답\\]*\\)*\\b*"; # [정답] 정답 형태 -> [정답] 으로 다 바꿔
    hwp.HParameterSet.HFindReplace.ReplaceString = "[정답] "; # 바꾸는 문자는 정규식 안먹음
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); 
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "@"; # @ 다 지워
    hwp.HParameterSet.HFindReplace.ReplaceString = ""; 
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    
    hwp.MovePos(2)


def tabdiv_commonpro_sol(i): # 공통지문 -> 1탭에 문제 / i는 몇번째 공통지문 / 작동 후 1탭에 맨처음
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("WindowNextTab")
    find_commonproble()
    hwp.HAction.Run("MoveSelPageDown")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Copy")
    time.sleep(0.1)
    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("PasteOriginal")
    time.sleep(0.1)
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)
    hwp.MovePos(2)
    global start_num, last_num
    start_num = common_problems[i-1][0]
    last_num = common_problems[i-1][1]
    hwp.HAction.Run("MoveSelNextParaBegin");
    hwp.HAction.Run("DeleteBack");
    hwp.MovePos(2)
    
def tabdiv_sonpro_sol(): # 새끼문제 -> 1탭에 문제/ 작동 후 1탭에 맨처음
    hwp.HAction.Run("FileNewTab")
    hwp.HAction.Run("WindowNextTab")
    find_sonproble() # 새끼문제 @ 찾아가기 (@앞으로)
    hwp.HAction.Run("MoveSelPageDown")
    hwp.HAction.Run("MoveSelLeft")
    hwp.HAction.Run("Cut")
    time.sleep(0.1)
    hwp.HAction.Run("Cancel")
    hwp.HAction.Run("WindowNextTab")
    hwp.HAction.Run("Paste")
    hwp.HAction.Run("PasteOriginal")
    time.sleep(0.1)
    # page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    # multicolumn_1() # 페이지 1단으로 
    time.sleep(0.1)
    hwp.MovePos(2)
    
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); 
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "@?\\b*\\(\\d+\\)\\b*"; # @ 다 지워
    hwp.HParameterSet.HFindReplace.ReplaceString = " "; 
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    hwp.MovePos(2)
    



def image_merge(imgname): # 해설 이미지 여러쪽이면 병합하고 남는거 지우기 
    im_list = []
    for i in range(1, solution_page+1):
        im_list.append(imgname+str(i).zfill(3)+".png")


    images = [Image.open(x) for x in im_list]

    # size -> size[0] : width, size[1] : height
    widths = [x.size[0] for x in images]
    heights = [x.size[1] for x in images]

    # 가장 width가 넓고 heights를 다 합친 큰 스케치북을 준비
    # 먼저 최대 넓이, 전체 높이 구해오기
    max_width, total_height = max(widths), sum(heights)

    # 스케치북 준비
    result_img = Image.new("RGB", (max_width, total_height)) 
    y_offset = 0 # y 위치 정보
    for img in images:
        result_img.paste(img, (0, y_offset))
        y_offset += img.size[1] # height 값 만큼 더해줌
    result_img.save(imgname+".png", "png")
    time.sleep(0.2)
    for i in range(1, solution_page+1):
        os.remove(imgname+str(i).zfill(3)+".png")


def save_sol_hml(num):  # 해설 hml파일 저장할거야 num = 문제번호
    global sol_hml_name
    sol_hml_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [1hml]" # 확장자 없는 이름
    while os.path.isfile(sol_hml_name+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = sol_hml_name+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_sol_png(num): # 해설 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global sol_png_name
    sol_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [2해설] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = sol_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(1)


def save_sol_hwp(num):  # 해설 hwp파일 저장할거야 num = 문제번호
    global sol_hwp_name
    sol_hwp_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [2해설] [3hwp]" # 확장자 없는 이름
    while os.path.isfile(sol_hwp_name+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = sol_hwp_name+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)


def save_pro_hml(num):  # 문제 hml파일 저장할거야 num = 문제번호
    global pro_hml_name
    pro_hml_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [1문제] [1hml]"
    while os.path.isfile(pro_hml_name+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = pro_hml_name+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_pro_png(num): # 문제 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global pro_png_name
    pro_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [1문제] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = pro_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)
    
    os.rename(pro_png_name+"001"+".png", pro_png_name+".png")
    
    
def save_commonpro_hml(start_num, last_num):  # 공통문제 hml파일 저장할거야 start_num=시작번호, last_num=끝번호
    global commonpro_hml_name
    commonpro_hml_name = file_fullname.replace(".hwp", "") + " "+ str(start_num).zfill(3) +"-"+str(last_num).zfill(3)+ "번 [0공통문제] [1hml]"
    while os.path.isfile(commonpro_hml_name+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = commonpro_hml_name+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

def save_commonpro_png(start_num, last_num): # 공통문제 이미지(png)파일 저장할거야 start_num=시작번호, last_num=끝번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global commonpro_png_name
    commonpro_png_name = file_fullname.replace(".hwp", "") + " "+ str(start_num).zfill(3) +"-"+str(last_num).zfill(3)+ "번 [0공통문제] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = commonpro_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)
    
    os.rename(commonpro_png_name+"001"+".png", commonpro_png_name+".png")

def save_presol_hwp(num, type):  # 정답 hwp파일 저장할거야 num = 문제번호, type = 정답종류
    global presol_hwp_name
    presol_hwp_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) +  "번 [3정답] [1hwp]." + str(type) # 확장자 없는 이름
    while os.path.isfile(presol_hwp_name+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = presol_hwp_name+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)
    


def save_presol_png(num): # 정답 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global presol_png_name
    presol_png_name = file_fullname.replace(".hwp", "") + " "+ str(num).zfill(3) + "번 [3정답] [2png]"
    hwp.HParameterSet.HFileOpenSave.filename = presol_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(0.2)
    
    os.rename(presol_png_name+"001"+".png", presol_png_name+".png")



def save_sonsol_hwp(num, type):  # 새끼문제정답 hwp파일 저장할거야 num = 문제번호, type = 정답종류
    global sonsol_hwp_name
    sonsol_hwp_name = son_file_name.replace(".hwp", "") + "-"+str(num)+"번" + str(type) # 확장자 없는 이름
    while os.path.isfile(sonsol_hwp_name+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = sonsol_hwp_name+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)                    

def save_sonsol_png(num): # 정답 이미지(png)파일 저장할거야 num = 문제번호
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    global sonsol_png_name
    sonsol_png_name = son_file_name.replace(".hwp", "") + "-"+str(num)+"번"
    sonsol_png_name = sonsol_png_name.replace("1hwp", "2png")
    hwp.HParameterSet.HFileOpenSave.filename = sonsol_png_name+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(1)
    
    os.rename(sonsol_png_name+"001"+".png", sonsol_png_name+".png")


def find_equation(): #내 옆 수식을 텍스트로
    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
    hwp.HParameterSet.HGotoE.HSet.SetItem('DialogResult', 37)
    hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
    hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
    



def extract_eqn(hwp):  # 수식을 선택한 후 수식안 텍스트 추출
    Act = hwp.CreateAction("EquationModify")
    Set = Act.CreateSet()
    Pset = Set.CreateItemSet("EqEdit", "EqEdit")
    Act.GetDefault(Pset)
    return Pset.Item("String") # Pset.Item("String") 여기에 수식 안 내용 저장됨


"""모든 수식 텍스트 차례로 dict로 얻기.
키는 (List, Para, Pos), 값은 eqn_string"""

def equation_to_text_all(hwp): # 한글 파일 열어서 모든 수식 text화
    eqn_dict = {}  # 사전 형식의 자료 생성 예정 ##
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    global sum_reST
    sum_reST = [] # reST들을 하나의 리스트에 담기 = 한 문제 정답에 여러 수식이 있을 수 있음

    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = position.Item("List"), position.Item("Para"), position.Item("Pos")
            hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            Act = hwp.CreateAction("EquationModify")
            Set = Act.CreateSet()
            Pset = Set.CreateItemSet("EqEdit", "EqEdit")
            Act.GetDefault(Pset) # Pset.Item("String") 여기에 수식 안 내용 저장됨
            
            #임시 해보기
            ST = re.sub('`|~|\s|\n|rm|it|bold|\{|\}', "", Pset.Item("String"), flags=re.I) # 수식 안 내용 중 `,~,띄어,엔터,rm,it,{,} 이거 제외 / flags=re.I 이거는 대소문자 구분 없이
            ST1 = re.sub('`|~|bold|\s|\n', "", Pset.Item("String"), flags=re.I) # 수식 안 내용 중 `,~,띄어,엔터 이거 제외
            ST2 = re.sub('\{*rm\{*(\d+)\}*\}*', r'\1', ST1) # ST1 내용 중 {rm{숫자}} -> 숫자로만 빼냄
            
            reST = re.findall('[^\d+-]', ST) # ST안 내용 중 숫자,+,-를 제외한 것들 리스트에 담아
            for i in reST:
                sum_reST.append(i)

            hwp.HAction.Run("Delete") ## 여기부터는 한글에서 수식을 다 텍스트로

            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            # hwp.HParameterSet.HInsertText.Text = Pset.Item("String");
            hwp.HParameterSet.HInsertText.Text = "$"+ST2+"$";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            eqn_dict[position] = Pset.Item("String")  # 좌표가 key이고, 수식문자열이 value인 사전 생성 ##
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.Run("Cancel")  # 완료했으면 선택해제
    # print(sum_reST)

    # for key, value in eqn_dict.items(): ##
    #     print(f"{key} : {value}")

def equation_to_text_first(hwp): # 한글 파일 열어서 첫수식 text화
    eqn_dict = {}  # 사전 형식의 자료 생성 예정 ##
    
    hwp.MovePos(2)
    
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.

    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            position = ctrl.GetAnchorPos(0)  # 해당 컨트롤의 좌표를 position 변수에 저장
            position = position.Item("List"), position.Item("Para"), position.Item("Pos")
            hwp.SetPos(*position)  # 해당 컨트롤 앞으로 캐럿(커서)을 옮김
            hwp.FindCtrl()  # 해당 컨트롤 선택
            Act = hwp.CreateAction("EquationModify")
            Set = Act.CreateSet()
            Pset = Set.CreateItemSet("EqEdit", "EqEdit")
            Act.GetDefault(Pset) 
            hwp.HAction.Run("Delete") ## 여기부터는 한글에서 수식을 다 텍스트로
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = Pset.Item("String");
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            eqn_dict[position] = Pset.Item("String")  # 좌표가 key이고, 수식문자열이 value인 사전 생성 ##
            break
        ctrl = nextctrl  # 다음 컨트롤 탐색
    hwp.Run("Cancel")  # 완료했으면 선택해제

    # for key, value in eqn_dict.items(): ##
    #     print(f"{key} : {value}")


def count_eqed(hwp): # 수식개수 세기 / cnt_eqed 라는 변수에 저장
    ctrl = hwp.HeadCtrl  # 첫 번째 컨트롤(HeadCtrl)부터 탐색 시작.
    global cnt_eqed
    cnt_eqed = 0
    while ctrl != None:  # 끝까지 탐색을 마치면 ctrl이 None을 리턴하므로.
        nextctrl = ctrl.Next  # 미리 nextctrl을 지정해 두고,
        if ctrl.CtrlID == "eqed":  # 현재 컨트롤이 "수식eqed"인 경우
            cnt_eqed = cnt_eqed+1
        ctrl = nextctrl  # 다음 컨트롤 탐색
    # print("수식개수는 " + str(cnt_eqed))



def preview_sol_hwp(): # 빠른정답만들기(각파일로) (새탭이 없을때 사용해야함.)
    hwp.MovePos(2)
    for i in range(1, cnt_mizu+1):
    # for i in range(1, 2):
        find_mizunum()
        hwp.HAction.Run("MoveRight")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
        hwp.HParameterSet.HInsertText.Text = " ";
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Select");
        hwp.HAction.Run("Copy");
        hwp.HAction.Run("Cancel");
        hwp.HAction.Run("FileNewTab")
        hwp.HAction.Run("Paste")
        hwp.HAction.Run("PasteOriginal")
        time.sleep(0.1)
        page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
        multicolumn_1() # 페이지 1단으로 
        time.sleep(0.1)
        hwp.MovePos(2)
        find_mizunum()
        hwp.HAction.Run("MoveSelRight")
        hwp.HAction.Run("DeleteBack")
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); #정답 없애기
        hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
        hwp.HParameterSet.HFindReplace.FindString = "\\b*\\[*\\(*정답\\]*\\)*\\b*"; # 정답 형태를 [정답] 으로 다 바꿔
        hwp.HParameterSet.HFindReplace.ReplaceString = " "; # 바꾸는 문자는 정규식 안먹음
        hwp.HParameterSet.HFindReplace.FindRegExp = 1;
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
        time.sleep(0.1)
        save_presol_png(i) # 정답 이미지(png)파일 저장할거야
        count_eqed(hwp) # 수식개수 세기 / cnt_eqed 라는 변수에 저장
        
        hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); # 엔터없애기
        hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
        hwp.HParameterSet.HFindReplace.FindString = "^n";
        hwp.HParameterSet.HFindReplace.ReplaceString = "";
        hwp.HParameterSet.HFindReplace.FindRegExp = 1;
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);

        hwp.MovePos(2)
        # hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
        # hwp.HParameterSet.HInsertText.Text = " ";
        # hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

        hwp.HAction.Run("SelectAll")
        real_text = hwp.GetTextFile("TEXT","saveblock"); # real_text란 변수에 전체 텍스트 넣기
        hwp.HAction.Run("Cancel");
        # hwp.MovePos(2)
        # hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
        # hwp.HParameterSet.HInsertText.Text = " ";
        # hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
        hwp.MovePos(2)
        hwp.HAction.Run("MoveSelNextWord");
        hwp.HAction.Run("Delete");
        
        
        re_real_text = re.sub('\s+|,', "", real_text) # 수식을 제외한 글자들 중 띄어쓰기/,를 제외한 글자
        
        cnt_circ = len(re.findall(r'[①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 개수
        cnt_circ_no = len(re.findall(r'[^①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 외 문자 개수
        
        # cnt_OX = len(re.findall(r'[OX]', re_real_text)) # text 중 대문자 O, X 의 개수
        # cnt_TF = len(re.findall(r'[TF]', re_real_text)) # text 중 대문자 T, F 의 개수
        # cnt_TF_ko = len(re.findall(r'참|거짓', re_real_text)) # text 중 참 거짓 의 개수
        cnt_OXTF = len(re.findall(r'[OX]|[TF]|참|거짓', re_real_text)) # text 중 대문자 O, X, T, F, 참, 거짓 의 개수
        cnt_OXTF_no_text = len(re.sub('O|X|T|F|참|거짓', "", re_real_text)) # re_real_text 문자열 중 O, X, T, F, 참, 거짓 를 제외한 문자
        
        global cnt_son
        cnt_son = len(re.findall(r'@', re_real_text)) # text 중 @의 개수
        if cnt_son !=0 : 
            cnt_son_lists.append(cnt_son)
        
        if cnt_son == 0: # 만약 @(새끼문제 기호) 개수가 0이면 -> 새끼문제는 아니네
            if "해설참조" not in re_real_text : # 해설참조 없을때 
                if cnt_eqed == 0: # 만약 수식 개수가 0이면
                    if len(re_real_text) == 0 : # 텍스트가 없으면 [빈해설파일] -> 수식도 없고 텍스트도 없으면 빈해설이지
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_presol_hwp(i, "[빈해설파일]")
                    elif cnt_circ_no == 0: # 원문자 외 문자가 없으면
                        if cnt_circ == 1 : # 원문자가 1개 이면 [객관식(선다-단일)]
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_presol_hwp(i, "[객관식(선다-단일)]")
                        elif cnt_circ > 1 : # 원문자가 2개 이상이면 
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_presol_hwp(i, "[객관식(선다-다중)]")
                        else : # 혹시 print 해보기
                            progress_head.delete(1.0, END)
                            progress_head.insert(END, f"원문자 외 문자가 없는데 원문자가 없어...이건 뭔상황일까???")
                            progress_head.update()
                    else: # 원문자 외 문자가 있으면
                        if cnt_OXTF_no_text == 0: # O, X, T, F, 참, 거짓 외 글자가 존재 X ([정답]/띄어쓰기/콤마제외)
                            if cnt_OXTF == 1: # O, X, T, F, 참, 거짓 -> 1개만 
                                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                                save_presol_hwp(i, "[객관식(OX-단일)]")
                            elif cnt_OXTF > 1: # O, X, T, F, 참, 거짓 -> 2개 이상 
                                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                                save_presol_hwp(i, "[객관식(OX-다중)]")
                            else : # 혹시 print 해보기
                                progress_head.delete(1.0, END)
                                progress_head.insert(END, f"O, X, T, F, 참, 거짓 외 글자가 없는데 저게 없어...이건 뭔상황일까???")
                                progress_head.update()
                        else : # O, X, T, F, 참, 거짓 외 글자가 존재 O
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_presol_hwp(i, "[주관식(자판)]")
                elif cnt_eqed == 1: # 만약 수식 개수가 1개 이면
                    equation_to_text_all(hwp)
                    if len(sum_reST) == 0: # +-숫자 외의 문자가 없으면 -> 정수
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_presol_hwp(i, "[주관식(정수)]")
                    else : # +-숫자 외의 문자가 있으면 -> 정수 X -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_presol_hwp(i, "[주관식(iink)]")
                else: # 만약 수식 개수가 2개 이상이면 -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                    equation_to_text_all(hwp)
                    Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                    save_presol_hwp(i, "[주관식(iink)]")
            else: # 해설참조 있으면 [증명문제]
                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                save_presol_hwp(i, "[증명문제]") 
        else: # 만약 @(새끼문제 기호) 개수가 0이면 -> 새끼문제이네
            save_presol_hwp(i, "[새끼문제]") # 일단 새끼문제는 저장한 후 다시 처

        time.sleep(0.2)
        # print(f"{cnt_mizu}번 중 {i}번 정답 추출 완료")
        progress_condi.insert(END, f"{cnt_mizu}번 중 {i}번 정답 추출 완료\n")
        progress_condi.see(END)
        progress_condi.update()
        hwp.XHwpDocuments.Item(1).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)

def Divide_files(file_fullname):
    
    count_common_problem(file_fullname) # common_problems -> 공통지문 리스트 / cnt_common_problems -> 공통지문 개수 / 한글 열기 전에 써야함
    
    
    hwp.Open(file_fullname,"HWP","forceopen:True")
    time.sleep(2)

    count_mizu() # 미주 개수 세기 -> 변수 cnt_mizu 에 저장
    # print(f"미주 개수는 {cnt_mizu}개 입니다.")
    progress_condi.insert(END, f"미주 개수는 {cnt_mizu}개 입니다.\n")
    progress_condi.see(END)
    progress_condi.update()
    
    global cnt_son_lists
    cnt_son_lists = []
    
    preview_sol_hwp() # 빠른정답만들기(각파일로) (새탭이 없을때 사용해야함.)

    onepageoneproblem() # 한쪽에 한문제 (첫페이지는 표지)
    # print("한쪽에 한문제 나누기 완료")
    progress_condi.insert(END, "한쪽에 한문제 나누기 완료\n")
    progress_condi.see(END)
    progress_condi.update()
    
    onepage_commonproble() # 공통지문 형식 찾아가서 ctrl+enter
    progress_condi.insert(END, f"공통지문({cnt_common_problems}개) 나누기 완료\n")
    progress_condi.see(END)
    progress_condi.update() 


    ### 문제 해설 나눠서 각 파일로 저장
    for i in range(1, cnt_mizu+1):
    # for i in range(1, 3): # 실험용
        tabdiv_pro_sol() # 원본 문제는 지우고 1탭에 문제만 2탭에 해설만 / 작동 후 해설 맨 처음
        save_sol_hml(i)
        save_sol_png(i)
        # equation_to_text_first(hwp) #첫수식 text화
        # save_sol_hwp(i)
        time.sleep(0.2)
        ## 해설 페이지가 1이면 이름만 바꾸고 2이상이면 병합할꺼야
        if solution_page==1:
            global sol_png_name
            sol_png_name = file_fullname.replace(".hwp", "") + " "+ str(i).zfill(3) + "번 [2해설] [2png]"
            os.rename(sol_png_name+"001"+".png", sol_png_name+".png")
        else:
            image_merge(sol_png_name)

        time.sleep(0.2)
        hwp.HAction.Run("WindowNextTab")
        hwp.HAction.Run("WindowNextTab")
        save_pro_hml(i)
        save_pro_png(i)
        time.sleep(0.2)
        hwp.HAction.Run("WindowNextTab")
        hwp.XHwpDocuments.Item(2).Close(isDirty=False)
        hwp.XHwpDocuments.Item(1).Close(isDirty=False)
        # print(f"{cnt_mizu}번 중 {i}번 나누기 완료")
        progress_condi.insert(END, f"{cnt_mizu}번 중 {i}번 나누기 완료\n")
        progress_condi.see(END)
        progress_condi.update()
    
    hwp.MovePos(2)
    if cnt_common_problems == 0: # 만약 공통지문 개수가 0이면 넘어가
        pass
    else: # 만약 공통지문 개수가 1 이상이면 개수만큼 돌려
        for i in range(1, cnt_common_problems+1): 
            tabdiv_commonpro_sol(i) # 공통지문 -> 1탭에 문제 / i는 몇번째 공통지문 / 작동 후 1탭에 맨처음
            save_commonpro_hml(start_num, last_num)
            save_commonpro_png(start_num, last_num)
            time.sleep(0.2)
            # hwp.HAction.Run("WindowNextTab")
            hwp.XHwpDocuments.Item(1).Close(isDirty=False)
            progress_condi.insert(END, f"공통지문 {cnt_common_problems}번 중 {i}번 나누기 완료\n")
            progress_condi.see(END)
            progress_condi.update()

    hwp.XHwpDocuments.Item(0).Close(isDirty=False)

    time.sleep(1)



def Divide_son_files(son_lists):
    for ind, son_list in enumerate(son_lists) :
        reind = ind+1
        global son_file_name
        son_file_name = os.path.join(dir, son_list)
        hwp.Open(son_file_name,"HWP","forceopen:True")
        time.sleep(2)
        hwp.MovePos(3)
        hwp.HAction.Run("BreakPage")
        hwp.MovePos(2)
        # print(cnt_son_lists[ind])
        for k in range(1, cnt_son_lists[ind]+1): # @ 앞 ctrl+enter
            find_sonproble() # 새끼문제 @ 찾아가기 (@앞으로)
            hwp.HAction.Run("BreakPage")
            hwp.HAction.Run("MoveRight")
        hwp.MovePos(2)
        for i in range(1, cnt_son_lists[ind]+1): # 새끼문제 개수만큼 돌려
            tabdiv_sonpro_sol() # 새끼문제 -> 1탭에 문제/ 작동 후 1탭에 맨처음
            save_sonsol_png(i) # 정답 이미지(png)파일 저장할거야
            count_eqed(hwp) # 수식개수 세기 / cnt_eqed 라는 변수에 저장
            
            hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); # 엔터없애기
            hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
            hwp.HParameterSet.HFindReplace.FindString = "^n";
            hwp.HParameterSet.HFindReplace.ReplaceString = "";
            hwp.HParameterSet.HFindReplace.FindRegExp = 1;
            hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
            
            hwp.MovePos(2)
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = " ";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);

            hwp.HAction.Run("SelectAll")
            real_text = hwp.GetTextFile("TEXT","saveblock"); # real_text란 변수에 전체 텍스트 넣기
            hwp.HAction.Run("Cancel");
            hwp.MovePos(2)

            hwp.HAction.Run("MoveSelNextWord");
            hwp.HAction.Run("Delete");
            
            
            re_real_text = re.sub('\s+|,', "", real_text) # 수식을 제외한 글자들 중 띄어쓰기/,를 제외한 글자
            
            cnt_circ = len(re.findall(r'[①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 개수
            cnt_circ_no = len(re.findall(r'[^①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 외 문자 개수
            cnt_OXTF = len(re.findall(r'[OX]|[TF]|참|거짓', re_real_text)) # text 중 대문자 O, X, T, F, 참, 거짓 의 개수
            cnt_OXTF_no_text = len(re.sub('O|X|T|F|참|거짓', "", re_real_text)) # re_real_text 문자열 중 O, X, T, F, 참, 거짓 를 제외한 문자
            
            if "해설참조" not in re_real_text : # 해설참조 없을때 
                if cnt_eqed == 0: # 만약 수식 개수가 0이면
                    if len(re_real_text) == 0 : # 텍스트가 없으면 [빈해설파일] -> 수식도 없고 텍스트도 없으면 빈해설이지
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_sonsol_hwp(i, "[빈해설파일]")
                    elif cnt_circ_no == 0: # 원문자 외 문자가 없으면
                        if cnt_circ == 1 : # 원문자가 1개 이면 [객관식(선다-단일)]
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_sonsol_hwp(i, "[객관식(선다-단일)]")
                        elif cnt_circ > 1 : # 원문자가 2개 이상이면 
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_sonsol_hwp(i, "[객관식(선다-다중)]")
                        else : # 혹시 print 해보기
                            progress_head.delete(1.0, END)
                            progress_head.insert(END, f"원문자 외 문자가 없는데 원문자가 없어...이건 뭔상황일까???")
                            progress_head.update()
                    else: # 원문자 외 문자가 있으면
                        if cnt_OXTF_no_text == 0: # O, X, T, F, 참, 거짓 외 글자가 존재 X ([정답]/띄어쓰기/콤마제외)
                            if cnt_OXTF == 1: # O, X, T, F, 참, 거짓 -> 1개만 
                                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                                save_sonsol_hwp(i, "[객관식(OX-단일)]")
                            elif cnt_OXTF > 1: # O, X, T, F, 참, 거짓 -> 2개 이상 
                                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                                save_sonsol_hwp(i, "[객관식(OX-다중)]")
                            else : # 혹시 print 해보기
                                progress_head.delete(1.0, END)
                                progress_head.insert(END, f"O, X, T, F, 참, 거짓 외 글자가 없는데 저게 없어...이건 뭔상황일까???")
                                progress_head.update()
                        else : # O, X, T, F, 참, 거짓 외 글자가 존재 O
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_sonsol_hwp(i, "[주관식(자판)]")
                elif cnt_eqed == 1: # 만약 수식 개수가 1개 이면
                    equation_to_text_all(hwp)
                    if len(sum_reST) == 0: # +-숫자 외의 문자가 없으면 -> 정수
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_sonsol_hwp(i, "[주관식(정수)]")
                    else : # +-숫자 외의 문자가 있으면 -> 정수 X -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_sonsol_hwp(i, "[주관식(iink)]")
                else: # 만약 수식 개수가 2개 이상이면 -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                    equation_to_text_all(hwp)
                    Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                    save_sonsol_hwp(i, "[주관식(iink)]")
            else: # 해설참조 있으면 [증명문제]
                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                save_sonsol_hwp(i, "[증명문제]")    
            time.sleep(0.2)
            hwp.XHwpDocuments.Item(1).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)
            time.sleep(0.2)
        hwp.XHwpDocuments.Item(0).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)
        time.sleep(1)
        progress_condi.insert(END, f"새끼문제 {sons}개 중 {reind}개 완료.\n")
        progress_condi.see(END)
        progress_condi.update()
        
        
        
        os.remove(os.path.join(dir, son_list))
        try:
            os.remove(os.path.join(dir, son_list.replace("[1hwp][새끼문제].hwp", "[2png].png")))
        except FileNotFoundError:
            pass
        # try:
        #     os.remove(os.path.join(dir, son_list.replace("[3정답][새끼문제].hwp", "[2png].png")))
        # except FileNotFoundError:
        #     pass
        try:
            os.remove(os.path.join(dir, son_list.replace("[3정답][새끼문제].hwp", "[3정답].png")))
        except FileNotFoundError:
            pass
        
        time.sleep(1)
    time.sleep(1)

def result_div(): #파일명 나누기 + 쪼개기
    file_fullnames = [os.path.split(x) for x in list_file.get(0, END)] # os.path.split -> 폴더, 파일명 나눠주는거 튜플로 나눠줌

    for idx, fullname in enumerate(file_fullnames) :
        progress_condi.delete(1.0, END)
        progress_condi.update()
        global dir
        dir = r'{0}'.format(fullname[0]).replace("/", "\\")
        name = fullname[1]
        name_only = re.sub(r".hwp", "", name)
        reidx = idx+1
        progress_head.delete(1.0, END)
        progress_head.insert(END, f"{reidx}번째 파일을 진행합니다.")
        progress_head.update()
        global file_fullname
        file_fullname = os.path.join(dir, name)
        # print(file_fullname)
        Divide_files(file_fullname) # 각 문제 파일 쪼개
        # progress_condi.delete(1.0, END)
        # progress_condi.update()
        
        global son_lists, sons
        son_lists = [i for i in os.listdir(dir) if "새끼문제" in i] # son_lists 리스트 안에 dir 안의 파일 중 "새끼문제"가 있는 파일 이름 저장
        sons = len(cnt_son_lists)
        
        progress_condi.insert(END, f"새끼문제 개수는 {sons}개 입니다.\n")
        progress_condi.see(END)
        progress_condi.update()

        # print(cnt_son_lists) # cnt_son_lists는 첫번째 새끼문제 부터 각 새끼문제의 개수 넣어둠.
        
        
        # if len(son_lists) != 0: # 새끼문제 있으면
        #     Divide_son_files(son_lists) #새끼문제 돌려
            
        time.sleep(1)
        lists = [i for i in os.listdir(dir) if i.startswith(f'{name_only}')] # name_only로 시작하는 파일 찾아라
        
        try: # 폴더 만들어 
            if not os.path.exists(os.path.join(dir, name_only)): os.makedirs(os.path.join(dir, name_only)) 
        except OSError:
            print("Error: Cannot create the directory {}".format(os.path.join(dir, name_only)))
        time.sleep(1)
        
        for list in lists: # 지금 작업하는 파일 이름(name_only)이 있는 파일 다 옮겨라 
            try:
                if os.path.join(dir, list) != file_fullname:
                    shutil.move(os.path.join(dir, list), os.path.join(dir, name_only))
            except PermissionError:
                print(f"{list}이 파일이 오류나네??")
                # os.remove(os.path.join(dir, list))

def Divide_one_pro(file_fullname):

    hwp.Open(file_fullname)
    time.sleep(2)
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 

    re_file_fullname = file_fullname.replace(".hml", "") +" (수정)"

    while os.path.isfile(re_file_fullname+".hml")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(1)
    
    os.rename(re_file_fullname+"001"+".png", re_file_fullname+".png")
    hwp.XHwpDocuments.Item(0).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)
    
    


def result_div_one_pro(): # 한 문제 hml 저장
    file_fullnames = [os.path.split(x) for x in list_file.get(0, END)] # os.path.split -> 폴더, 파일명 나눠주는거 튜플로 나눠줌

    for idx, fullname in enumerate(file_fullnames) :
        progress_condi.delete(1.0, END)
        progress_condi.update()
        global dir
        dir = r'{0}'.format(fullname[0]).replace("/", "\\")
        name = fullname[1]
        name_only = re.sub(r".hml", "", name)
        reidx = idx+1
        progress_head.delete(1.0, END)
        progress_head.insert(END, f"{reidx}번째 파일을 진행합니다.")
        progress_head.update()
        global file_fullname
        file_fullname = os.path.join(dir, name)
        Divide_one_pro(file_fullname) # 수정한 문제파일 조정
        time.sleep(1)
        lists = [i for i in os.listdir(dir) if i.startswith(f'{name_only}')] # name_only로 시작하는 파일 찾아라
        
        try: # 폴더 만들어 
            if not os.path.exists(os.path.join(dir, name_only)): os.makedirs(os.path.join(dir, name_only)) 
        except OSError:
            print("Error: Cannot create the directory {}".format(os.path.join(dir, name_only)))
        time.sleep(1)
        
        for list in lists: # 지금 작업하는 파일 이름(name_only)이 있는 파일 다 옮겨라 
            try:
                shutil.move(os.path.join(dir, list), os.path.join(dir, name_only))
            except PermissionError:
                print(f"{list}이 파일이 오류나네??")
                # os.remove(os.path.join(dir, list))        



def result_div_one_sol(): # 한 해설 hml 저장
    file_fullnames = [os.path.split(x) for x in list_file.get(0, END)] # os.path.split -> 폴더, 파일명 나눠주는거 튜플로 나눠줌

    for idx, fullname in enumerate(file_fullnames) :
        progress_condi.delete(1.0, END)
        progress_condi.update()
        global dir
        dir = r'{0}'.format(fullname[0]).replace("/", "\\")
        name = fullname[1]
        name_only = re.sub(r".hml", "", name)
        reidx = idx+1
        progress_head.delete(1.0, END)
        progress_head.insert(END, f"{reidx}번째 파일을 진행합니다.")
        progress_head.update()
        global file_fullname
        file_fullname = os.path.join(dir, name)
        Divide_one_sol(file_fullname) # 수정한 문제파일 조정
        
        progress_condi.insert(END, f"{reidx}번째 파일 - 나누기 작업 완료\n")
        progress_condi.see(END)
        progress_condi.update()
        
        global son_lists, sons
        son_lists = [i for i in os.listdir(dir) if "새끼문제" in i] # son_lists 리스트 안에 dir 안의 파일 중 "새끼문제"가 있는 파일 이름 저장
        sons = len(cnt_son_lists)
        
        progress_condi.insert(END, f"새끼문제 개수는 {sons}개 입니다.\n")
        progress_condi.see(END)
        progress_condi.update()

        # print(cnt_son_lists) # cnt_son_lists는 첫번째 새끼문제 부터 각 새끼문제의 개수 넣어둠.
        
        
        if len(son_lists) != 0: # 새끼문제 있으면
            Divide_son_files(son_lists) #새끼문제 돌려
            
        time.sleep(1)
        lists = [i for i in os.listdir(dir) if i.startswith(f'{name_only}')] # name_only로 시작하는 파일 찾아라
        
        try: # 폴더 만들어 
            if not os.path.exists(os.path.join(dir, name_only)): os.makedirs(os.path.join(dir, name_only)) 
        except OSError:
            print("Error: Cannot create the directory {}".format(os.path.join(dir, name_only)))
        time.sleep(1)
        
        for list in lists: # 지금 작업하는 파일 이름(name_only)이 있는 파일 다 옮겨라 
            try:
                shutil.move(os.path.join(dir, list), os.path.join(dir, name_only))
            except PermissionError:
                print(f"{list}이 파일이 오류나네??")
                # os.remove(os.path.join(dir, list))
        

def Divide_one_sol(file_fullname):
    
    hwp.Open(file_fullname)
    time.sleep(2)
    page_size_set() # 페이지 크기 정보 106.5 , 1100 / 여백은 다 0
    multicolumn_1() # 페이지 1단으로 

    hwp.MovePos(3) # 해설 맨 마지막 끝으로가서
    
    global solution_page
    solution_page = hwp.KeyIndicator()[3] # 현재 커서의 페이지 번호 저장

    global cnt_son_lists
    cnt_son_lists = []

    global re_file_fullname
    re_file_fullname = file_fullname.replace(".hml", "") +" (수정)"
    
    hwp.MovePos(2)
    
    hwp.HAction.Run("Select");
    hwp.HAction.Run("Select");
    hwp.HAction.Run("Select");
    text = hwp.GetTextFile("TEXT","saveblock");
    hwp.HAction.Run("Cancel");
    
    cnt_rhfqoddl = len(re.findall(r'@', text)) # text 중 @의 개수

    hwp.MovePos(2)
    
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); 
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "@"; # @ 다 지워
    hwp.HParameterSet.HFindReplace.ReplaceString = ""; 
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);

    while os.path.isfile(re_file_fullname+".hml")==False: # 해설 hml 저장해
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname+".hml"
        hwp.HParameterSet.HFileOpenSave.Format = "HWPML2X"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)

    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet) # 해설 png 저장해
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname+".png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    time.sleep(1)
    
    if solution_page==1:
        os.rename(re_file_fullname+"001"+".png", re_file_fullname+".png")
    else:
        image_merge(re_file_fullname)
    
    time.sleep(0.2)
    
    hwp.MovePos(2)
    
    hwp.HAction.Run("Select");
    hwp.HAction.Run("Select");
    hwp.HAction.Run("Select");
    hwp.HAction.Run("Cancel");
    hwp.HAction.Run("MoveSelTopLevelEnd");
    hwp.HAction.Run("DeleteBack");
    
    Allreplace_rhfqoddl() # 괄호숫자 앞에 골뱅이 붙여 바꾸기

    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); #정답 없애기
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "\\b*\\[*\\(*정답\\]*\\)*\\b*"; # 정답 형태를 [정답] 으로 다 바꿔
    hwp.HParameterSet.HFindReplace.ReplaceString = " "; # 바꾸는 문자는 정규식 안먹음
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);
    time.sleep(0.1)
    
    hwp.HAction.GetDefault("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HParameterSet.HPrint.PrinterName = "그림으로 저장하기"
    hwp.HParameterSet.HPrint.PrintAutoFootNote = 0
    hwp.HParameterSet.HPrint.PrintAutoHeadNote = 0
    hwp.HParameterSet.HPrint.PrintMethod = hwp.PrintType("Nomal")
    hwp.HParameterSet.HPrint.Collate = 1
    hwp.HParameterSet.HPrint.UserOrder = 0
    hwp.HParameterSet.HPrint.PrintToFile = 0
    hwp.HParameterSet.HPrint.NumCopy = 1
    hwp.HParameterSet.HPrint.OverlapSize = 0
    hwp.HParameterSet.HPrint.PrintCropMark = 0
    hwp.HParameterSet.HPrint.BinderHoleType = 0
    hwp.HParameterSet.HPrint.ZoomX = 100
    hwp.HParameterSet.HPrint.UsingPagenum = 1
    hwp.HParameterSet.HPrint.ReverseOrder = 0
    hwp.HParameterSet.HPrint.Pause = 0
    hwp.HParameterSet.HPrint.PrintImage = 1
    hwp.HParameterSet.HPrint.PrintDrawObj = 1
    hwp.HParameterSet.HPrint.PrintClickHere = 0
    hwp.HParameterSet.HPrint.EvenOddPageType = 0
    hwp.HParameterSet.HPrint.PrintWithoutBlank = 0
    hwp.HParameterSet.HPrint.PrintAutoFootnoteLtext = "^f"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteCtext = "^t"
    hwp.HParameterSet.HPrint.PrintAutoFootnoteRtext = "^P쪽 중 ^p쪽"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteLtext = "^c"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteCtext = "^n"
    hwp.HParameterSet.HPrint.PrintAutoHeadnoteRtext = "^p"
    hwp.HParameterSet.HPrint.ZoomY = 100
    hwp.HParameterSet.HPrint.PrintFormObj = 1
    hwp.HParameterSet.HPrint.PrintMarkPen = 0
    hwp.HParameterSet.HPrint.PrintBarcode = 1
    hwp.HParameterSet.HPrint.Device = hwp.PrintDevice("Image")
    hwp.HParameterSet.HPrint.PrintPronounce = 0

    hwp.HAction.Execute("FileSaveAsImage", hwp.HParameterSet.HPrint.HSet)
    hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)

    hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname+"[3정답].png"
    hwp.HParameterSet.HFileOpenSave.Format = "PNG"
    hwp.HParameterSet.HFileOpenSave.Attributes = 0
    hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    
    os.rename(re_file_fullname+"[3정답]001"+".png", re_file_fullname+"[3정답].png")
    
    count_eqed(hwp) # 수식개수 세기 / cnt_eqed 라는 변수에 저장
    
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet); # 엔터없애기
    hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("AllDoc");
    hwp.HParameterSet.HFindReplace.FindString = "^n";
    hwp.HParameterSet.HFindReplace.ReplaceString = "";
    hwp.HParameterSet.HFindReplace.FindRegExp = 1;
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet);

    hwp.HAction.Run("SelectAll")
    real_text = hwp.GetTextFile("TEXT","saveblock"); # real_text란 변수에 전체 텍스트 넣기
    hwp.HAction.Run("Cancel");
    hwp.MovePos(2)
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
    hwp.HParameterSet.HInsertText.Text = " ";
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
    hwp.MovePos(2)
    hwp.HAction.Run("MoveSelNextWord");
    hwp.HAction.Run("Delete");
    
    
    re_real_text = re.sub('\s+|,', "", real_text) # 수식을 제외한 글자들 중 띄어쓰기/,를 제외한 글자
    
    cnt_circ = len(re.findall(r'[①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 개수
    cnt_circ_no = len(re.findall(r'[^①②③④⑤⑥⑦⑧⑨]', re_real_text)) # text 중 원문자 외 문자 개수
    
    # cnt_OX = len(re.findall(r'[OX]', re_real_text)) # text 중 대문자 O, X 의 개수
    # cnt_TF = len(re.findall(r'[TF]', re_real_text)) # text 중 대문자 T, F 의 개수
    # cnt_TF_ko = len(re.findall(r'참|거짓', re_real_text)) # text 중 참 거짓 의 개수
    cnt_OXTF = len(re.findall(r'[OX]|[TF]|참|거짓', re_real_text)) # text 중 대문자 O, X, T, F, 참, 거짓 의 개수
    cnt_OXTF_no_text = len(re.sub('O|X|T|F|참|거짓', "", re_real_text)) # re_real_text 문자열 중 O, X, T, F, 참, 거짓 를 제외한 문자
    
    global cnt_son
    cnt_son = cnt_rhfqoddl # text 중 @의 개수
    if cnt_son !=0 : 
        cnt_son_lists.append(cnt_son)
    
    if cnt_son == 0: # 만약 @(새끼문제 기호) 개수가 0이면 -> 새끼문제는 아니네
        if "해설참조" not in re_real_text : # 해설참조 없을때 
            if cnt_eqed == 0: # 만약 수식 개수가 0이면
                if len(re_real_text) == 0 : # 텍스트가 없으면 [빈해설파일] -> 수식도 없고 텍스트도 없으면 빈해설이지
                    Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                    save_onesol_hwp("[빈해설파일]")
                elif cnt_circ_no == 0: # 원문자 외 문자가 없으면
                    if cnt_circ == 1 : # 원문자가 1개 이면 [객관식(선다-단일)]
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_onesol_hwp("[객관식(선다-단일)]")
                    elif cnt_circ > 1 : # 원문자가 2개 이상이면 
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_onesol_hwp("[객관식(선다-다중)]")
                    else : # 혹시 print 해보기
                        progress_head.delete(1.0, END)
                        progress_head.insert(END, f"원문자 외 문자가 없는데 원문자가 없어...이건 뭔상황일까???")
                        progress_head.update()
                else: # 원문자 외 문자가 있으면
                    if cnt_OXTF_no_text == 0: # O, X, T, F, 참, 거짓 외 글자가 존재 X ([정답]/띄어쓰기/콤마제외)
                        if cnt_OXTF == 1: # O, X, T, F, 참, 거짓 -> 1개만 
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_onesol_hwp("[객관식(OX-단일)]")
                        elif cnt_OXTF > 1: # O, X, T, F, 참, 거짓 -> 2개 이상 
                            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                            save_onesol_hwp("[객관식(OX-다중)]")
                        else : # 혹시 print 해보기
                            progress_head.delete(1.0, END)
                            progress_head.insert(END, f"O, X, T, F, 참, 거짓 외 글자가 없는데 저게 없어...이건 뭔상황일까???")
                            progress_head.update()
                    else : # O, X, T, F, 참, 거짓 외 글자가 존재 O
                        Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                        save_onesol_hwp("[주관식(자판)]")
            elif cnt_eqed == 1: # 만약 수식 개수가 1개 이면
                equation_to_text_all(hwp)
                if len(sum_reST) == 0: # +-숫자 외의 문자가 없으면 -> 정수
                    Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                    save_onesol_hwp("[주관식(정수)]")
                else : # +-숫자 외의 문자가 있으면 -> 정수 X -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                    Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                    save_onesol_hwp("[주관식(iink)]")
            else: # 만약 수식 개수가 2개 이상이면 -> 단 <보기> 정답 a,b,c 이런 경우는 사람이 고쳐야함
                equation_to_text_all(hwp)
                Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
                save_onesol_hwp("[주관식(iink)]")
        else: # 해설참조 있으면 [증명문제]
            Allreplace_circ() # 원숫자 다 괄호숫자로 바꾸기
            save_onesol_hwp("[증명문제]") 
    else: # 만약 @(새끼문제 기호) 개수가 0이면 -> 새끼문제이네
        save_onesol_hwp("[새끼문제]") # 일단 새끼문제는 저장한 후 다시 처

    time.sleep(0.2)

    hwp.XHwpDocuments.Item(0).Close(isDirty=False) # 새창 닫아(저장할지 물어보지 말고)

    
def save_onesol_hwp(type):  # 수정해설 정답 hwp파일 저장할거야 num = 문제번호, type = 정답종류
    # re_file_fullname + str(type) # 확장자 없는 이름
    while os.path.isfile(re_file_fullname +"[3정답]"+ str(type)+".hwp")==False:
        hwp.HAction.GetDefault("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = re_file_fullname +"[3정답]"+ str(type)+".hwp"
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HParameterSet.HFileOpenSave.Attributes = 0
        hwp.HAction.Execute("FileSave_S", hwp.HParameterSet.HFileOpenSave.HSet)
    time.sleep(0.2)




# 파일 추가
def add_file():
    files = filedialog.askopenfilenames(title="분할할 hwp파일을 선택하세요", \
        filetypes=(("HWP 파일", "*.hwp"), ("HML 파일", "*.hml"), ("모든 파일", "*.*")), \
        initialdir=os.path.expanduser(r"~\Desktop"))
        # 최초에 사용자가 지정한 경로를 보여줌
    
    # 사용자가 선택한 파일 목록
    for file in files:
        list_file.insert(END, file)
        list_file.see(END)

# 선택 삭제
def del_file():
    #print(list_file.curselection())
    for index in reversed(list_file.curselection()):
        list_file.delete(index)


# 전체 나누기(시작)
def start_all():
    # 파일 목록 확인
    if list_file.size() == 0:
        msgbox.showwarning("경고", "분할할 hwp파일을 선택하세요")
        return
    else:
        result_div()
        msgbox.showinfo("알림", "작업이 완료되었습니다.")
        progress_head.delete(1.0, END)
        

# 수정해설저장
def start_one_sol():
    if list_file.size() == 0:
        msgbox.showwarning("경고", "서버에서 다운받아 수정한 해설 hml파일을 선택하세요")
        return
    else:
        result_div_one_sol()
        msgbox.showinfo("알림", "작업이 완료되었습니다.")
        progress_head.delete(1.0, END)


# 수정문제저장
def start_one_pro():
    if list_file.size() == 0:
        msgbox.showwarning("경고", "서버에서 다운받아 수정한 문제 hml파일을 선택하세요")
        return
    else:
        result_div_one_pro()
        msgbox.showinfo("알림", "작업이 완료되었습니다.")
        progress_head.delete(1.0, END)
        







        


        



# 파일 프레임 (파일 추가, 선택 삭제)
file_frame = Frame(root)
file_frame.pack(fill="x", padx=5, pady=5) # 간격 띄우기

btn_add_file = Button(file_frame, padx=5, pady=5, width=12, text="파일추가", command=add_file)
btn_add_file.pack(side="left")

btn_del_file = Button(file_frame, padx=5, pady=5, width=12, text="선택삭제", command=del_file)
btn_del_file.pack(side="right")

# 리스트 프레임
list_frame = LabelFrame(root, width=100, text="선택된 파일")
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=10, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)



# 진행 상태 확인
progress_frame = LabelFrame(root, width=100, height=5, text="진행상황")
progress_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(progress_frame)
scrollbar.pack(side="right", fill="y")

progress_head = Text(progress_frame, width=100, height=1)
progress_head.pack()

progress_condi = Text(progress_frame, width=100, height=30, yscrollcommand=scrollbar.set)
progress_condi.pack(side="left", fill="both", expand=True)
scrollbar.config(command=progress_condi.yview)



# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start1 = Button(frame_run, padx=5, pady=5, text="전체나누기", width=12, command=start_all)
btn_start1.pack(side="left", padx=5, pady=5)

btn_start2 = Button(frame_run, padx=5, pady=5, text="수정해설저장", width=12, command=start_one_sol)
btn_start2.pack(side="left", padx=5, pady=5)

btn_start3 = Button(frame_run, padx=5, pady=5, text="수정문제저장", width=12, command=start_one_pro)
btn_start3.pack(side="left", padx=5, pady=5)


# root.geometry('1000x1100')
# root.resizable(True, True)
root.resizable(False, False)
root.mainloop()



hwp.Quit()