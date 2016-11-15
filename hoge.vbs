Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"


Dim strInputData 

' 入力ダイアログを出力 
strInputData = InputBox("回線番号を入力してください " & vbNewLine & " plz input cct number","Bridge（仮）") 

' 入力ダイアログで入力された値をメッセージボックスで出力 
MsgBox "あなたの好きな食べ物は" & strInputData & "ですね！"

MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",0,0,strInputData
MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",100,100,strInputData

Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.AppActivate "Explorer"
 


Function MakeWindow(URL,top,left,strInputData)

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True

objIE.FullScreen = False
objIE.Top = top
objIE.Left = left
objIE.Width = 1280
objIE.Height = 964

objIE.Toolbar = True
objIE.MenuBar = True
objIE.AddressBar = True
objIE.StatusBar = True

objIE.Navigate2 ""+URL, navOpenInNewWindow
Do Until objIE.Busy = False
WScript.sleep(250)
Loop
objIE.Document.Forms(0).BTX0010.value = strInputData
objIE.Document.Forms(0).S.checked = False
objIE.Document.Forms(0).BPW0020.value = strInputData
'objIE.Document.Forms(0).forward_BSM2010.Click

End Function