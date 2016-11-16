Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"


Dim strInputData 

' 入力ダイアログを出力 
strInputData = InputBox("回線番号を入力してください " & vbNewLine & " plz input cct number","Bridge（仮）") 

' 入力ダイアログで入力された値をメッセージボックスで出力 
MsgBox "あなたの好きな食べ物は" & strInputData & "ですね！"

MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",0,0,strInputData
'MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",100,100,strInputData


'Set oApp = CreateObject("PowerPoint.Application")
'oApp.Presentations.Open("C:\Users\Keita Yamamoto\Desktop\プレゼンテーション.ppt")

WindowSearchLink "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398","すべての商品・サービス一覧"

Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.AppActivate "Explorer"
 


'-------------------------------------------------------------------------------
Function MakeWindow(URL,top,left,strInputData)

Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True

objIE.FullScreen = False
objIE.Top = top
objIE.Left = left
objIE.Width = 1280
objIE.Height = 964

objIE.Toolbar = False
objIE.MenuBar = False
objIE.AddressBar = False
objIE.StatusBar = False

objIE.Navigate2 ""+URL, navOpenInNewWindow
Do Until objIE.Busy = False
WScript.sleep(250)
Loop
objIE.Document.Forms(0).BTX0010.value = strInputData
objIE.Document.Forms(0).S.checked = False
objIE.Document.Forms(0).BPW0020.value = strInputData
'objIE.Document.Forms(0).forward_BSM2010.Click

MakeWindow = objIE
End Function
'-------------------------------------------------------------------------------
Function WindowSearchLink(URL,LinkString)
'複数立ち上がったIEから IPAT　投票メニューを見つける
   Set  objIE = CreateObject("InternetExplorer.Application")
   objIE.Visible = True
   objIE.Navigate2 URL, navOpenInNewWindow
   Do Until objIE.Busy = False
   WScript.sleep(250)
   Loop
'↑上で見つけたIPAT　投票メニューから 入出金メニュー を 押す

    'Aのタグを集める .getElementsByTagName("A")を使用
    Set objA = objIE.Document.getElementsByTagName("A")

    'ループで頭から表示してみる
    For n = 0 To objA.Length - 1
        '※.InnerHTMLじゃなくて、.OuterHTMLでAの全体を見る
        '入出金メニューのリンクを探す、ソースの文字を探す
        If InStr(objA(n).OuterHTML, ""+LinkString) > 0 Then
            objA(n).Click  'クリックする
            Exit For  'ループを抜ける
        End If
    Next

    Set objA = Nothing  'オブジェクト変数解放
End Function



'-------------------------------------------------------------------------------








