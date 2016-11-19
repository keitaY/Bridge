Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"


Dim strInputData 

' 入力ダイアログを出力 
strInputData = InputBox("例の番号を入力してください " & vbNewLine & " plz input cct number","Bridge（仮）") 

' 入力ダイアログで入力された値をメッセージボックスで出力 
If (strInputData <> "" ) Then

openMS strInputData
openPI strInputData
'openGN strInputData

'Set oApp = CreateObject("PowerPoint.Application")
'oApp.Presentations.Open("C:\Users\Keita Yamamoto\Desktop\プレゼンテーション.ppt")

End If

Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.AppActivate "Explorer"
 
'-------------------------------------------------------------------------------
Function openMS(strInputData)

Set objIE = openWindow("https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",0,0,1000,1600)

objIE.Document.Forms(0).BTX0010.value = strInputData
objIE.Document.Forms(0).S.checked = False
objIE.Document.Forms(0).BPW0020.value = strInputData
'objIE.Document.Forms(0).forward_BSM2010.Click
clickLink objIE,"すべての商品・サービス一","A"

openMS = objIE
End Function
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Function openPI(strInputData)

Set objIE = openWindow("http://yume.hacca.jp/koiki/form/button-link.htm",50,50,1000,1600)

'objIE.Document.Forms(0).forward_BSM2010.Click
clickLink objIE,"ボタンでリンク３","Button"

openPI = objIE
End Function
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Function openGN(strInputData)

Set objIE = openWindow("https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",100,100,1000,1600)

objIE.Document.Forms(0).BTX0010.value = strInputData
objIE.Document.Forms(0).S.checked = False
objIE.Document.Forms(0).BPW0020.value = strInputData
'objIE.Document.Forms(0).forward_BSM2010.Click
'clickLink objIE,"すべての商品・サービス一"

openGN = objIE
End Function
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Function openWindow(URL,top,left,height,width)
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.FullScreen = False
objIE.Top = top
objIE.Left = left
objIE.Width = width
objIE.Height = height
objIE.Toolbar = false
objIE.MenuBar = false
objIE.AddressBar = false
objIE.StatusBar = false

objIE.Navigate2 ""+URL, navOpenInNewWindow

Do Until objIE.Busy = False
WScript.sleep(250)
Loop

 Set openWindow = objIE
End Function
'-------------------------------------------------------------------------------
Function clickLink(objIE,linkString,tagType)
   Do Until objIE.Busy = False
   WScript.sleep(250)
   Loop
    Set objA = objIE.Document.getElementsByTagName(tagType)
    For n = 0 To objA.Length - 1
        If InStr(objA(n).OuterHTML, ""+linkString) > 0 Then
            objA(n).Click  
            Exit For  
        End If
    Next
    Set objA = Nothing  
	set clickLink = objIE
End Function
'-------------------------------------------------------------------------------