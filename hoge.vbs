Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"


Dim strInputData 

' ���̓_�C�A���O���o�� 
strInputData = InputBox("����ԍ�����͂��Ă������� " & vbNewLine & " plz input cct number","Bridge�i���j") 

' ���̓_�C�A���O�œ��͂��ꂽ�l�����b�Z�[�W�{�b�N�X�ŏo�� 
MsgBox "���Ȃ��̍D���ȐH�ו���" & strInputData & "�ł��ˁI"

MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",0,0,strInputData
'MakeWindow "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398",100,100,strInputData


'Set oApp = CreateObject("PowerPoint.Application")
'oApp.Presentations.Open("C:\Users\Keita Yamamoto\Desktop\�v���[���e�[�V����.ppt")

WindowSearchLink "https://www.ib2.aozorabank.co.jp/ib/index.do?PT=BS&CCT0080=0398","���ׂĂ̏��i�E�T�[�r�X�ꗗ"

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
'���������オ����IE���� IPAT�@���[���j���[��������
   Set  objIE = CreateObject("InternetExplorer.Application")
   objIE.Visible = True
   objIE.Navigate2 URL, navOpenInNewWindow
   Do Until objIE.Busy = False
   WScript.sleep(250)
   Loop
'����Ō�����IPAT�@���[���j���[���� ���o�����j���[ �� ����

    'A�̃^�O���W�߂� .getElementsByTagName("A")���g�p
    Set objA = objIE.Document.getElementsByTagName("A")

    '���[�v�œ�����\�����Ă݂�
    For n = 0 To objA.Length - 1
        '��.InnerHTML����Ȃ��āA.OuterHTML��A�̑S�̂�����
        '���o�����j���[�̃����N��T���A�\�[�X�̕�����T��
        If InStr(objA(n).OuterHTML, ""+LinkString) > 0 Then
            objA(n).Click  '�N���b�N����
            Exit For  '���[�v�𔲂���
        End If
    Next

    Set objA = Nothing  '�I�u�W�F�N�g�ϐ����
End Function



'-------------------------------------------------------------------------------








