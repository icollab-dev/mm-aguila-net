Attribute VB_Name = "Module1"
'----------------------------------------1-----------------------
''Windows API
'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpReturnedString As String, ByVal lpFileName As String) As Long
'----------------------------------------2-----------------------
Public Declare Function dCOMPILE Lib ".\dll\CALCOF.dll" (ByRef lui As Long) As Long
Public Declare Function dCALCULATE Lib ".\dll\CALCOF.dll" (ByRef ID As Long) As Long
Public Declare Function dQGET Lib ".\dll\CALCOF.dll" (ByRef ID As Long) As Long
Public Declare Function dQSET Lib ".\dll\CALCOF.dll" (ByRef ID As Long, ByRef Q As Long) As Long
Public Declare Function dVGET Lib ".\dll\CALCOF.dll" (ByRef ID As Long) As Double
Public Declare Function dVSET Lib ".\dll\CALCOF.dll" (ByRef ID As Long, ByRef VAL As Double) As Long
Public Declare Function dVGETT Lib ".\dll\CALCOF.dll" (ByRef buf As String) As Double
Public Declare Function dVSETT Lib ".\dll\CALCOF.dll" (ByRef buf As String, ByRef VAL As Double) As Long
Public Declare Function dCGETT Lib ".\dll\CALCOF.dll" (ByRef buf As String, ByRef comment As String) As Long
Public Declare Function dNGET Lib ".\dll\CALCOF.dll" (ByRef ID As Long) As Long
Public Declare Function dLUPCLS Lib ".\dll\CALCOF.dll" (ByRef ID As Long) As Long
Public Declare Function dCGET Lib ".\dll\CALCOF.dll" (ByRef ID As Long, ByRef buf As String) As Long
Public Declare Function dTGET Lib ".\dll\CALCOF.dll" (ByRef ID As Long, ByRef buf As String) As Long

Public OutFile1 As String
Public OutFile2 As String
Public Dl1Tag(4) As String
Public Dl1Frm(4) As Integer
Public Dl2Tag As String
Public Dl2Frm As Integer
Public LANGINI As String

Public Function frmget(ByVal ii As Integer) As String
    Dim j As Integer
    frmget = "#####0"
    If ii > 0 Then
      frmget = frmget & "."
      For j = 1 To ii
        frmget = frmget & "0"
      Next j
    End If
End Function

Public Sub profileset()
    Dim fPath   As String
    Dim Section As String
    Dim Gbuf    As String
    Dim frm As Integer
    Dim InpTag As String
    Dim i As Integer
    Dim IKey As String
    fPath = App.Path + "\CTUNG.ini"
'INPUT TAG Setting
    Section = "INPUT"
    For i = 0 To 7
      Gbuf = Form1.Text1(i).tag
      IKey = "TAG" & Format(i, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
      Gbuf = Str(Form1.Text1(i).DataField)
      IKey = "FORM" & Format(i, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    Next i
'DISPLAY TAG Setting
    Section = "DISPLAY"
    For i = 0 To 19
      Gbuf = Form1.Label4(i).tag
      IKey = "TAG" & Format(i, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
      Gbuf = Str(Form1.Label4(i).DataField)
      IKey = "FORM" & Format(i, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    Next i
'DIALOG1 TAG Setting
    Section = "DIALOG1"
    Gbuf = Dialog1.Text1.tag
    IKey = "TAG" & Format(0, "0")
    sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    Gbuf = Str(Dialog1.Text1.DataField)
    IKey = "FORM" & Format(0, "0")
    sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    For i = 0 To 2
      Gbuf = Dialog1.Label1(i).tag
      IKey = "TAG" & Format(i + 1, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
      Gbuf = Str(Dialog1.Label1(i).DataField)
      IKey = "FORM" & Format(i + 1, "0")
      sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    Next i
'DIALOG2 TAG Setting
    Section = "DIALOG2"
    Gbuf = Dialog2.Text1.tag
    IKey = "TAG" & Format(0, "0")
    sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)
    Gbuf = Str(Dialog2.Text1.DataField)
    IKey = "FORM" & Format(0, "0")
    sts = WritePrivateProfileString(Section, IKey, Gbuf, fPath)

End Sub

Public Sub ProfileGet()
    Dim fPath   As String
    Dim Section As String
    Dim Gbuf    As String
    Dim DefRtn As String
    Dim GIntMAX As Integer
    Dim frm As Integer
    Dim InpTag As String
    Dim i As Integer
    Dim IKey As String
    fPath = App.Path + "\CTUNG.ini"
'OUTPUT FILENAME Setting
    Section = "FILE"
      DefRtn = "./data/CALOUT.csv"
      GIntMAX = 50
      Gbuf = Space$(GIntMAX)
      IKey = "CALOUT"
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      OutFile1 = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
      DefRtn = "./data/VALUE.DAT"
      GIntMAX = 50
      Gbuf = Space$(GIntMAX)
      IKey = "VALOUT"
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      OutFile2 = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)

'INPUT TAG Setting
    Section = "INPUT"
    For i = 0 To 7
      DefRtn = Form1.Text1(i).tag
      GIntMAX = 12
      Gbuf = Space$(GIntMAX)
      IKey = "TAG" & Format(i, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Text1(i).tag = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
      GIntMAX = 4
      Gbuf = Space$(GIntMAX)
      IKey = "FORM" & Format(i, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Text1(i).DataField = VAL(Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1))
    Next i
'DISPLAY TAG Setting
    Section = "DISPLAY"
    For i = 0 To 19
      DefRtn = Form1.Label4(i).tag
      GIntMAX = 12
      Gbuf = Space$(GIntMAX)
      IKey = "TAG" & Format(i, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Label4(i).tag = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
      DefRtn = Str(Form1.Label4(i).DataField)
      GIntMAX = 4
      Gbuf = Space$(GIntMAX)
      IKey = "FORM" & Format(i, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Label4(i).DataField = VAL(Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1))
    Next i
'DIALOG1 TAG Setting
    Dialog1.Show
    Section = "DIALOG1"
    DefRtn = Dialog1.Text1.tag
    GIntMAX = 12
    Gbuf = Space$(GIntMAX)
    IKey = "TAG" & Format(0, "0")
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Dl1Tag(0) = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    DefRtn = Str(Dialog1.Text1.DataField)
    GIntMAX = 4
    Gbuf = Space$(GIntMAX)
    IKey = "FORM" & Format(0, "0")
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Dl1Frm(0) = VAL(Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1))
    For i = 0 To 2
      DefRtn = Dialog1.Label1(i).tag
      GIntMAX = 12
      Gbuf = Space$(GIntMAX)
      IKey = "TAG" & Format(i + 1, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Dl1Tag(i + 1) = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
      DefRtn = Str(Dialog1.Label1(i).DataField)
      GIntMAX = 4
      Gbuf = Space$(GIntMAX)
      IKey = "FORM" & Format(i + 1, "0")
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Dl1Frm(i + 1) = VAL(Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1))
    Next i
    Unload Dialog1
'DIALOG2 TAG Setting
    Dialog2.Show
    Section = "DIALOG2"
    DefRtn = Dialog2.Text1.tag
    GIntMAX = 12
    Gbuf = Space$(GIntMAX)
    IKey = "TAG" & Format(0, "0")
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Dl2Tag = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    DefRtn = Str(Dialog2.Text1.DataField)
    GIntMAX = 4
    Gbuf = Space$(GIntMAX)
    IKey = "FORM" & Format(0, "0")
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Dl2Frm = VAL(Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1))
    Unload Dialog2
'LANGUAGE SELCTION
    Section = "FILE"
    DefRtn = "SPANISH"
    GIntMAX = 12
    Gbuf = Space$(GIntMAX)
    IKey = "LANGUAGE"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    LANGINI = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FORM1 LANGUAGE SETTING
    fPath = App.Path + "\" & LANGINI & ".ini"
'FILE/READ
    Section = "MENU"
    DefRtn = "FILE"
    GIntMAX = 12
    Gbuf = Space$(GIntMAX)
    IKey = "FILE"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.FILE.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    DefRtn = "READ"
    GIntMAX = 12
    Gbuf = Space$(GIntMAX)
    IKey = "READ"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.READ.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FORM1
    Section = "FORM1"
    DefRtn = "CTUNG"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "TITLE"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'LABEL1-
    For i = 0 To 26
      If i = 16 Then GoTo NEXT1
      DefRtn = "LABEL1" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "LABEL1" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Label1(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
NEXT1:
    Next i
'LABEL2-
    For i = 0 To 14
      DefRtn = "LABEL2" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "LABEL2" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Label2(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    Next i
'LABEL3
    DefRtn = "LABEL3"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "LABEL3"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Label3.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FRAME2
    DefRtn = "FRAME2"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "FRAME2"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Frame2.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FRAME3
    DefRtn = "FRAME3"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "FRAME3"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Frame3.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FRAME4-
    For i = 0 To 1
      DefRtn = "FRAME4" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "FRAME4" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Frame4(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    Next i
'FRAME5
    DefRtn = "FRAME5"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "FRAME5"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Frame5.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'FRAME6
    DefRtn = "FRAME6"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "FRAME6"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Form1.Frame6.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'COMMAND1-
    For i = 0 To 1
      DefRtn = "COMMAND1" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "COMMAND1" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Command1(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    Next i
'COMMAND2-
    For i = 0 To 1
      DefRtn = "COMMAND2" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "COMMAND2" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Form1.Command2(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    Next i
End Sub

