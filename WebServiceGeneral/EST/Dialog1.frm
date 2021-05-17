VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Capacidad Neta Demonstrada(KC)"
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      DataField       =   "2"
      Height          =   270
      Left            =   1200
      TabIndex        =   2
      Tag             =   "KC100"
      Text            =   "495.00"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "CANCELAR"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "MW"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "MW"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "MW"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "MW"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "25%"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "50%"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "75%"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "100%"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      DataField       =   "2"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Tag             =   "KC75"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      DataField       =   "2"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Tag             =   "KC75"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BorderStyle     =   1  '実線
      DataField       =   "2"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Tag             =   "KC75"
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim VV As Double
Dim ist As Long
Dim frm As String

Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Dim fPath   As String
    Dim Section As String
    Dim Gbuf    As String
    Dim DefRtn As String
    Dim GIntMAX As Integer
    Dim i As Integer
    Dim sts As Long
    Dim IKey As String
    fPath = App.Path + "\" & LANGINI & ".ini"
'FORM1
    Section = "DIALOG1"
    DefRtn = "DIALOG1"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "TITLE"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Me.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'LABEL2/3-
    For i = 0 To 3
      DefRtn = "LABEL2" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "LABEL2" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Me.Label2(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
      DefRtn = "LABEL3" & VAL(i)
      GIntMAX = 80
      Gbuf = Space$(GIntMAX)
      IKey = "LABEL3" & VAL(i)
      sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
      Me.Label3(i).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    Next i
'OKBUTTON
    DefRtn = "OKBUTTON"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "OKBUTTON"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    OKButton.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'CANCELBUTTON
    DefRtn = "CANCELBUTTON"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "CANCELBUTTON"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    CancelButton.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)

   If Dl1Tag(0) <> "" Then Text1.tag = Dl1Tag(0)
   If Dl1Frm(0) <> 0 Then Text1.DataField = Dl1Frm(0)
   For i = 0 To 2
     If Dl1Tag(i + 1) <> "" Then Label1(i).tag = Dl1Tag(i + 1)
     If Dl1Frm(i + 1) <> 0 Then Label1(i).DataField = Dl1Frm(i + 1)
   Next i
   Call dlgdsp1
End Sub
Private Sub dlgdsp1()
    Dim tag As String
    Dim i, j, ii As Integer
    tag = Text1.tag
    ii = Text1.DataField
    frm = frmget(ii)
    VV = dVGETT(tag)
    Text1.Text = Format(VV, frm)
    For i = 0 To 2
      tag = Label1(i).tag
      ii = Label1(i).DataField
      frm = frmget(ii)
      VV = dVGETT(tag)
      Label1(i).Caption = Format(VV, frm)
    Next i
End Sub


Private Sub OKButton_Click()
    Dim tag As String
    Dim ii As Integer
    tag = Text1.tag
    VV = VAL(Text1.Text)
    ii = Text1.DataField
    frm = frmget(ii)
    Form1.Label4(0).Caption = Format(VV, frm)
    Form1.Label4(0).ForeColor = vbRed
    ist = dVSETT(tag, VV)
    Unload Me
End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    Dim VALUE As double
'    tag = Text1(Index).tag
'    If (KeyAscii <> &HD) Then Exit Sub
'    VALUE = VAL(Text1(Index).Text)
'    Text1.ForeColor = vbRed
''    ist = dVSETT(tag, VALUE)
'End Sub
