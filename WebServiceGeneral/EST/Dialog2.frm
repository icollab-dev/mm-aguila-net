VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "Energia Disponible"
   ClientHeight    =   1215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  '‰E‘µ‚¦
      DataField       =   "2"
      Height          =   270
      Left            =   1200
      TabIndex        =   2
      Tag             =   "EDh"
      Text            =   "495.00"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  '‰E‘µ‚¦
      Caption         =   "EDh"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "MW"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Dialog2"
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
    Section = "DIALOG2"
    DefRtn = "DIALOG2"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "TITLE"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Me.Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
'LABEL2/3-
    DefRtn = "LABEL20"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "LABEL20"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Me.Label2(0).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
    DefRtn = "LABEL30"
    GIntMAX = 80
    Gbuf = Space$(GIntMAX)
    IKey = "LABEL30"
    sts = GetPrivateProfileString(Section, IKey, DefRtn, Gbuf, GIntMAX, fPath)
    Me.Label3(0).Caption = Left$(Gbuf, InStr(1, Gbuf, Chr(0)) - 1)
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

   If Dl2Tag <> "" Then Text1.tag = Dl2Tag
   If Dl2Frm <> 0 Then Text1.DataField = Dl2Frm
   Call dlgdsp2
End Sub
Private Sub dlgdsp2()
    Dim tag As String
    Dim i, j, ii As Integer
    tag = Text1.tag
    ii = Text1.DataField
    frm = frmget(ii)
    VV = dVGETT(tag)
    Text1.Text = Format(VV, frm)
End Sub


Private Sub OKButton_Click()
    Dim tag As String
    Dim ii As Integer
    tag = Text1.tag
    VV = VAL(Text1.Text)
    ii = Text1.DataField
    frm = frmget(ii)
    Form1.Label4(1).Caption = Format(VV, frm)
    Form1.Label4(1).ForeColor = vbRed
    ist = dVSETT(tag, VV)
    Unload Me
End Sub

