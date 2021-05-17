VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modelo Matematico Ver5.0.0"
   ClientHeight    =   8865
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   10170
   Icon            =   "CTUNG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Correccion de CTUNG"
      Height          =   3255
      Index           =   1
      Left            =   7440
      TabIndex        =   53
      Top             =   4680
      Width           =   2655
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   960
         TabIndex        =   100
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   960
         TabIndex        =   99
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   960
         TabIndex        =   98
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   97
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   960
         TabIndex        =   96
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   960
         TabIndex        =   95
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   18
         Left            =   960
         TabIndex        =   78
         Tag             =   "ChrPF"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   17
         Left            =   960
         TabIndex        =   77
         Tag             =   "ChrCW"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   76
         Tag             =   "ChrLH"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   15
         Left            =   960
         TabIndex        =   75
         Tag             =   "ChrRH"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   14
         Left            =   960
         TabIndex        =   74
         Tag             =   "ChrBP"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   13
         Left            =   960
         TabIndex        =   73
         Tag             =   "ChrAT"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_AT"
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_BP"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_RH"
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_LHV"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_CW"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Chr_PF"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Energia Disponible"
      Height          =   975
      Left            =   120
      TabIndex        =   43
      Top             =   6360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   83
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modifique"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   46
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500"
         DataField       =   "2"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   61
         Tag             =   "EDh"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   45
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "en el Punto de interconexion(EDh)"
         Height          =   255
         Index           =   18
         Left            =   480
         TabIndex        =   44
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Capacidad Neta Demonstrada(KC)"
      Height          =   975
      Left            =   120
      TabIndex        =   40
      Top             =   5280
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   82
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modifique"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   42
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "495.00"
         DataField       =   "2"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   60
         Tag             =   "KC100"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   10
         Left            =   2400
         TabIndex        =   41
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   2
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Correccion de Potencia"
      Height          =   3255
      Index           =   0
      Left            =   4680
      TabIndex        =   33
      Top             =   4680
      Width           =   2655
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   960
         TabIndex        =   94
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   960
         TabIndex        =   93
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   960
         TabIndex        =   92
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   960
         TabIndex        =   91
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   90
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   960
         TabIndex        =   89
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   12
         Left            =   960
         TabIndex        =   72
         Tag             =   "CpwPF100"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   11
         Left            =   960
         TabIndex        =   71
         Tag             =   "CpwCW100"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   10
         Left            =   960
         TabIndex        =   70
         Tag             =   "CpwLH100"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   69
         Tag             =   "CpwRH100"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   68
         Tag             =   "CpwBP100"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "15"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   67
         Tag             =   "CpwAT100"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_PF"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_CW"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   38
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_LHV"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_RH"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_BP"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cpw_AT"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resultad del Calculo"
      Height          =   3855
      Left            =   4680
      TabIndex        =   28
      Top             =   720
      Width           =   5415
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   3000
         TabIndex        =   101
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3000
         TabIndex        =   88
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   87
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   86
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   85
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   84
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "kJ/h"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   80
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   79
         Tag             =   "CTRH"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   66
         Tag             =   "CTRH"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   65
         Tag             =   "NPxx"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   64
         Tag             =   "CShsum"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   63
         Tag             =   "KCact100"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "2"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   62
         Tag             =   "LOAD"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "CTUNG"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   52
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "kJ/kWh"
         Height          =   255
         Index           =   14
         Left            =   4560
         TabIndex        =   51
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "CTUNG a Carga XX%[NPxx]"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   50
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "kJ/kWh"
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   49
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Capacidad Actual Correqid a la Condicion de Diseno de Verano"
         Height          =   495
         Index           =   17
         Left            =   240
         TabIndex        =   48
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   11
         Left            =   4560
         TabIndex        =   47
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "100% de la Capacidad Neta en las Cond. Act. de Operation"
         Height          =   495
         Index           =   9
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Calga Actual de la Planta"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "CTOV"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   81
         Top             =   3360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Entrada"
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   4455
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   6
         Left            =   2280
         TabIndex        =   10
         Tag             =   "PF"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   5
         Left            =   2280
         TabIndex        =   9
         Tag             =   "CW"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   4
         Left            =   2280
         TabIndex        =   8
         Tag             =   "PCI"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Tag             =   "RH"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Tag             =   "BP"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Tag             =   "AT"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         Height          =   270
         Index           =   7
         Left            =   2280
         TabIndex        =   4
         Tag             =   "CShman"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "3"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Tag             =   "CSh"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   27
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Factor de Potencia"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "degC"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   25
         Top             =   2900
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Temp. de Agua del Mar"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "kJ/kg"
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   23
         Top             =   2420
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Poder Calori. Inf"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   21
         Top             =   1940
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Humedad Relativa"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "bar(A)"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   19
         Top             =   1460
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Presion Atomosferica"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "degC"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   17
         Top             =   980
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Temp. de Bulbo Seco"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Entrada Manual de Capacidad Neta"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "MW"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   13
         Top             =   500
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Capacidad Actual"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculo"
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   0
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Modelo Matematico de C.C.C TUXPAN V Electricidad Sol de Tuxpan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin VB.Menu FILE 
      Caption         =   "Archivos"
      Begin VB.Menu READ 
         Caption         =   "Leer"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VV As Double
Dim icount As Long
Dim ist As Long
Dim CmdStr As String

Private Type FData
    ID As Integer
    Frmt As String
End Type
Dim NFData() As FData

Private Sub valinp()
    Dim tag As String
    Dim i As Integer
    For i = 0 To 7
      VV = VAL(Text1(i).Text)
      tag = Text1(i).tag
      ist = dVSETT(tag, VV)
    Next i
End Sub
Private Sub valint()
    Dim tag As String
    Dim comment As String
    Dim frm As String
    Dim i, j, ii As Integer
    For i = 0 To 7
      tag = Text1(i).tag
      ii = Text1(i).DataField
      frm = frmget(ii)
      VV = dVGETT(tag)
      comment = Space(54)
      ist = dCGETT(tag, comment)
      If (comment = Space(54)) Then
        Text1(i).ToolTipText = tag
      Else
        Text1(i).ToolTipText = RTrim(comment)
      End If
      Text1(i).Text = Format(VV, frm)
    Next i
    For i = 0 To 19
      tag = Label4(i).tag
      comment = Space(54)
      ist = dCGETT(tag, comment)
      If (comment = Space(54)) Then
        Label4(i).ToolTipText = tag
      Else
        Label4(i).ToolTipText = RTrim(comment)
      End If
    Next i
End Sub


Private Sub valdsp()
    Dim tag As String
    Dim frm As String
    Dim i, j, ii As Integer
    For i = 0 To 19
      tag = Label4(i).tag
      ii = Label4(i).DataField
      frm = frmget(ii)
      If i >= 13 And i <= 18 Then
        If CDbl(Label4(2).Caption) < 2 Then
          If i = 13 Then
            VV = dVGETT("ChrCTOV")
            Label4(i).Caption = Format(VV, frm)
            Text2(i).Text = Format(VV, frm)
          Else
            Label4(i).Caption = Format(1, frm)
            Text2(i).Text = Format(1, frm)
          End If
        Else
          VV = dVGETT(tag)
          Label4(i).Caption = Format(VV, frm)
          Text2(i).Text = Format(VV, frm)
        End If
      Else
        VV = dVGETT(tag)
        Label4(i).Caption = Format(VV, frm)
        Text2(i).Text = Format(VV, frm)
      End If
    Next i
    If CDbl(Label4(2).Caption) < 2 Then
      Label4(5).ForeColor = vbMenuBar
      Text2(5).Text = ""
      Label4(6).ForeColor = vbMenuBar
      Text2(6).Text = ""
      Label4(19).ForeColor = vbWindowText
    Else
      Label4(5).ForeColor = vbWindowText
      Label4(6).ForeColor = vbWindowText
      Label4(19).ForeColor = vbMenuBar 'vbMenuBar "desaparece" valor en GUI (Solo visualmente)
      Text2(19).Text = ""
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim outbuf As String
    Dim i As Long
    
'EXIT PROGRAM
    If (Index = 1) Then
      Call profileset
      Call Form_Unload(0)
    End If

'VALUE INPUT
    Call valinp

'START CALCULATION
    ist = dCALCULATE(3)
       
'VALUE DISPLAY
    Call valint
    Call valdsp
    For i = 0 To 7
      Text1(i).ForeColor = vbBlack
      If (i < 2) Then Label4(i).ForeColor = vbBlack
    Next i
End Sub

Private Sub Command2_Click(Index As Integer)
    If Index = 0 Then Dialog1.Show
    If Index = 1 Then Dialog2.Show
    Call valint
    Call valdsp
End Sub

Private Sub Form_Load()
    Dim buf As String
    Call ProfileGet
    ist = dCOMPILE(3)
'VALUE INPUT
    Call valint
'START CALCULATION
    ist = dCALCULATE(3)
'VALUE DISPLAY
    Call valdsp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim outbuf As String
    Dim ist As Long
    ist = dLUPCLS(1)
    Open OutFile2 For Output Shared As #1
    For i = 1 To 74
        outbuf = ""
        If i = 70 Then
          outbuf = Str(i) & "," & "0.0"
        Else
          outbuf = Str(i) & "," & CStr(dVGET(i))
        End If
        Print #1, outbuf
    Next i
   Close #1
   End
End Sub


Private Sub READ_Click()
    Dim rval(7) As Double
    Dim fullname As String
    Dim outbuf As String
    Dim ist As Long
    Dim i As Integer
    Dim tag(7) As String
    Dim CNO As String
    fullname = GetOpenFileName("*.csv")
    If fullname = "" Then Exit Sub
    If fullname = "*.csv" Then Exit Sub
    On Error GoTo NEXT1
    Open fullname For Input As #2
    Open OutFile1 For Output Shared As #1
    icount = 0
    CNO = ""
    Input #2, CNO, tag(0), tag(1), tag(2), tag(3), tag(4), tag(5), tag(6)
    outbuf = CNO
    For i = 0 To 6
      outbuf = outbuf & "," & Text1(i).tag
      If (i = 0) Then outbuf = outbuf & "," & Text1(7).tag
    Next i
    For i = 0 To 19
      outbuf = outbuf & "," & Label4(i).tag
    Next i
    Print #1, outbuf
    While 1
      CNO = ""
      Input #2, CNO, rval(0), rval(1), rval(2), rval(3), rval(4), rval(5), rval(6)
      For i = 0 To 6
       ist = dVSETT(tag(i), rval(i))
      Next i
'VALUE INPUT
      Call valint
'START CALCULATION
      ist = dCALCULATE(3)
'VALUE DISPLAY
      Call valdsp
      icount = icount + 1
      outbuf = CNO
      outbuf = Str(icount)
      For i = 0 To 6
        outbuf = outbuf & "," & Text1(i).Text
        If (i = 0) Then outbuf = outbuf & "," & Text1(7).Text
      Next i
      For i = 0 To 19
        If i = 5 Or i = 6 Or i = 19 Then
          If i = 5 Or i = 6 Then
            If CDbl(Label4(2).Caption) < 2 Then
              outbuf = outbuf & ","
            Else
              outbuf = outbuf & "," & Label4(i).Caption
            End If
          Else
            If CDbl(Label4(2).Caption) < 2 Then
              outbuf = outbuf & "," & Label4(i).Caption
            Else
              outbuf = outbuf & ","
            End If
          End If
        Else
          outbuf = outbuf & "," & Label4(i).Caption
        End If
      Next i
      Print #1, outbuf
    Wend
NEXT1:
    Close #1
    Close #2
End Sub

Private Function GetOpenFileName(FileFilter As String) As String
    On Error Resume Next
    With CommonDialog1
        .Flags = &H1000 Or &H4
        .Filter = FileFilter
        .FileName = FileFilter
        .ShowOpen
        DoEvents
        If Err = 0 Then GetOpenFileName = .FileName
    End With
End Function

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim VALUE As Double
    tag = Text1(Index).tag
    If (KeyAscii <> &HD) Then Exit Sub
    VALUE = VAL(Text1(Index).Text)
    Text1(Index).ForeColor = vbRed
    ist = dVSETT(tag, VALUE)
    Call valint
    Call valdsp
End Sub

