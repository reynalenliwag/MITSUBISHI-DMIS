VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmToolsCalculator 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7035
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "ToolCalculator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ToolCalculator.frx":08CA
   ScaleHeight     =   4275
   ScaleWidth      =   7035
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TFocus 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Focus"
      Top             =   5640
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cdbColor 
      Left            =   -480
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose color"
      Orientation     =   2
   End
   Begin VB.CommandButton cmdExit 
      DownPicture     =   "ToolCalculator.frx":3606
      Height          =   615
      Left            =   3690
      Picture         =   "ToolCalculator.frx":3A48
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4530
      Width           =   1095
   End
   Begin VB.CommandButton cmdColor 
      DownPicture     =   "ToolCalculator.frx":3E8A
      Height          =   615
      Left            =   2250
      Picture         =   "ToolCalculator.frx":4754
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4530
      Width           =   1095
   End
   Begin VB.TextBox txtSecond 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "0"
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtFirst 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "0"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdGenPower 
      DownPicture     =   "ToolCalculator.frx":501E
      Height          =   495
      Left            =   3570
      MouseIcon       =   "ToolCalculator.frx":58E8
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":5A3A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdGenRoot 
      DownPicture     =   "ToolCalculator.frx":6304
      Height          =   495
      Left            =   3570
      MouseIcon       =   "ToolCalculator.frx":6BCE
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":6D20
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdSquare 
      DownPicture     =   "ToolCalculator.frx":75EA
      Height          =   495
      Left            =   2850
      MouseIcon       =   "ToolCalculator.frx":7EB4
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":8006
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSqrRoot 
      DownPicture     =   "ToolCalculator.frx":88D0
      Height          =   495
      Left            =   2850
      MouseIcon       =   "ToolCalculator.frx":919A
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":92EC
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdNatLog 
      DownPicture     =   "ToolCalculator.frx":9BB6
      Height          =   495
      Left            =   6060
      MouseIcon       =   "ToolCalculator.frx":A480
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":A5D2
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdComLog 
      DownPicture     =   "ToolCalculator.frx":AE9C
      Height          =   495
      Left            =   330
      MouseIcon       =   "ToolCalculator.frx":B766
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":B8B8
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdExp_e 
      DownPicture     =   "ToolCalculator.frx":C182
      Height          =   495
      Left            =   5370
      MouseIcon       =   "ToolCalculator.frx":CA4C
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":CB9E
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdExp_ten 
      DownPicture     =   "ToolCalculator.frx":D468
      Height          =   495
      Left            =   1050
      MouseIcon       =   "ToolCalculator.frx":DD32
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":DE84
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdSum 
      DownPicture     =   "ToolCalculator.frx":E74E
      Height          =   495
      Left            =   4890
      MouseIcon       =   "ToolCalculator.frx":F018
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":F16A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdPermutation 
      Caption         =   "P(x,y)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      MouseIcon       =   "ToolCalculator.frx":FA34
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdCombination 
      Caption         =   "C(x,y)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1290
      MouseIcon       =   "ToolCalculator.frx":FB86
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "x!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5820
      MouseIcon       =   "ToolCalculator.frx":FCD8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdContagent 
      Caption         =   "cot x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3690
      MouseIcon       =   "ToolCalculator.frx":FE2A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdTangent 
      Caption         =   "tan x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2490
      MouseIcon       =   "ToolCalculator.frx":FF7C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdCosine 
      Caption         =   "cos x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3690
      MouseIcon       =   "ToolCalculator.frx":100CE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdSine 
      Caption         =   "sin x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2490
      MouseIcon       =   "ToolCalculator.frx":10220
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdDivision 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      MouseIcon       =   "ToolCalculator.frx":10372
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdMultiplication 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2850
      MouseIcon       =   "ToolCalculator.frx":104C4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdSubtraction 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      MouseIcon       =   "ToolCalculator.frx":10616
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdAddition 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2850
      MouseIcon       =   "ToolCalculator.frx":10768
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3090
      MouseIcon       =   "ToolCalculator.frx":108BA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRBSp 
      DownPicture     =   "ToolCalculator.frx":10A0C
      Height          =   495
      Left            =   5250
      MouseIcon       =   "ToolCalculator.frx":10E4E
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":10FA0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdLBSp 
      DownPicture     =   "ToolCalculator.frx":113E2
      Height          =   495
      Left            =   450
      MouseIcon       =   "ToolCalculator.frx":11824
      MousePointer    =   99  'Custom
      Picture         =   "ToolCalculator.frx":11976
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdR9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      MouseIcon       =   "ToolCalculator.frx":11DB8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdR8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5250
      MouseIcon       =   "ToolCalculator.frx":11F0A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdR7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      MouseIcon       =   "ToolCalculator.frx":1205C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdR6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      MouseIcon       =   "ToolCalculator.frx":121AE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdR5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5250
      MouseIcon       =   "ToolCalculator.frx":12300
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdR4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      MouseIcon       =   "ToolCalculator.frx":12452
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdR3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      MouseIcon       =   "ToolCalculator.frx":125A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdR2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5250
      MouseIcon       =   "ToolCalculator.frx":126F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdR1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      MouseIcon       =   "ToolCalculator.frx":12848
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdR0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      MouseIcon       =   "ToolCalculator.frx":1299A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdL9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      MouseIcon       =   "ToolCalculator.frx":12AEC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdL8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      MouseIcon       =   "ToolCalculator.frx":12C3E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdL7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      MouseIcon       =   "ToolCalculator.frx":12D90
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdL6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      MouseIcon       =   "ToolCalculator.frx":12EE2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdL5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      MouseIcon       =   "ToolCalculator.frx":13034
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdL4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      MouseIcon       =   "ToolCalculator.frx":13186
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdL3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      MouseIcon       =   "ToolCalculator.frx":132D8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdL2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      MouseIcon       =   "ToolCalculator.frx":1342A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdL1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      MouseIcon       =   "ToolCalculator.frx":1357C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton cmdL0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      MouseIcon       =   "ToolCalculator.frx":136CE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1770
      TabIndex        =   48
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Menu mnuCalculator 
      Caption         =   "&Calculator"
      Begin VB.Menu mnuCalculatorReset 
         Caption         =   "Reset"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculatorExit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy Result"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFirst 
         Caption         =   "Paste in First Box"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditSecond 
         Caption         =   "Paste in Second Box"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewThemes 
         Caption         =   "Themes"
         Begin VB.Menu mnuViewThemesDefault 
            Caption         =   "Default"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewThemesSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewThemesReddish 
            Caption         =   "Reddish"
         End
         Begin VB.Menu mnuViewThemesNight 
            Caption         =   "Night"
         End
         Begin VB.Menu mnuViewThemesSeaside 
            Caption         =   "Seaside"
         End
         Begin VB.Menu mnuViewThemesSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewThemesColor 
            Caption         =   "Pick a Color"
         End
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "Mode"
         Begin VB.Menu mnuModeSimple 
            Caption         =   "Simple"
         End
         Begin VB.Menu mnuModeAdvanced 
            Caption         =   "Advanced"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmToolsCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dblFirst As Double
Dim dblSecond As Double
Dim dblResult As Double
Dim dblBSpL As Double, dblBSpR As Double
Dim intMessage As Integer
Dim strTheme As String, strMode As String

Private Sub cmdAddition_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    dblResult = dblFirst + dblSecond
    lblResult.Caption = dblResult
End Sub

Private Sub cmdColor_Click()
TFocus.SetFocus
    cdbColor.ShowColor
    Me.BackColor = cdbColor.Color
    
    cmdL0.BackColor = &H8000000F
    cmdL1.BackColor = &H8000000F
    cmdL2.BackColor = &H8000000F
    cmdL3.BackColor = &H8000000F
    cmdL4.BackColor = &H8000000F
    cmdL5.BackColor = &H8000000F
    cmdL6.BackColor = &H8000000F
    cmdL7.BackColor = &H8000000F
    cmdL8.BackColor = &H8000000F
    cmdL9.BackColor = &H8000000F
    cmdLBSp.BackColor = &H8000000F
    
    cmdR0.BackColor = &H8000000F
    cmdR1.BackColor = &H8000000F
    cmdR2.BackColor = &H8000000F
    cmdR3.BackColor = &H8000000F
    cmdR4.BackColor = &H8000000F
    cmdR5.BackColor = &H8000000F
    cmdR6.BackColor = &H8000000F
    cmdR7.BackColor = &H8000000F
    cmdR8.BackColor = &H8000000F
    cmdR9.BackColor = &H8000000F
    cmdRBSp.BackColor = &H8000000F
    
    txtFirst.BackColor = &H0&
    txtFirst.ForeColor = vbWhite
    txtSecond.BackColor = &H0&
    txtSecond.ForeColor = vbWhite
    lblResult.BackColor = &H800000
    lblResult.ForeColor = vbYellow
    
    cmdAddition.BackColor = &H8000000F
    cmdCombination.BackColor = &H8000000F
    cmdComLog.BackColor = &H8000000F
    cmdContagent.BackColor = &H8000000F
    cmdCosine.BackColor = &H8000000F
    cmdDivision.BackColor = &H8000000F
    cmdExit.BackColor = &H8000000F
    cmdExp_e.BackColor = &H8000000F
    cmdExp_ten.BackColor = &H8000000F
    cmdFactorial.BackColor = &H8000000F
    cmdGenPower.BackColor = &H8000000F
    cmdGenRoot.BackColor = &H8000000F
    cmdMultiplication.BackColor = &H8000000F
    cmdNatLog.BackColor = &H8000000F
    cmdPermutation.BackColor = &H8000000F
    cmdReset.BackColor = &H8000000F
    cmdSine.BackColor = &H8000000F
    cmdSqrRoot.BackColor = &H8000000F
    cmdSquare.BackColor = &H8000000F
    cmdSubtraction.BackColor = &H8000000F
    cmdSum.BackColor = &H8000000F
    cmdTangent.BackColor = &H8000000F
    
    mnuViewThemesDefault.Checked = False
    mnuViewThemesReddish.Checked = False
    mnuViewThemesNight.Checked = False
    mnuViewThemesSeaside.Checked = False
    mnuViewThemesColor.Checked = True
End Sub

Private Sub cmdCombination_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    
    If dblSecond > dblFirst Then
        lblResult.Caption = "Invalid input for function"
    Else
        If dblFirst > 170 Then
            lblResult.Caption = "Result is out of range"
        Else
            dblResult = Combination(dblFirst, dblSecond)
            lblResult.Caption = dblResult
        End If
    End If
End Sub

Private Sub cmdComLog_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst <= 0 Then
        lblResult.Caption = "Invalid input for function"
    Else
        dblResult = CommonLog(dblFirst)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdContagent_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst = 0 Then
        lblResult.Caption = "Undefined"
    Else
        dblResult = Cot(dblFirst)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdCosine_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblResult = Cos((dblFirst * Pi) / 180)
    lblResult.Caption = dblResult
End Sub

Private Sub cmdDivision_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    If dblSecond = 0 Then
        lblResult.Caption = "Cannot divide by 0"
    Else
        dblResult = dblFirst / dblSecond
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdExit_Click()
TFocus.SetFocus
    Call mnuCalculatorExit_Click
End Sub

Private Sub cmdExp_e_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst > 709 Then
        lblResult.Caption = "Result is out of range"
    Else
        dblResult = Exp(dblFirst)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdExp_ten_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst > 308 Then
        lblResult.Caption = "Result is out of range"
    Else
        dblResult = 10 ^ dblFirst
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdFactorial_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst > 170 Then
        lblResult.Caption = "Result is out of range"
    Else
        dblResult = Factorial(dblFirst)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdGenPower_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    dblResult = dblFirst ^ dblSecond
    lblResult.Caption = dblResult
End Sub

Private Sub cmdGenRoot_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    If dblSecond = 0 Then
        lblResult.Caption = "Invalid input for function"
    Else
        dblResult = dblFirst ^ (1 / dblSecond)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdL0_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 0
    End If
End Sub

Private Sub cmdL1_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 1
    End If
End Sub

Private Sub cmdL2_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 2
    End If
End Sub

Private Sub cmdL3_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 3
    End If
End Sub

Private Sub cmdL4_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 4
    End If
End Sub

Private Sub cmdL5_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 5
    End If
End Sub

Private Sub cmdL6_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 6
    End If
End Sub

Private Sub cmdL7_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 7
    End If
End Sub

Private Sub cmdL8_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 8
    End If
End Sub

Private Sub cmdL9_Click()
TFocus.SetFocus
    If Len(txtFirst.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtFirst.Text = txtFirst.Text * 10 + 9
    End If
End Sub

Private Sub cmdLBSp_Click()
TFocus.SetFocus
    dblBSpL = txtFirst.Text
    dblBSpR = Int(dblBSpL / 10)
    txtFirst.Text = dblBSpR
End Sub

Private Sub cmdMultiplication_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    dblResult = dblFirst * dblSecond
    lblResult.Caption = dblResult
End Sub

Private Sub cmdNatLog_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    If dblFirst <= 0 Then
        lblResult.Caption = "Invalid input for function"
    Else
        dblResult = Log(dblFirst)
        lblResult.Caption = dblResult
    End If
End Sub

Private Sub cmdPermutation_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    
    If dblSecond > dblFirst Then
        lblResult.Caption = "Invalid input for function"
    Else
        If dblFirst > 170 Then
            lblResult.Caption = "Result is out of range"
        Else
            dblResult = Permutation(dblFirst, dblSecond)
            lblResult.Caption = dblResult
        End If
    End If
End Sub

Private Sub cmdR0_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 0
    End If
End Sub

Private Sub cmdR1_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 1
    End If
End Sub

Private Sub cmdR2_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 2
    End If
End Sub

Private Sub cmdR3_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 3
    End If
End Sub

Private Sub cmdR4_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 4
    End If
End Sub

Private Sub cmdR5_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 5
    End If
End Sub

Private Sub cmdR6_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 6
    End If
End Sub

Private Sub cmdR7_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 7
    End If
End Sub

Private Sub cmdR8_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 8
    End If
End Sub

Private Sub cmdR9_Click()
TFocus.SetFocus
    If Len(txtSecond.Text) >= 15 Then
        intMessage = MsgBox("Number may have at most 15 digits", vbOKOnly, "Max # of digits")
    Else
        txtSecond.Text = txtSecond.Text * 10 + 9
    End If
End Sub

Private Sub cmdRBSp_Click()
TFocus.SetFocus
    dblBSpL = txtSecond.Text
    dblBSpR = Int(dblBSpL / 10)
    txtSecond.Text = dblBSpR
End Sub

Private Sub cmdReset_Click()
TFocus.SetFocus
    txtFirst.Text = 0
    txtSecond.Text = 0
End Sub

Private Sub cmdSine_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblResult = Sin((dblFirst * Pi) / 180)
    lblResult.Caption = dblResult
End Sub

Private Sub cmdSqrRoot_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblResult = Sqr(dblFirst)
    lblResult.Caption = dblResult
End Sub

Private Sub cmdSquare_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblResult = dblFirst ^ 2
    lblResult.Caption = dblResult
End Sub

Private Sub cmdSubtraction_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = CDbl(txtSecond.Text)
    dblResult = dblFirst - dblSecond
    lblResult.Caption = dblResult
End Sub

Private Sub cmdSum_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblSecond = 1
    lblResult.Caption = "Result is out of range"
    dblResult = Sum(dblSecond, dblFirst)
    lblResult.Caption = dblResult
End Sub

Private Sub cmdTangent_Click()
TFocus.SetFocus
    dblFirst = CDbl(txtFirst.Text)
    dblResult = Tan((dblFirst * Pi) / 180)
    lblResult.Caption = dblResult
End Sub

Private Sub Form_Load()
    strTheme = "Default"
    strMode = "ModeSimple"
    Select Case strTheme
        Case "Default"
            Call mnuViewThemesDefault_Click
        Case "Reddish"
            Call mnuViewThemesReddish_Click
        Case "Night"
            Call mnuViewThemesNight_Click
        Case "Seaside"
            Call mnuViewThemesSeaside_Click
        Case Else
            Me.BackColor = strTheme
        End Select
    
    Select Case strMode
        Case "ModeAdvanced"
            Call mnuModeAdvanced_Click
        Case "ModeSimple"
            Call mnuModeSimple_Click
    End Select
    CenterMe frmMain, Me, 1
End Sub

Private Sub mnuCalculatorExit_Click() ' End application subroutine
    Unload Me
End Sub

Private Sub mnuCalculatorReset_Click()
TFocus.SetFocus
    Call cmdReset_Click
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetText (lblResult.Caption)
End Sub

Private Sub mnuEditFirst_Click()
    txtFirst.Text = Clipboard.GetText
End Sub

Private Sub mnuEditSecond_Click()
    txtSecond.Text = Clipboard.GetText
End Sub

Private Sub mnuHelpContents_Click()
MsgBox "for help on this calculator see the readme.txt file"
End Sub

Private Sub mnuModeAdvanced_Click()
strMode = "ModeAdvanced"
    strTheme = "Default"
    strMode = "ModeSimple"
    mnuModeSimple.Checked = False
    mnuModeAdvanced.Checked = True
    
    Me.Height = 6030
    
    cmdColor.Top = 4600
    cmdExit.Top = 4600
    
    cmdAddition.Visible = True
    cmdCombination.Visible = True
    cmdComLog.Visible = True
    cmdContagent.Visible = True
    cmdCosine.Visible = True
    cmdDivision.Visible = True
    cmdExit.Visible = True
    cmdExp_e.Visible = True
    cmdExp_ten.Visible = True
    cmdFactorial.Visible = True
    cmdGenPower.Visible = True
    cmdGenRoot.Visible = True
    cmdMultiplication.Visible = True
    cmdNatLog.Visible = True
    cmdPermutation.Visible = True
    cmdReset.Visible = True
    cmdSine.Visible = True
    cmdSqrRoot.Visible = True
    cmdSquare.Visible = True
    cmdSubtraction.Visible = True
    cmdSum.Visible = True
    cmdTangent.Visible = True
End Sub

Private Sub mnuModeSimple_Click()
    strTheme = "Default"
    strMode = "ModeSimple"
    mnuModeSimple.Checked = True
    mnuModeAdvanced.Checked = False
    Me.Height = 4935
    cmdColor.Top = 3520
    cmdExit.Top = 3520
    
    cmdAddition.Visible = True
    cmdCombination.Visible = False
    cmdComLog.Visible = False
    cmdContagent.Visible = False
    cmdCosine.Visible = False
    cmdDivision.Visible = True
    cmdExit.Visible = True
    cmdExp_e.Visible = False
    cmdExp_ten.Visible = False
    cmdFactorial.Visible = False
    cmdGenPower.Visible = True
    cmdGenRoot.Visible = True
    cmdMultiplication.Visible = True
    cmdNatLog.Visible = False
    cmdPermutation.Visible = False
    cmdReset.Visible = True
    cmdSine.Visible = False
    cmdSqrRoot.Visible = True
    cmdSquare.Visible = True
    cmdSubtraction.Visible = True
    cmdSum.Visible = False
    cmdTangent.Visible = False
End Sub

Private Sub mnuViewThemesColor_Click()
Call cmdColor_Click

strTheme = Me.BackColor
End Sub

Private Sub mnuViewThemesDefault_Click()
strTheme = "Default"
    Me.BackColor = &H8000000F
    
    cmdL0.BackColor = &H8000000F
    cmdL1.BackColor = &H8000000F
    cmdL2.BackColor = &H8000000F
    cmdL3.BackColor = &H8000000F
    cmdL4.BackColor = &H8000000F
    cmdL5.BackColor = &H8000000F
    cmdL6.BackColor = &H8000000F
    cmdL7.BackColor = &H8000000F
    cmdL8.BackColor = &H8000000F
    cmdL9.BackColor = &H8000000F
    cmdLBSp.BackColor = &H8000000F
    
    cmdR0.BackColor = &H8000000F
    cmdR1.BackColor = &H8000000F
    cmdR2.BackColor = &H8000000F
    cmdR3.BackColor = &H8000000F
    cmdR4.BackColor = &H8000000F
    cmdR5.BackColor = &H8000000F
    cmdR6.BackColor = &H8000000F
    cmdR7.BackColor = &H8000000F
    cmdR8.BackColor = &H8000000F
    cmdR9.BackColor = &H8000000F
    cmdRBSp.BackColor = &H8000000F
    
    txtFirst.BackColor = &H0&
    txtFirst.ForeColor = vbWhite
    txtSecond.BackColor = &H0&
    txtSecond.ForeColor = vbWhite
    lblResult.BackColor = &H800000
    lblResult.ForeColor = vbYellow
    
    cmdAddition.BackColor = &H8000000F
    cmdCombination.BackColor = &H8000000F
    cmdComLog.BackColor = &H8000000F
    cmdContagent.BackColor = &H8000000F
    cmdCosine.BackColor = &H8000000F
    cmdDivision.BackColor = &H8000000F
    cmdExit.BackColor = &H8000000F
    cmdExp_e.BackColor = &H8000000F
    cmdExp_ten.BackColor = &H8000000F
    cmdFactorial.BackColor = &H8000000F
    cmdGenPower.BackColor = &H8000000F
    cmdGenRoot.BackColor = &H8000000F
    cmdMultiplication.BackColor = &H8000000F
    cmdNatLog.BackColor = &H8000000F
    cmdPermutation.BackColor = &H8000000F
    cmdReset.BackColor = &H8000000F
    cmdSine.BackColor = &H8000000F
    cmdSqrRoot.BackColor = &H8000000F
    cmdSquare.BackColor = &H8000000F
    cmdSubtraction.BackColor = &H8000000F
    cmdSum.BackColor = &H8000000F
    cmdTangent.BackColor = &H8000000F
    
    mnuViewThemesDefault.Checked = True
    mnuViewThemesReddish.Checked = False
    mnuViewThemesNight.Checked = False
    mnuViewThemesSeaside.Checked = False
    mnuViewThemesColor.Checked = False
End Sub

Private Sub mnuViewThemesNight_Click()
strTheme = "Night"
    Me.BackColor = vbBlack
    
    cmdL0.BackColor = vbWhite
    cmdL1.BackColor = vbWhite
    cmdL2.BackColor = vbWhite
    cmdL3.BackColor = vbWhite
    cmdL4.BackColor = vbWhite
    cmdL5.BackColor = vbWhite
    cmdL6.BackColor = vbWhite
    cmdL7.BackColor = vbWhite
    cmdL8.BackColor = vbWhite
    cmdL9.BackColor = vbWhite
    cmdLBSp.BackColor = vbWhite
    
    cmdR0.BackColor = vbWhite
    cmdR1.BackColor = vbWhite
    cmdR2.BackColor = vbWhite
    cmdR3.BackColor = vbWhite
    cmdR4.BackColor = vbWhite
    cmdR5.BackColor = vbWhite
    cmdR6.BackColor = vbWhite
    cmdR7.BackColor = vbWhite
    cmdR8.BackColor = vbWhite
    cmdR9.BackColor = vbWhite
    cmdRBSp.BackColor = vbWhite
    
    txtFirst.BackColor = vbWhite
    txtFirst.ForeColor = vbBlack
    txtSecond.BackColor = vbWhite
    txtSecond.ForeColor = vbBlack
    lblResult.BackColor = vbBlack
    lblResult.ForeColor = vbWhite
    
    cmdAddition.BackColor = vbWhite
    cmdCombination.BackColor = vbWhite
    cmdComLog.BackColor = vbWhite
    cmdContagent.BackColor = vbWhite
    cmdCosine.BackColor = vbWhite
    cmdDivision.BackColor = vbWhite
    cmdExit.BackColor = vbWhite
    cmdExp_e.BackColor = vbWhite
    cmdExp_ten.BackColor = vbWhite
    cmdFactorial.BackColor = vbWhite
    cmdGenPower.BackColor = vbWhite
    cmdGenRoot.BackColor = vbWhite
    cmdMultiplication.BackColor = vbWhite
    cmdNatLog.BackColor = vbWhite
    cmdPermutation.BackColor = vbWhite
    cmdReset.BackColor = vbWhite
    cmdSine.BackColor = vbWhite
    cmdSqrRoot.BackColor = vbWhite
    cmdSquare.BackColor = vbWhite
    cmdSubtraction.BackColor = vbWhite
    cmdSum.BackColor = vbWhite
    cmdTangent.BackColor = vbWhite
    
    mnuViewThemesDefault.Checked = False
    mnuViewThemesReddish.Checked = False
    mnuViewThemesNight.Checked = True
    mnuViewThemesSeaside.Checked = False
    mnuViewThemesColor.Checked = False
End Sub

Private Sub mnuViewThemesReddish_Click()
strTheme = "Reddish"
    Me.BackColor = vbRed
    
    cmdL0.BackColor = vbRed
    cmdL1.BackColor = vbRed
    cmdL2.BackColor = vbRed
    cmdL3.BackColor = vbRed
    cmdL4.BackColor = vbRed
    cmdL5.BackColor = vbRed
    cmdL6.BackColor = vbRed
    cmdL7.BackColor = vbRed
    cmdL8.BackColor = vbRed
    cmdL9.BackColor = vbRed
    cmdLBSp.BackColor = vbRed
    
    cmdR0.BackColor = vbRed
    cmdR1.BackColor = vbRed
    cmdR2.BackColor = vbRed
    cmdR3.BackColor = vbRed
    cmdR4.BackColor = vbRed
    cmdR5.BackColor = vbRed
    cmdR6.BackColor = vbRed
    cmdR7.BackColor = vbRed
    cmdR8.BackColor = vbRed
    cmdR9.BackColor = vbRed
    cmdRBSp.BackColor = vbRed
    
    txtFirst.BackColor = vbYellow
    txtFirst.ForeColor = vbBlue
    txtSecond.BackColor = vbYellow
    txtSecond.ForeColor = vbBlue
    lblResult.BackColor = vbCyan
    lblResult.ForeColor = vbBlack
    
    cmdAddition.BackColor = vbRed
    cmdCombination.BackColor = vbRed
    cmdComLog.BackColor = vbRed
    cmdContagent.BackColor = vbRed
    cmdCosine.BackColor = vbRed
    cmdDivision.BackColor = vbRed
    cmdExit.BackColor = vbRed
    cmdExp_e.BackColor = vbRed
    cmdExp_ten.BackColor = vbRed
    cmdFactorial.BackColor = vbRed
    cmdGenPower.BackColor = vbRed
    cmdGenRoot.BackColor = vbRed
    cmdMultiplication.BackColor = vbRed
    cmdNatLog.BackColor = vbRed
    cmdPermutation.BackColor = vbRed
    cmdReset.BackColor = vbRed
    cmdSine.BackColor = vbRed
    cmdSqrRoot.BackColor = vbRed
    cmdSquare.BackColor = vbRed
    cmdSubtraction.BackColor = vbRed
    cmdSum.BackColor = vbRed
    cmdTangent.BackColor = vbRed
    
    mnuViewThemesDefault.Checked = False
    mnuViewThemesReddish.Checked = True
    mnuViewThemesNight.Checked = False
    mnuViewThemesSeaside.Checked = False
    mnuViewThemesColor.Checked = False
End Sub

Private Sub mnuViewThemesSeaside_Click()
strTheme = "Seaside"
    Me.BackColor = vbBlue
    
    cmdL0.BackColor = vbCyan
    cmdL1.BackColor = vbCyan
    cmdL2.BackColor = vbCyan
    cmdL3.BackColor = vbCyan
    cmdL4.BackColor = vbCyan
    cmdL5.BackColor = vbCyan
    cmdL6.BackColor = vbCyan
    cmdL7.BackColor = vbCyan
    cmdL8.BackColor = vbCyan
    cmdL9.BackColor = vbCyan
    cmdLBSp.BackColor = vbCyan
    
    cmdR0.BackColor = vbCyan
    cmdR1.BackColor = vbCyan
    cmdR2.BackColor = vbCyan
    cmdR3.BackColor = vbCyan
    cmdR4.BackColor = vbCyan
    cmdR5.BackColor = vbCyan
    cmdR6.BackColor = vbCyan
    cmdR7.BackColor = vbCyan
    cmdR8.BackColor = vbCyan
    cmdR9.BackColor = vbCyan
    cmdRBSp.BackColor = vbCyan
    
    txtFirst.BackColor = vbYellow
    txtFirst.ForeColor = vbBlue
    txtSecond.BackColor = vbYellow
    txtSecond.ForeColor = vbBlue
    lblResult.BackColor = vbCyan
    lblResult.ForeColor = vbBlack
    
    cmdAddition.BackColor = &H8000000F
    cmdCombination.BackColor = &H8000000F
    cmdComLog.BackColor = &H8000000F
    cmdContagent.BackColor = &H8000000F
    cmdCosine.BackColor = &H8000000F
    cmdDivision.BackColor = &H8000000F
    cmdExit.BackColor = &H8000000F
    cmdExp_e.BackColor = &H8000000F
    cmdExp_ten.BackColor = &H8000000F
    cmdFactorial.BackColor = &H8000000F
    cmdGenPower.BackColor = &H8000000F
    cmdGenRoot.BackColor = &H8000000F
    cmdMultiplication.BackColor = &H8000000F
    cmdNatLog.BackColor = &H8000000F
    cmdPermutation.BackColor = &H8000000F
    cmdReset.BackColor = &H8000000F
    cmdSine.BackColor = &H8000000F
    cmdSqrRoot.BackColor = &H8000000F
    cmdSquare.BackColor = &H8000000F
    cmdSubtraction.BackColor = &H8000000F
    cmdSum.BackColor = &H8000000F
    cmdTangent.BackColor = &H8000000F
    
    mnuViewThemesDefault.Checked = False
    mnuViewThemesReddish.Checked = False
    mnuViewThemesNight.Checked = False
    mnuViewThemesSeaside.Checked = True
    mnuViewThemesColor.Checked = False
End Sub

Private Sub txtFirst_Click()
TFocus.SetFocus
End Sub

Private Sub txtSecond_Click()
TFocus.SetFocus
End Sub
