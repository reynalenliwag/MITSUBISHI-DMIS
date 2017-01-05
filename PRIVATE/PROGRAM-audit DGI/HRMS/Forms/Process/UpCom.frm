VERSION 5.00
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmHRMSUpCom 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Update Commission"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   2895
   ControlBox      =   0   'False
   ForeColor       =   &H00D8E9EC&
   Icon            =   "UpCom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "UpCom.frx":030A
   ScaleHeight     =   2160
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   780
      Width           =   885
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   1845
   End
   Begin VB.ComboBox cboQuensina 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DEDFDE&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1500
      MouseIcon       =   "UpCom.frx":3046
      MousePointer    =   99  'Custom
      Picture         =   "UpCom.frx":3198
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Width           =   885
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Generate"
      Height          =   825
      Left            =   510
      MouseIcon       =   "UpCom.frx":41DA
      MousePointer    =   99  'Custom
      Picture         =   "UpCom.frx":432C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   885
   End
End
Attribute VB_Name = "frmHRMSUpCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEMPINFO, rsDeductions, rsPAYROLL As ADODB.Recordset
Dim FromDate, ToDate As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo ErrorCode
Dim MM, ddFROM, YY As String
MM = What_month(cboMonth)
YY = cboYear.Text
If cboQuensina.Text = "1st Quensena" Then
   FromDate = DateSerial(YY, MM, 1)
   ToDate = DateSerial(YY, MM, 15)
Else
   FromDate = DateSerial(YY, MM, 16)
   ToDate = lastDay(FromDate)
End If
GENFROM = Format(FromDate, "Short Date")
GENTO = Format(ToDate, "Short Date")
frmHRMSUpDateCom.Show vbModal
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
wizMacApp1.MacCaption = Me.Caption
wizMacApp1.Buttons = CloseMinimize
cboQuensina.AddItem "1st Quensena"
cboQuensina.AddItem "2nd Quensena"
fillcbomonth cboMonth
FillcboYear cboYear
If Day(LOGDATE) > 15 Then
   cboQuensina.Text = "2nd Quensena"
Else
   cboQuensina.Text = "1st Quensena"
End If
cboYear.Text = Year(LOGDATE)
cboMonth.Text = The_month(Month(LOGDATE))
DrawXPCtl Me
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadForm Me
End Sub
