VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmHRMSUpDateCom 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Update Commission"
   ClientHeight    =   1860
   ClientLeft      =   1500
   ClientTop       =   2850
   ClientWidth     =   5700
   ControlBox      =   0   'False
   ForeColor       =   &H00D8E9EC&
   Icon            =   "UpdateCom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "UpdateCom.frx":0442
   ScaleHeight     =   1860
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DEDFDE&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   795
      Left            =   4680
      MouseIcon       =   "UpdateCom.frx":317E
      MousePointer    =   99  'Custom
      Picture         =   "UpdateCom.frx":32D0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   990
      Width           =   945
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Go"
      Height          =   795
      Left            =   3810
      MouseIcon       =   "UpdateCom.frx":4312
      MousePointer    =   99  'Custom
      Picture         =   "UpdateCom.frx":4464
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   990
      Width           =   885
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Done"
      Height          =   795
      Left            =   4680
      Picture         =   "UpdateCom.frx":4D2E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   990
      Width           =   945
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   -30
      Picture         =   "UpdateCom.frx":5170
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   300
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   4
         Top             =   750
         Width           =   3615
         Begin VB.Label labName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   5
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg gauProgress 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "UpdateCom.frx":7EAC
         BackColor       =   14215660
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "UpdateCom.frx":7EC8
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00D8E9EC&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         Picture         =   "UpdateCom.frx":7EE4
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   7
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   8
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "UpdateCom.frx":AC20
         End
      End
      Begin VB.Label lblPercent 
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmHRMSUpDateCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEMPINFO, rsCommission As ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()
On Error GoTo ErrorCode
cmdGO.Visible = False
DoEvents
cmdCancel.Visible = False
cmdDone.Enabled = False
Dim TotCommission, TotCommissionTax As Double
Dim i, Cnt As Integer
Dim amt As Double
Dim KABALI As Boolean
KABALI = False
Set rsEMPINFO = New ADODB.Recordset
    rsEMPINFO.Open "select * from HRMS_EmpInfo where (datehired <= '" & Format(GENTO, "Short Date") & "') order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsEMPINFO.EOF And Not rsEMPINFO.BOF Then
   Screen.MousePointer = 11
   rsEMPINFO.MoveFirst
   i = 0
   Do While Not rsEMPINFO.EOF
         Set rsCommission = New ADODB.Recordset
             rsCommission.Open "select SUM(amount) AS TOTALCOMMISSION, SUM(tax) AS TOTALCOMMISSIONTAX from HRMS_Commission where empno = " & N2Str2Null(rsEMPINFO!empno) & " AND " & _
                                "(deyt BETWEEN '" & CDate(GENFROM) & "' AND '" & CDate(GENTO) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
         TotCommission = 0
         TotCommissionTax = 0
         If Not rsCommission.EOF And Not rsCommission.BOF Then
            TotCommission = N2Str2Zero(rsCommission!TOTALCOMMISSION)
            TotCommissionTax = N2Str2Zero(rsCommission!TOTALCOMMISSIONTAX)
         End If
         gconDMIS.Execute "update HRMS_Payroll set " & _
                          "commission = " & TotCommission & ", " & _
                          "commissionTax = " & TotCommissionTax & _
                          " where empno = " & N2Str2Null(rsEMPINFO!empno) & " AND (PAYDATEFROM = '" & GENFROM & "') AND (PAYDATETO = '" & GENTO & "')"
      i = i + 1
      gauProgress.Value = (i / rsEMPINFO.RecordCount) * 100
      lblPercent.Caption = Int(gauProgress.Value) & "%"
      rsEMPINFO.MoveNext
      DoEvents
   Loop
      
   labName.Caption = ""
   Screen.MousePointer = 0
End If
cmdDone.Enabled = True
Screen.MousePointer = 0
Exit Sub

ErrorCode:
ShowVBError
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
wizMacApp1.MacCaption = Me.Caption
wizMacApp1.Buttons = CloseMinimize
labName.Caption = ""
DrawXPCtl Me
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadForm Me
End Sub
