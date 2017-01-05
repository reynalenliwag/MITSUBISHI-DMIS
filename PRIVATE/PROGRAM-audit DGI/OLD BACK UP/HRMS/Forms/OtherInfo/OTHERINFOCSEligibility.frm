VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOCSEligibility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CS Eligibility"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1230
      Left            =   150
      ScaleHeight     =   1170
      ScaleWidth      =   4125
      TabIndex        =   1
      Top             =   120
      Width           =   4185
      Begin VB.CommandButton cmdCancel 
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
         Height          =   660
         Left            =   3330
         MouseIcon       =   "OTHERINFOCSEligibility.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOCSEligibility.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel"
         Top             =   450
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2640
         MouseIcon       =   "OTHERINFOCSEligibility.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOCSEligibility.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save Entry"
         Top             =   450
         Width           =   705
      End
      Begin VB.ComboBox cboCSEligibility 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Text            =   "cboEmpStatus"
         Top             =   60
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Eligibility"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   975
      End
   End
   Begin wizButton.cmd cmdDependents 
      Height          =   1350
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   2381
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERINFOCSEligibility.frx":0932
   End
End
Attribute VB_Name = "frmOTHERINFOCSEligibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMPINFO                                                         As ADODB.Recordset

Sub rsRefreshStoreMemVars()
    Set RSEMPINFO = New ADODB.Recordset
    Set RSEMPINFO = gconDMIS.Execute("Select CSEligibility from HRMS_EmpInfo Where EmpNo = " & EMPLOYEE_NO)
    If Not RSEMPINFO.EOF And Not RSEMPINFO.BOF Then
        cboCSEligibility.Text = Null2String(RSEMPINFO!CSEligibility)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:55
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    gconDMIS.Execute "update HRMS_EmpInfo Set" & _
                   " CSEligibility = " & N2Str2Null(cboCSEligibility.Text) & _
                   " Where EmpNo = " & EMPLOYEE_NO

    Call LogAudit("E", "UPDATE EMPLOYEE ELIGIBILITY RECORD", EMPLOYEE_NO)
    Call ShowSuccessFullyUpdated
    Unload Me





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    cboCSEligibility.Clear
    cboCSEligibility.AddItem "2ND GRADE/FIRST"
    cboCSEligibility.AddItem "NOT ELIGIBLE"
    cboCSEligibility.AddItem "CS PROFESSIONAL"
    cboCSEligibility.AddItem "CS SUB-PROF"
    cboCSEligibility.AddItem "R.A. 1080"
    rsRefreshStoreMemVars
End Sub

