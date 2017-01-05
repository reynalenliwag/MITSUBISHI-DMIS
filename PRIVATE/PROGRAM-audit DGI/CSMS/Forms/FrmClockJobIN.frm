VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSClockJobIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Clock / JobClock  Login"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FrmClockJobIN.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   8100
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   50266113
      CurrentDate     =   38939
   End
   Begin VB.CommandButton cmdunloackFR 
      Height          =   375
      Left            =   7740
      Picture         =   "FrmClockJobIN.frx":0C02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3180
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   3990
      ScaleHeight     =   3030
      ScaleWidth      =   6465
      TabIndex        =   12
      Top             =   990
      Width           =   6465
      Begin VB.TextBox tlHrs 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5220
         TabIndex        =   26
         Top             =   1710
         Width           =   1005
      End
      Begin VB.TextBox txtStdHr 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5190
         TabIndex        =   21
         Text            =   "txtStdHr"
         Top             =   3120
         Width           =   1005
      End
      Begin VB.TextBox txtJobCode 
         Height          =   315
         Left            =   5220
         TabIndex        =   20
         Text            =   "txtJobCode"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1470
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   780
         Width           =   4755
      End
      Begin VB.TextBox txtAccessCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1470
         TabIndex        =   0
         Top             =   2580
         Width           =   2205
      End
      Begin VB.TextBox txtRO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2115
         Width           =   2205
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clock  In/Out (AM)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         TabIndex        =   14
         Top             =   30
         Width           =   3015
         Begin VB.CheckBox chkTimeInAM 
            Caption         =   "Clock IN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   1
            Top             =   300
            Width           =   1155
         End
         Begin VB.CheckBox chkTimeOutAM 
            Caption         =   "Clock OUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1500
            TabIndex        =   2
            Top             =   330
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Time In/Out (PM)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3210
         TabIndex        =   13
         Top             =   30
         Width           =   3045
         Begin VB.CheckBox chkTimeOutPM 
            Caption         =   "Clock OUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1530
            TabIndex        =   4
            Top             =   330
            Width           =   1425
         End
         Begin VB.CheckBox chkTimeInPM 
            Caption         =   "Clock IN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   3
            Top             =   300
            Width           =   1245
         End
      End
      Begin MSComCtl2.DTPicker dtPromised 
         Height          =   375
         Left            =   1470
         TabIndex        =   28
         Top             =   1710
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
         Format          =   50266115
         CurrentDate     =   38936
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Promise Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   27
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Std Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4410
         TabIndex        =   23
         Top             =   1740
         Width           =   885
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4350
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         TabIndex        =   16
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         TabIndex        =   15
         Top             =   2220
         Width           =   1245
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9600
      Top             =   450
   End
   Begin MSComctlLib.ListView lblJob4Service 
      Height          =   1890
      Left            =   120
      TabIndex        =   17
      Top             =   4905
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmClockJobIN.frx":118C
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Technician"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "R/O"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Clock In am"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Clock Out am"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Clock In pm"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Clock Out pm"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Hour(s)"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lblTech 
      Height          =   4740
      Left            =   120
      TabIndex        =   25
      Top             =   90
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   8361
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmClockJobIN.frx":12EE
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " Technician"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
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
      Height          =   795
      Left            =   9780
      MouseIcon       =   "FrmClockJobIN.frx":1450
      MousePointer    =   99  'Custom
      Picture         =   "FrmClockJobIN.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cancel"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9060
      MouseIcon       =   "FrmClockJobIN.frx":18E0
      MousePointer    =   99  'Custom
      Picture         =   "FrmClockJobIN.frx":1A32
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ok"
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label labtime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   8520
      TabIndex        =   11
      Top             =   60
      Width           =   2025
   End
   Begin VB.Label labdate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   8520
      TabIndex        =   10
      Top             =   450
      Width           =   2055
   End
   Begin VB.Label lblIS_TECH_STATUS 
      BackStyle       =   0  'Transparent
      Caption         =   "Clock In"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Left            =   6570
      TabIndex        =   9
      Top             =   540
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Clock"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   5430
      TabIndex        =   7
      Top             =   -30
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   4320
      Picture         =   "FrmClockJobIN.frx":1CCD
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
End
Attribute VB_Name = "frmCSMSClockJobIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xhour                                                             As String
Dim xampm                                                             As String
Dim rsSearch                                                          As ADODB.Recordset
Dim rsPas                                                             As ADODB.Recordset
Dim xTechName                                                         As String
Dim xteccode                                                          As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Function SeeEmpNo(XXX As Variant)
    Dim rsSeePass                                                     As ADODB.Recordset
    Set rsSeePass = New ADODB.Recordset
    Set rsSeePass = gconDMIS.Execute("Select empno,Tech_Name from [CSMS_vw_Technician] where [empno] = '" & XXX & "'")
    If Not rsSeePass.EOF And Not rsSeePass.BOF Then
        SeeEmpNo = Null2String(rsSeePass![empno])
        xteccode = Null2String(rsSeePass![empno])
        xTechName = Null2String(rsSeePass![Tech_Name])
    End If
    Set rsSeePass = Nothing
End Function

Private Sub cmdOk_Click()

    If Trim(txtAccessCode) <> Trim(SeeEmpNo(UCase(txtAccessCode))) Then
        MsgBox "Invalid Technician Code!"
        Exit Sub
    End If
    Dim xRO_No                                                        As String
    Dim xTechnician, xTranDate, xTime_In_Am, xTime_Out_Am, xTime_In_Pm, xTime_Out_Pm As String
    xRO_No = N2Str2Null(txtro)
    Set rsSearch = New ADODB.Recordset
    Set rsSearch = gconDMIS.Execute("select Technician,[Tech_Name], time_In_Am, Time_Out_Am, Time_In_Pm, Time_Out_Pm from CSMS_JobClock Where Technician = '" & txtAccessCode & "' and TranDate = '" & CDate(labdate.Caption) & "' and RO_No = '" & txtro & "'")
    If Not (rsSearch.EOF And rsSearch.BOF) Then
        If IsNull(rsSearch![Time_out_am]) = True Then
            xTime_Out_Am = N2Str2Null(Now)
            If chkTimeOutAM.Value = 1 Then
                gconDMIS.Execute "update CSMS_JobClock set [STATUS]='Available',Time_Out_Am = " & xTime_Out_Am & " Where Technician = '" & txtAccessCode & "' and trandate = '" & CDate(labdate.Caption) & "'"
                gconDMIS.Execute "update HRMS_EmpInfo set  [IS_TECH_STATUS] = 'Available' where empno = '" & Trim(xteccode) & "'"
            Else
                MsgBox "Time-out failed... please check appropriate check box..."
                Exit Sub
            End If
            frmCSMSFinishorTobeCont.labTech.Caption = xTechName
            frmCSMSFinishorTobeCont.Show 1
        ElseIf IsNull(rsSearch![Time_in_pm]) = True Then
            xTime_In_Pm = N2Str2Null(Now)
            If chkTimeInPM.Value = 1 Then
                gconDMIS.Execute "update CSMS_JobClock set [STATUS]='Working',Time_In_Pm = " & xTime_In_Pm & " Where Technician = '" & txtAccessCode & "' and trandate = '" & CDate(labdate.Caption) & "'"
                gconDMIS.Execute "update HRMS_EmpInfo set  [IS_TECH_STATUS] = 'Working' where empno = " & xteccode & ""
            Else
                MsgBox "Time-out failed... please check appropriate check box..."
                Exit Sub
            End If
        ElseIf IsNull(rsSearch![Time_out_pm]) = True Then
            xTime_Out_Pm = N2Str2Null(Now)
            If chkTimeOutPM.Value = 1 Then
                gconDMIS.Execute "update CSMS_JobClock set [STATUS]='Available',Time_Out_Pm = " & xTime_Out_Pm & " Where Technician = '" & txtAccessCode & "' and trandate = '" & CDate(labdate.Caption) & "'"
                gconDMIS.Execute "update HRMS_EmpInfo set  [IS_TECH_STATUS] = 'Available' where empno = " & xteccode & ""
            Else
                MsgBox "Time-out failed... please check appropriate check box..."
                Exit Sub
            End If
            frmCSMSFinishorTobeCont.labTech.Caption = xTechName
            frmCSMSFinishorTobeCont.Show 1
        End If
    Else
        chkTimeInAM.Value = 1
        If chkTimeInAM.Value = 1 Then
            xTechnician = N2Str2Null(txtAccessCode)
            xTranDate = Null2String(DTPicker1)
            xTime_In_Am = N2Str2Null(Now)
            xTime_Out_Am = N2Str2Null("")
            xTime_In_Pm = N2Str2Null("")
            xTime_Out_Pm = N2Str2Null("")
            gconDMIS.Execute "Insert into CSMS_JobClock ([STATUS],stdhrs,DETCDE,RO_No,[Tech_Name],Technician, TranDate, Time_In_Am, Time_Out_Am, Time_In_Pm, Time_Out_Pm) " & _
                           " values ('Working'," & Val(txtStdHr) & ",'" & txtJobCode & "'," & xRO_No & ",'" & xTechName & "'," & xTechnician & ",'" & Format(DTPicker1, "MM/dd/yyyy") & "'," & xTime_In_Am & "," & xTime_Out_Am & "," & xTime_In_Pm & "," & xTime_Out_Pm & ")"
            gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Working',PromiseDate= '" & dtPromised & "' where RO_No = '" & txtro & "'"
            gconDMIS.Execute "update [HRMS_EmpInfo] set  [IS_TECH_STATUS] = 'Working' where [EmpNo] = '" & Trim(xteccode) & "'"
            gconDMIS.Execute "update CSMS_RO_Det set  TECHNICIAN = '" & xTechName & "',TECHCODE='" & xteccode & "' where REP_OR = '" & txtro & "' and DETCDE = '" & txtJobCode & "'"

        End If
    End If
    txtAccessCode = ""
    Unload Me
End Sub



Private Sub Form_Activate()
    Dim rsUpload                                                      As ADODB.Recordset
    lblTech.Sorted = False: lblTech.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select EmpNo,Tech_Name,IS_TECH_STATUS from [CSMS_vw_Technician] Order by [IS_TECH_STATUS] Asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lblTech.ListItems, rsUpload
    End If
    ViewClockIn
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " " & App.Major & App.Minor & App.Revision
    labtime.Caption = Format(Time, "hh:mm:ss ampm")
    labdate.Caption = Format(Now, "MM/dd/yyyy")
    xhour = Left(Trim(labtime.Caption), 2)
    xampm = Right(Trim(labtime.Caption), 2)
    If Val(xhour) >= 6 And Val(xhour) <= 10 Then
        chkTimeInAM.Value = 1
        chkTimeOutAM.Value = 0
        chkTimeInPM.Value = 0
        chkTimeOutPM.Value = 0
    ElseIf Val(xhour) >= 11 And Val(xhour) <= 12 Then
        chkTimeInAM.Value = 0
        chkTimeOutAM.Value = 1
        chkTimeInPM.Value = 0
        chkTimeOutPM.Value = 0
    ElseIf Val(xhour) >= 1 And Val(xhour) <= 2 Then
        chkTimeInAM.Value = 0
        chkTimeOutAM.Value = 0
        chkTimeInPM.Value = 1
        chkTimeOutPM.Value = 0
    ElseIf Val(xhour) >= 3 And Val(xhour) <= 6 Then
        chkTimeInAM.Value = 0
        chkTimeOutAM.Value = 0
        chkTimeInPM.Value = 0
        chkTimeOutPM.Value = 1
    End If

End Sub
Sub ViewClockIn()
    Dim rsUpload                                                      As ADODB.Recordset
    lblJob4Service.Sorted = False: lblJob4Service.ListItems.Clear
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select Technician,Tech_Name,RO_No,Time_in_am,Time_out_am,Time_in_pm,Time_out_pm,NumHrs from CSMS_JobClock where RO_No = '" & txtro & "'")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lblJob4Service.ListItems, rsUpload
    End If
End Sub

Private Sub lblJob4Service_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAccessCode = lblJob4Service.SelectedItem
End Sub


Private Sub lblTech_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAccessCode = lblTech.SelectedItem
End Sub

Private Sub Timer1_Timer()
    labtime.Caption = Format(Time, "hh:mm:ss ampm")
    labdate.Caption = Format(Now, "MM/dd/yyyy")
    DTPicker1.Value = Format(Now, "MM/dd/yyyy")
End Sub
Private Sub chkTimeInAM_Click()
    If chkTimeInAM.Value = 1 Then
        chkTimeOutAM.Value = 0
        chkTimeInPM.Value = 0
        chkTimeOutPM.Value = 0
        lblIS_TECH_STATUS.Caption = "Clock In"
    End If
End Sub
Private Sub chkTimeOutAM_Click()
    If chkTimeOutAM.Value = 1 Then
        chkTimeInAM.Value = 0
        chkTimeInPM.Value = 0
        chkTimeOutPM.Value = 0
        lblIS_TECH_STATUS.Caption = "Clock Out"
    End If
End Sub
Private Sub chkTimeInPM_Click()
    If chkTimeInPM.Value = 1 Then
        chkTimeInAM.Value = 0
        chkTimeOutAM.Value = 0
        chkTimeOutPM.Value = 0
        lblIS_TECH_STATUS.Caption = "Clock In"
    End If
End Sub
Private Sub chkTimeOutPM_Click()
    If chkTimeOutPM.Value = 1 Then
        chkTimeInAM.Value = 0
        chkTimeOutAM.Value = 0
        chkTimeInPM.Value = 0
        lblIS_TECH_STATUS.Caption = "Clock Out"
    End If
End Sub
Private Sub txtAccessCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOk.Value = True
    End If
End Sub
