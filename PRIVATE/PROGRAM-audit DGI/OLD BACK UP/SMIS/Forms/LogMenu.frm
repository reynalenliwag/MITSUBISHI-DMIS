VERSION 5.00
Begin VB.Form frmSMIS_Log_Menu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " LOG MENU"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "LogMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLogProspect 
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   0
      ScaleHeight     =   4665
      ScaleWidth      =   7965
      TabIndex        =   0
      Top             =   -150
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CommandButton cmdInquiryProspect 
         Height          =   780
         Left            =   3540
         Picture         =   "LogMenu.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "1196 "
         ToolTipText     =   "PROSPECT LOG INQUIRY"
         Top             =   2910
         Width           =   825
      End
      Begin VB.CommandButton Command1 
         Height          =   780
         Left            =   150
         Picture         =   "LogMenu.frx":08C3
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "1088"
         ToolTipText     =   "LOG A LETTER"
         Top             =   330
         Width           =   825
      End
      Begin VB.CommandButton Command2 
         Height          =   780
         Left            =   150
         Picture         =   "LogMenu.frx":0F28
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "1088"
         ToolTipText     =   "LOG A EMAIL"
         Top             =   2970
         Width           =   825
      End
      Begin VB.CommandButton Command5 
         Height          =   780
         Left            =   150
         Picture         =   "LogMenu.frx":1820
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "1088"
         ToolTipText     =   "LOG VISIT"
         Top             =   1230
         Width           =   825
      End
      Begin VB.CommandButton CMDlOG 
         Height          =   780
         Left            =   120
         Picture         =   "LogMenu.frx":2048
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1088"
         ToolTipText     =   "LOG A CALL"
         Top             =   2085
         Width           =   825
      End
      Begin VB.CommandButton Command3 
         Height          =   780
         Left            =   3540
         Picture         =   "LogMenu.frx":2870
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1088"
         ToolTipText     =   "SALES APPOINTMENT"
         Top             =   2085
         Width           =   825
      End
      Begin VB.CommandButton Command4 
         Height          =   780
         Left            =   3540
         Picture         =   "LogMenu.frx":2F3F
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1088"
         ToolTipText     =   "TEST DRIVE"
         Top             =   1230
         Width           =   825
      End
      Begin VB.CommandButton Command6 
         Height          =   780
         Left            =   3540
         Picture         =   "LogMenu.frx":3826
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "1088"
         ToolTipText     =   "QUOTATION"
         Top             =   330
         Width           =   825
      End
      Begin VB.CommandButton Command12 
         Height          =   780
         Left            =   150
         Picture         =   "LogMenu.frx":3FE3
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "1088"
         ToolTipText     =   "LOG A CALL"
         Top             =   3840
         Width           =   825
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PROSPECT LOG INQUIRY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4500
         TabIndex        =   18
         Top             =   3180
         Width           =   3570
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A LETTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1125
         TabIndex        =   17
         Top             =   533
         Width           =   1950
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A EMAIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1110
         TabIndex        =   16
         Top             =   3180
         Width           =   3825
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG VISIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1125
         TabIndex        =   15
         Top             =   1440
         Width           =   2040
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A CALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1095
         TabIndex        =   14
         Top             =   2295
         Width           =   2280
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES APPOINTMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4500
         TabIndex        =   13
         Top             =   2295
         Width           =   2535
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "TEST DRIVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4500
         TabIndex        =   12
         Top             =   1440
         Width           =   3390
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "QUOTATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4500
         TabIndex        =   11
         Top             =   533
         Width           =   2400
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "PROSPECT REMINDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1170
         TabIndex        =   10
         Top             =   4095
         Width           =   4050
      End
   End
   Begin VB.PictureBox picLogCustomer 
      BorderStyle     =   0  'None
      Height          =   4635
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   8535
      TabIndex        =   19
      Top             =   -150
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdInquiryCustomer 
         Height          =   780
         Left            =   3390
         Picture         =   "LogMenu.frx":480B
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "1196 "
         ToolTipText     =   "CUSTOMER LOG INQUIRY"
         Top             =   390
         Width           =   825
      End
      Begin VB.CommandButton Command10 
         Height          =   780
         Left            =   180
         Picture         =   "LogMenu.frx":50C2
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "1088"
         ToolTipText     =   "LOG A CALL"
         Top             =   2085
         Width           =   825
      End
      Begin VB.CommandButton Command9 
         Height          =   780
         Left            =   180
         Picture         =   "LogMenu.frx":58EA
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "1088"
         ToolTipText     =   "LOG VISIT"
         Top             =   1230
         Width           =   825
      End
      Begin VB.CommandButton Command8 
         Height          =   780
         Left            =   180
         Picture         =   "LogMenu.frx":6112
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "1088"
         ToolTipText     =   "LOG A EMAIL"
         Top             =   2940
         Width           =   825
      End
      Begin VB.CommandButton Command7 
         Height          =   780
         Left            =   180
         Picture         =   "LogMenu.frx":6A0A
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "1088"
         ToolTipText     =   "LOG A LETTER"
         Top             =   330
         Width           =   825
      End
      Begin VB.CommandButton Command11 
         Height          =   780
         Left            =   180
         Picture         =   "LogMenu.frx":706F
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "1088"
         ToolTipText     =   "LOG A CALL"
         Top             =   3780
         Width           =   825
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER LOG INQUIRY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4350
         TabIndex        =   31
         Top             =   600
         Width           =   3570
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A CALL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1140
         TabIndex        =   30
         Top             =   2340
         Width           =   4050
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG VISIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1125
         TabIndex        =   29
         Top             =   1455
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A EMAIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1125
         TabIndex        =   28
         Top             =   3210
         Width           =   3825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LOG A LETTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1125
         TabIndex        =   27
         Top             =   533
         Width           =   1950
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER REMINDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         TabIndex        =   26
         Top             =   4035
         Width           =   4050
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents FormSearch                                         As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Private LOGACTION                                                     As String

Private Sub cmdInquiryCustomer_Click()
    On Error GoTo ErrorCode:
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:LOGINQ"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdInquiryProspect_Click()
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects(vbNullString)
    Else
        Call FormSearch.SearchForProspects(" USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "PROS:LOGINQ"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub CMDlOG_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "CALL"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "LETTER"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command10_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:CALL"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command11_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    frmSMIS_Log_CustomerReminder.Show
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Command12_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    frmSMIS_Log_ProspectReminder.Show
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "EMAIL"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0 AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "SALESAPPOINTMENT"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command4_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0 AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "TESTDRIVE"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command5_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If LOGSAE = "" Then
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    Else
        Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND USERCODE='" & LOGSAE & "'")
    End If
    LOGACTION = "VISIT"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command6_Click()
    If Module_Access(LOGID, "QUOTATION", "TRANSACTION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    On Error Resume Next
    frmSMIS_Trans_Quotation.Show
    '    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    '   LOGACTION = "QUOTATION"
    '  FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:LETTER"
    FormSearch.Show 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command8_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:EMAIL"
    FormSearch.Show 1





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command9_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:VISIT"
    FormSearch.Show 1





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set FormSearch = New frmSMIS_Mis_SearchMaster
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Unload FormSearch
    Select Case LOGACTION
        Case "CALL"
            Call frmSMIS_Log_Call.AddCall(oCusRs!PROSPECTID, vbNullString)
            frmSMIS_Log_Call.Show
        Case "LETTER"
            Call frmSMIS_Log_Letter.AddLetter(oCusRs!PROSPECTID, vbNullString)
            frmSMIS_Log_Letter.Show
        Case "EMAIL"
            Call frmSMIS_Log_Email.AddEmail(oCusRs!PROSPECTID, vbNullString)
            frmSMIS_Log_Email.Show
        Case "SALESAPPOINTMENT"
            frmSMIS_Log_SalesAppointment.AddSalesAppointment (oCusRs!PROSPECTID)
            frmSMIS_Log_SalesAppointment.Show
        Case "TESTDRIVE"
            frmSMIS_Log_TestDriveAppointment.AddTestDriveAppointment (oCusRs!PROSPECTID)
            frmSMIS_Log_TestDriveAppointment.Show
        Case "VISIT"
            Call frmSMIS_Log_Visits.AddVisit(oCusRs!PROSPECTID, vbNullString)
            frmSMIS_Log_Visits.Show
        Case "QUOTATION"
            frmSMIS_Trans_Quotation.AddNewQuotation (oCusRs!PROSPECTID)
            frmSMIS_Trans_Quotation.Show
        Case "CUS:LETTER"
            Call frmSMIS_Log_Letter.AddLetter(0, oCusRs!CUSCDE)
            frmSMIS_Log_Letter.Show
        Case "CUS:VISIT"
            Call frmSMIS_Log_Visits.AddVisit(0, oCusRs!CUSCDE)
            frmSMIS_Log_Visits.Show
        Case "CUS:CALL"
            Call frmSMIS_Log_Call.AddCall(0, oCusRs!CUSCDE)
            frmSMIS_Log_Call.Show
        Case "CUS:EMAIL"
            Call frmSMIS_Log_Email.AddEmail(0, oCusRs!CUSCDE)
            frmSMIS_Log_Email.Show
        Case "CUS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWCUSTOMERLOG(oCusRs!CUSCDE, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
        Case "PROS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(oCusRs!PROSPECTID, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
    End Select

End Sub

