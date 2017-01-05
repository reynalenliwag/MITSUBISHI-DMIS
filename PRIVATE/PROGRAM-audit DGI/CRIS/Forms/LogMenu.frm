VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmSMIS_Log_Menu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " LOG MENU"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "LogMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5325
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8685
      _Version        =   655364
      _ExtentX        =   15319
      _ExtentY        =   9393
      _StockProps     =   64
      AllowReorder    =   -1  'True
      Appearance      =   1
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Prospect Log Menu"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "picLogProspect"
      Item(1).Caption =   "Customer Log Menu"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "picLogCustomer"
      Begin VB.PictureBox picLogCustomer 
         BorderStyle     =   0  'None
         Height          =   4365
         Left            =   60
         ScaleHeight     =   4365
         ScaleWidth      =   8535
         TabIndex        =   18
         Top             =   390
         Width           =   8535
         Begin VB.CommandButton Command7 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "1088"
            Top             =   330
            Width           =   825
         End
         Begin VB.CommandButton Command8 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":0671
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "1088"
            Top             =   2940
            Width           =   825
         End
         Begin VB.CommandButton Command9 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":0F69
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "1088"
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton Command10 
            Height          =   780
            Left            =   120
            Picture         =   "LogMenu.frx":1791
            Style           =   1  'Graphical
            TabIndex        =   20
            Tag             =   "1088"
            Top             =   2085
            Width           =   825
         End
         Begin VB.CommandButton cmdAction 
            Height          =   780
            Index           =   39
            Left            =   3390
            Picture         =   "LogMenu.frx":1FB9
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "1196 "
            Top             =   330
            Width           =   825
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
            TabIndex        =   28
            Top             =   533
            Width           =   1950
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
            TabIndex        =   27
            Top             =   3210
            Width           =   3825
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
            TabIndex        =   26
            Top             =   1455
            Width           =   2040
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
            TabIndex        =   25
            Top             =   2340
            Width           =   4050
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
            TabIndex        =   24
            Top             =   533
            Width           =   3570
         End
      End
      Begin VB.PictureBox picLogProspect 
         BorderStyle     =   0  'None
         Height          =   4365
         Left            =   -69940
         ScaleHeight     =   4365
         ScaleWidth      =   7965
         TabIndex        =   1
         Top             =   390
         Visible         =   0   'False
         Width           =   7965
         Begin VB.CommandButton Command6 
            Height          =   780
            Left            =   3540
            Picture         =   "LogMenu.frx":2870
            Style           =   1  'Graphical
            TabIndex        =   9
            Tag             =   "1088"
            Top             =   330
            Width           =   825
         End
         Begin VB.CommandButton Command4 
            Height          =   780
            Left            =   3540
            Picture         =   "LogMenu.frx":302D
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "1088"
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton Command3 
            Height          =   780
            Left            =   3540
            Picture         =   "LogMenu.frx":3914
            Style           =   1  'Graphical
            TabIndex        =   7
            Tag             =   "1088"
            Top             =   2085
            Width           =   825
         End
         Begin VB.CommandButton CMDlOG 
            Height          =   780
            Left            =   120
            Picture         =   "LogMenu.frx":3FE3
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "1088"
            Top             =   2085
            Width           =   825
         End
         Begin VB.CommandButton Command5 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":480B
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "1088"
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton Command2 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":5033
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "1088"
            Top             =   2970
            Width           =   825
         End
         Begin VB.CommandButton Command1 
            Height          =   780
            Left            =   150
            Picture         =   "LogMenu.frx":592B
            Style           =   1  'Graphical
            TabIndex        =   3
            Tag             =   "1088"
            Top             =   330
            Width           =   825
         End
         Begin VB.CommandButton cmdAction 
            Height          =   780
            Index           =   0
            Left            =   3540
            Picture         =   "LogMenu.frx":5F90
            Style           =   1  'Graphical
            TabIndex        =   2
            Tag             =   "1196 "
            Top             =   2970
            Width           =   825
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
            TabIndex        =   17
            Top             =   540
            Width           =   2400
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
            Height          =   375
            Left            =   4500
            TabIndex        =   16
            Top             =   1440
            Width           =   3390
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
            TabIndex        =   15
            Top             =   2295
            Width           =   2535
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
            Height          =   375
            Left            =   1125
            TabIndex        =   13
            Top             =   1440
            Width           =   2040
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
            TabIndex        =   12
            Top             =   3180
            Width           =   3825
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
            TabIndex        =   11
            Top             =   540
            Width           =   1950
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
            TabIndex        =   10
            Top             =   3180
            Width           =   3570
         End
      End
   End
End
Attribute VB_Name = "frmSMIS_Log_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FormSearch          As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Private LOGACTION                      As String

Private Sub cmdAction_Click(Index As Integer)
    Call FormSearch.SearchForProspects(vbNullString)
    LOGACTION = "PROS:LOGINQ"
    FormSearch.Show 1
End Sub

Private Sub CMDlOG_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    LOGACTION = "CALL"
    FormSearch.Show 1
End Sub

Private Sub Command1_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    LOGACTION = "LETTER"
    FormSearch.Show 1
End Sub

Private Sub Command10_Click()
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:CALL"
    FormSearch.Show 1
End Sub

Private Sub Command2_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    LOGACTION = "EMAIL"
    FormSearch.Show 1
End Sub

Private Sub Command3_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    'Call FormSearch.SearchForProspects("'P','C','G','I'", "isdate(logso)=0")
    LOGACTION = "SALESAPPOINTMENT"
    FormSearch.Show 1
End Sub

Private Sub Command4_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I') AND ISDATE(LOGSO)=0")
    LOGACTION = "TESTDRIVE"
    FormSearch.Show 1
End Sub

Private Sub Command5_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    LOGACTION = "VISIT"
    FormSearch.Show 1
End Sub

Private Sub Command6_Click()
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    LOGACTION = "QUOTATION"
    FormSearch.Show 1
End Sub

Private Sub Command7_Click()
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:LETTER"
    FormSearch.Show 1
End Sub

Private Sub Command8_Click()
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:EMAIL"
    FormSearch.Show 1
End Sub

Private Sub Command9_Click()
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:VISIT"
    FormSearch.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
'    Select Case Menu
 '       Case "PROSPECT"
  '          picLogCustomer.Visible = False
   '         picLogProspect.Visible = True
    '    Case "CUSTOMER"
     '       picLogCustomer.Visible = True
      '      picLogProspect.Visible = False
    'End Select
    Set FormSearch = New frmSMIS_Mis_SearchMaster
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Menu = ""
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Unload FormSearch
    Select Case LOGACTION
        Case "CALL"
            Call frmSMIS_Log_Call.AddCall(oCusRs!ProspectID, vbNullString)
            frmSMIS_Log_Call.Show
        Case "LETTER"
            Call frmSMIS_Log_Letter.AddLetter(oCusRs!ProspectID, vbNullString)
            frmSMIS_Log_Letter.Show
        Case "EMAIL"
            Call frmSMIS_Log_Email.AddEmail(oCusRs!ProspectID, vbNullString)
            frmSMIS_Log_Email.Show
        Case "SALESAPPOINTMENT"
            frmSMIS_Log_SalesAppointment.AddSalesAppointment (oCusRs!ProspectID)
            frmSMIS_Log_SalesAppointment.Show
        Case "TESTDRIVE"
            frmSMIS_Log_TestDriveAppointment.AddTestDriveAppointment (oCusRs!ProspectID)
            frmSMIS_Log_TestDriveAppointment.Show
        Case "VISIT"
            Call frmSMIS_Log_Visits.AddVisit(oCusRs!ProspectID, vbNullString)
            frmSMIS_Log_Visits.Show
        Case "QUOTATION"
            frmSMIS_Trans_Quotation.AddNewQuotation (oCusRs!ProspectID)
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
            Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(oCusRs!ProspectID, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
    End Select

End Sub

