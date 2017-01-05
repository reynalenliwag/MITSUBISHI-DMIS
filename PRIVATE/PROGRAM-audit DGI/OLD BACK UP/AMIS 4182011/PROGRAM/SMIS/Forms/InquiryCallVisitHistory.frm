VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmSMIS_Inquiry_CallVisit_History 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Customer Calls and Visit History inquiry"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "InquiryCallVisitHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin VB.CommandButton Command1 
      Caption         =   "INQUIRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10800
      TabIndex        =   14
      ToolTipText     =   "Inquire"
      Top             =   2760
      Width           =   1305
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8880
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5700
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2790
      Width           =   2355
   End
   Begin MSComctlLib.ListView lvCustomers 
      Height          =   1935
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "InquiryCallVisitHistory.frx":058A
      NumItems        =   0
   End
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   4245
      Left            =   30
      TabIndex        =   1
      Top             =   3150
      Width           =   12135
      _Version        =   655364
      _ExtentX        =   21405
      _ExtentY        =   7488
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   150
      PaintManager.MinTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Call History"
      Item(0).Tooltip =   "Call History"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lvCalls"
      Item(1).Caption =   "Visit"
      Item(1).Tooltip =   "Visit"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lvVisits"
      Item(2).Caption =   "Letter"
      Item(2).Tooltip =   "Letter"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lvLetter"
      Item(3).Caption =   "Email"
      Item(3).Tooltip =   "Email"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lvEmails"
      Item(4).Caption =   "Reminders"
      Item(4).Tooltip =   "Reminders"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lvReminder"
      Begin MSComctlLib.ListView lvCalls 
         Height          =   3705
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   6535
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "InquiryCallVisitHistory.frx":06EC
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvVisits 
         Height          =   3735
         Left            =   -69940
         TabIndex        =   8
         Top             =   390
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "InquiryCallVisitHistory.frx":084E
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvLetter 
         Height          =   3765
         Left            =   -69970
         TabIndex        =   9
         Top             =   390
         Visible         =   0   'False
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   6641
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "InquiryCallVisitHistory.frx":09B0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvEmails 
         Height          =   3765
         Left            =   -69970
         TabIndex        =   10
         Top             =   390
         Visible         =   0   'False
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   6641
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "InquiryCallVisitHistory.frx":0B12
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvReminder 
         Height          =   3765
         Left            =   -69970
         TabIndex        =   11
         Top             =   390
         Visible         =   0   'False
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   6641
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "InquiryCallVisitHistory.frx":0C74
         NumItems        =   0
      End
   End
   Begin VB.Frame fraAllCustomer 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   12105
      Begin VB.OptionButton Otp 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Contact Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2130
         TabIndex        =   6
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Cuscde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4110
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox txtSearchKey_All 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5370
         TabIndex        =   4
         Top             =   210
         Width           =   3705
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4950
      TabIndex        =   16
      Top             =   2850
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8130
      TabIndex        =   15
      Top             =   2820
      Width           =   2055
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_CallVisit_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim TheCuscde                                                         As String

Sub Fill_AllCustomerSearch()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim Keyword                                                       As String
    On Error GoTo ErrorCode:

    lvCustomers.Enabled = False

    SQL = "SELECT TOP 100 CUSCDE, CUSTOMERNAME, ADDRESS, EMAIL, PHONE, MOBILE FROM CRIS_vw_AllProfile"

    Keyword = RTrim(LTrim(txtSearchKey_All.Text))

    If Otp(0).Value = True Then
        SQL = SQL & " WHERE CUSTOMERNAME LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(1).Value = True Then
        SQL = SQL & " WHERE CONTACTPERSON LIKE'" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(2).Value = True Then
        SQL = SQL & " WHERE cuscde like '" & ReplaceQuote(Keyword) & "%'"
    End If

    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvCustomers.Enabled = True
    End If

    flex_FillListView RS, lvCustomers





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillCalls()
    Dim RS                                                            As ADODB.Recordset
    Dim SQL                                                           As String
    Dim MANTH
    On Error GoTo ErrorCode:

    lvCalls.Enabled = False

    MANTH = What_month(Combo1.Text)
    If Combo1.Text <> "ALL" Then
        SQL = " where  YEAR(DateTimeCall)=" & Combo2.Text & " AND  MONTH(DateTimeCall)=" & MANTH & " AND CSCDE=" & N2Str2Null(TheCuscde)
    Else
        SQL = " where  YEAR(DateTimeCall)=" & Combo2.Text & " AND  CSCDE=" & N2Str2Null(TheCuscde)
    End If
    SQL = "Select  Convert(varchar,DateTimeCall,101) as Date , Convert(varchar,DateTimeCall,108) as Time , Duration, Subject, Comments, Bound, CalledBy, PhoneNo from CRIS_Prospect_Calls " & SQL



    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvCalls.Enabled = True
    End If


    Listview_Loadval lvCalls.ListItems, RS



    Set RS = Nothing





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillEmail()
    Dim RS                                                            As ADODB.Recordset
    Dim SQL                                                           As String
    Dim MANTH
    On Error GoTo ErrorCode:

    lvEmails.Enabled = False

    MANTH = What_month(Combo1.Text)
    If Combo1.Text <> "ALL" Then
        SQL = " where  YEAR(DateEmail)=" & Combo2.Text & " AND  MONTH(DateEmail)=" & MANTH & " AND CSCDE=" & N2Str2Null(TheCuscde)
    Else
        SQL = " where  YEAR(DateEmail)=" & Combo2.Text & " AND  CSCDE=" & N2Str2Null(TheCuscde)
    End If
    SQL = "Select  Convert(varchar,DateEmail,101) as Date , EmailFrom, EmailTO, Subject, Bound from CRIS_Prospect_Email " & SQL
    Set RS = gconDMIS.Execute(SQL)
    flex_FillListView RS, lvEmails

    If Not RS.EOF And Not RS.BOF Then
        lvEmails.Enabled = True
    End If

    Set RS = Nothing





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillLetter()
    Dim RS                                                            As ADODB.Recordset
    Dim SQL                                                           As String
    Dim MANTH
    On Error GoTo ErrorCode:

    lvLetter.Enabled = False


    MANTH = What_month(Combo1.Text)
    If Combo1.Text <> "ALL" Then
        SQL = " WHERE YEAR(DateLetter)=" & Combo2.Text & " AND  MONTH(DateLetter)=" & MANTH & " AND CSCDE=" & N2Str2Null(TheCuscde)
    Else
        SQL = " WHERE YEAR(DateLetter)=" & Combo2.Text & " AND  CSCDE=" & N2Str2Null(TheCuscde)
    End If
    SQL = "Select  Convert(varchar,DateLetter,101) as Date , LetterFrom , LetterTo, Subject, Bound  from CRIS_Prospect_Letter " & SQL
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvLetter.Enabled = True
    End If


    flex_FillListView RS, lvLetter




    Set RS = Nothing





    Exit Sub
ErrorCode:
    ShowVBError
End Sub
'REMINDERS

Sub FillReminders()
    Dim RS                                                            As ADODB.Recordset
    Dim SQL                                                           As String
    Dim MANTH
    Dim WhereStateMent                                                As String
    On Error GoTo ErrorCode:

    lvReminder.Enabled = False

    MANTH = What_month(Combo1.Text)
    If Combo1.Text <> "ALL" Then
        WhereStateMent = " WHERE YEAR(DateTimeRemind)=" & Combo2.Text & " AND  MONTH(DateTimeRemind)=" & MANTH & " AND CSCDE=" & N2Str2Null(TheCuscde)
    Else
        WhereStateMent = " WHERE YEAR(DateTimeRemind)=" & Combo2.Text & " AND  CSCDE=" & N2Str2Null(TheCuscde)
    End If


    'SELECT ReminderType, EntityType, DateTimeRemind, ReminderNotes, Subject, Snoozed, ID, NextTime, LOGID FROM DMIS.dbo.CRIS_Reminders
    SQL = "Select  Convert(varchar,DateTimeRemind,101) as Date , " _
        & " CASE when DateTimeRemind<=getdate() then Cast( datediff(day, DateTimeRemind,getdate()) as varchar)  + ' Day' Else '0 Days' end ,  " _
        & " ReminderType, " _
        & " Subject  " _
        & " FROM CRIS_Reminders " & WhereStateMent


    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvReminder.Enabled = True
    End If



    Listview_Loadval lvReminder.ListItems, RS



    Set RS = Nothing






    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Sub FillVisit()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim MANTH
    On Error GoTo ErrorCode:

    lvVisits.Enabled = False

    MANTH = What_month(Combo1.Text)


    If Combo1.Text <> "ALL" Then
        SQL = " WHERE YEAR(DateTimeVisit)=" & Combo2.Text & " AND  MONTH(DateTimeVisit)=" & MANTH & " AND CSCDE=" & N2Str2Null(TheCuscde)
    Else
        SQL = " WHERE YEAR(DateTimeVisit)=" & Combo2.Text & " AND  CSCDE=" & N2Str2Null(TheCuscde)
    End If
    SQL = "Select Convert(varchar,DateTimeVisit,101) as Date ,Comments,Results from CRIS_Prospect_Visits " & SQL
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvVisits.Enabled = True
    End If


    flex_FillListView RS, lvVisits



    Set RS = Nothing





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub fillcbos()
    fillcbomonth Combo1
    Combo1.AddItem "ALL", 0
    Combo1.Text = "ALL"
    FillCboMoreYear Combo2
    
    'fillcombo_up Combo2
End Sub

Sub FillData()
    If TabControl.SelectedItem = 0 Then
        Call FillCalls
    ElseIf TabControl.SelectedItem = 1 Then
        Call FillVisit
    ElseIf TabControl.SelectedItem = 2 Then
        Call FillLetter
    ElseIf TabControl.SelectedItem = 3 Then
        Call FillEmail
    ElseIf TabControl.SelectedItem = 4 Then
        Call FillReminders
    End If


End Sub

Private Sub CmdAll_Click()
    Fill_AllCustomerSearch
    If lvCustomers.ListItems.Count > 0 Then
        lvCustomers.ListItems(1).Selected = True
        lvCustomers.ListItems(1).EnsureVisible
        lvCustomers_ItemClick lvCustomers.SelectedItem

    Else
        lvCalls.ListItems.Clear
        lvVisits.ListItems.Clear
    End If
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub Command1_Click()
    FillData
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    'CUSCDE, CUSTOMERNAME, ADDRESS, CONTACTPERSON, EMAIL, PHONE, MOBILE
    AddColumnHeader " CODE, NAME, ADDRESS, EMAIL, PHONE, MOBILE ", lvCustomers
    ResizeColumnHeader lvCustomers, "8,25,30,12,12,12"
    'CALLS
    'Date , Time, Bound, Duration , Subject, Comment, CalledBy , Phone Number
    AddColumnHeader "Date , Time, Duration  , Subject ,Comment, Bound, CalledBy , Phone Number", lvCalls
    ResizeColumnHeader lvCalls, "8,8,8,20,20,20,10,10"
    'VISITS
    '"Date ,Comments,Results"
    AddColumnHeader "Date ,Comments,Results", lvVisits
    ResizeColumnHeader lvVisits, "10,40,40"

    'LETTERS
    'DateLetter, LetterFrom , LetterTo, Subject, Bound
    AddColumnHeader "Date , From , To , Subject, Bound", lvLetter
    ResizeColumnHeader lvLetter, "10,20,20,37,10"

    'EMAILS
    'Convert(varchar,DateEmail,101) as Date , EmailFrom, EmailTO, Subject, Bound
    AddColumnHeader "Date ,EmailFrom , EmailTo , Subject, Bound ", lvEmails
    ResizeColumnHeader lvEmails, "10,20,20,28,20"

    'REMINDERS
    'DateTimeRemind, ReminderType, Subject, NextTime FROM DMIS.dbo.CRIS_Reminders
    AddColumnHeader "Date ,OverDue , ReminderType,Subject   ", lvReminder
    ResizeColumnHeader lvReminder, "10,10,40,38"
    fillcbos
    CmdAll_Click
End Sub

Private Sub lvCustomers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TheCuscde = lvCustomers.SelectedItem.Text
    FillData
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    FillData
End Sub

Private Sub Otp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub txtSearchKey_All_Change()
    Fill_AllCustomerSearch
    If lvCustomers.ListItems.Count > 0 Then
        lvCustomers.ListItems(1).Selected = True
        lvCustomers.ListItems(1).EnsureVisible
        lvCustomers_ItemClick lvCustomers.SelectedItem
    Else
        lvCalls.ListItems.Clear
        lvVisits.ListItems.Clear
    End If
End Sub

