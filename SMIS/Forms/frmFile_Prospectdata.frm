VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmFile_Prospectdata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prospect Data Information "
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFile_Prospectdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport ProspectSheet 
      Left            =   12930
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   6900
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.PictureBox picCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   4710
         ScaleHeight     =   1695
         ScaleWidth      =   3705
         TabIndex        =   10
         Top             =   2250
         Width           =   3735
         Begin VB.ComboBox cboyear 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmFile_Prospectdata.frx":08CA
            Left            =   840
            List            =   "frmFile_Prospectdata.frx":08E6
            TabIndex        =   19
            Top             =   900
            Width           =   2775
         End
         Begin VB.ComboBox cbomonth 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmFile_Prospectdata.frx":091A
            Left            =   840
            List            =   "frmFile_Prospectdata.frx":091C
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   540
            Width           =   2775
         End
         Begin wizButton.cmd CmdCost 
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   1260
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            TX              =   "Ok"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "frmFile_Prospectdata.frx":091E
         End
         Begin wizButton.cmd CmdCostCancel 
            Height          =   375
            Left            =   2490
            TabIndex        =   12
            Top             =   1260
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   661
            TX              =   "Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "frmFile_Prospectdata.frx":093A
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   465
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Range"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   300
            TabIndex        =   15
            Top             =   540
            Width           =   705
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   345
            Left            =   -30
            TabIndex        =   14
            Top             =   0
            Width           =   3705
            _Version        =   655364
            _ExtentX        =   6535
            _ExtentY        =   609
            _StockProps     =   14
            ForeColor       =   16579836
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
            ForeColor       =   16579836
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Cost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   60
            Width           =   2205
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Prospect Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   60
         TabIndex        =   7
         Top             =   660
         Width           =   12975
         Begin MSComctlLib.ListView listProspect 
            Height          =   5595
            Left            =   60
            TabIndex        =   8
            Top             =   300
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   9869
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Prospect Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Model"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Description"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Color"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Clasification"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Lead Source"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Email"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Contact Person"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Address"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Tel No"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Mobile"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Status"
               Object.Width           =   2293
            EndProperty
         End
      End
      Begin VB.CommandButton CmdView 
         Appearance      =   0  'Flat
         Caption         =   "View"
         Height          =   375
         Left            =   11790
         TabIndex        =   6
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton CmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "Print"
         Height          =   375
         Left            =   10530
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
      Begin VB.ComboBox cbostatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFile_Prospectdata.frx":0956
         Left            =   5220
         List            =   "frmFile_Prospectdata.frx":0958
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   2415
      End
      Begin VB.ComboBox CboSA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   180
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4620
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "SA Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Information: To update the status of prospects it can be update to Sales diary or in Sales Monitoring"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   20
      Top             =   6870
      Width           =   9615
   End
End
Attribute VB_Name = "frmFile_Prospectdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theStatus                                                         As String
Dim theMonth                                                          As Integer

Sub FillSAE()
    Dim RsSAE                                                         As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim MIDDLE                                                        As String
    Dim X                                                             As String

    SQL = "SELECT lname,fname,middle FROM Smis_SalesTeam"


    Set RsSAE = New ADODB.Recordset
    Set RsSAE = gconDMIS.Execute(SQL)

    cboSA.Clear

    Do While Not RsSAE.EOF
        MIDDLE = Null2String(RsSAE!MIDDLE)
        X = Mid(MIDDLE, 1, 1)
        cboSA.AddItem Null2String(RsSAE!lname) + "," + Null2String(RsSAE!fname) + "." + X

        RsSAE.MoveNext
    Loop
    Set RsSAE = Nothing

End Sub

Sub FillStatus()
    cbostatus.Clear
    cbostatus.AddItem "OPEN"
    cbostatus.AddItem "ClOSE"
    cbostatus.AddItem "INACTIVE"
    cbostatus.AddItem "LOST SALES"
    cbostatus.AddItem "All"
End Sub

Sub DisplayInformation()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim Item                                                          As ListItem
    Dim cnt                                                           As Integer
    Dim theSA                                                         As String
    Dim Mstatus                                                       As String

    theSA = Trim(cboSA.Text)

    If cbostatus.Text = "All" Then
        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile,status from CRIS_prospects where SAE='" & theSA & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "

    Else
        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile,status from CRIS_prospects where SAE='" & theSA & "' and status='" & theStatus & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "

    End If



    NEW_LogAudit "V", "PROSPECT DATA REPORT", "", "", "", "SAE Name:" & cboSA & "Status:" & cbostatus, "", ""
    LogAudit "V", "PROSPECT DATA ENTRY VIEW" & " MONTH:" & cboMonth & " YEAR:" & cboYear
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listProspect.ListItems.Clear
    cnt = 0

    If RS.EOF And RS.BOF Then
        MsgBox "No Item Found..", vbInformation, "Information"
        Exit Sub
    End If
    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = listProspect.ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!AcctName)
        Item.SubItems(2) = Null2String(RS!loginitialinquiry)
        Item.SubItems(3) = Null2String(RS!Model)
        Item.SubItems(4) = Null2String(RS!Variant)
        Item.SubItems(5) = Null2String(RS!Color)
        Item.SubItems(6) = Null2String(RS!Classification)
        Item.SubItems(7) = Null2String(RS!LeadSource)
        Item.SubItems(8) = Null2String(RS!EMAIL)
        Item.SubItems(9) = Null2String(RS!ContactPerson)
        Item.SubItems(10) = Null2String(RS!Address)
        Item.SubItems(11) = Null2String(RS!Telephone)
        Item.SubItems(12) = Null2String(RS!Mobile)
        If Null2String(RS!STATUS) = "O" Then
            Mstatus = "OPEN"
            Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "C" Then Mstatus = "CLOSE": Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "L" Then Mstatus = "LOSE SALES": Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "I" Then Mstatus = "INACTIVE": Item.SubItems(13) = Mstatus
        Else
            Mstatus = "NO STATUS"
            Item.SubItems(13) = Mstatus
        End If
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub cbomonth_Click()
    If cboMonth.Text = "Jannuary" Then
        theMonth = 1
    ElseIf cboMonth.Text = "January" Then theMonth = 2
    ElseIf cboMonth.Text = "February" Then theMonth = 2
    ElseIf cboMonth.Text = "March" Then theMonth = 3
    ElseIf cboMonth.Text = "April" Then theMonth = 4
    ElseIf cboMonth.Text = "May" Then theMonth = 5
    ElseIf cboMonth.Text = "June" Then theMonth = 6
    ElseIf cboMonth.Text = "July" Then theMonth = 7
    ElseIf cboMonth.Text = "August" Then theMonth = 8
    ElseIf cboMonth.Text = "September" Then theMonth = 9
    ElseIf cboMonth.Text = "October" Then theMonth = 10
    ElseIf cboMonth.Text = "November" Then theMonth = 11
    ElseIf cboMonth.Text = "December" Then theMonth = 12
    Else
        cboMonth.AddItem "All Months"
    End If
End Sub

Private Sub cboStatus_Click()
    If cbostatus.Text = "OPEN" Then
        theStatus = "O"
    ElseIf cbostatus.Text = "CLOSE" Then theStatus = "C"
    ElseIf cbostatus.Text = "INACTIVE" Then theStatus = "I"
    ElseIf cbostatus.Text = "LOST SALES" Then theStatus = "L"

    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdCost_Click()
    If cboMonth.Text = "" Then
        MsgBox "Plese Select Month..", vbInformation, "Information"
        cboMonth.SetFocus
        Exit Sub
    End If
    If cboYear.Text = "" Then
        MsgBox "Plese Select Year..", vbInformation, "Information"
        cboYear.SetFocus
        Exit Sub
    End If


    DisplayInformation
    picCost.Visible = False
    LV_AutoSizeColumn listProspect
End Sub

Private Sub CmdCostCancel_Click()
    picCost.Visible = False
End Sub

Private Sub cmdPrint_Click()
    If Module_Access(LOGID, "PROSPECT DATA REPORT", "REPORTS") = False Then Exit Sub
    Dim rsProspect                                                    As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim theSA                                                         As String
    Dim thedate                                                       As String

    theSA = Trim(cboSA.Text)
    thedate = cboMonth.Text + "-" + cboYear.Text

    If cbostatus.Text = "All" Then
        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile from CRIS_prospects where SAE='" & theSA & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "
    Else
        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile from CRIS_prospects where SAE='" & theSA & "' and status='" & theStatus & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "
    End If

    Set rsProspect = New ADODB.Recordset
    Set rsProspect = gconDMIS.Execute(SQL)

    If Not rsProspect.EOF And Not rsProspect.BOF Then
        ProspectSheet.Formulas(0) = "Company_name='" & COMPANY_NAME & "'"
        ProspectSheet.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
        ProspectSheet.Formulas(2) = "thedate='" & thedate & "'"
        If cbostatus.Text = "All" Then
            PrintSQLReport ProspectSheet, SMIS_REPORT_PATH & "prospectsheet.rpt", "({CRIS_PROSPECTS.SAE})= '" & theSA & "'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport ProspectSheet, SMIS_REPORT_PATH & "prospectsheet.rpt", "({CRIS_PROSPECTS.SAE})= '" & theSA & "' and ({CRIS_Prospects.Status})='" & theStatus & "'", DMIS_REPORT_Connection, 1
        End If
    End If

End Sub

Private Sub CmdView_Click()

    If cboSA.Text = "" Then
        MsgBox "Please select SAE..", vbInformation, "Information"
        cboSA.SetFocus
        Exit Sub
    End If

    If cbostatus.Text = "" Then
        MsgBox "Please Select A Status..", vbInformation, "Information"
        cbostatus.SetFocus
        Exit Sub
    End If

    picCost.Visible = True
    fillcbomonth cboMonth
    fillcbomoreyear cboYear
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillSAE
    FillStatus
    picCost.Visible = False
End Sub

