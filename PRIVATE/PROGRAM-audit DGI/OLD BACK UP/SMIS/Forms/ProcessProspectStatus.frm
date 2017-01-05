VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSMIS_Process_ProspectStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prospect Data Information "
   ClientHeight    =   6585
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
   Icon            =   "ProcessProspectStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   13200
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
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   13185
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   12150
         MouseIcon       =   "ProcessProspectStatus.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "ProcessProspectStatus.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print Report"
         Top             =   810
         Width           =   885
      End
      Begin VB.ComboBox cboSTATUSFILTER 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ProcessProspectStatus.frx":0EBB
         Left            =   4800
         List            =   "ProcessProspectStatus.frx":0EBD
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1410
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton optSelectSpecific 
         Caption         =   "Select All"
         Height          =   345
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1410
         Width           =   1185
      End
      Begin VB.OptionButton optSelectNone 
         Caption         =   "Select None"
         Height          =   345
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1410
         Width           =   1185
      End
      Begin VB.OptionButton optSelectALL 
         Caption         =   "Select All"
         Height          =   345
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1410
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ProcessProspectStatus.frx":0EBF
         Left            =   1260
         List            =   "ProcessProspectStatus.frx":0EC1
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1020
         Width           =   3105
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10020
         TabIndex        =   12
         Text            =   "0"
         Top             =   1050
         Width           =   1995
      End
      Begin VB.ComboBox cbomonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ProcessProspectStatus.frx":0EC3
         Left            =   10020
         List            =   "ProcessProspectStatus.frx":0EC5
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   255
         Width           =   2025
      End
      Begin VB.ComboBox cboyear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ProcessProspectStatus.frx":0EC7
         Left            =   10020
         List            =   "ProcessProspectStatus.frx":0EC9
         TabIndex        =   8
         Top             =   675
         Width           =   2025
      End
      Begin VB.Frame Frame2 
         Caption         =   "Prospect Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   60
         TabIndex        =   6
         Top             =   1770
         Width           =   13065
         Begin MSComctlLib.ListView listProspect 
            Height          =   4455
            Left            =   60
            TabIndex        =   7
            Top             =   210
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   7858
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
            NumItems        =   16
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
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Age"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.CommandButton CmdView 
         Appearance      =   0  'Flat
         Caption         =   "View"
         Height          =   645
         Left            =   12150
         Picture         =   "ProcessProspectStatus.frx":0ECB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   885
      End
      Begin VB.ComboBox cbostatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "ProcessProspectStatus.frx":134A
         Left            =   1260
         List            =   "ProcessProspectStatus.frx":134C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   3105
      End
      Begin VB.ComboBox CboSA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   195
         Width           =   3075
      End
      Begin Crystal.CrystalReport ProspectSheet 
         Left            =   4590
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Prospect Age From Initial Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7200
         TabIndex        =   13
         Top             =   1140
         Width           =   2745
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9420
         TabIndex        =   11
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9510
         TabIndex        =   10
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
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
         Left            =   60
         TabIndex        =   3
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "SA Name"
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
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   4230
      ScaleHeight     =   1485
      ScaleWidth      =   4005
      TabIndex        =   20
      Top             =   2490
      Visible         =   0   'False
      Width           =   4035
      Begin wizProgBar.Prg Prg1 
         Height          =   465
         Left            =   210
         TabIndex        =   21
         Top             =   270
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   820
         Picture         =   "ProcessProspectStatus.frx":134E
         ForeColor       =   0
         BarPicture      =   "ProcessProspectStatus.frx":136A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSMIS_Process_ProspectStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CheckCheck()

    Dim i
    For i = 1 To listProspect.ListItems.Count
        If listProspect.ListItems(i).Checked = True Then
            cmdPRINT.Enabled = True
            Exit Sub
        End If
    Next

End Sub

Sub SetColorX(colorx As OLE_COLOR, lstitem As ListItem)
    Dim i
    lstitem.ForeColor = colorx
    For i = 1 To lstitem.ListSubItems.Count - 1
        lstitem.ListSubItems(i).ForeColor = colorx
    Next

End Sub

Sub InitCbo()

    Combo_Loadval cboSA, gconDMIS.Execute("SELECT DISTINCT upper(SAE) as SAE FROM CRIS_PROSPECTS where isnull(SAE ,'')<>'' ")
    cboSA.AddItem "ALL", 0
    cboSA.ListIndex = 0

    cbostatus.Clear
    cbostatus.AddItem "OPEN"
    cbostatus.AddItem "CLOSE"
    cbostatus.AddItem "INACTIVE"
    cbostatus.AddItem "LOST SALES"
    cbostatus.AddItem "NO STATUS"
    cbostatus.AddItem "ALL", 0
    cbostatus.ListIndex = 0



    fillcbomonth cboMonth
    cboMonth.AddItem "ALL", 0
    cboMonth.ListIndex = 0

    Combo_Loadval cboYear, gconDMIS.Execute("Select distinct year(loginitialinquiry) from cris_prospects where isdate(loginitialinquiry)=1 order by 1 desc ")
    cboYear.AddItem "ALL", 0
    cboYear.ListIndex = 0

    Combo_Loadval Combo1, gconDMIS.Execute("Select distinct classification from cris_prospects where classification is not null")
    Combo1.AddItem "ALL", 0
    Combo1.ListIndex = 0

    cboSTATUSFILTER.Clear
    cboSTATUSFILTER.AddItem "OPEN"
    cboSTATUSFILTER.AddItem "CLOSE"
    cboSTATUSFILTER.AddItem "INACTIVE"
    cboSTATUSFILTER.AddItem "LOST SALES"
    cboSTATUSFILTER.AddItem "NO STATUS"
    cboSTATUSFILTER.ListIndex = 0



End Sub

Private Sub cbomonth_LostFocus()
    cboMonth.ListIndex = SelectCombo(cboMonth, cboMonth)
    If cboMonth.ListIndex = -1 Then
        cboMonth.ListIndex = 0
    End If
End Sub

Private Sub CboSA_LostFocus()
    cboSA.ListIndex = SelectCombo(cboSA, cboSA)
    If cboSA.ListIndex = -1 Then
        cboSA.ListIndex = 0
    End If
End Sub

Private Sub cbostatus_LostFocus()
    cbostatus.ListIndex = SelectCombo(cbostatus, cbostatus)
    If cbostatus.ListIndex = -1 Then
        cbostatus.ListIndex = 0
    End If
End Sub

Private Sub cboSTATUSFILTER_Click()
    Dim i

    For i = 1 To listProspect.ListItems.Count
        If UCase(listProspect.ListItems(i).ListSubItems(13).Text) = cboSTATUSFILTER Then
            listProspect.ListItems(i).Checked = True
            cmdPRINT.Enabled = True
        Else
            listProspect.ListItems(i).Checked = False
        End If
    Next

End Sub

Private Sub cboyear_LostFocus()
    cboYear.ListIndex = SelectCombo(cboYear, cboYear)
    If cboYear.ListIndex = -1 Then
        cboYear.ListIndex = 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Check3_Click()

End Sub

Private Sub cmdPrint_Click()
    '    ShowHidePictureBox2 Picture1, True, Frame1
    '    Prg1.Value = 0
    '    Prg1.Max = listProspect.ListItems.Count
    '    Dim i
    '
    '    For i = 1 To listProspect.ListItems.Count
    '    DoEvents
    '            Prg1.Value = Prg1.Value + 1
    '            Prg1.Text = (Prg1.Value / Prg1.Max) * 100
    '        If listProspect.ListItems(i).Checked = True Then
    '            gconDMIS.Execute "update CRIS_PROSPECTS SET STATUS='C' WHERE PROSPECTID=" & listProspect.ListItems(i).ListSubItems(15).Text
    '        End If
    '    Next
    '
    '    ShowHidePictureBox2 Picture1, False, Frame1
    '    CmdView.Value = True


    '    If Module_Access(LOGID, "PROSPECT DATA REPORT", "REPORTS") = False Then Exit Sub
    '    Dim rsProspect As New ADODB.Recordset
    '    Dim SQL As String
    '    Dim theSA As String
    '    Dim thedate As String
    '
    '    theSA = Trim(CboSA.Text)
    '    thedate = cboMonth.Text + "-" + cboYear.Text
    '
    '    If cbostatus.Text = "All" Then
    '        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile from CRIS_prospects where SAE='" & theSA & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "
    '        Else
    '        SQL = "SELECT logInitialInquiry,acctname,model,variant,Classification,color,leadsource,email,Contactperson,address,telephone,mobile from CRIS_prospects where SAE='" & theSA & "' and status='" & theStatus & "' and year(loginitialinquiry)= '" & cboYear & "' and month(loginitialinquiry) <='" & theMonth & "' "
    '    End If
    '
    '    Set rsProspect = New ADODB.Recordset
    '    Set rsProspect = gconDMIS.Execute(SQL)
    '
    '    If Not rsProspect.EOF And Not rsProspect.BOF Then
    '        ProspectSheet.Formulas(0) = "Company_name='" & COMPANY_NAME & "'"
    '        ProspectSheet.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
    '        ProspectSheet.Formulas(2) = "thedate='" & thedate & "'"
    '        If cbostatus.Text = "All" Then
    '            PrintSQLReport ProspectSheet, SMIS_REPORT_PATH & "prospectsheet.rpt", "({CRIS_PROSPECTS.SAE})= '" & theSA & "'", DMIS_REPORT_Connection, 1
    '            Else
    '            PrintSQLReport ProspectSheet, SMIS_REPORT_PATH & "prospectsheet.rpt", "({CRIS_PROSPECTS.SAE})= '" & theSA & "' and ({CRIS_Prospects.Status})='" & theStatus & "'", DMIS_REPORT_Connection, 1
    '        End If
    '    End If
End Sub

Private Sub CmdView_Click()

    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim Item                                                          As ListItem
    Dim cnt                                                           As Integer
    Dim theSA                                                         As String
    Dim theStatus                                                     As String
    Dim theYear                                                       As String
    Dim theMonth                                                      As String
    Dim theClassification                                             As String
    Dim Mstatus
    If cboSA = "ALL" Then
        theSA = ""
    Else
        theSA = " AND SAE=" & N2Str2Null(cboSA)
    End If

    If cbostatus = "ALL" Then
        theStatus = ""
    ElseIf cbostatus = "NO STATUS" Then
        theStatus = " AND STATUS is null "
    Else
        theStatus = " AND STATUS=" & N2Str2Null(Left(cbostatus, 1))
    End If
    If cboYear = "ALL" Then
        theYear = ""
    Else
        theYear = " AND YEAR(LOGINITIALINQUIRY)=" & (cboYear)
    End If

    If cboMonth = "ALL" Then
        theMonth = ""
    Else
        theMonth = " AND MONTH(LOGINITIALINQUIRY)=" & What_month(cboMonth)
    End If
    If Combo1 = "ALL" Then
        theClassification = ""
    Else
        theClassification = " AND CLASSIFICATION='" & Combo1 & "'"
    End If
    If theSA = "" And theStatus = "" And theYear = "" And theMonth = "" And theClassification = "" Then
        SQL = "SELECT LOGINITIALINQUIRY,ACCTNAME,MODEL,VARIANT,CLASSIFICATION,COLOR,LEADSOURCE,EMAIL,CONTACTPERSON,ADDRESS,TELEPHONE,MOBILE,STATUS , PROSPECTID FROM CRIS_PROSPECTS WHERE DATEDIFF(DAY,LOGINITIALINQUIRY,GETDATE())>=" & txtAge
    Else
        SQL = "SELECT LOGINITIALINQUIRY,ACCTNAME,MODEL,VARIANT,CLASSIFICATION,COLOR,LEADSOURCE,EMAIL,CONTACTPERSON,ADDRESS,TELEPHONE,MOBILE,STATUS , PROSPECTID FROM CRIS_PROSPECTS WHERE DATEDIFF(DAY,LOGINITIALINQUIRY,GETDATE())>=" & txtAge & theSA & theStatus & theYear & theMonth & theClassification
    End If
    listProspect.Sorted = False




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
        If IsDate(RS!loginitialinquiry) = True Then
            Item.SubItems(2) = Format(RS!loginitialinquiry, "mm/dd/yyyy")
            Item.SubItems(14) = DateDiff("d", RS!loginitialinquiry, LOGDATE)
        End If

        Item.SubItems(3) = UCase(Null2String(RS!Model))
        Item.SubItems(4) = UCase(Null2String(RS!Variant))
        Item.SubItems(5) = UCase(Null2String(RS!Color))
        Item.SubItems(6) = UCase(Null2String(RS!Classification))
        Item.SubItems(7) = UCase(Null2String(RS!LeadSource))
        Item.SubItems(8) = Null2String(RS!EMAIL)
        Item.SubItems(9) = Null2String(RS!ContactPerson)
        Item.SubItems(10) = Null2String(RS!Address)
        Item.SubItems(11) = Null2String(RS!Telephone)
        Item.SubItems(12) = Null2String(RS!Mobile)

        Item.SubItems(15) = RS!PROSPECTID


        If Null2String(RS!STATUS) = "O" Then
            Mstatus = "OPEN"
            Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "C" Then Mstatus = "CLOSE": Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "L" Then Mstatus = "LOST SALES": Item.SubItems(13) = Mstatus
        ElseIf Null2String(RS!STATUS) = "I" Then Mstatus = "INACTIVE": Item.SubItems(13) = Mstatus
        Else
            Mstatus = "NO STATUS"
            Item.SubItems(13) = Mstatus
        End If
        RS.MoveNext
    Loop
    Set RS = Nothing
    Dim i
    For i = 1 To listProspect.ListItems.Count
        If listProspect.ListItems(i).ListSubItems(13).Text = "OPEN" Then
            SetColorX &H800000, listProspect.ListItems(i)
        ElseIf listProspect.ListItems(i).ListSubItems(13).Text = "CLOSE" Then
            SetColorX &H4000&, listProspect.ListItems(i)
        ElseIf listProspect.ListItems(i).ListSubItems(13).Text = "LOST SALES" Then
            SetColorX vbRed, listProspect.ListItems(i)
        ElseIf listProspect.ListItems(i).ListSubItems(13).Text = "INACTIVE" Then
            SetColorX &H40C0&, listProspect.ListItems(i)
        Else
            SetColorX &H4080&, listProspect.ListItems(i)
        End If
    Next
    '    If listProspect.ListItems.Count > 0 Then
    '        optSelectALL.Visible = True
    '        optSelectNone.Visible = True
    '        optSelectSpecific.Visible = True
    '        optSelectNone.Value = True
    '
    '    Else
    '        optSelectALL.Visible = False
    '        optSelectNone.Visible = False
    '        optSelectSpecific.Visible = False
    '        cboSTATUSFILTER.Visible = False
    '    End If
End Sub

Private Sub Combo1_LostFocus()
    Combo1.ListIndex = SelectCombo(Combo1, Combo1)
    If Combo1.ListIndex = -1 Then
        Combo1.ListIndex = 0
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
    ShowHidePictureBox2 Picture1, False, Frame1
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitCbo
    optSelectALL.Visible = False
    optSelectNone.Visible = False
    optSelectSpecific.Visible = False
    cboSTATUSFILTER.Visible = False
End Sub

Private Sub listProspect_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With listProspect
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub optSelectALL_Click()
    cboSTATUSFILTER.Visible = False
    Dim i

    If optSelectALL.Value = True Then
        For i = 1 To listProspect.ListItems.Count
            listProspect.ListItems(i).Checked = True
            cmdPRINT.Enabled = True
        Next
    End If

End Sub

Private Sub optSelectNone_Click()
    If optSelectNone.Value = True Then

        cboSTATUSFILTER.Visible = False
        Dim i
        For i = 1 To listProspect.ListItems.Count
            listProspect.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub optSelectSpecific_Click()
    Dim i
    If optSelectSpecific.Value = True Then
        cboSTATUSFILTER.Visible = True

        For i = 1 To listProspect.ListItems.Count
            If UCase(listProspect.ListItems(i).ListSubItems(13).Text) = cboSTATUSFILTER Then
                cmdPRINT.Enabled = True
                listProspect.ListItems(i).Checked = True
            Else
                listProspect.ListItems(i).Checked = False
            End If
        Next


    Else
        cboSTATUSFILTER.Visible = False

    End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAge_LostFocus()
    If IsNumeric(txtAge) = False Then txtAge.Text = "0"
End Sub

