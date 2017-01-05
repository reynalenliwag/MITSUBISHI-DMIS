VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "WIZMACBUT.OCX"
Begin VB.Form frmAMISMYOBLEDGERAccounts 
   Appearance      =   0  'Flat
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MYOB Accounts General Ledger"
   ClientHeight    =   6165
   ClientLeft      =   585
   ClientTop       =   330
   ClientWidth     =   11850
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MYOBAccountsLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   11850
   Begin wizMacBut.MacBut cmdExit 
      Height          =   345
      Left            =   9570
      TabIndex        =   4
      Top             =   5790
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Exit"
      Caption_Xpos    =   600
   End
   Begin wizMacBut.MacBut cmdPrint 
      Height          =   345
      Left            =   7860
      TabIndex        =   3
      Top             =   5790
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Print"
      Caption_Xpos    =   600
   End
   Begin wizMacBut.MacBut cmdFind 
      Height          =   345
      Left            =   6150
      TabIndex        =   2
      Top             =   5790
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Search"
      Caption_Xpos    =   400
   End
   Begin wizMacBut.MacBut cmdNext 
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   5790
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Next >>"
      Caption_Xpos    =   400
   End
   Begin wizMacBut.MacBut cmdPrev 
      Height          =   345
      Left            =   2730
      TabIndex        =   0
      Top             =   5790
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "<< Previous"
      Caption_Xpos    =   100
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      TabIndex        =   14
      Top             =   0
      Width           =   9135
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   180
         Width           =   1665
      End
      Begin VB.TextBox txtCode3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   2790
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtCode2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtCode1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtAccount_Name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   3270
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   180
         Width           =   5775
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2670
         TabIndex        =   20
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2070
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   17
         Top             =   210
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   180
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         Left            =   30
         TabIndex        =   15
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   2640
      TabIndex        =   18
      Top             =   570
      Width           =   9135
      Begin Crystal.CrystalReport rptGeneralLedger 
         Left            =   90
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "G E N E R A L  L E D G E R"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin wizButton.cmd cmdShow 
         Height          =   285
         Left            =   7350
         TabIndex        =   13
         Top             =   150
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         TX              =   "Show Ledger"
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
         MICON           =   "MYOBAccountsLedger.frx":030A
      End
      Begin VB.TextBox txtTotalBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   7260
         MaxLength       =   20
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   4740
         Width           =   1425
      End
      Begin VB.TextBox txtTotalDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   4500
         MaxLength       =   20
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   4740
         Width           =   1395
      End
      Begin VB.TextBox txtTotalCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   4740
         Width           =   1395
      End
      Begin MSFlexGridLib.MSFlexGrid grdAccountsLedger 
         Height          =   4155
         Left            =   60
         TabIndex        =   10
         Top             =   510
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7329
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   16744448
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         SelectionMode   =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MYOBAccountsLedger.frx":0326
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   2820
         TabIndex        =   11
         Top             =   150
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   20774915
         CurrentDate     =   38148
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   5250
         TabIndex        =   12
         Top             =   150
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   20774915
         CurrentDate     =   38148
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   4860
         TabIndex        =   34
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Left            =   2250
         TabIndex        =   33
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Date Range:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   60
         TabIndex        =   32
         Top             =   150
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Height          =   255
         Left            =   3390
         TabIndex        =   30
         Top             =   4770
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   6030
      Left            =   60
      ScaleHeight     =   5970
      ScaleWidth      =   2505
      TabIndex        =   25
      Top             =   90
      Width           =   2565
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   0
         Picture         =   "MYOBAccountsLedger.frx":0640
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   6105
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   2565
      Begin VB.OptionButton optCode 
         BackColor       =   &H00DEDFDE&
         Caption         =   "By &Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   390
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton optAccount_Name 
         BackColor       =   &H00DEDFDE&
         Caption         =   "By &Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   660
         Width           =   1725
      End
      Begin VB.TextBox TextSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   990
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4665
         Left            =   60
         TabIndex        =   24
         Top             =   1380
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MYOBAccountsLedger.frx":14FB8
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ACCOUNTS"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label12 
         BackColor       =   &H00DEDFDE&
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   27
         Top             =   150
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAMISMYOBLEDGERAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMYOBChartAccounts, rsAccType As ADODB.Recordset
Dim rsvwJOURNAL_HD, rsvwJOURNAL_HDDet, rsProfile As ADODB.Recordset
Dim AddorEdit, ORDER_BY As String
Dim TUTAL_DEBIT, TUTAL_CREDIT, TUTAL_BALANCE, BEGINNING_BALANCE As Double

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Unload Me
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Frame2.ZOrder 0
TextSearch.SetFocus
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
rsMYOBChartAccounts.MoveNext
If rsMYOBChartAccounts.EOF Then
   rsMYOBChartAccounts.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
rsMYOBChartAccounts.MovePrevious
If rsMYOBChartAccounts.BOF Then
   rsMYOBChartAccounts.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If MsgBox("Print General Ledger for this Account?", vbYesNo + vbQuestion, "Print: " & txtAccount_Name.Text) = vbYes Then
'   rptGeneralLedger.Reset
'   rptGeneralLedger.Formulas(3) = "BEG_DATE = '" & Format(dtFrom, "MMM-DD-YYYY") & "'"
'   rptGeneralLedger.Formulas(4) = "BEGINNING = " & BEGINNING_BALANCE
'   rptGeneralLedger.Formulas(5) = "REPORTDATE = '" & Format(dtTo, "LONG DATE") & "'"
'   rptGeneralLedger.ReportTitle = "G E N E R A L  L E D G E R"
'   Dim rsProfile As ADODB.Recordset
'   Set rsProfile = New ADODB.Recordset
'   Set rsProfile = gconAmis.Execute("Select * from Profile")
'   If Not (rsProfile.EOF And rsProfile.BOF) Then
'      'CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
'      rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
'      rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
'      'ShowReport "AccountGeneralLedger", "Ledgers", "{vwJOURNAL_HD.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {vwJOURNAL_HD.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {MYOBChartAccounts.Account_No} = '" & txtCode.Text & "'", "G E N E R A L  L E D G E R", "FROM: " & dtFrom & " TO: " & dtTo, True
'   End If
'   PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{vwJOURNAL_HD.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {vwJOURNAL_HD.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {MYOBChartAccounts.Account_No} = '" & txtCode.Text & "'", AMIS_REPORT_Connection, 1
End If
End Sub

Private Sub cmdShow_Click()
FillGrids
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 2 Then
   If KeyCode = vbKeyC Then
      optCode.Value = True
      optCode_Click
      TextSearch.SetFocus
   End If
   If KeyCode = vbKeyD Then
      optAccount_Name.Value = True
      optAccount_Name_Click
      TextSearch.SetFocus
   End If
End If
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
rsRefresh
TextSearch.Text = ""
initMemvars
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsMYOBChartAccounts = New ADODB.Recordset
    rsMYOBChartAccounts.Open "select * from MYOBChartAccounts order by Account_No asc", gconAmis, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
Frame1.Enabled = True
txtCode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
txtAccount_Name.Text = "":
txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
txtTotalBalance.Text = ZERO:
End Sub

Sub StoreMemvars()
If Not rsMYOBChartAccounts.EOF And Not rsMYOBChartAccounts.BOF Then
   Frame1.Enabled = False
   labID.Caption = rsMYOBChartAccounts!ID
   txtCode.Text = Null2String(rsMYOBChartAccounts!account_no)
   txtAccount_Name.Text = Null2String(rsMYOBChartAccounts!account_name)
   Set rsvwJOURNAL_HDDet = New ADODB.Recordset
       rsvwJOURNAL_HDDet.Open "select MIN(vwJournal_Det.JDate) AS MinimumDate, MAX(vwJournal_Det.JDate) AS MaximumDate from vwJournal_Det inner Join vwJOURNAL_HD on vwJournal_Det.Journal_No = vwJOURNAL_HD.Journal_No where vwJournal_Det.Account_No = '" & txtCode.Text & "'", gconAmis
   If Not rsvwJOURNAL_HDDet.EOF And Not rsvwJOURNAL_HDDet.BOF Then
      If IsNull(rsvwJOURNAL_HDDet!MinimumDate) = True Then
         cmdShow.Enabled = False
         dtFrom.Enabled = False
         dtTo.Enabled = False
      Else
         dtFrom.Enabled = True: dtTo.Enabled = True: cmdShow.Enabled = True
         'dtFrom = Null2Date(rsvwJOURNAL_HDDet!MinimumDate)
         'Set rsProfile = New ADODB.Recordset
         'Set rsProfile = gconAmis.Execute("Select PeriodMonth,PeriodYear from Profile")
         'If Not rsProfile.EOF And Not rsProfile.BOF Then dtFrom = DateSerial(Null2String(rsProfile!periodyear), Null2String(rsProfile!periodmonth), "1")
         dtFrom = firstDay(Null2Date(rsvwJOURNAL_HDDet!MaximumDate))
         dtTo = Null2Date(rsvwJOURNAL_HDDet!MaximumDate)
      End If
   Else
      dtFrom = LOGDATE: dtTo = LOGDATE: cmdShow.Enabled = False
      dtFrom.Enabled = False: dtTo.Enabled = False
   End If
End If
End Sub

Sub initGrid()
With grdAccountsLedger
   .Rows = 2
   .ColWidth(0) = 1200: .ColWidth(1) = 1300: .ColWidth(2) = 2000
   .ColWidth(3) = 1400: .ColWidth(4) = 1400: .ColWidth(5) = 1400
   .ColWidth(6) = 25000: .ColWidth(7) = 1: .Row = 0
   .Col = 0: .Text = "DOCDATE"
   .Col = 1: .Text = "REFERENCE"
   .Col = 2: .Text = "REFERENCE NAME"
   .Col = 3: .Text = "DEBIT"
   .Col = 4: .Text = "CREDIT"
   .Col = 5: .Text = "BALANCE"
   .Col = 6: .Text = "PARTICULARS"
   .Col = 7: .Text = "ID"
End With
End Sub

Sub FillGrids()
Dim OUTBALANCE As Double
Dim REFERENCE, REFERENCE_NAME As String
Dim Cnt As Integer
cleargrid grdAccountsLedger: initGrid
Set rsvwJOURNAL_HDDet = New ADODB.Recordset
    'rsvwJOURNAL_HDDet.Open "select Journal_Det.ID,Journal_Det.Journal_No,Journal_Det.JDate,Journal_Det.JType,Journal_Det.Debit,Journal_Det.Credit,Journal_Det.Journal_No,vwJOURNAL_HD.VendorCode,vwJOURNAL_HD.CustomerCode,vwJOURNAL_HD.Journal_No,vwJOURNAL_HD.Remarks from Journal_Det inner Join vwJOURNAL_HD on Journal_Det.Journal_No = vwJOURNAL_HD.Journal_No where Journal_Det.Jdate >= '" & dtFrom & "' and Journal_Det.Jdate <= '" & dtTo & "' and Journal_Det.Status='P' and Journal_Det.Account_No = '" & txtCode.Text & "' order by Journal_Det.JDate asc,Journal_Det.ID asc", gconAmis
    rsvwJOURNAL_HDDet.Open "select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from vwJOURNAL_DET where Jdate < '" & dtFrom & "' and Account_No = '" & txtCode.Text & "'", gconAmis
TUTAL_BALANCE = 0: TUTAL_BALANCE = TUTAL_BALANCE: Cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0: BEGINNING_BALANCE = 0
If Not rsvwJOURNAL_HDDet.EOF And Not rsvwJOURNAL_HDDet.BOF Then
   OUTBALANCE = N2Str2Zero(rsvwJOURNAL_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsvwJOURNAL_HDDet!TOTAL_CREDIT)
   BEGINNING_BALANCE = OUTBALANCE
   If BEGINNING_BALANCE > 0 Then
      grdAccountsLedger.AddItem dtFrom & Chr(9) & _
                                "" & Chr(9) & _
                                "BEGINNING BALANCE" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                "0.00" & Chr(9) & _
                                ToDoubleNumber(OUTBALANCE) & Chr(9) & "" & Chr(9)
   End If
End If
Set rsvwJOURNAL_HDDet = New ADODB.Recordset
    'rsvwJOURNAL_HDDet.Open "select * from vLEDGER where Jdate >= '" & dtFrom & "' and Jdate <= '" & dtTo & "' and Account_No = '" & txtCode.Text & "' order by JDate asc,ID asc", gconAmis
    rsvwJOURNAL_HDDet.Open "select * from vwJOURNAL_DET where Jdate >= '" & dtFrom & "' and Jdate <= '" & dtTo & "' and Account_No = '" & txtCode.Text & "' order by JDate asc, Journal_No asc", gconAmis
If Not rsvwJOURNAL_HDDet.EOF And Not rsvwJOURNAL_HDDet.BOF Then
   rsvwJOURNAL_HDDet.MoveFirst
   Screen.MousePointer = 11: DoEvents
   grdAccountsLedger.MousePointer = flexHourglass
   Do While Not rsvwJOURNAL_HDDet.EOF
      Cnt = Cnt + 1
      REFERENCE = Null2String(rsvwJOURNAL_HDDet!Journal_No)
      REFERENCE_NAME = Null2String(rsvwJOURNAL_HDDet!Memo)
      OUTBALANCE = OUTBALANCE + (N2Str2Zero(rsvwJOURNAL_HDDet!Debit) - N2Str2Zero(rsvwJOURNAL_HDDet!Credit))
      grdAccountsLedger.AddItem Null2String(rsvwJOURNAL_HDDet!jdate) & Chr(9) & _
                                REFERENCE & Chr(9) & _
                                REFERENCE_NAME & Chr(9) & _
                                ToDoubleNumber(N2Str2Zero(rsvwJOURNAL_HDDet!Debit)) & Chr(9) & _
                                ToDoubleNumber(N2Str2Zero(rsvwJOURNAL_HDDet!Credit)) & Chr(9) & _
                                ToDoubleNumber(OUTBALANCE) & Chr(9) & REFERENCE_NAME & Chr(9) & rsvwJOURNAL_HDDet!ID
      TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsvwJOURNAL_HDDet!Debit)
      TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsvwJOURNAL_HDDet!Credit)
      rsvwJOURNAL_HDDet.MoveNext
   Loop
   If Cnt > 0 Then grdAccountsLedger.RemoveItem 1
   txtTotalDebit.Text = ToDoubleNumber(TUTAL_DEBIT)
   txtTotalCredit.Text = ToDoubleNumber(TUTAL_CREDIT)
   txtTotalBalance.Text = ToDoubleNumber(TUTAL_BALANCE + OUTBALANCE)
   Screen.MousePointer = 0: grdAccountsLedger.MousePointer = flexCustom
Else
   txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO:   txtTotalBalance.Text = ZERO
   cleargrid grdAccountsLedger
End If
Set rsvwJOURNAL_HDDet = Nothing
End Sub

Private Sub grdAccountsLedger_DblClick()
grdAccountsLedger.Row = grdAccountsLedger.Row
grdAccountsLedger.Col = 7
Dim VarJournal_No As String
VarJournal_No = grdAccountsLedger.Text
MYOB_JTYPE = VarJournal_No
Screen.MousePointer = 11
On Error Resume Next
Unload frmAMISMYOBJournalEntry
frmAMISMYOBJournalEntry.Show
frmAMISMYOBJournalEntry.StoreSearch (VarJournal_No)
Screen.MousePointer = 0
End Sub

Private Sub optAccount_Name_Click()
If TextSearch = "" Then FillGrid2 Else FillSearchGrid2 (TextSearch.Text)
TextSearch.SetFocus
End Sub

Private Sub optCode_Click()
If TextSearch = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
TextSearch.SetFocus
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
If optCode.Value = True Then
   rsMYOBChartAccounts.Bookmark = rsFind(rsMYOBChartAccounts.Clone, "Account_No", lstAccounts.SelectedItem).Bookmark
Else
   rsMYOBChartAccounts.Bookmark = rsFind(rsMYOBChartAccounts.Clone, "Account_No", lstAccounts.SelectedItem.SubItems(1)).Bookmark
End If
cleargrid grdAccountsLedger
initGrid
StoreMemvars
End Sub

Private Sub lstAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstAccounts
     .Sorted = True
     If .SortKey = ColumnHeader.Index - 1 Then
        If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
     Else
        .SortOrder = lvwAscending
        .SortKey = ColumnHeader.Index - 1
     End If
End With
End Sub

Private Sub lstAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TextSearch.SetFocus
End Sub

Private Sub textSearch_Change()
If optCode.Value = True Then
   If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
Else
   If Trim(TextSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (TextSearch.Text)
End If
End Sub

Sub FillGrid()
Dim rsMYOBChartAccountss As ADODB.Recordset
lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
Set rsMYOBChartAccountss = New ADODB.Recordset
Set rsMYOBChartAccountss = gconAmis.Execute("select Account_No from MYOBChartAccounts order by Account_No asc")
If Not (rsMYOBChartAccountss.EOF And rsMYOBChartAccountss.BOF) Then
   Listview_Loadval Me.lstAccounts.ListItems, rsMYOBChartAccountss
   lstAccounts.Refresh
   lstAccounts.Enabled = True
Else
   lstAccounts.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsMYOBChartAccountss As ADODB.Recordset
lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
Set rsMYOBChartAccountss = New ADODB.Recordset
Set rsMYOBChartAccountss = gconAmis.Execute("select Account_No from MYOBChartAccounts where Account_No like'" & XXX & "%' order by Account_No asc")
If Not (rsMYOBChartAccountss.EOF And rsMYOBChartAccountss.BOF) Then
   Listview_Loadval Me.lstAccounts.ListItems, rsMYOBChartAccountss
   lstAccounts.Refresh
   lstAccounts.Enabled = True
Else
   lstAccounts.Enabled = False
End If
End Sub

Sub FillGrid2()
Dim rsMYOBChartAccountss As ADODB.Recordset
lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
Set rsMYOBChartAccountss = New ADODB.Recordset
Set rsMYOBChartAccountss = gconAmis.Execute("select Account_Name,Account_No from MYOBChartAccounts order by Account_Name asc")
If Not (rsMYOBChartAccountss.EOF And rsMYOBChartAccountss.BOF) Then
   Listview_Loadval Me.lstAccounts.ListItems, rsMYOBChartAccountss
   lstAccounts.Refresh
   lstAccounts.Enabled = True
Else
   lstAccounts.Enabled = False
End If
End Sub

Sub FillSearchGrid2(XXX As String)
Dim rsMYOBChartAccountss As ADODB.Recordset
lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
Set rsMYOBChartAccountss = New ADODB.Recordset
Set rsMYOBChartAccountss = gconAmis.Execute("select Account_Name,Account_No from MYOBChartAccounts where Account_Name like'" & XXX & "%' order by Account_Name asc")
If Not (rsMYOBChartAccountss.EOF And rsMYOBChartAccountss.BOF) Then
   Listview_Loadval Me.lstAccounts.ListItems, rsMYOBChartAccountss
   lstAccounts.Refresh
   lstAccounts.Enabled = True
Else
   lstAccounts.Enabled = False
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Frame2.ZOrder 0
If KeyCode = vbKeyDown Then lstAccounts.SetFocus
End Sub
