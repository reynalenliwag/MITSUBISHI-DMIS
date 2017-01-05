VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCMIS_Tools_BankDepo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Deposit Tools"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCMIS_Tools_BankDepo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   7845
   Begin VB.ComboBox cboBanks 
      Height          =   330
      Left            =   2100
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   180
      Width           =   3105
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   465
      Left            =   6210
      TabIndex        =   3
      Top             =   4770
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   60
      ScaleHeight     =   4065
      ScaleWidth      =   7695
      TabIndex        =   1
      Top             =   630
      Width           =   7725
      Begin wizProgBar.Prg Prg 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   3300
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   503
         Picture         =   "frmCMIS_Tools_BankDepo.frx":1082
         ForeColor       =   0
         BarPicture      =   "frmCMIS_Tools_BankDepo.frx":109E
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
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   3270
         Width           =   1185
      End
      Begin MSComctlLib.ListView lsvList 
         Height          =   2895
         Left            =   30
         TabIndex        =   5
         Top             =   360
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR no."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Deposit"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cut Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboCorrect 
         Height          =   330
         Left            =   2040
         TabIndex        =   7
         Top             =   3660
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label labCAP 
         AutoSize        =   -1  'True
         Caption         =   "Deposit to Correct Bank"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   3750
         Width           =   1965
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -90
         TabIndex        =   4
         Top             =   -30
         Width           =   7815
         _Version        =   655364
         _ExtentX        =   13785
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "   LIST OF DEPOSIT TRANSACTION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorDark=   4210752
      End
   End
   Begin VB.CommandButton cmdCorrect 
      Caption         =   "Process"
      Height          =   465
      Left            =   4650
      TabIndex        =   2
      Top             =   4770
      Width           =   1575
   End
   Begin VB.Label labCAP 
      Caption         =   "List of Banks Without Links to AMIS Banks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   570
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   2055
   End
End
Attribute VB_Name = "frmCMIS_Tools_BankDepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBanks_Click()
    Call FillGrid
End Sub

Sub FillGrid()
    Dim rsTMP                                           As New ADODB.Recordset
    Dim ITEM                                            As ListItem
        
    Prg.Value = 0
    Screen.MousePointer = 11
    Set rsTMP = gconDMIS.Execute("SELECT OR_NUM, TYPE, DATDEPOSIT, CUTDATE, DEPOSIT, ID FROM CMIS_BANKDEPO WHERE DEPOSIT_TO = " & N2Str2Null(cboBanks.Text) & "")
    lsvList.ListItems.Clear
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            Set ITEM = lsvList.ListItems.Add(, , Null2String(rsTMP!OR_NUM))
            If rsTMP!Type = 1 Then
                ITEM.SubItems(1) = "CASH"
            ElseIf rsTMP!Type = 1 Then
                ITEM.SubItems(1) = "CHECK"
            Else
                ITEM.SubItems(1) = "CARD"
            End If
            ITEM.SubItems(2) = Null2String(rsTMP!DATDEPOSIT)
            ITEM.SubItems(3) = Null2String(rsTMP!CUTDATE)
            ITEM.SubItems(4) = Format(NumericVal(rsTMP!DEPOSIT), MAXIMUM_DIGIT)
            ITEM.SubItems(5) = rsTMP!Id
            
            rsTMP.MoveNext
            DoEvents
        Loop
    End If
    Screen.MousePointer = 0
    Set rsTMP = Nothing
End Sub

Private Sub cboBanks_DblClick()
    Call FillGrid
End Sub

Private Sub cboBanks_LostFocus()
    Call FillGrid
End Sub

Private Sub Check1_Click()
    If lsvList.ListItems.Count = 0 Then Exit Sub
    
    Dim X                                                   As Long
    If Check1.Value = 1 Then
        For X = 1 To lsvList.ListItems.Count
            lsvList.ListItems(X).Checked = True
        Next
    Else
        For X = 1 To lsvList.ListItems.Count
            lsvList.ListItems(X).Checked = False
        Next
    End If
End Sub

Private Sub cmdCorrect_Click()
    If lsvList.ListItems.Count = 0 Then Exit Sub
        
    If cboCorrect.Text = "" Then
        Call ShowIsRequiredMsg("Choose the bank to be Moved")
        cboCorrect.SetFocus
        Exit Sub
    End If
    
    Dim X                                       As Long
    For X = 1 To lsvList.ListItems.Count
        If lsvList.ListItems(X).Checked = True Then
            GoTo Procced_Moving
        End If
    Next
    
    MsgBox "Choose a transaction to move", vbInformation, "Info."
    Exit Sub
    
Procced_Moving:
    If MsgBox("Proceed the Process, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    Prg.Max = lsvList.ListItems.Count
    Prg.Value = 0
    
    Screen.MousePointer = 11
    For X = 1 To lsvList.ListItems.Count
        If lsvList.ListItems(X).Checked = True Then
            gconDMIS.Execute ("UPDATE CMIS_BANKDEPO SET DEPOSIT_TO = " & N2Str2Null(cboCorrect.BoundText) & " WHERE ID = " & lsvList.ListItems(X).ListSubItems(5) & "")
            Prg.Text = "OR no. " & lsvList.ListItems(X).Text
            Prg.ForeColor = vbWhite
        End If
        
        Prg.Value = Prg.Value + 1
        DoEvents
    Next
    Screen.MousePointer = 0
    MsgBox "Process Complete", vbInformation, "Info."
    
    Call InitData
    lsvList.ListItems.Clear
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call InitData
End Sub

Sub InitData()
    LoadDataCombo cboCorrect, "SELECT BANKNAME, BANKCODE FROM ALL_BANKS", "BANKCODE", "BANKNAME"
    
    Dim rsTMP                                               As New ADODB.Recordset
    Set rsTMP = gconDMIS.Execute("SELECT DISTINCT(CMIS_BANKDEPO.DEPOSIT_TO) FROM CMIS_BANKDEPO " & _
        " WHERE DEPOSIT_TO NOT IN ( SELECT bankcode FROM ALL_BANKS )")
    Call Combo_Loadval(cboBanks, rsTMP)
End Sub

Sub LoadDataCombo(cbo As DataCombo, xxxSQL As String, showingcolumn As String, datacolumn As String)
    Dim RSCOMBO As ADODB.Recordset
    Set RSCOMBO = gconDMIS.Execute(xxxSQL)
        'Set cboUSER_DEALER.DataSource = RSCOMBO
        Set cbo.RowSource = RSCOMBO
        cbo.BoundColumn = showingcolumn
        cbo.ListField = datacolumn
        Set cbo.DataSource = RSCOMBO
    Set RSCOMBO = Nothing
End Sub


'cmd.Parameters.Append cmd.CreateParameter("@DEP_ID", adInteger, adParamInput, , cboDepartments.BoundText)
