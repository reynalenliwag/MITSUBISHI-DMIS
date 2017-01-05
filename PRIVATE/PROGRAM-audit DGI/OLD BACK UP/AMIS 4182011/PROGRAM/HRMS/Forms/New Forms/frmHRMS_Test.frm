VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMS_TEST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmHRMS_TEST"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Test.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   10680
   Begin MSComctlLib.ListView lsvPayroll 
      Height          =   825
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   1455
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.ComboBox cboEmpno 
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4425
   End
   Begin MSComctlLib.ListView lsvPH 
      Height          =   885
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvSSS 
      Height          =   885
      Left            =   150
      TabIndex        =   3
      Top             =   2910
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvPI 
      Height          =   885
      Left            =   120
      TabIndex        =   4
      Top             =   4140
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvTAX 
      Height          =   885
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvPHd 
      Height          =   885
      Left            =   5490
      TabIndex        =   10
      Top             =   1740
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvSSSd 
      Height          =   885
      Left            =   5520
      TabIndex        =   11
      Top             =   2910
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvPId 
      Height          =   885
      Left            =   5490
      TabIndex        =   12
      Top             =   4140
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvTAXd 
      Height          =   885
      Left            =   5490
      TabIndex        =   13
      Top             =   5400
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvATM 
      Height          =   885
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvATMd 
      Height          =   885
      Left            =   5490
      TabIndex        =   15
      Top             =   6600
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvLOAN 
      Height          =   885
      Left            =   120
      TabIndex        =   17
      Top             =   7770
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date start"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "last update"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsvLOANd 
      Height          =   885
      Left            =   5490
      TabIndex        =   18
      Top             =   7770
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   1561
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LOAN"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   7500
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ATM"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   6330
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PHIL HEALTH"
      Height          =   255
      Index           =   3
      Left            =   150
      TabIndex        =   9
      Top             =   1500
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SSS"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   2670
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PAG-IBIG"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   3900
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TAX"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   5130
      Width           =   375
   End
End
Attribute VB_Name = "frmHRMS_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboEmpno_Change()
    Call FillListview
End Sub

Sub FillListview()
    Dim rsTmp As New ADODB.Recordset
    Dim Item As ListItem
    Dim LSV As ListView
    
    Set LSV = lsvSSS
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_SSS Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!DateStart)
            Item.SubItems(1) = rsTmp!EmployeeShare
            Item.SubItems(2) = rsTmp!LastDateCont
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing

    Set LSV = lsvPI
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_PagIbig Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!DateStart)
            Item.SubItems(1) = rsTmp!EmployeeShare
            Item.SubItems(2) = rsTmp!LastDateCont
            
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvPH
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_PhilHealth Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!DateStart)
            Item.SubItems(1) = rsTmp!EmployeeShare
            Item.SubItems(2) = rsTmp!LastDateCont
            
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvTAX
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_Tin Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!DateStart)
            Item.SubItems(1) = rsTmp!Deduction
            Item.SubItems(2) = rsTmp!LastDateCont
            
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing

    Set LSV = lsvTAXd
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_TinDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!AMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing

    Set LSV = lsvSSSd
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_SSSDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!EmployeeAMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvPId
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_PagIbigDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!EmployeeAMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvPHd
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_PhilhealthDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!EmployeeAMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvATM
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_ATM Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!NetAMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvATMd
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_AtmDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!NetAMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing

    Set LSV = lsvLOAN
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_LoanMas Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!DateGranted)
            Item.SubItems(1) = rsTmp!AMOUNTLoaned
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
    
    Set LSV = lsvLOANd
    Set rsTmp = New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_LoanMasDet Where Empno = '" & cboEmpno & "'")
    LSV.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set Item = LSV.ListItems.Add(, , rsTmp!Deyt)
            Item.SubItems(1) = rsTmp!AMOUNT
        
            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
End Sub


Private Sub cboEmpno_Click()
    Call FillListview
End Sub

Private Sub cboEmpno_LostFocus()
    Call FillListview
End Sub

Private Sub Form_Load()
    
    Call CenterMe(frmMain, Me, 1)
    
    Call FillCBO
End Sub

Sub FillCBO()
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = gconDMIS.Execute("SElect Distinct Empno From HRMS_EMPINFO ORder BY Empno asc")
    cboEmpno.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboEmpno.AddItem rsTmp!EMPNO
            
            rsTmp.MoveNext
        Loop
        cboEmpno.ListIndex = 0
    End If
    
    Set rsTmp = Nothing
End Sub
