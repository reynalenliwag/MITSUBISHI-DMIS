VERSION 5.00
Begin VB.Form frmTaxComputer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmTaxComputer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6075
   Begin VB.ComboBox cboStatus 
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
      ItemData        =   "frmTaxComputer.frx":058A
      Left            =   360
      List            =   "frmTaxComputer.frx":05B2
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   510
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1125
      Left            =   60
      TabIndex        =   3
      Top             =   2160
      Width           =   5865
      Begin VB.Label lblTax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2070
         TabIndex        =   4
         Top             =   360
         Width           =   3525
      End
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "compute"
      Height          =   525
      Left            =   4410
      TabIndex        =   2
      Top             =   1440
      Width           =   1545
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   1140
      TabIndex        =   1
      Top             =   930
      Width           =   4755
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmTaxComputer.frx":05EC
      Left            =   360
      List            =   "frmTaxComputer.frx":05FC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   5535
   End
End
Attribute VB_Name = "frmTaxComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim RSTAX As New ADODB.Recordset
    Dim COLNO As Integer
    Dim EMPSAL_GROSS As Currency
    Dim RESULT_TAX As Currency
    Dim RESULT As Currency
    
    EMPSAL_GROSS = CCur(txtAmount.Text)
    
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_TaxTableDetails Where TaxBasis = '" & cboType & "' aND tAXcODE ='" & cboStatus & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        If 2 > EMPSAL_GROSS Then RESULT_TAX = rsTmp!Col1: COLNO = 1
        If EMPSAL_GROSS >= rsTmp!Col2 And EMPSAL_GROSS <= (rsTmp!Col3 - 1) Then RESULT_TAX = rsTmp!Col2: COLNO = 2
        If EMPSAL_GROSS >= rsTmp!Col3 And EMPSAL_GROSS <= (rsTmp!Col4 - 1) Then RESULT_TAX = rsTmp!Col3: COLNO = 3
        If EMPSAL_GROSS >= rsTmp!Col4 And EMPSAL_GROSS <= (rsTmp!Col5 - 1) Then RESULT_TAX = rsTmp!Col4: COLNO = 4
        If EMPSAL_GROSS >= rsTmp!Col5 And EMPSAL_GROSS <= (rsTmp!Col6 - 1) Then RESULT_TAX = rsTmp!Col5: COLNO = 5
        If EMPSAL_GROSS >= rsTmp!Col6 And EMPSAL_GROSS <= (rsTmp!Col7 - 1) Then RESULT_TAX = rsTmp!Col6: COLNO = 6
        If EMPSAL_GROSS >= rsTmp!Col7 And EMPSAL_GROSS <= (rsTmp!Col8 - 1) Then RESULT_TAX = rsTmp!Col7: COLNO = 7
        If EMPSAL_GROSS >= rsTmp!Col8 Then RESULT_TAX = rsTmp!Col8: COLNO = 8
        
        Set RSTAX = gconDMIS.Execute("Select * From HRMS_taxTable Where TaxBasis = '" & cboType & "'")
        If Not (RSTAX.BOF And RSTAX.EOF) Then
            If COLNO = 1 Then RESULT = 1
            If COLNO = 2 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per2) + RSTAX!EXp2
            If COLNO = 3 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per3) + RSTAX!EXp3
            If COLNO = 4 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per4) + RSTAX!EXp4
            If COLNO = 5 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per5) + RSTAX!EXp5
            If COLNO = 6 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per6) + RSTAX!EXp6
            If COLNO = 7 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per7) + RSTAX!EXp7
            If COLNO = 8 Then RESULT = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per8) + RSTAX!EXp8
        End If
    End If
    
    lblTax.Caption = Format(RESULT, "#,###,##0.00")
    
    Set RSTAX = Nothing
    Set rsTmp = Nothing
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub
