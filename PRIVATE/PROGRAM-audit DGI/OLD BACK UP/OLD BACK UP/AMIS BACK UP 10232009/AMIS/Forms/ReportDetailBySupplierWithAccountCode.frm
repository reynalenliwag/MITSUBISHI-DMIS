VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISDetailBySupplierWithAccountCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Detail By Account Code"
   ClientHeight    =   2670
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4815
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ReportDetailBySupplierWithAccountCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4815
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2340
      MouseIcon       =   "ReportDetailBySupplierWithAccountCode.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportDetailBySupplierWithAccountCode.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close Window"
      Top             =   1740
      Width           =   885
   End
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
      Height          =   825
      Left            =   1470
      MouseIcon       =   "ReportDetailBySupplierWithAccountCode.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportDetailBySupplierWithAccountCode.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print Report"
      Top             =   1740
      Width           =   885
   End
   Begin VB.ComboBox cboVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   90
      Width           =   4695
   End
   Begin VB.ComboBox cboAcct_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   360
      Left            =   1260
      TabIndex        =   1
      Top             =   450
      Width           =   3465
   End
   Begin Crystal.CrystalReport rptAMISrange 
      Left            =   120
      Top             =   2100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   810
      TabIndex        =   5
      Top             =   1230
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53149697
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3060
      TabIndex        =   7
      Top             =   1230
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   53149697
      CurrentDate     =   38216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2580
      TabIndex        =   6
      Top             =   1290
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1290
      Width           =   675
   End
   Begin VB.Label Label34 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Account No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   510
      Width           =   1305
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2730
      TabIndex        =   10
      Top             =   2130
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISDetailBySupplierWithAccountCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                                      As ADODB.Recordset
Dim rsChartAccount                                                    As ADODB.Recordset
Dim rsVENDOR                                                          As ADODB.Recordset

Function SetAccountName(VVV As Variant) As String
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountName = Null2String(rsChartAccount!AcctCode)
    End If
End Function

Function SetVendorCode(VVV As Variant) As String
    If REPORT_RANGETYPE = "APJ" Then
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "Select Code,NameOfVendor from ALL_Vendor where NameOfVendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            SetVendorCode = Null2String(rsVENDOR!code)
        End If
    ElseIf REPORT_RANGETYPE = "REC_REGISTER" Then
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "Select Code,NameOfVendor from ALL_Vendor where NameOfVendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            SetVendorCode = Null2String(rsVENDOR!code)
        End If
    Else
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "Select CustCode,AcctName from ALL_CUSTMASTER_AMIS where AcctName = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            SetVendorCode = Null2String(rsVENDOR!CUSTCODE)
        End If
    End If
End Function

Sub InitCbo()
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select Description from AMIS_ChartAccount order by acctcode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
    If REPORT_RANGETYPE = "APJ" Then
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "select NameOfVendor from ALL_Vendor Where NameOfVendor <> '' order by NameOfVendor asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            Combo_Loadval cboVendor, rsVENDOR
        End If
    ElseIf REPORT_RANGETYPE = "REC_REGISTER" Then
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "select NameOfVendor from ALL_Vendor Where NameOfVendor <> '' order by NameOfVendor asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            Combo_Loadval cboVendor, rsVENDOR
        End If
    Else
        Set rsVENDOR = New ADODB.Recordset
        rsVENDOR.Open "select AcctName from ALL_CUSTMASTER_AMIS where AcctName <> '' order by AcctName asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            Combo_Loadval cboVendor, rsVENDOR
        End If
    End If
End Sub

Private Sub cboAcct_Code_Change()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_Click()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_GotFocus()
    If cboAcct_Code.Text = "" Then
        Call VBComBoBoxDroppedDown(cboAcct_Code)
    End If
End Sub

Private Sub cboAcct_Code_LostFocus()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

Private Sub cboVendor_GotFocus()
    If cboVendor.Text = "" Then
        Call VBComBoBoxDroppedDown(cboVendor)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200713:02
Private Sub cmdPrint_Click()

    On Error GoTo Errorcode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        If REPORT_RANGETYPE = "APJ" Then
            If Function_Access(LOGID, "Acess_Print", "ACCOUNTS DETAIL BY SUPPLIERS") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "APJDetailBySupplier", "AccountDetail", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' AND {Journal_Hd.VendorCode} = '" & SetVendorCode(cboVendor.Text) & "'", "Account Detail Report By Supplier", False
            LogAudit "V", "ACCOUNTS DETAIL BY SUPPLIERS", cboVendor & ":" & cboAcct_Code & dtpFrom & "-" & dtpTo
            Call NEW_LogAudit("G", "ACCOUNTS DETAIL BY SUPPLIERS", "", "", "", cboAcct_Code & " " & dtpFrom & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "SJ" Then
            If Function_Access(LOGID, "Acess_Print", "ACCOUNTS DETAIL BY CUSTOMER") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "SJDetailByCustomer", "AccountDetail", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' AND {Journal_Hd.CustomerCode} = '" & SetVendorCode(cboVendor.Text) & "'", "Account Detail Report By Customer", False
            Call NEW_LogAudit("V", "ACCOUNTS DETAIL BY CUSTOMER", "", "", "", cboAcct_Code & " " & dtpFrom & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "REC_REGISTER" Then
            If Function_Access(LOGID, "Acess_Print", "RECEIVING REPORT REGISTER") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "ReceiptsRegisters", "Registers", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' AND {Journal_Hd.VendorCode} = '" & SetVendorCode(cboVendor.Text) & "' and {Journal_Hd.jtype} = 'APJ' ", "RECEIVING REPORT REGISTERS", False
            Call NEW_LogAudit("V", "RECEIVING REPORT REGISTER", "", "", "", cboAcct_Code & " " & dtpFrom & dtpTo, "", "")
        End If
    Else
        ShowNoRecord
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
     Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If REPORT_RANGETYPE = "APJ" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCOUNTS DETAIL BY SUPPLIERS)"
                Call frmALL_AuditInquiry.DisplayHistory("", "ACCOUNTS DETAIL BY SUPPLIERS", "PRINTING")
            ElseIf REPORT_RANGETYPE = "SJ" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCOUNTS DETAIL BY CUSTOMER)"
                Call frmALL_AuditInquiry.DisplayHistory("", "ACCOUNTS DETAIL BY CUSTOMER", "PRINTING")
            ElseIf REPORT_RANGETYPE = "REC_REGISTER" Then
               frmALL_AuditInquiry.Caption = "Audit Inquiry (RECEIVING REPORT REGISTER)"
               Call frmALL_AuditInquiry.DisplayHistory("", "RECEIVING REPORT REGISTER", "PRINTING")
            End If
        End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    InitCbo
    Screen.MousePointer = 0
End Sub

