VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISProcessUnusedInvoices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unused Invoices"
   ClientHeight    =   1875
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4515
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ProcessUnusedInvoices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4515
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
      Left            =   2220
      MouseIcon       =   "ProcessUnusedInvoices.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUnusedInvoices.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   945
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
      Left            =   1350
      MouseIcon       =   "ProcessUnusedInvoices.frx":0E67
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUnusedInvoices.frx":0FB9
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   945
      Width           =   885
   End
   Begin VB.ComboBox cboInvoiceType 
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
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   4245
   End
   Begin Crystal.CrystalReport rptAMISSnusedInvoices 
      Left            =   120
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Unused Invoices"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   510
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48758785
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   510
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48758785
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   540
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   675
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   7
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISProcessUnusedInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoiceType                                      As ADODB.Recordset
Dim rsINVOICES                                         As ADODB.Recordset

Function SetInvCode(XXX As String)
    Dim SQL                                            As String
    Dim rsInvoiceType                                  As New ADODB.Recordset

    On Error Resume Next

    SQL = "Select InvCode from ALL_InvoiceType Where InvType = '" & XXX & "'"

    Set rsInvoiceType = New ADODB.Recordset
    Set rsInvoiceType = gconDMIS.Execute(SQL)
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvCode = Null2String(rsInvoiceType!InvCode)
    End If
End Function

Private Sub cboInvoiceType_Click()
    On Error Resume Next
    Set rsINVOICES = New ADODB.Recordset
    Set rsINVOICES = gconDMIS.Execute("Select MIN(jdate) as FirstInvNo, MAX(jdate) as LastInvNo from AMIS_Journal_HD Where Jtype = 'SJ' and InvoiceType = '" & SetInvCode(cboInvoiceType) & "' and status = 'P'")
    If Not rsINVOICES.EOF And Not rsINVOICES.BOF Then
        cmdPrint.Enabled = True
        If IsNull(rsINVOICES!FirstInvNo) = False Then
            dtpFrom = Null2String(rsINVOICES!FirstInvNo)
        End If

        If IsNull(rsINVOICES!LastInvNo) = False Then
            dtpTo = Null2String(rsINVOICES!LastInvNo)
        End If

    Else
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:01
Private Sub cmdPrint_Click()
'If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub

    Dim InvoiceCode, Findings                          As String
    On Error GoTo ErrorCode:

    InvoiceCode = SetInvCode(cboInvoiceType)
    Dim CurrentSeries                                  As String
    Dim ISeries                                        As Long
    Dim rsFINDINGS, rsFINDINGS2                        As ADODB.Recordset
    Set rsINVOICES = New ADODB.Recordset
    Set rsINVOICES = gconDMIS.Execute("Select invoiceno,jno from AMIS_Journal_HD where status = 'P' and jtype = 'SJ' and invoicetype = '" & InvoiceCode & "' and jdate >= '" & dtpFrom & "' and jdate <= '" & dtpTo & "' order by invoiceno asc")

    If Not rsINVOICES.EOF And Not rsINVOICES.BOF Then
        rsINVOICES.MoveFirst
        CurrentSeries = Null2String(rsINVOICES!INVOICENO)
        ISeries = NumericVal(CurrentSeries)
        Screen.MousePointer = 11
        If InvoiceCode = "SI" Then
            'Dim gconDMIS As ADODB.Connection
            'Set gconDMIS = New ADODB.Connection
            'deCSMIOS.deConnCSMIOS.ConnectionString = CSMIOS_Connection
            'Set gconDMIS = deCSMIOS.deConnCSMIOS
            'gconDMIS.Open
        End If
        If InvoiceCode = "CI" Or InvoiceCode = "CGI" Then
            'Dim gconDMIS As ADODB.Connection
            'Set gconDMIS = New ADODB.Connection
            'dePMIOS.deConnPMIOS.ConnectionString = PMIOS_Connection
            'Set gconDMIS = dePMIOS.deConnPMIOS
            'gconDMIS.Open
        End If
        If InvoiceCode = "VI" Then
            '            Dim gconDMIS As ADODB.Connection
            '            Set gconDMIS = New ADODB.Connection
            '            deSMIS.deConnSMIS.ConnectionString = SMIS_Connection
            '            Set gconDMIS = deSMIS.deConnSMIS
            '            gconDMIS.Open
        End If

        gconDMIS.Execute ("delete from AMIS_UnUsedInvoices where invcode = '" & InvoiceCode & "'")
        Do While Not rsINVOICES.EOF
            CurrentSeries = Null2String(rsINVOICES!INVOICENO)
            If Format(ISeries, "000000") <> Format(CurrentSeries, "000000") Then
                Set rsFINDINGS = New ADODB.Recordset
                If InvoiceCode = "SI" Then
                    Set rsFINDINGS = gconDMIS.Execute("Select Invoice from CSMS_REPOR Where Invoice = '" & Format(ISeries, "000000") & "'")
                    If rsFINDINGS.EOF And rsFINDINGS.BOF Then
                        Findings = "UNUSED"
                    Else
                        Findings = "MISSING"
                    End If
                End If
                If InvoiceCode = "CI" Then
                    Set rsFINDINGS = gconDMIS.Execute("Select status from PMIS_ord_hd Where tranno = '" & Format(ISeries, "000000") & "' and trantype = 'CSH'")
                    If rsFINDINGS.EOF And rsFINDINGS.BOF Then
                        Set rsFINDINGS2 = gconDMIS.Execute("Select status from PMIS_ord_hist Where tranno = '" & Format(ISeries, "000000") & "' and trantype = 'CSH'")
                        If rsFINDINGS2.EOF And rsFINDINGS2.BOF Then
                            Findings = "UNUSED"
                        Else
                            If rsFINDINGS2!Status = "C" Then
                                Findings = "CANCELLED"
                            Else
                                Findings = "MISSING"
                            End If
                        End If
                    Else
                        If rsFINDINGS!Status = "C" Then
                            Findings = "CANCELLED"
                        Else
                            Findings = "MISSING"
                        End If
                    End If
                End If
                If InvoiceCode = "CGI" Then
                    Set rsFINDINGS = gconDMIS.Execute("Select status from PMIS_ord_hd Where tranno = '" & Format(ISeries, "000000") & "' and trantype = 'CHG'")
                    If rsFINDINGS.EOF And rsFINDINGS.BOF Then
                        Set rsFINDINGS2 = gconDMIS.Execute("Select status from PMIS_ord_hist Where tranno = '" & Format(ISeries, "000000") & "' and trantype = 'CHG'")
                        If rsFINDINGS2.EOF And rsFINDINGS2.BOF Then
                            Findings = "UNUSED"
                        Else
                            If rsFINDINGS2!Status = "C" Then
                                Findings = "CANCELLED"
                            Else
                                Findings = "MISSING"
                            End If
                        End If
                    Else
                        If rsFINDINGS!Status = "C" Then
                            Findings = "CANCELLED"
                        Else
                            Findings = "MISSING"
                        End If
                    End If
                End If
                If InvoiceCode = "VI" Then
                    Set rsFINDINGS = gconDMIS.Execute("Select VI_NO from SMIS_PURCHAGREE Where VI_NO = '" & Format(ISeries, "000000") & "'")
                    If rsFINDINGS.EOF And rsFINDINGS.BOF Then
                        Findings = "UNUSED"
                    Else
                        Findings = "MISSING"
                    End If
                End If
                gconDMIS.Execute ("Insert into AMIS_UnUsedInvoices (Invoiceno,InvCode,Findings) values ('" & Format(ISeries, "000000") & "','" & InvoiceCode & "','" & Findings & "')")
                ISeries = NumericVal(CurrentSeries)
            Else
                Set rsFINDINGS = New ADODB.Recordset
                Set rsFINDINGS = gconDMIS.Execute("Select jno from AMIS_Journal_HD Where Jno = '" & rsINVOICES!JNo & "' and status = 'N'")
                If Not rsFINDINGS.EOF And Not rsFINDINGS.BOF Then
                    Findings = "UNPOSTED"
                    gconDMIS.Execute ("Insert into AMIS_UnUsedInvoices (Invoiceno,InvCode,Findings) values ('" & Format(ISeries, "000000") & "','" & InvoiceCode & "','" & Findings & "')")
                End If
                Set rsFINDINGS = New ADODB.Recordset
                Set rsFINDINGS = gconDMIS.Execute("Select jno from AMIS_Journal_HD Where Jno = '" & rsINVOICES!JNo & "' and status = 'C'")
                If Not rsFINDINGS.EOF And Not rsFINDINGS.BOF Then
                    Findings = "CANCELLED"
                    gconDMIS.Execute ("Insert into AMIS_UnUsedInvoices (Invoiceno,InvCode,Findings) values ('" & Format(ISeries, "000000") & "','" & InvoiceCode & "','" & Findings & "')")
                End If
            End If
            rsINVOICES.MoveNext
            ISeries = ISeries + 1
        Loop
        '        If InvoiceCode = "SI" Then gconDMIS.Close
        '        If InvoiceCode = "CI" Then gconDMIS.Close
        '        If InvoiceCode = "CGI" Then gconDMIS.Close
        'If InvoiceCode = "VI" Then gconDMIS.Close
        ShowRangeReport dtpFrom, dtpTo, "UnusedInvoices", "InvoicesReport", "{UnusedInvoices.InvCode} = '" & InvoiceCode & "'", cboInvoiceType.Text & " Unused Invoices", False
        'PrintSQLReport rptAMISSnusedInvoices, AMIS_REPORT_PATH & "\InvoicesReport\UnusedInvoices.rpt", "{UnusedInvoices.InvCode} = '" & InvoiceCode & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("V", "UNUSED INVOICES", "", "", "", dtpFrom & " " & dtpFrom, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry

        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "Audit Inquiry (UNUSED INVOICES)"
        Call frmALL_AuditInquiry.DisplayHistory("", "UNUSED INVOICES", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 11
    Set rsInvoiceType = New ADODB.Recordset
    Set rsInvoiceType = gconDMIS.Execute("Select InvType from ALL_InvoiceType order by id asc")
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        Combo_Loadval cboInvoiceType, rsInvoiceType
    End If
    cmdPrint.Enabled = False
    Screen.MousePointer = 0
End Sub

