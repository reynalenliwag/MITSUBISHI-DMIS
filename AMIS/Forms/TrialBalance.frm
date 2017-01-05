VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISTrialBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Balance"
   ClientHeight    =   3015
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4470
   ForeColor       =   &H00FFFFFF&
   Icon            =   "TrialBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdPrint_ASof 
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
      Left            =   1320
      MouseIcon       =   "TrialBalance.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "TrialBalance.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print Report"
      Top             =   2130
      Width           =   885
   End
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
      MouseIcon       =   "TrialBalance.frx":1433
      MousePointer    =   99  'Custom
      Picture         =   "TrialBalance.frx":1585
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Close Window"
      Top             =   2130
      Width           =   885
   End
   Begin VB.Frame picPeriod 
      Height          =   585
      Left            =   150
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   7
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56295425
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2730
         TabIndex        =   9
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61931521
         CurrentDate     =   38216
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   2220
         TabIndex        =   8
         Top             =   210
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   210
         Width           =   675
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Trial Balance for the Period"
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
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1140
      Width           =   2925
   End
   Begin VB.OptionButton optBalances 
      Caption         =   "Trial Balance of Balances"
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
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   570
      Value           =   -1  'True
      Width           =   2715
   End
   Begin VB.OptionButton optTotals 
      Caption         =   "Trial Balance of Totals"
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
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   870
      Width           =   2715
   End
   Begin Crystal.CrystalReport rptAMISTrialBalance 
      Left            =   300
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Income From Insurance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   1740
      TabIndex        =   1
      Top             =   90
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
      Format          =   62259201
      CurrentDate     =   38216
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
      MouseIcon       =   "TrialBalance.frx":19D0
      MousePointer    =   99  'Custom
      Picture         =   "TrialBalance.frx":1B22
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print Report"
      Top             =   2130
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "As Of:"
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
      Left            =   900
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As ADODB.Recordset

Function Setacctname(XXX As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount where AcctCode = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Setacctname = Null2String(rsChartAccount!DESCRIPTION)
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Printing for as of
'optBalances
'optTotals
Private Sub cmdPrint_ASof_Click()
    On Error GoTo ErrorCode:
    Dim DEBIT_BALANCE                                       As Double
    Dim CREDIT_BALANCE                                      As Double
    Dim TOTAL_DEBIT_BALANCE                                 As Double
    Dim TOTAL_CREDIT_BALANCE                                As Double
    
    If IsDate(dtpAsOF) = False Then
        MsgSpeechBox "Error In As of date"
        Exit Sub
    End If
        
        If Option1.Value = True Then
        If COMPANY_CODE <> "HCE" Then
            If CheckUnposted(dtpFrom, dtpTo) = True Then
                MessagePop InfoWarning, "Unposted Transactions", "Unable to generate report. Please check for the unposted transactions."
            End If
            Exit Sub
                       
            Else
                gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                Set rsJournal_HD = New ADODB.Recordset
                rsJournal_HD.Open "select * from AMIS_Journal_Det where (jdate >= '" & dtpFrom & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenDynamic
                If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
                    Set rsJournal_HD = New ADODB.Recordset
                    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','CCM','COB','VPJ','CDM','VDM','VCM') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    ElseIf COMPANY_CODE = "DSSC" Then
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype <> 'CLO' and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    Else
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','COB','VPJ') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    End If
                    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                        rsJournal_HD.MoveFirst
                        TOTAL_DEBIT_BALANCE = 0: TOTAL_CREDIT_BALANCE = 0
                        Screen.MousePointer = 11
                        Do While Not rsJournal_HD.EOF
                            If NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) Then
                                DEBIT_BALANCE = NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL))
                                CREDIT_BALANCE = 0
                            Else
                                If NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) Then
                                    CREDIT_BALANCE = NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL))
                                    DEBIT_BALANCE = 0
                                Else
                                    CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                                End If
                            End If
                            TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                            TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                            gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                             " Debit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)), 2) & "," & _
                                             " Credit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)), 2) & "," & _
                                             " DebitBalance = " & DEBIT_BALANCE & "," & _
                                             " CreditBalance = " & CREDIT_BALANCE & _
                                             " Where AcctCode = '" & Null2String(rsJournal_HD!ACCT_CODE) & "'"
                            rsJournal_HD.MoveNext
                        Loop
                        Screen.MousePointer = 0
                        ShowReport "TrialBalanceOfBalances", "FinancialStatement", "", "Trial Balance for the Period", "From: " & dtpFrom & " To: " & dtpTo, True
                    End If
                Else
                    ShowNoRecord
                End If
                Call NEW_LogAudit("V", "TrialBalanceOfBalances", "", "", "", dtpFrom & " " & dtpTo, "", "")
                
            End If
        Else

            If CheckUnpostedAsOf(dtpAsOF) = True Then
'                If COMPANY_CODE = "HCE" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                    GoTo skipvalidation
'                Else
'                    MessagePop InfoWarning, "Unposted Transactions", "Unable to generate report. Please check for the unposted transactions."
'                Exit Sub
'                End If
           Else
skipvalidation:
        
                gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                Set rsJournal_HD = New ADODB.Recordset
                'rsJournal_Hd.Open "select * from AMIS_Journal_Det where (jdate <= '" & CDate(dtpAsOF) & "') and year(jdate) = " & Year(dtpAsOF), gconDMIS, adOpenForwardOnly, adLockReadOnly
                rsJournal_HD.Open "select * from AMIS_Journal_Det where (jdate <= '" & CDate(dtpAsOF) & "')", gconDMIS, adOpenDynamic
                If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
                    If optBalances.Value = True Then
                        Screen.MousePointer = 11
                        Set rsJournal_HD = New ADODB.Recordset
                        'rsJournal_Hd.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "' and year(jdate) = " & Year(dtpTo) & ") group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                        'If COMPANY_CODE = "JMC" Then
                            If CheckYearEndClosing(dtpAsOF.Value) = True Then
                                If COMPANY_CODE = "DGI" Then
                                    rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('COB','VPJ','CCM','CDM','VCM','VDM') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                                ElseIf COMPANY_CODE = "DSSC" Then
                                     'rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype <> 'CLO' and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                                    'OLD QUERY FOR DSSC
                                    rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('COB','VPJ') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
'Update by Julie: to correct the result of TB
'                                ElseIf COMPANY_CODE = "DJM" Then
'                                rsJournal_HD.Open "(select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE " & _
'                                            "From AMIS_Journal_Det " & _
'                                            "where jtype NOT IN ('CCM','COB','VPJ','CLO') and Status = 'P' AND (jdate <= ('" & CDate(dtpAsOF) & "')) " & _
'                                            "group by ACCT_CODE )order by ACCT_CODE asc", gconDMIS, adOpenDynamic

                                Else
                                    rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where LEFT(ACCT_CODE,1) <= 3 AND jtype NOT IN ('CCM','COB','VPJ') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                                End If
                        'End If
                        Else
                            If COMPANY_CODE = "HCE" Then
                                rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('COB','VPJ') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                            ElseIf COMPANY_CODE = "DSSC" Then
                                'rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype <> 'CLO' and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                                'OLD QUERY FOR DSSC
                                rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('COB','VPJ') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                            ElseIf COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
                                rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('COB','VPJ','CCM','CDM','VCM','VDM') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                            ElseIf COMPANY_CODE = "HMR" Then
                                rsJournal_HD.Open "(select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det " & _
                                            "where jtype NOT IN ('CCM','COB','VPJ') and Status = 'P' AND Acct_Code = (select ACCTCODE from AMIS_CHARTACCOUNT where TRANTYPE1 = 'REU') and  (Year(jdate) <= '" & Year(dtpAsOF) & "'-1) " & _
                                            "group by ACCT_CODE) " & _
                                            "Union All " & _
                                            "(select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE " & _
                                            "From AMIS_Journal_Det " & _
                                            "where jtype NOT IN ('CCM','COB','VPJ') and Status = 'P' AND Acct_Code not in (select ACCTCODE from AMIS_CHARTACCOUNT where TRANTYPE1 = 'REU') and  (jdate <= ('" & CDate(dtpAsOF) & "')) " & _
                                            "group by ACCT_CODE )order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                            
                            Else
                                rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CCM','COB','VPJ') and Status = 'P' AND (jdate <= '" & CDate(dtpAsOF) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                            End If
                        End If
                        If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                            rsJournal_HD.MoveFirst
                            TOTAL_DEBIT_BALANCE = 0
                            TOTAL_CREDIT_BALANCE = 0
                            Screen.MousePointer = 11
                            Do While Not rsJournal_HD.EOF
                                'MsgBox "CODE : " & Null2String(rsJournal_HD!ACCT_CODE) & vbCrLf & _
                                 '       "NAME : " & SetAcctName(Null2String(rsJournal_HD!ACCT_CODE)) & vbCrLf & _
                                 '       "DEBIT: " & N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) & " CREDIT: " & N2Str2Zero(rsJournal_HD!CREDIT_TOTAL)
                                    If N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) > N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) Then
                                    DEBIT_BALANCE = N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) - N2Str2Zero(rsJournal_HD!CREDIT_TOTAL)
                                    CREDIT_BALANCE = 0
                                Else
                                    If N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) > N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) Then
                                        CREDIT_BALANCE = N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) - N2Str2Zero(rsJournal_HD!DEBIT_TOTAL)
                                        DEBIT_BALANCE = 0
                                    Else
                                        CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                                    End If
                                End If
                                TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                                TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                                gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                                 " Debit_Total = " & N2Str2Zero(rsJournal_HD!DEBIT_TOTAL) & "," & _
                                                 " Credit_Total = " & N2Str2Zero(rsJournal_HD!CREDIT_TOTAL) & "," & _
                                                 " DebitBalance = " & DEBIT_BALANCE & "," & _
                                                 " CreditBalance = " & CREDIT_BALANCE & _
                                                 " Where AcctCode = '" & Null2String(rsJournal_HD!ACCT_CODE) & "'"
                                'gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                 '                 " Debit_Total = " & DEBIT_BALANCE & "," & _
                                 '                 " Credit_Total = " & CREDIT_BALANCE & "," & _
                                 '                 " DebitBalance = " & DEBIT_BALANCE & "," & _
                                 '                 " CreditBalance = " & CREDIT_BALANCE & _
                                 '                 " Where AcctCode = '" & Null2String(rsJOURNAL_HD!acct_code) & "'"
                                rsJournal_HD.MoveNext
                            Loop
                            Screen.MousePointer = 0
                        End If
                        Screen.MousePointer = 0
                        
                        If COMPANY_CODE = "DMI" Then
                            ShowReport "TrialBalanceOfBalances", "FinancialStatement", "", "Trial Balance", "AS OF: " & dtpAsOF, True
                        Else
                            ShowReport "TrialBalanceOfBalances", "FinancialStatement", "", "Trial Balance of Balances", "AS OF: " & dtpAsOF, True
                        End If
                        
                    Else
                        'ShowReport "TrialBalanceOfTotals", "FinancialStatement", "({AMIS_Journal_Det.jdate} <= date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")) and year({AMIS_Journal_Det.jdate}) = " & Year(dtpTo), "Trial Balance of Totals", "AS OF: " & dtpAsOF, True
                        ShowReport "TrialBalanceOfTotals", "FinancialStatement", "({AMIS_Journal_Det.jdate} <= date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & "))", "Trial Balance of Totals", "AS OF: " & dtpAsOF, True
                    End If
                    'Unload Me
                Else
                    ShowNoRecord
                End If
            End If
            
        End If
          
    
    Call NEW_LogAudit("V", "TTrialBalanceOfTotals", "", "", "", dtpFrom & " " & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub


'Upating Code       : AXP-0713200714:19
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    Dim DEBIT_BALANCE                                       As Double
    Dim CREDIT_BALANCE                                      As Double
    Dim TOTAL_DEBIT_BALANCE                                 As Double
    Dim TOTAL_CREDIT_BALANCE                                As Double
    
    If IsDate(dtpAsOF) = False Then
        MsgSpeechBox "Error In As of date"
        Exit Sub
    End If
    
    
        If Option1.Value = True Then
        If COMPANY_CODE <> "HCE" And COMPANY_CODE <> "DGI" And COMPANY_CODE <> "HMH" And COMPANY_CODE <> "DSSC" Then
            If CheckUnposted(dtpFrom, dtpTo) = True Then
                'SJR 09042014
                If COMPANY_CODE = "HCA" Then
                GoTo skipvalidation
                Else
                MessagePop InfoWarning, "Unposted Transactions", "Unable to generate report. Please check for the unposted transactions."
                Exit Sub
                End If
                       
            Else
                gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                Set rsJournal_HD = New ADODB.Recordset
                rsJournal_HD.Open "select * from AMIS_Journal_Det where (jdate >= '" & dtpFrom & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenDynamic
                If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
                    Set rsJournal_HD = New ADODB.Recordset
                    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','CCM','COB','VPJ','CDM','VDM','VCM') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    Else
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','COB','VPJ') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    End If
                    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                        rsJournal_HD.MoveFirst
                        TOTAL_DEBIT_BALANCE = 0: TOTAL_CREDIT_BALANCE = 0
                        Screen.MousePointer = 11
                        Do While Not rsJournal_HD.EOF
                            If NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) Then
                                DEBIT_BALANCE = NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL))
                                CREDIT_BALANCE = 0
                            Else
                                If NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) Then
                                    CREDIT_BALANCE = NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL))
                                    DEBIT_BALANCE = 0
                                Else
                                    CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                                End If
                            End If
                            TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                            TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                            gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                             " Debit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)), 2) & "," & _
                                             " Credit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)), 2) & "," & _
                                             " DebitBalance = " & DEBIT_BALANCE & "," & _
                                             " CreditBalance = " & CREDIT_BALANCE & _
                                             " Where AcctCode = '" & Null2String(rsJournal_HD!ACCT_CODE) & "'"
                            'gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                             '                 " Debit_Total = " & DEBIT_BALANCE & "," & _
                             '                 " Credit_Total = " & CREDIT_BALANCE & "," & _
                             '                 " DebitBalance = " & DEBIT_BALANCE & "," & _
                             '                 " CreditBalance = " & CREDIT_BALANCE & _
                             '                 " Where AcctCode = '" & Null2String(rsJOURNAL_HD!acct_code) & "'"
                            rsJournal_HD.MoveNext
                        Loop
                        Screen.MousePointer = 0
                        ShowReport "TrialBalanceOfBalances", "FinancialStatement", "", "Trial Balance for the Period", "From: " & dtpFrom & " To: " & dtpTo, True
                    End If
                Else
                    ShowNoRecord
                End If
                Call NEW_LogAudit("V", "TrialBalanceOfBalances", "", "", "", dtpFrom & " " & dtpTo, "", "")
                
            End If
        Else

            If CheckUnpostedAsOf(dtpAsOF) = True Then
                If COMPANY_CODE = "HCE" Or COMPANY_CODE = "HCA" Or COMPANY_CODE <> "DGI" Or COMPANY_CODE <> "HMH" Or COMPANY_CODE <> "DSSC" Then
                    GoTo skipvalidation
                Else
                    MessagePop InfoWarning, "Unposted Transactions", "Unable to generate report. Please check for the unposted transactions."
                Exit Sub
                End If
            
            Else
skipvalidation:
                gconDMIS.Execute "update AMIS_ChartAccount Set Debit_Total = 0,Credit_Total = 0,DebitBalance = 0,CreditBalance = 0"
                Set rsJournal_HD = New ADODB.Recordset
                rsJournal_HD.Open "select * from AMIS_Journal_Det where (jdate >= '" & dtpFrom & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenDynamic
                If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
                    Set rsJournal_HD = New ADODB.Recordset
                    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','CCM','COB','VPJ','CDM','VDM','VCM') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    Else
                        rsJournal_HD.Open "select SUM(ISNULL(DEBIT,0)) AS DEBIT_TOTAL, SUM(ISNULL(CREDIT,0)) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where jtype NOT IN ('CLO','COB','VPJ') and Status = 'P' AND (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "') group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
                    End If
                    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                        rsJournal_HD.MoveFirst
                        TOTAL_DEBIT_BALANCE = 0: TOTAL_CREDIT_BALANCE = 0
                        Screen.MousePointer = 11
                        Do While Not rsJournal_HD.EOF
                            If NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) Then
                                DEBIT_BALANCE = NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL))
                                CREDIT_BALANCE = 0
                            Else
                                If NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) Then
                                    CREDIT_BALANCE = NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL))
                                    DEBIT_BALANCE = 0
                                Else
                                    CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                                End If
                            End If
                            TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                            TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                            gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                                             " Debit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)), 2) & "," & _
                                             " Credit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)), 2) & "," & _
                                             " DebitBalance = " & DEBIT_BALANCE & "," & _
                                             " CreditBalance = " & CREDIT_BALANCE & _
                                             " Where AcctCode = '" & Null2String(rsJournal_HD!ACCT_CODE) & "'"
                            rsJournal_HD.MoveNext
                        Loop
                        Screen.MousePointer = 0
                        ShowReport "TrialBalanceOfBalances", "FinancialStatement", "", "Trial Balance for the Period", "From: " & dtpFrom & " To: " & dtpTo, True
                    End If
                Else
                    ShowNoRecord
                End If
                Call NEW_LogAudit("V", "TrialBalanceOfBalances", "", "", "", dtpFrom & " " & dtpTo, "", "")
                
            End If
           End If
'
      End If
    Call NEW_LogAudit("V", "TTrialBalanceOfTotals", "", "", "", dtpFrom & " " & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry

        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        If Option1.Value = True Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TrialBalanceOfBalances)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TrialBalanceOfBalances", "PRINTING")
        Else
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TTrialBalanceOfTotals)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TTrialBalanceOfTotals", "PRINTING")
        End If

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub optBalances_Click()
    cmdPrint_ASof.ZOrder 0
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub

Private Sub Option1_Click()
    cmdPrint_ASof.ZOrder 1
    If Option1.Value = True Then
        picPeriod.Enabled = True
        dtpAsOF.Enabled = False
        dtpFrom.Enabled = True
        dtpTo.Enabled = True
    Else
        picPeriod.Enabled = False
        dtpAsOF.Enabled = True
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
    End If
End Sub

Private Sub optTotals_Click()
    cmdPrint_ASof.ZOrder 0
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub



