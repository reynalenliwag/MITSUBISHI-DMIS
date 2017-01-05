VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISRecurringEntries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Recurring Journal Entries"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4665
   Begin VB.CheckBox chkReverse 
      Caption         =   "Reverse Transaction"
      Height          =   345
      Left            =   90
      TabIndex        =   13
      Top             =   900
      Width           =   2385
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3000
      TabIndex        =   9
      Top             =   3690
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   1380
      TabIndex        =   8
      Top             =   3690
      Width           =   1545
   End
   Begin VB.Frame Frame 
      Caption         =   "Date Range"
      Height          =   1755
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   4485
      Begin VB.TextBox txtOccurences 
         Height          =   285
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   1350
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   345
         Left            =   2670
         TabIndex        =   4
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   119144449
         CurrentDate     =   40877
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   345
         Left            =   2670
         TabIndex        =   5
         Top             =   930
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         Format          =   119144449
         CurrentDate     =   40877
      End
      Begin VB.Label lblJDATE 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "occurence(s)"
         Height          =   195
         Index           =   3
         Left            =   3090
         TabIndex        =   11
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End On:"
         Height          =   195
         Index           =   2
         Left            =   1710
         TabIndex        =   6
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Transaction Date:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   2250
      End
   End
   Begin VB.ComboBox cboOption 
      Height          =   315
      ItemData        =   "frmAMISRecurringEntries.frx":0000
      Left            =   2160
      List            =   "frmAMISRecurringEntries.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   2115
   End
   Begin VB.Label lblRecur 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This transaction will recur on the 15th every month."
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3300
      Width           =   4440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "How often do you want to recur this transactions?"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Width           =   4245
   End
End
Attribute VB_Name = "frmAMISRecurringEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xJOURNALTYPE As String
Dim xVOUCHERNO As String

Private Sub cboOption_Click()
    If cboOption.Text = "Quarterly" Then
        dtTo.Value = DateAdd("yyyy", 1, dtFrom.Value)
        txtOccurences.Text = "4"
    ElseIf cboOption.Text = "Twice a year" Then
        dtTo.Value = DateAdd("yyyy", 1, dtFrom.Value)
        txtOccurences.Text = "2"
    ElseIf cboOption.Text = "Yearly" Then
        dtTo.Value = DateAdd("yyyy", 1, dtFrom.Value)
        txtOccurences.Text = "1"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
     If dtTo.Value <= dtFrom.Value Then
        MsgBox "Please check date selected.", vbInformation, "Date Range"
        Exit Sub
    End If
    Call SaveJournalEntries(xJOURNALTYPE, xVOUCHERNO)
End Sub

Private Sub dtFrom_Change()
    lblRecur.Caption = "This transaction will recur on the " & dtFrom.Day & "th every month."
End Sub

Private Sub dtTo_Change()
    txtOccurences.Text = DateDiff("m", dtFrom.Value, dtTo.Value)
End Sub

Private Sub dtTo_LostFocus()
    If dtTo.Value < dtFrom.Value Then
        MsgBox "Please check date selected.", vbInformation, "Date Range"
        dtTo.SetFocus
        Exit Sub
    End If
    
    If chkReverse.Value = False Then
        If dtTo.Day <= 31 Then
            dtTo.Value = lastDay(dtTo.Value)
        End If
    Else
        If dtTo.Day <= 31 Then
            dtTo.Value = firstDay(dtTo.Value)
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    cboOption.ListIndex = 0
    Screen.MousePointer = 0
End Sub

Sub LoadJournalDate(XXX As Date)
    dtFrom.Value = XXX
    dtTo.Value = XXX
    lblRecur.Caption = "This transaction will recur on the " & dtFrom.Day & "th every month."
End Sub

Sub LOADJOURNAL(JTYPE As String, VOUCHERNO As String)
    xJOURNALTYPE = JTYPE
    xVOUCHERNO = VOUCHERNO
End Sub

Sub SaveJournalEntries(xJType As String, xVOUCHERNO As String)
Dim xNEWVOUCHERNO  As String
Dim xNEWJNO As String
Dim Occurence As Integer

For Occurence = 1 To Val(txtOccurences.Text)
    xNEWVOUCHERNO = N2Str2Null(GetVoucherNo(xJType))
    xNEWJNO = N2Str2Null(GetJNo(xJType))
    If cboOption.Text = "Monthly" Then
        If chkReverse.Value = False Then
            lblJDATE.Caption = N2Str2Null(lastDay(DateAdd("m", Occurence, dtFrom.Value)))
        Else
            lblJDATE.Caption = N2Str2Null(firstDay(DateAdd("m", Occurence, dtFrom.Value)))
        End If
    ElseIf cboOption.Text = "Quarterly" Then
        If chkReverse.Value = False Then
            lblJDATE.Caption = N2Str2Null(lastDay(DateAdd("q", Occurence, dtFrom.Value)))
        Else
            lblJDATE.Caption = N2Str2Null(firstDay(DateAdd("q", Occurence, dtFrom.Value)))
        End If
    ElseIf cboOption.Text = "Twice a year" Then
        If chkReverse.Value = False Then
            lblJDATE.Caption = N2Str2Null(lastDay(DateAdd("m", 6 * Occurence, dtFrom.Value)))
        Else
            lblJDATE.Caption = N2Str2Null(firstDay(DateAdd("m", 6 * Occurence, dtFrom.Value)))
        End If
    ElseIf cboOption.Text = "Yearly" Then
        lblJDATE.Caption = N2Str2Null(DateAdd("yyyy", Occurence, dtFrom.Value))
    End If
    
    
    SQL_STATEMENT = ""
    SQL_STATEMENT = "INSERT INTO AMIS_JOURNAL_HD(JDATE,VOUCHERNO,JTYPE,VENDORCODE,NEW_CUSTOMERCODE,CUSTOMERCODE,CUSTOMERNAME,PAYEECODE,BANKCODE,CHECKNO,"
    SQL_STATEMENT = SQL_STATEMENT & "CHECKDATE,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,DUEDATE,PAYTYPE,REFNO,REFDATE,TERMS,DEALER,AMOUNTTOPAY,"
    SQL_STATEMENT = SQL_STATEMENT & "AMOUNTPAID,BALANCE,REMARKS,JNO,DEBIT,CREDIT,TAX,OUTBALANCE,STATUS,PAIDSTATUS,RECEIVESTATUS,RECONSTATUS,AR_DATEGEN,"
    SQL_STATEMENT = SQL_STATEMENT & "AR_BALANCE,BANK,REFERENCENO,ENTITY_CLASS,USERCODE,LASTUPDATE,DATEPOSTED,DATECANCELLED)"
    SQL_STATEMENT = SQL_STATEMENT & "SELECT " & lblJDATE.Caption & "," & xNEWVOUCHERNO & ",JTYPE,VENDORCODE,NEW_CUSTOMERCODE,CUSTOMERCODE,CUSTOMERNAME,PAYEECODE,BANKCODE,CHECKNO,"
    SQL_STATEMENT = SQL_STATEMENT & "CHECKDATE,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,DUEDATE,PAYTYPE,REFNO,REFDATE,TERMS,DEALER,AMOUNTTOPAY,"
    SQL_STATEMENT = SQL_STATEMENT & "AMOUNTPAID,BALANCE,REMARKS," & xNEWJNO & ",DEBIT,CREDIT,TAX,OUTBALANCE,STATUS,PAIDSTATUS,RECEIVESTATUS,RECONSTATUS,AR_DATEGEN,"
    SQL_STATEMENT = SQL_STATEMENT & "AR_BALANCE , BANK, REFERENCENO, ENTITY_CLASS, USERCODE, LASTUPDATE, DATEPOSTED, DATECANCELLED FROM AMIS_JOURNAL_HD WHERE JTYPE='" & xJType & "' AND VOUCHERNO='" & xVOUCHERNO & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    SQL_STATEMENT = ""
    SQL_STATEMENT = "INSERT INTO AMIS_JOURNAL_DET(JITEMNO,JDATE,JTYPE,JNO,VOUCHERNO,ACCT_CODE,ACCT_NAME,GROSSAMT,NETAMT,DEBIT,CREDIT,TAX,STATUS,ATC,RATE,TAXBASE,"
    SQL_STATEMENT = SQL_STATEMENT & "REFERENCENO,ENTITY,INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,ADJ_REMARKS,IS_OTHERS,USERCODE,LASTUPDATE,JOURNAL_HD_ID,CHARTACCOUNT_ID)"
    SQL_STATEMENT = SQL_STATEMENT & "SELECT JITEMNO," & lblJDATE.Caption & ",JTYPE," & xNEWJNO & "," & xNEWVOUCHERNO & ",ACCT_CODE,ACCT_NAME,GROSSAMT,NETAMT,DEBIT,CREDIT,TAX,STATUS,ATC,RATE,TAXBASE,"
    SQL_STATEMENT = SQL_STATEMENT & "REFERENCENO,ENTITY,INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,ADJ_REMARKS,IS_OTHERS,USERCODE,LASTUPDATE,JOURNAL_HD_ID,CHARTACCOUNT_ID FROM AMIS_JOURNAL_DET WHERE JTYPE='" & xJType & "' AND VOUCHERNO='" & xVOUCHERNO & "'"
    gconDMIS.Execute SQL_STATEMENT
Next Occurence

    MessagePop RecSave, "System Message!", "Recurring Entries successfully created!"
    Unload Me
    Unload frmAMISJournalEntry_GJ
    Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
    FormExistsShow frmAMISJournalEntry_GJ
End Sub

Function GetVoucherNo(XXX As String) As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select MAX(VOUCHERNO) VOUCHERNO from AMIS_Journal_HD Where Jtype = '" & XXX & "'")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function GetJNo(XXX As String) As String
    Dim rsJournal_HD                                     As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select MAX(JNO) JNO from AMIS_Journal_HD Where Jtype = '" & XXX & "'")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetJNo = Format(N2Str2Zero(rsJournal_HD!JNo) + 1, "000000")
    Else
        GetJNo = "000001"
    End If
End Function
