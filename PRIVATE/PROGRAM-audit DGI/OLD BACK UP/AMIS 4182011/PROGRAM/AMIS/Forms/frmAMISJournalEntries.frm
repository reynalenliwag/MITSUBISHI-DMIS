VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMISJournalEntries 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9210
   Icon            =   "frmAMISJournalEntries.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraAddJournal 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   30
      ScaleHeight     =   1635
      ScaleWidth      =   9105
      TabIndex        =   0
      Top             =   30
      Width           =   9135
      Begin VB.CommandButton cmdJournalCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   8315
         MouseIcon       =   "frmAMISJournalEntries.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   765
         Width           =   705
      End
      Begin VB.CommandButton cmdJournalSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7620
         MouseIcon       =   "frmAMISJournalEntries.frx":1512
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries.frx":1664
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   765
         Width           =   705
      End
      Begin VB.Frame fraComp 
         Height          =   915
         Left            =   2340
         TabIndex        =   19
         Top             =   660
         Width           =   4365
         Begin VB.TextBox txtGrossAmt 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   150
            MaxLength       =   10
            TabIndex        =   22
            Top             =   510
            Width           =   1300
         End
         Begin VB.TextBox txtTax 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   21
            Top             =   510
            Width           =   1300
         End
         Begin VB.TextBox txtNetAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2910
            MaxLength       =   10
            TabIndex        =   20
            Top             =   510
            Width           =   1300
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Amt."
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
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label labTax 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Output Tax"
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
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Net Amount"
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
            Left            =   2910
            TabIndex        =   23
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame fraATC 
         Height          =   915
         Left            =   2340
         TabIndex        =   11
         Top             =   660
         Width           =   4365
         Begin VB.TextBox txtTaxBase 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2550
            MaxLength       =   15
            TabIndex        =   14
            Top             =   510
            Width           =   1725
         End
         Begin VB.TextBox txtRATE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   13
            Top             =   510
            Width           =   615
         End
         Begin VB.ComboBox cboATC 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   60
            TabIndex        =   12
            Top             =   510
            Width           =   1425
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Taxbase Amt."
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
            Left            =   2550
            TabIndex        =   18
            Top             =   240
            Width           =   1725
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RATE"
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
            Left            =   1380
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ATC Code"
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label41 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2190
            TabIndex        =   15
            Top             =   540
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   735
         Left            =   2310
         TabIndex        =   5
         Top             =   -30
         Width           =   4425
         Begin RichTextLib.RichTextBox txtAcct_Name 
            Height          =   315
            Left            =   30
            TabIndex        =   6
            Top             =   360
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            TextRTF         =   $"frmAMISJournalEntries.frx":19B4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label33 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Left            =   60
            TabIndex        =   7
            Top             =   90
            Width           =   2205
         End
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6780
         MaxLength       =   15
         TabIndex        =   4
         Top             =   330
         Width           =   1100
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7950
         MaxLength       =   15
         TabIndex        =   3
         Top             =   330
         Width           =   1100
      End
      Begin VB.CommandButton cmdJournalDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         MouseIcon       =   "frmAMISJournalEntries.frx":1A47
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries.frx":1B99
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   765
         Width           =   705
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   60
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   330
         Width           =   2235
      End
      Begin VB.TextBox txtJItemNo 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   690
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox txtAcctID 
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   330
         Width           =   585
      End
      Begin VB.Label labPartNo 
         Caption         =   "Description"
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
         Height          =   315
         Left            =   2340
         TabIndex        =   33
         Top             =   420
         Width           =   2685
      End
      Begin VB.Label labDetID 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   930
         TabIndex        =   32
         Top             =   390
         Width           =   915
      End
      Begin VB.Label labPrevOrdQty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item No."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Left            =   8130
         TabIndex        =   30
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
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
         Left            =   7050
         TabIndex        =   29
         Top             =   60
         Width           =   885
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
         TabIndex        =   28
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
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
         Left            =   390
         TabIndex        =   27
         Top             =   390
         Width           =   855
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1725
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   9255
      _Version        =   655364
      _ExtentX        =   16325
      _ExtentY        =   3043
      _StockProps     =   14
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
   End
End
Attribute VB_Name = "frmAMISJournalEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsChartAccount                                     As New ADODB.Recordset
Dim rsATC                                              As New ADODB.Recordset
Dim xJOURNALTYPE                                       As String

Private Sub cboAcct_Code_Change()
    Dim DEALER_ITW_COMPENSATION                        As String
    Dim DEALER_ITW_EXPANDED                            As String
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    GettheTaxBaseAmnt
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        '        fraATC.Visible = True
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
    End If
End Sub

Private Sub cboAcct_Code_Click()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cmdJournalCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
            frmAMISJournalEntry_Chart.Caption = frmAMISJournalEntry_APJ.Caption
            frmAMISJournalEntry_Chart.Show 1
            DoEvents
            On Error Resume Next
            frmAMISJournalEntry_Chart.txtSearch.SetFocus
        End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitJournal
    InitCbo
    fraATC.Visible = False
    fraComp.Visible = False
End Sub


Sub InitJournal()
'    txtJItemNo.Text = Format(kcnt + 1, "0000")
    cboAcct_Code.Text = ""
    txtAcct_Name.Text = ""
    txtDebit.Text = ZERO
    txtCredit.Text = ZERO
    txtTax.Text = ZERO
    txtGrossAmt.Text = ZERO
    txtNetAmt.Text = ZERO
    '    txtSearch.Text = ""
    If xJOURNALTYPE = "APJ" Then
        cboATC.Text = ""
        txtRATE.Text = "0"
        txtTaxBase.Text = ZERO
    End If
End Sub

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Sub InitCbo()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select acctcode from AMIS_ChartAccount order by acctcode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
    Set rsChartAccount = Nothing
    If xJOURNALTYPE = "APJ" Then
        Set rsATC = New ADODB.Recordset
        Set rsATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
        If Not rsATC.EOF And Not rsATC.BOF Then
            'Combo_Loadval cboATC, rsATC
            rsATC.MoveFirst: cboATC.AddItem ""
            Do While Not rsATC.EOF
                cboATC.AddItem Null2String(rsATC!ATC)
                rsATC.MoveNext
            Loop
        End If
        Set rsATC = Nothing
    End If
End Sub

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
End Function

Function ReturnWithholdingTax(XXX As String)
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnWithholdingTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Sub GettheTaxBaseAmnt()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    If xJOURNALTYPE = "APJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & frmAMISJournalEntry_APJ.txtVoucherNo & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'APJ'"
    Else
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & frmAMISJournalEntry_APJ.txtVoucherNo & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'CDJ'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTaxBase.Text = NumericVal(RS!SumDebit)
    End If
    Set RS = Nothing
End Sub

Function ReturnInPutTax()
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If COMPANY_CODE = "HCC" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where DESCRIPTION = 'INPUT TAX'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function
