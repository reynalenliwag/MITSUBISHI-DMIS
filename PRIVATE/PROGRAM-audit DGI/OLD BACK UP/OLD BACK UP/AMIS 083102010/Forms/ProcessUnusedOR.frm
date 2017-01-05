VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmAMISProcessUnusedOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unused OR"
   ClientHeight    =   2370
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4515
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ProcessUnusedOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4515
   Begin wizProgBar.Prg Prg 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   556
      Picture         =   "ProcessUnusedOR.frx":08CA
      ForeColor       =   0
      BarPicture      =   "ProcessUnusedOR.frx":08E6
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
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
      MouseIcon       =   "ProcessUnusedOR.frx":0902
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUnusedOR.frx":0A54
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
      MouseIcon       =   "ProcessUnusedOR.frx":0E9F
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUnusedOR.frx":0FF1
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   945
      Width           =   885
   End
   Begin VB.ComboBox cboORType 
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
      Format          =   51838977
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
      Format          =   51838977
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
Attribute VB_Name = "frmAMISProcessUnusedOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsINVOICES                                    As ADODB.Recordset

Private Sub cboORType_Click()
    Set rsINVOICES = New ADODB.Recordset
    If cboORType.Text = "VAT" Then
        Set rsINVOICES = gconDMIS.Execute("Select MIN(jdate) as FirstInvNo, MAX(jdate) as LastInvNo from AMIS_Journal_HD Where Jtype = 'CRJ' and Left(Invoiceno,2) <> 'NV'")
    Else
        Set rsINVOICES = gconDMIS.Execute("Select MIN(jdate) as FirstInvNo, MAX(jdate) as LastInvNo from AMIS_Journal_HD Where Jtype = 'CRJ' and Left(Invoiceno,2) = 'NV'")
    End If
    If Not rsINVOICES.EOF And Not rsINVOICES.BOF Then
        cmdPrint.Enabled = True
        dtpFrom = rsINVOICES!FirstInvNo
        dtpTo = rsINVOICES!LastInvNo
    Else
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    Dim Findings                                  As String
    Dim CurrentSeries                             As String
    Dim ISeries                                   As Long
    Dim Chat                                      As Integer
    Dim rsFINDINGS                                As ADODB.Recordset
    On Error GoTo Errorcode:

    Set rsINVOICES = New ADODB.Recordset
    If cboORType.Text = "VAT" Then
        Set rsINVOICES = gconDMIS.Execute("Select invoiceno,jno from AMIS_Journal_HD where jtype = 'CRJ' and left(invoiceno,2) <> 'NV' and jdate >= '" & dtpFrom & "' and jdate <= '" & dtpTo & "' order by invoiceno asc")
    Else
        Set rsINVOICES = gconDMIS.Execute("Select invoiceno,jno from AMIS_Journal_HD where status = 'P' and jtype = 'CRJ' and left(invoiceno,2) = 'NV' and jdate >= '" & dtpFrom & "' and jdate <= '" & dtpTo & "' order by invoiceno asc")
    End If

    If Not rsINVOICES.EOF And Not rsINVOICES.BOF Then
        rsINVOICES.MoveFirst
        If cboORType.Text = "NON-VAT" Then
            CurrentSeries = Right(Null2String(rsINVOICES!INVOICENO), 6)
        Else
            CurrentSeries = Null2String(rsINVOICES!INVOICENO)
        End If
        ISeries = NumericVal(CurrentSeries)
        Screen.MousePointer = 11
        gconDMIS.Execute ("Delete from AMIS_UnUsedOR")
        Prg.Max = rsINVOICES.RecordCount
        Do While Not rsINVOICES.EOF
            If cboORType.Text = "NON-VAT" Then
                CurrentSeries = Right(Null2String(rsINVOICES!INVOICENO), 6)
            Else
                CurrentSeries = Null2String(rsINVOICES!INVOICENO)
            End If
            If Format(ISeries, "000000") <> Format(CurrentSeries, "000000") Then
                Findings = "UNUSED"
                For Chat = NumericVal(ISeries) To NumericVal(CurrentSeries) - 1
                    Set rsFINDINGS = New ADODB.Recordset
                    Set rsFINDINGS = gconDMIS.Execute("Select invoiceno from AMIS_Journal_HD Where jtype = 'CRJ' and invoiceno = '" & Format(Chat, "000000") & "'")
                    If rsFINDINGS.EOF And rsFINDINGS.BOF Then
                        gconDMIS.Execute ("Insert into AMIS_UnUsedOR (ORno,Findings) values ('" & Format(Chat, "000000") & "','" & Findings & "')")
                    End If
                Next
                ISeries = NumericVal(CurrentSeries)
            Else
                Set rsFINDINGS = New ADODB.Recordset
                Set rsFINDINGS = gconDMIS.Execute("Select jno from AMIS_Journal_HD Where Jtype = 'CRJ' and VoucherNo = '" & rsINVOICES!VOUCHERNO & "' and status = 'N'")
                If Not rsFINDINGS.EOF And Not rsFINDINGS.BOF Then
                    Findings = "UNPOSTED"
                    gconDMIS.Execute ("Insert into AMIS_UnUsedOR (ORno,Findings) values ('" & Format(ISeries, "000000") & "','" & Findings & "')")
                End If
                Set rsFINDINGS = New ADODB.Recordset
                Set rsFINDINGS = gconDMIS.Execute("Select jno from AMIS_Journal_HD Where JTYPE = 'CRJ' and VoucherNo = '" & rsINVOICES!VOUCHERNO & "' and status = 'C'")
                If Not rsFINDINGS.EOF And Not rsFINDINGS.BOF Then
                    Findings = "CANCELLED"
                    gconDMIS.Execute ("Insert into AMIS_UnUsedOR (ORno,Findings) values ('" & Format(ISeries, "000000") & "','" & Findings & "')")
                End If
            End If
            rsINVOICES.MoveNext
            ISeries = ISeries + 1
            Prg.Value = rsINVOICES.AbsolutePosition
            Prg.Text = Round((rsINVOICES.AbsolutePosition / rsINVOICES.RecordCount) * 100, 0) & "%"
            DoEvents
        Loop
        If cboORType.Text = "VAT" Then
            ShowRangeReport dtpFrom, dtpTo, "UnusedOR", "InvoicesReport", "", "Unused OR (VAT)", False
        Else
            ShowRangeReport dtpFrom, dtpTo, "UnusedOR", "InvoicesReport", "", "Unused OR (NON-VAT)", False
        End If
        Screen.MousePointer = 0
    End If
    LogAudit "V", "UNUSED OR", cboORType & ": " & dtpFrom & "-" & dtpTo
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cboORType.Clear
    cboORType.AddItem "VAT"
    cboORType.AddItem "NON-VAT"
    cmdPrint.Enabled = False
    Screen.MousePointer = 0
End Sub
