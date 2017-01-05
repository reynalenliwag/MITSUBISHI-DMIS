VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISRangeReceipts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving Summary Report"
   ClientHeight    =   1950
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4830
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ReportRangeReceipts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4830
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
      Left            =   2460
      MouseIcon       =   "ReportRangeReceipts.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeReceipts.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1050
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
      Left            =   1590
      MouseIcon       =   "ReportRangeReceipts.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeReceipts.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   1050
      Width           =   885
   End
   Begin VB.ComboBox cboNameofVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Text            =   "cboRecvd_Desc"
      Top             =   120
      Width           =   4590
   End
   Begin Crystal.CrystalReport rptAMISrange 
      Left            =   870
      Top             =   1440
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
      Left            =   780
      TabIndex        =   2
      Top             =   540
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
      Format          =   48758785
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
      TabIndex        =   4
      Top             =   540
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
      Format          =   48758785
      CurrentDate     =   38216
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
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   675
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
      Left            =   2550
      TabIndex        =   3
      Top             =   600
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   7
      Top             =   3060
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISRangeReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                                      As ADODB.Recordset

Function SetVendorCode(XXX As String) As String
    Dim rsVENDOR                                                      As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("Select Code from ALL_Vendor where NameOfVendor = '" & XXX & "'")
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!code)
    End If
End Function

Sub FillCboVendor()
    Dim rsVENDOR                                                      As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("Select NameOfVendor from ALL_Vendor order by NameOfVendor asc")
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        rsVENDOR.MoveFirst: cboNameofVendor.Clear
        Do While Not rsVENDOR.EOF
            cboNameofVendor.AddItem Null2String(rsVENDOR!nameofvendor)
            rsVENDOR.MoveNext
        Loop
    End If
    Set rsVENDOR = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:10
Private Sub cmdPrint_Click()
    'If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub

    On Error GoTo Errorcode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If REPORT_RANGETYPE = "REC_REGISTER" Then
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'APJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "ReceiptsRegisters", "Registers", "{Journal_Hd.VendorCode} = '" & SetVendorCode(cboNameofVendor.Text) & "' AND  {Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "RECEIVING REPORT REGISTERS", False
            Unload Me
        Else
            ShowNoRecord
        End If
    End If
    LogAudit "V", "RECEIVING SUMMARY REPORT", cboNameofVendor & ": " & dtpFrom & "-" & dtpTo
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    FillCboVendor
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

