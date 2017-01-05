VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMIS_UniportedReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Un-Imported Reports"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmAMIS_UniportedReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   3570
   Begin Crystal.CrystalReport rptUnimported 
      Left            =   2850
      Top             =   1860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1770
      MouseIcon       =   "frmAMIS_UniportedReports.frx":6852
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_UniportedReports.frx":69A4
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   1680
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
      Left            =   900
      MouseIcon       =   "frmAMIS_UniportedReports.frx":6DEF
      MousePointer    =   99  'Custom
      Picture         =   "frmAMIS_UniportedReports.frx":6F41
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   1680
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   570
      TabIndex        =   4
      Top             =   1290
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   129499137
      CurrentDate     =   39986
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "Un-Imported Cash Receipts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3045
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Un-Imported Sales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   810
         Width           =   3045
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Un-Imported Purchases"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Value           =   -1  'True
         Width           =   3045
      End
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   2190
      TabIndex        =   7
      Top             =   1290
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   129499137
      CurrentDate     =   39986
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   1350
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   5
      Top             =   1320
      Width           =   915
   End
End
Attribute VB_Name = "frmAMIS_UniportedReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim rsTEMP                                              As ADODB.Recordset

    Screen.MousePointer = 11

    rptUnimported.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptUnimported.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptUnimported.Formulas(2) = "fromdate ='" & dtFrom & "'"
    rptUnimported.Formulas(3) = "todate ='" & dtTo & "'"

    If Option1.Value = True Then
        'This is unimported Purchases
        rptUnimported.WindowTitle = "Un-Imported Purchasess"
        Set rsTEMP = gconDMIS.Execute("Select * from AMIS_VW_UNIMPORTED_PURCHASE  where [Date] >= '" & CDate(dtFrom) & "' and [Date] <= '" & CDate(dtTo) & "' ")
        If Not rsTEMP.EOF And Not rsTEMP.BOF Then
            PrintSQLReport rptUnimported, AMIS_REPORT_PATH & "UnimportedReports\Unimported Transaction.rpt", "{AMIS_Header.Date} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {AMIS_Header.Date} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & " )", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo NO_RECORD
        End If
    ElseIf Option2.Value = True Then
        'This is for unimported Cash Receipts
        rptUnimported.WindowTitle = "Un-Imported Cash  Receipts"
        Set rsTEMP = gconDMIS.Execute("Select * from AMIS_VW_UNIMPORTED_OR  where [Date] >= '" & CDate(dtFrom) & "' and [Date] <= '" & CDate(dtTo) & "' ")
        If Not rsTEMP.EOF And Not rsTEMP.BOF Then
            PrintSQLReport rptUnimported, AMIS_REPORT_PATH & "UnimportedReports\Unimported Transaction-OR.rpt", "{AMIS_Header.Date} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {AMIS_Header.Date} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & " )", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo NO_RECORD
        End If

    Else
        'This is fofr unimported Sales
        rptUnimported.WindowTitle = "Un-Imported Sales"
        Set rsTEMP = gconDMIS.Execute("Select * from AMIS_VW_UNIMPORTED_ISSUANCES  where [Date] >= '" & CDate(dtFrom) & "' and [Date] <= '" & CDate(dtTo) & "' ")
        If Not rsTEMP.EOF And Not rsTEMP.BOF Then
            PrintSQLReport rptUnimported, AMIS_REPORT_PATH & "UnimportedReports\Unimported Transaction-Issuances.rpt", "{AMIS_Header.Date} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {AMIS_Header.Date} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & " )", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo NO_RECORD
        End If
    End If

NO_RECORD:
    MessagePop InfoFriend, "INFORMATION", "No such records"
    Screen.MousePointer = 0
    Set rsTEMP = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

