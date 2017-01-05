VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReprintReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Print Transaction Report"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4710
   Icon            =   "ReprintReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboModule 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "ReprintReport.frx":058A
      Left            =   810
      List            =   "ReprintReport.frx":059D
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   3795
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
      Left            =   3735
      MouseIcon       =   "ReprintReport.frx":05BF
      MousePointer    =   99  'Custom
      Picture         =   "ReprintReport.frx":0711
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1155
      Width           =   855
   End
   Begin Crystal.CrystalReport rptReprint 
      Left            =   570
      Top             =   1410
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
      TabIndex        =   2
      Top             =   180
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   131858433
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   2940
      TabIndex        =   3
      Top             =   180
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   131858433
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
      Left            =   2880
      MouseIcon       =   "ReprintReport.frx":0B5C
      MousePointer    =   99  'Custom
      Picture         =   "ReprintReport.frx":0CAE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Module "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   780
      Width           =   765
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
      Left            =   2490
      TabIndex        =   5
      Top             =   300
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
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   675
   End
End
Attribute VB_Name = "frmReprintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    If cboModule.Text = "" Then
        MsgBox "Please Select module", vbInformation, "Information"
        Exit Sub
    End If

    SQL = "SELECT * from ALL_Reprint_TRANSACTION where date_Reprint>='" & CDate(dtpFrom) & "' and date_Reprint<='" & CDate(dtpTo) & "' and module_name='" & cboModule.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Screen.MousePointer = 11
        rptReprint.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReprint.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptReprint.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
        rptReprint.Formulas(3) = "tojdate ='" & dtpTo & "'"
        rptReprint.WindowTitle = "Cancelled Report"
        PrintSQLReport rptReprint, AMIS_REPORT_PATH & "ReprintReport.rpt", "{reprint.DATE_Reprint} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {reprint.DATE_Reprint} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {reprint.module_name}='" & cboModule.Text & "'", DMIS_REPORT_Connection, 1
        'PrintSQLReport rptReprint, AMIS_REPORT_PATH & "ReprintReport.rpt", "{All_reprint_transaction.DATE_Reprint} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {All_reprint_transaction.DATE_Reprint} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {All_reprint_transaction.module_name}=" & cboModule.Text & "", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
        LogAudit "V", "Reprint Report", dtpFrom & "-" & dtpTo
    Else
        ShowNoRecord
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0
End Sub

