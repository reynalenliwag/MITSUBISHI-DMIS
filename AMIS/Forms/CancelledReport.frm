VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCancelledReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelled Report"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4755
   ForeColor       =   &H8000000F&
   Icon            =   "CancelledReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4755
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
      Left            =   2910
      MouseIcon       =   "CancelledReport.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "CancelledReport.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1095
      Width           =   885
   End
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
      ItemData        =   "CancelledReport.frx":0B7B
      Left            =   780
      List            =   "CancelledReport.frx":0B8E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   630
      Width           =   3915
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
      Left            =   3795
      MouseIcon       =   "CancelledReport.frx":0BA9
      MousePointer    =   99  'Custom
      Picture         =   "CancelledReport.frx":0CFB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1095
      Width           =   885
   End
   Begin Crystal.CrystalReport rptcancel 
      Left            =   630
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
      Left            =   780
      TabIndex        =   2
      Top             =   120
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
      Format          =   131858433
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Top             =   120
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
      Format          =   131858433
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
      TabIndex        =   6
      Top             =   270
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
      TabIndex        =   5
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "Type"
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
      Left            =   180
      TabIndex        =   4
      Top             =   690
      Width           =   555
   End
End
Attribute VB_Name = "frmCancelledReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim SQL1                                                As String
    Dim RS                                                  As New ADODB.Recordset

    If cboModule.Text = "" Then
        MsgBox "Please Select module", vbInformation, "Information"
        Exit Sub
    End If

    SQL1 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='" & cboModule.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL1)

    If Not RS.EOF And Not RS.BOF Then
        Screen.MousePointer = 11
        rptcancel.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptcancel.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptcancel.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
        rptcancel.Formulas(3) = "tojdate ='" & dtpTo & "'"
        rptcancel.WindowTitle = "Cancelled Report"
        PrintSQLReport rptcancel, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='" & cboModule.Text & "' ", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
        LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
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

