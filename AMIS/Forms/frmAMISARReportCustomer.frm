VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmARReportCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   Icon            =   "frmAMISARReportCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   10050
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8460
      Top             =   3690
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7755
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   15015
      TabIndex        =   0
      Top             =   30
      Width           =   15015
      Begin VB.TextBox txtEntity 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   10
         Text            =   "txtEntity"
         Top             =   450
         Width           =   4545
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9210
         MouseIcon       =   "frmAMISARReportCustomer.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISARReportCustomer.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Export"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   735
         Left            =   8490
         MouseIcon       =   "frmAMISARReportCustomer.frx":2256
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISARReportCustomer.frx":23A8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print Report"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdInquire 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7770
         MouseIcon       =   "frmAMISARReportCustomer.frx":2847
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISARReportCustomer.frx":2999
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "View"
         Top             =   120
         Width           =   735
      End
      Begin VB.ComboBox cboOption 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmAMISARReportCustomer.frx":2CE0
         Left            =   90
         List            =   "frmAMISARReportCustomer.frx":2CEA
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   450
         Width           =   2985
      End
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   5850
         ScaleHeight     =   435
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   3600
         Visible         =   0   'False
         Width           =   2595
         Begin VB.Label lblLoading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait while loading"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   8
            Top             =   120
            Width           =   2145
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   435
            Left            =   0
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   2535
            _Version        =   655364
            _ExtentX        =   4471
            _ExtentY        =   767
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
            SubItemCaption  =   -1  'True
            RightToLeftReading=   -1  'True
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6765
         Left            =   30
         TabIndex        =   6
         Top             =   960
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   11933
         Appearance      =   0
         BackColor2      =   16573135
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         Rows            =   1
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Option"
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   2895
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FAF1DC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00F5D8BC&
         FillStyle       =   0  'Solid
         Height          =   885
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   14925
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "Export"
      Visible         =   0   'False
      Begin VB.Menu mnuExcel 
         Caption         =   "Export to Excel"
      End
      Begin VB.Menu mnuPDF 
         Caption         =   "Export to PDF"
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "Export to HTML"
      End
   End
End
Attribute VB_Name = "frmARReportCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPORTS                                               As ADODB.Recordset
Dim i                                                       As Integer
Dim ACCT_CODE                                               As String
Dim DESCRIPTION                                             As String
Dim REPORTTYPE                                              As String
Dim CMD                                                     As ADODB.Command
Dim BILLING_TYPE                                            As String
Dim xlsWorkSheet                                            As Excel.Worksheet

Private Sub cmdExport_Click()
    PopupMenu mnuExport
End Sub


Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    initGrid
    
    Screen.MousePointer = 0
    Grid1.Rows = 1
End Sub

Sub initGrid()
    With Grid1
        .Cols = 4
        .Rows = 1
        '        .FixedCols = 4
        .Cell(0, 0).Text = "L/N"

        .Cell(0, 1).Text = "CUSTOMER CODE"
        .Column(1).Width = 100
        .Column(1).FormatString = "mm/dd/yyyy"

        .Cell(0, 2).Text = "CUSTOMER NAME"
        .Column(2).Alignment = cellCenterCenter
        .Column(2).Width = 350

        .Cell(0, 3).Text = "BALANCE"
        .Column(3).Alignment = cellRightCenter
        .Column(3).Width = 110
    End With
End Sub

Sub BILLING_DUE_REPORT()
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_BILLING_DATERANGE"

    With CMD.Parameters
        '.Append CMD.CreateParameter("@ACCT_CODE", adVarChar, adParamInput, 12, ACCT_CODE)
        '.Append CMD.CreateParameter("@JDATE1", adDate, adParamInput, 8, dtFrom)
        '.Append CMD.CreateParameter("@JDATE2", adDate, adParamInput, 8, dtTo)
        '.Append CMD.CreateParameter("@REPORTTYPE", adVarChar, adParamInput, 8, REPORTTYPE)
    End With
    Set rsREPORTS = CMD.Execute
    FILLREPORTS
End Sub

Sub FILLREPORTS()
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    Picture1.Enabled = False
    Picture2.Visible = True

    If Not rsREPORTS.EOF And Not rsREPORTS.BOF Then
        While Not rsREPORTS.EOF
            If REPORTTYPE = "AP" Then
                Grid1.AddItem _
                        rsREPORTS!VENDOR_NAME & Chr(9) & rsREPORTS!VOUCHERNO & Chr(9) & _
                                              rsREPORTS!INVOICENO & Chr(9) & rsREPORTS!invoicedate & Chr(9) & _
                                              Null2String(rsREPORTS!DUEDATE) & Chr(9) & ToDoubleNumber(rsREPORTS!AMOUNT2PAY) & Chr(9) & rsREPORTS!ACCT_CODE & Chr(9) & _
                                              rsREPORTS!DESCRIPTION
            Else
                Grid1.AddItem _
                        rsREPORTS!CUSTOMERNAME & Chr(9) & rsREPORTS!SJVoucherno & Chr(9) & _
                                               rsREPORTS!INVOICENO & Chr(9) & rsREPORTS!invoicedate & Chr(9) & _
                                               Null2String(rsREPORTS!DUEDATE) & Chr(9) & ToDoubleNumber(rsREPORTS!AR_TOPAY) & Chr(9) & rsREPORTS!ACCT_CODE & Chr(9) & _
                                               rsREPORTS!DESCRIPTION
            End If
            rsREPORTS.MoveNext
            Loading
        Wend
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    Picture1.Enabled = True
    Picture2.Visible = False
    cmdPrint.Enabled = True
    cmdExport.Enabled = True
    Set rsREPORTS = Nothing
End Sub

Sub Loading()
    If lblLoading.Caption = "Please wait while loading" Then
        lblLoading.Caption = "Please wait while loading."
    ElseIf lblLoading.Caption = "Please wait while loading." Then
        lblLoading.Caption = "Please wait while loading.."
    ElseIf lblLoading.Caption = "Please wait while loading.." Then
        lblLoading.Caption = "Please wait while loading..."
    ElseIf lblLoading.Caption = "Please wait while loading..." Then
        lblLoading.Caption = "Please wait while loading...."
    ElseIf lblLoading.Caption = "Please wait while loading...." Then
        lblLoading.Caption = "Please wait while loading....."
    ElseIf lblLoading.Caption = "Please wait while loading....." Then
        lblLoading.Caption = "Please wait while loading."
    End If
End Sub

Private Sub mnuExcel_Click()
    Grid1.ExportToExcel ("")
End Sub

Private Sub mnuHTML_Click()
    Grid1.ExportToHTML ("")
End Sub

Private Sub mnuPDF_Click()
    Grid1.ExportToPDF ("")
End Sub
