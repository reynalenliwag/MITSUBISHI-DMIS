VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_InvControl31 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Trend By Model"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ForeColor       =   &H00FCFCFC&
   Icon            =   "frmSMIS_Report_SalesTrend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
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
      Left            =   2160
      MouseIcon       =   "frmSMIS_Report_SalesTrend.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_SalesTrend.frx":015E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   690
      Width           =   885
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   3210
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "DMC Monthly Inventory Control"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   555
      Left            =   3420
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "9999"
      Top             =   90
      Width           =   1005
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
      Left            =   1290
      MouseIcon       =   "frmSMIS_Report_SalesTrend.frx":05A9
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_SalesTrend.frx":06FB
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   690
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2610
      TabIndex        =   2
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_InvControl31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                            As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:43
Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT") = False Then Exit Sub
    On Error GoTo ErrorCode:


PrintPOExcel cboMonth, txtYear

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


Sub PrintPOExcel(xMonth, xyear)
    Dim xlApp                           As Excel.Application
    Dim xlBook                          As Excel.Workbook
    Dim xlSheet                         As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\SALESTREND1.xls")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                         As ADODB.Recordset
    Dim vmodel                          As String
    Dim i                               As Integer
    Dim j As Integer
    Dim rsCountProspect As ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select DISTINCT MODEL from ALL_MODEL where LEN(MODEL)>0")
    If Not rsModel.EOF Or Not rsModel.BOF Then
        While Not rsModel.EOF
            i = i + 1
            vmodel = Null2String(rsModel("MODEL"))
            xlSheet.Cells(i + 4, 1) = vmodel
            
            For j = 1 To 31
                Set rsCountProspect = gconDMIS.Execute("select COUNT(*) from CRIS_PROSPECTS  where MODEL='" & vmodel & "' AND DAY(LOGINITIALINQUIRY)=" & j & " AND MONTH(LOGINITIALINQUIRY)=" & xMonth & " AND YEAR(LOGINITIALINQUIRY)=" & xyear)
                
            If Not rsCountProspect.EOF Or Not rsCountProspect.BOF Then
                xlSheet.Cells(i + 4, 1 + j) = rsCountProspect.Fields(0).Value
            End If
            
                Set rsCountProspect = Nothing
            Next
            rsModel.MoveNext
        Wend
        xlApp.Visible = True
        
        Set xlApp = Nothing

    End If
End Sub
