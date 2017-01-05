VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmOSMSReportSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issuance By Supplier"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "BySupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   4815
   Begin VB.ComboBox cboYear 
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   1245
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   630
      Width           =   2355
   End
   Begin VB.ComboBox cboSupplier 
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   4635
   End
   Begin Crystal.CrystalReport rptBySupplier 
      Left            =   3540
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Issuance By Department"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Height          =   795
      Left            =   2340
      MouseIcon       =   "BySupplier.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "BySupplier.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   735
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
      Height          =   795
      Left            =   1620
      MouseIcon       =   "BySupplier.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "BySupplier.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   735
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
      TabIndex        =   3
      Top             =   630
      Width           =   825
   End
End
Attribute VB_Name = "frmOSMSReportSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplier As ADODB.Recordset
Dim rsrrHEADER As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode
    If cboSupplier.Text = "" Then
        MsgBoxXP "Invalid Supplier Name", "Warning", XP_OKOnly, msg_Critical
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set rsrrHEADER = New ADODB.Recordset
    rsrrHEADER.Open "select OSMS_RRHEADER.rrDATE,OSMS_RRHEADER.Supplier_CODE,OSMS_SUPPLIER.Supplier_CODE from OSMS_RRHEADER  INNER JOIN OSMS_SUPPLIER ON OSMS_RRHEADER.SUPPLIER_CODE = OSMS_SUPPLIER.SUPPLIER_CODE WHERE OSMS_SUPPLIER.SUPPLIER_CODE = '" & SetSupplierCode(cboSupplier.Text) & "' and month(OSMS_RRHEADER.rrDATE) = " & What_month(cboMonth.Text) & " and year(OSMS_RRHEADER.rrDATE) = " & cboYear.Text, gconDMIS
    If Not rsrrHEADER.EOF And Not rsrrHEADER.BOF Then
        PrintSQLReport rptBySupplier, OSMS_REPORT_PATH & "BySupplier.rpt", "{RRHEADER.Supplier_CODE} = '" & SetSupplierCode(cboSupplier.Text) & "' and month({rrHEADER.rrDATE}) = " & What_month(cboMonth.Text) & " and year({rrHEADER.rrDATE}) = " & cboYear.Text, OSMS_DataConn, 1
        rptBySupplier.PageZoom 89
    Else
        Screen.MousePointer = 0
        MsgBoxXP "No Issuance made to " & cboSupplier.Text & vbCrLf & _
                 "for " & cboMonth.Text & ", " & cboYear.Text, "No Record", XP_OKOnly, msg_Information
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select SUPPLIER_NAME from OSMS_Supplier order by SUPPLIER_NAME asc", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        cboSupplier.Clear
        Do While Not rsSupplier.EOF
            cboSupplier.AddItem Null2String(rsSupplier!SUPPLIER_NAME)
            rsSupplier.MoveNext
        Loop
    End If
    FillcboYear cboYear: fillcbomonth cboMonth
    cboYear.Text = Year(LOGDATE): cboMonth.Text = The_month(Month(LOGDATE))
    Screen.MousePointer = 0
End Sub

Function SetSupplierCode(XXX As String) As String
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select * from  OSMS_Supplier where SUPPLIER_NAME = '" & XXX & "'", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupplierCode = Null2String(rsSupplier!Supplier_code)
    End If
End Function
