VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmOSMSReportDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issuance By Department"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "ByDepartment.frx":0000
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
   Begin VB.ComboBox cboDepartment 
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
   Begin Crystal.CrystalReport rptByDepartment 
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
      MouseIcon       =   "ByDepartment.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "ByDepartment.frx":045C
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
      MouseIcon       =   "ByDepartment.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "ByDepartment.frx":09F9
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
Attribute VB_Name = "frmOSMSReportDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment As ADODB.Recordset
Dim rsIssuance_Header As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim mysql As String
    ' On Error GoTo ErrorCode
    If cboDepartment.Text = "" Then
        MsgBoxXP "Invalid Department Name", "Warning", XP_OKOnly, msg_Critical
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set rsIssuance_Header = New ADODB.Recordset

    mysql = "SELECT TRANS_DATE, DEPARTMENT_CODE FROM OSMS_ISSUANCE_HEADER INNER JOIN OSMS_EMPLOYEE ON ISSUED_TO = EMPLOYEE_ID WHERE DEPARTMENT_CODE = '" & SetDeptCode(cboDepartment.Text) & "'  and month(TRANS_DATE) =  " & What_month(cboMonth.Text) & " and year(TRANS_DATE) = " & cboYear.Text

    '  mysql = "select ISSUANCE_HEADER.TRANS_DATE, EMPLOYEE.DEPARTMENT_CODE FROM OSMS_ISSUANCE_HEADER INNER JOIN OSMS_EMPLOYEE ON OSMS_ISSUANCE_HEADER.ISSUED_TO = EMPLOYEE.EMPLOYEE_ID WHERE EMPLOYEE.DEPARTMENT_CODE = '" & SetDeptCode(cboDepartment.Text) & "' and month(ISSUANCE_HEADER.TRANS_DATE) = " & What_month(cboMonth.Text) & " and year(ISSUANCE_HEADER.TRANS_DATE) = " & cboYear.Text
    Debug.Print mysql
    rsIssuance_Header.Open mysql, gconDMIS
    If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
        PrintSQLReport rptByDepartment, OSMS_REPORT_PATH & "ByDepartment.rpt", "{EMPLOYEE.DEPARTMENT_CODE} = '" & SetDeptCode(cboDepartment.Text) & "' and month({ISSUANCE_HEADER.TRANS_DATE}) = " & What_month(cboMonth.Text) & " and year({ISSUANCE_HEADER.TRANS_DATE}) = " & cboYear.Text, OSMS_DataConn, 1
        rptByDepartment.PageZoom 89
    Else
        Screen.MousePointer = 0
        MsgBoxXP "No Issuance made to " & cboDepartment.Text & vbCrLf & _
                 "for " & cboMonth.Text & ", " & cboYear.Text, "No Record", XP_OKOnly, msg_Information
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    MsgBox Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "select dept_description  from  OSMS_department order by dept_description asc", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        rsDepartment.MoveFirst
        cboDepartment.Clear
        Do While Not rsDepartment.EOF
            cboDepartment.AddItem Null2String(rsDepartment!dept_description)
            rsDepartment.MoveNext
        Loop
    End If
    FillcboYear cboYear: fillcbomonth cboMonth
    cboYear.Text = Year(LOGDATE): cboMonth.Text = The_month(Month(LOGDATE))
    Screen.MousePointer = 0
End Sub

Function SetDeptCode(XXX As String) As String
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "select *  from  OSMS_department where dept_description = '" & XXX & "'", gconDMIS
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SetDeptCode = Null2String(rsDepartment!DEPARTMENT_CODE)
    End If
End Function
