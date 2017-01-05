VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTallyTool 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer's Tally Tool"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   Icon            =   "frmTallyTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7320
   ScaleWidth      =   13500
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   6960
      TabIndex        =   9
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118358019
      CurrentDate     =   42241
   End
   Begin VB.ComboBox cmbChoose 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmTallyTool.frx":07AA
      Left            =   240
      List            =   "frmTallyTool.frx":07B4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   6615
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXPORT"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "GENERATE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdSLGL 
      Height          =   5745
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10134
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PRB 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   6840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblOf 
      BackStyle       =   0  'Transparent
      Caption         =   "OF"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMax 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmTallyTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       ´$$$$`                             ,,,
'       ´$$$$$$$`                         ´$$$`
'        `$$$$$$$`      ,,        ,,      ´$$$$´
'         `$$$$$$$`    ´$$`     ´$$`    ´$$$$$´
'          `$$$$$$$`_´$$$$$`_´$$$$$`__´$$$$$$$´
'           `$$$$$$$_$$$$$$$_$$$$$$$_´$$$$$$$´
'            `$$$$$$_$$$$$$$_$$$$$$$`´$$$$$$´
'   ,,,,,    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
' ´$$$$$`    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
'´$$$$$$$$$`´$$$$$$$_SARAJOY_$$$$$$$_$$$$$´
'´$$$$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$$_$$$$$´
'   `$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$_$$$$$$´
'      `$$$$$$$$$$$$$_$$$$$__$$_$$$$$$_$$´
'       `$$$$$$$$$$$$$__,$$$$,_____,$$$$$´
'         `$$$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'          `$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'            `$$$$$$$$CREATED$ON$$$$$$´
'               `$$$$$JULY242015$$$$$´
'
'              ROCK 'N ROLL TO THE WORLD

Dim rsUSP As ADODB.Recordset

Private Sub Form_Load()
    Screen.MousePointer = 0
    CenterMe frmMain, Me, 1
    cmbChoose.Text = "ACCOUNTS RECEIVABLE"
    initGrid
    storedlookup
End Sub

Sub initGrid()
    With grdSLGL
        .Rows = 1
        .ColWidth(0) = 1900: .ColWidth(1) = 4200: .ColWidth(2) = 1700
        .ColWidth(3) = 1700: .ColWidth(4) = 1600: .ColWidth(5) = 1800
        .Row = 0
        .Col = 0: .Text = "ACCOUNT CODE"
        .Col = 1: .Text = "ACCOUNT DESCRIPTION"
        .Col = 2: .Text = "SL"
        .Col = 3: .Text = "GL"
        .Col = 4: .Text = "DIFFERENCE"
        .Col = 5: .Text = "REMARKS"
    End With
    cmdPrint.Enabled = False
End Sub

Sub storedlookup()
On Error GoTo NoStored
    Set rsUSP = Nothing
    Set CMD = New ADODB.Command
    
    With CMD
        .ActiveConnection = gconDMIS
        .CommandType = adCmdStoredProc
        .CommandText = "XSP_TALLY"
        .Parameters.Append CMD.CreateParameter("@ReportDate", adDate, adParamInput, 8, DTPicker1.Value)
        Set rsUSP = .Execute
    End With
    Exit Sub
    
NoStored:
    MsgBox "Cannot find stored procedure 'XSP_TALLY' and XSP_TALLY2'." & vbCrLf & "Contact SJR to get procedure.", vbExclamation, "Could not find file"
End Sub

Private Sub cmbChoose_Click()
    cmdPrint.Enabled = False
End Sub

Private Sub cmdGenerate_Click()
    load_USP
    FillTallyToolGrid
End Sub

Sub load_USP()
    Set rsUSP = Nothing
    Set CMD = New ADODB.Command
    
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    
    If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
        CMD.CommandText = "XSP_TALLY"
    ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
        CMD.CommandText = "XSP_TALLY2"
    End If
    CMD.CommandTimeout = 1000
    CMD.Parameters.Append CMD.CreateParameter("@ReportDate", adDate, adParamInput, 8, DTPicker1.Value)

    Set rsUSP = CMD.Execute
    
End Sub

Sub FillTallyToolGrid()
On Error GoTo ErrorMessage
    
    Dim cnt As Integer
    cnt = 0
    grdSLGL.Rows = 1
    
    lblMax(2).Visible = True: lblCurrent(1).Visible = True: lblOf(0).Visible = True
    
    If Not rsUSP.EOF And Not rsUSP.BOF Then
        rsUSP.MoveFirst
        PRB.Value = 0
        PRB.Max = rsUSP.RecordCount
        Do While Not rsUSP.EOF
        cnt = cnt + 1
            If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
                With grdSLGL
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsUSP!xACCOUNTCODE
                    .TextMatrix(.Rows - 1, 1) = rsUSP!xDescription
                    .TextMatrix(.Rows - 1, 2) = ToDoubleNumber(N2Str2Zero(rsUSP!XSL))
                    .TextMatrix(.Rows - 1, 3) = ToDoubleNumber(N2Str2Zero(rsUSP!XGL))
                    .TextMatrix(.Rows - 1, 4) = ToDoubleNumber(N2Str2Zero(rsUSP!XDIFF))
                    .TextMatrix(.Rows - 1, 5) = rsUSP!xREMARKS
                End With
            ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
                With grdSLGL
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = rsUSP!xAcctCode
                    .TextMatrix(.Rows - 1, 1) = rsUSP!xDescription
                    .TextMatrix(.Rows - 1, 2) = ToDoubleNumber(N2Str2Zero(rsUSP!XSL))
                    .TextMatrix(.Rows - 1, 3) = ToDoubleNumber(N2Str2Zero(rsUSP!XGL))
                    .TextMatrix(.Rows - 1, 4) = ToDoubleNumber(N2Str2Zero(rsUSP!XDIFF))
                    .TextMatrix(.Rows - 1, 5) = rsUSP!xREMARKS
                End With
            End If
            
            If rsUSP!xREMARKS = "NOT BALANCED" Then
                    grdSLGL.Col = 5
                    grdSLGL.Row = cnt
                    grdSLGL.CellFontBold = True
                    grdSLGL.CellForeColor = RGB(255, 0, 0)
            End If
            
            rsUSP.MoveNext
            
            If PRB.Value = PRB.Max Then
                PRB.Enabled = False
            Else
                PRB.Value = PRB.Value + 1
                lblMax(2).Caption = rsUSP.RecordCount
                lblCurrent(1).Caption = cnt
            End If
            
        Loop
    End If
    
    grdSLGL.Refresh
    cmdPrint.Enabled = True
    'Set rsUSP = Nothing
    Exit Sub
ErrorMessage:
        MsgBox "Error on generating Developer's Tally Tool." & vbCrLf & "Contact SJR. Rock en roll!", vbCritical, "Error"
End Sub

Private Sub cmdPrint_Click()
'CRYSTAL REPORT
    If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
        ShowReport "SLGLREPORT_AR", "Summary", "", "SL-GL REPORT", DTPicker1.Value, 0
    Else
        ShowReport "SLGLREPORT_AP", "Summary", "", "SL-GL REPORT", DTPicker1.Value, 0
    End If
    
'EXCEL REPORT
    Dim sFileName As String
    sFileName = AMIS_REPORT_PATH & COMPANY_CODE & "_GL-SL STATUS AS OF " & Format(DTPicker1.Value, "mm.dd.yyyy") & ".xls"
    
    Dim oXLApp As Object
    Dim oXLBook As Object
    Dim oXLSheet As Object
    
    Set oXLApp = CreateObject("Excel.Application")
    
    If Dir(sFileName) = "" Then
      Set oXLBook = oXLApp.Workbooks.Add
    Else
      Set oXLBook = oXLApp.Workbooks.Open(sFileName)
    End If
    
    Set oXLSheet = oXLBook.Worksheets(1)

    oXLSheet.UsedRange.Clear
    oXLSheet.Range("A1:F1").Merge
    oXLSheet.Range("A2:F2").Merge
    oXLSheet.Range("A3:F3").Merge
    oXLSheet.Range("A4:F4").Merge
    
    oXLSheet.Range("A1:F1").Font.Bold = True
    oXLSheet.Range("A2:F2").Font.Bold = True
    oXLSheet.Range("A3:F3").Font.Bold = True
    oXLSheet.Range("A4").Font.Italic = True
    
    oXLSheet.Range("A1") = COMPANY_NAME
    oXLSheet.Range("A2") = COMPANY_ADDRESS
    oXLSheet.Range("A3") = cmbChoose.Text & " GL-SL STATUS AS OF " & DTPicker1.Value
    oXLSheet.Range("A5").CopyFromRecordset rsUSP
    
    oXLSheet.Columns("A:A").EntireColumn.AutoFit
    oXLSheet.Columns("B:B").EntireColumn.AutoFit
    oXLSheet.Columns("C:C").EntireColumn.AutoFit
    oXLSheet.Columns("D:D").EntireColumn.AutoFit
    oXLSheet.Columns("E:E").EntireColumn.AutoFit
    oXLSheet.Columns("F:F").EntireColumn.AutoFit
    
    oXLBook.SaveAs sFileName
    oXLApp.Visible = True
    Set oXLApp = Nothing
    Set oXLBook = Nothing
    Set oXLSheet = Nothing
End Sub

