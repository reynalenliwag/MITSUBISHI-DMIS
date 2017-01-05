VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSMSTechnicianWorkshopReport 
   Caption         =   "Tech Workshop Report"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3900
   Icon            =   "frmCSMSTechnicianWorkshopReport.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1650
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   825
      Left            =   2970
      MouseIcon       =   "frmCSMSTechnicianWorkshopReport.frx":06EA
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianWorkshopReport.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   750
      Width           =   705
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   825
      Left            =   2280
      MouseIcon       =   "frmCSMSTechnicianWorkshopReport.frx":0C87
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianWorkshopReport.frx":0DD9
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   750
      Width           =   705
   End
   Begin MSComCtl2.DTPicker dtpFROM 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96337921
      CurrentDate     =   39646
   End
   Begin Crystal.CrystalReport rptTechnician_Performance_Report 
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Performance Report"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpTO 
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   240
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96337921
      CurrentDate     =   39646
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   2025
      TabIndex        =   4
      Top             =   0
      Width           =   210
   End
End
Attribute VB_Name = "frmCSMSTechnicianWorkshopReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorCode

    Screen.MousePointer = 11

    If DTPTO.value < DTPFROM.value Then
        MsgBox "Invalid Range Format", vbInformation, "CSMS"
        DTPFROM.SetFocus
        Exit Sub
    End If
    Dim VRANGE                                         As String
    
    VRANGE = "From " & DTPFROM.value & " To " & DTPTO.value
                   

        rptTechnician_Performance_Report.Formulas(1) = "CompanyName = '" & Company_name & "'"
        rptTechnician_Performance_Report.Formulas(2) = "CompanyAddress = '" & Company_Address & "'"
        rptTechnician_Performance_Report.Formulas(3) = "Printedby = '" & LOGNAME & "'"
        rptTechnician_Performance_Report.Formulas(4) = "RANGE = '" & VRANGE & "' "
        rptTechnician_Performance_Report.WindowTitle = "TECHNICIAN WORKSHOP REPORT"

        PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianWorkshopSales.rpt", "{RO.dte_comp} >= date(" & Year(DTPFROM.value) & "," & Month(DTPFROM.value) & "," & Day(DTPFROM.value) & ") AND {RO.DTE_COMP} <= DATE(" & Year(DTPTO.value) & "," & Month(DTPTO.value) & "," & Day(DTPTO.value) & ")", CSMS_REPORT_CONNECTION, 1

    Screen.MousePointer = 0
    
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    DTPFROM.value = firstDay(Date)
    DTPTO.value = Date
    
    Screen.MousePointer = 0
End Sub

