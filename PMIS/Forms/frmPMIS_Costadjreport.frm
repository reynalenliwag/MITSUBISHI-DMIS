VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMIS_REPORT_ADJCOST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjusment Report"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "frmPMIS_Costadjreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   3630
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
      Left            =   1080
      MouseIcon       =   "frmPMIS_Costadjreport.frx":076A
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_Costadjreport.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1140
      Width           =   795
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
      Left            =   1860
      MouseIcon       =   "frmPMIS_Costadjreport.frx":0D5B
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_Costadjreport.frx":0EAD
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1140
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   630
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138739713
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   240
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138739713
      CurrentDate     =   39203
   End
   Begin Crystal.CrystalReport rptADJCOST 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   765
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "frmPMIS_REPORT_ADJCOST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'............./´?/)............. (\?`\
'............/....//..............\\....\
'.........../....//............ ....\\....\
'...../´?/..../´?\.........../? `\....\?`\
'.././.../..../..../.|_...._| .\....\....\...\.\..
'(.(....(....(..../.)..)..(..(. \....)....)....).)
'.\................\/.../....\. ..\/............/
'..\................. /........\.............../
'....\..............(...........\............./
'......\.............\...........\.........../
Public XTYPEREPORT                                  As String
Dim fdate                                           As String
Dim tdate                                           As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    fdate = CDate(dtpFromDate.Value)
    tdate = CDate(dtpToDate)
    If (gconDMIS.Execute("Select count(*) from PMIS_COSTADJ_HD  where STATUS = 'P' and TRANDATE between '" & dtpFromDate.Value & "' and '" & dtpToDate.Value & "'").Fields(0).Value) > 0 Then
        If MsgBox("Cost Adjustment will be printed.Are you Sure?", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        
        rptADJCOST.Formulas(1) = "Companyname = '" & COMPANY_NAME & "'"
        rptADJCOST.Formulas(2) = "Companyaddress = '" & COMPANY_NAME & "'"
        rptADJCOST.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
        PrintSQLReport rptADJCOST, PMIS_REPORT_PATH & "COSTADJ.RPT", "{PMIS_COSTADJ_HD.TYPE} = '" & XTYPEREPORT & "' AND {PMIS_COSTADJ_HD.TRANDATE} >= date(" & Year(fdate) & "," & Month(fdate) & "," & Day(fdate) & ") AND {PMIS_COSTADJ_HD.TRANDATE} <= date(" & Year(tdate) & "," & Month(tdate) & "," & Day(tdate) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Exit Sub
ErrorCode:
    MsgBox err.Description
    err.Clear
End Sub


Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Call CenterMe(frmMain, Me, 1)
    dtpFromDate.Value = firstDay(LOGDATE)
    dtpToDate.Value = LOGDATE
End Sub
