VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmcsms_report_voidro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "frmcsms_report_void.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdView 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   495
      Width           =   1545
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3120
      TabIndex        =   6
      Top             =   1185
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   3045
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
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
         Format          =   20316161
         CurrentDate     =   40249
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
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
         Format          =   20316161
         CurrentDate     =   40249
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
   End
   Begin Crystal.CrystalReport rptVOIDRO 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Vehicle By Model"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      _Version        =   655364
      _ExtentX        =   8493
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Void Repair Order Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
      GradientColorLight=   192
      GradientColorDark=   0
   End
End
Attribute VB_Name = "frmcsms_report_voidro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
Dim rschek                                     As ADODB.Recordset
Dim xtitle                                     As String
Dim fdate                                      As String
Dim tdate                                      As String

If DTPicker1.Value > DTPicker2.Value Then
    MsgBox "Invalid Date Range.", vbInformation + vbOKOnly
    DTPicker2.SetFocus
    Exit Sub
End If
fdate = CDate(DTPicker1.Value)
tdate = CDate(DTPicker2.Value)
xtitle = "Void R.0. from " + fdate + " to " + tdate
Set rschek = New ADODB.Recordset
Set rschek = gconDMIS.Execute("Select * from csms_repor where cancel_date >= '" & fdate & "' and cancel_date <= '" & tdate & "'")
If Not (rschek.EOF And rschek.BOF) Then
    rptVOIDRO.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
    rptVOIDRO.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"

    rptVOIDRO.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
    rptVOIDRO.Formulas(3) = "title = '" & xtitle & "'"

    PrintSQLReport rptVOIDRO, CSMS_REPORT_PATH & "VOIDRO_REPORT.rpt", "{CSMS_Repor.cancel_date} >= date(" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & ") AND {CSMS_Repor.cancel_date} <= date(" & Year(DTPicker2.Value) & "," & Month(DTPicker2.Value) & "," & Day(DTPicker2.Value) & ")", CSMS_REPORT_CONNECTION, 1

    Call NEW_LogAudit("V", "VOID R.O. REPORT", "", "", "", DTPicker1.Value & " - " & DTPicker2.Value, "", "")
Else
    ShowNoRecord
End If
End Sub

Private Sub Form_Load()
   Call CenterMe(frmMain, Me, 1)
    DTPicker1.Value = LOGDATE
    DTPicker2.Value = LOGDATE
End Sub

