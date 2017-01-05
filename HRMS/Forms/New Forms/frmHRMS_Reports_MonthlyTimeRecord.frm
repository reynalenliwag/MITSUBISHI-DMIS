VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Reports_MonthlyTimeRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Time Record"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Reports_MonthlyTimeRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   3465
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
      Height          =   705
      Left            =   2550
      MouseIcon       =   "frmHRMS_Reports_MonthlyTimeRecord.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_MonthlyTimeRecord.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   930
      Width           =   675
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
      Height          =   705
      Left            =   1890
      MouseIcon       =   "frmHRMS_Reports_MonthlyTimeRecord.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Reports_MonthlyTimeRecord.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   930
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52625409
      CurrentDate     =   40330
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   510
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52625409
      CurrentDate     =   40359
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "From :"
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
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "To :"
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
      Left            =   465
      TabIndex        =   4
      Top             =   600
      Width           =   300
   End
End
Attribute VB_Name = "frmHRMS_Reports_MonthlyTimeRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11

    Dim RSTMP                                           As New ADODB.Recordset
    Dim XXX                                             As String
    Dim xlApp                                           As Excel.Application
    Dim xlbook                                          As Excel.Workbook
    Dim xlsheet                                         As Excel.Worksheet
    Dim cmd                                             As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.NamedParameters = True
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_HRMS_MONTHLY_TIME_ATTANDANCE"
    cmd.ActiveConnection = gconDMIS
    cmd.Parameters.Append cmd.CreateParameter("@FROMDATE", adDBDate, adParamInput, , dtpFromDate.Value)
    cmd.Parameters.Append cmd.CreateParameter("@TODATE", adDBDate, adParamInput, , dtpToDate.Value)
    Set RSTMP = cmd.Execute
    
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        If Len(Dir(HRMS_REPORT_PATH & "Monthly Time Record.xlt")) = 0 Then
            MessagePop InfoStop, "Error", "Monthly Time Record.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Set xlApp = New Excel.Application
        Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "Monthly Time Record.xlt")
        Set xlsheet = xlbook.Worksheets(1)
        
        xlsheet.Range("A7").CopyFromRecordset RSTMP
        xlApp.Visible = True
        If Not xlbook Is Nothing Then
            Set xlbook = Nothing
            Set xlApp = Nothing
        End If
        Set xlApp = Nothing
    Else
        Call ShowNoRecord
    End If
    Set RSTMP = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    dtpFromDate.Value = firstDay(Date)
    dtpToDate.Value = Now
End Sub

