VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Deduction_Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Late Record"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Deduction_Report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   3285
   Begin VB.CheckBox chkabsent 
      Caption         =   "Print for Summary of Absence"
      Height          =   225
      Left            =   750
      TabIndex        =   7
      Top             =   1350
      Width           =   3345
   End
   Begin VB.CheckBox chklate 
      Caption         =   "Print for Summary of Late"
      Height          =   225
      Left            =   750
      TabIndex        =   6
      Top             =   1050
      Width           =   3345
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
      Height          =   705
      Left            =   2490
      MouseIcon       =   "frmHRMS_Deduction_Report.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Deduction_Report.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1650
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
      Left            =   1830
      MouseIcon       =   "frmHRMS_Deduction_Report.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmHRMS_Deduction_Report.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1650
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
      Format          =   20578305
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
      Format          =   20578305
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
Attribute VB_Name = "frmHRMS_Deduction_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS_Deduction       As ADODB.Recordset
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    If chklate.Value = 1 Then
        Call Late
    End If
    If chkabsent.Value = 1 Then
        Call Absent
    End If
End Sub

Sub Late()

Dim RSTMP                                           As New ADODB.Recordset
Dim xlApp As Excel.Application
Dim xlbook As Excel.Workbook
Dim xlsheet As Excel.Worksheet
Dim cmd As ADODB.Command
Set cmd = New ADODB.Command



        cmd.NamedParameters = True
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SP_HRMS_MONTHLY_LATE_RECORD"
        cmd.ActiveConnection = gconDMIS
        cmd.Parameters.Append cmd.CreateParameter("@FROMDATE", adDBDate, adParamInput, , dtpFromDate.Value)
        cmd.Parameters.Append cmd.CreateParameter("@TODATE", adDBDate, adParamInput, , dtpToDate.Value)
        Set RSTMP = cmd.Execute
        
        
        If Not (RSTMP.EOF And RSTMP.BOF) Then
                If Len(Dir(HRMS_REPORT_PATH & "Monthly Late Record.xlt")) = 0 Then
                    MessagePop InfoStop, "Error", "Monthly Late Record.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                Set xlApp = New Excel.Application
                Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "Monthly Late Record.xlt")
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

Sub Absent()

        Dim RSTMP                                           As New ADODB.Recordset
        Dim xlApp As Excel.Application
        Dim xlbook As Excel.Workbook
        Dim xlsheet As Excel.Worksheet
        Dim cmd As ADODB.Command
        Set cmd = New ADODB.Command
        
        
        cmd.NamedParameters = True
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SP_HRMS_MONTHLY_ABSENT_RECORD"
        cmd.ActiveConnection = gconDMIS
        cmd.Parameters.Append cmd.CreateParameter("@FROMDATE", adDBDate, adParamInput, , dtpFromDate.Value)
        cmd.Parameters.Append cmd.CreateParameter("@TODATE", adDBDate, adParamInput, , dtpToDate.Value)
        Set RSTMP = cmd.Execute
        
        
        If Not (RSTMP.EOF And RSTMP.BOF) Then
                If Len(Dir(HRMS_REPORT_PATH & "Monthly Absent Record.xlt")) = 0 Then
                    MessagePop InfoStop, "Error", "Monthly Absent Record.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                Set xlApp = New Excel.Application
                Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "Monthly Absent Record.xlt")
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

    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    dtpFromDate.Value = firstDay(Date)
    dtpToDate.Value = Now
    
    Screen.MousePointer = 0

End Sub
