VERSION 5.00
Begin VB.Form frmTechnicianReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Report"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSTechnicianReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   3705
   Begin VB.ComboBox cboTech 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCSMSTechnicianReport.frx":1082
      Left            =   90
      List            =   "frmCSMSTechnicianReport.frx":1095
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   330
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Technician Labor Sales Report"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   2850
      Width           =   3870
   End
   Begin VB.OptionButton optTechnicianAttendance 
      Caption         =   "Technician Attendance"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   2535
      Width           =   3870
   End
   Begin VB.OptionButton optTechnicianPerformance 
      Caption         =   "Technician Performance"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   1
      Top             =   2235
      Width           =   3870
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
      Height          =   825
      Left            =   2850
      MouseIcon       =   "frmCSMSTechnicianReport.frx":1110
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianReport.frx":1262
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2130
      MouseIcon       =   "frmCSMSTechnicianReport.frx":16AD
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianReport.frx":17FF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "View Report Window"
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmTechnicianReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cboTech.Text = "Technician Performance" Then
        frmCSMSTechnicianPerformanceReport.Show
    ElseIf cboTech.Text = "Technician Labor Sales" Then
        frmCSMSTechnicianLaborSales.Show
    ElseIf cboTech.Text = "Technician Productivity" Then
        frmCSMSTechnician_Efficiency.Show
    ElseIf cboTech.Text = "Technician Workshop" Then
'        frmCSMSTechnician_Efficiency.Show
        frmCSMSTechnicianWorkshopReport.Show
    Else
        frmCSMSTechnicianAttendance.Show
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub

            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (TECHNCIAN REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "TECHNICIAN REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0

    cboTech.Text = "Technician Performance"
End Sub

