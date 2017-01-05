VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCSMSWorkInProgressSublet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Work in Progress - Sublet"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3435
   Icon            =   "frmCSMSWorkInProgressSublet.frx":0000
   ScaleHeight     =   2580
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm1 
      Height          =   1095
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   3015
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmCSMSWorkInProgressSublet.frx":1082
         Left            =   840
         List            =   "frmCSMSWorkInProgressSublet.frx":1084
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Select month from the list"
         Top             =   240
         Width           =   2145
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select year from the list"
         Top             =   630
         Width           =   2145
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
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
         Left            =   0
         TabIndex        =   10
         Top             =   660
         Width           =   735
      End
   End
   Begin VB.Frame frm2 
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   3015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   245563393
         CurrentDate     =   41034
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.ComboBox cbotype 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmCSMSWorkInProgressSublet.frx":1086
      Left            =   360
      List            =   "frmCSMSWorkInProgressSublet.frx":1088
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      MouseIcon       =   "frmCSMSWorkInProgressSublet.frx":108A
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkInProgressSublet.frx":11DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      MouseIcon       =   "frmCSMSWorkInProgressSublet.frx":167B
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkInProgressSublet.frx":17CD
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1680
      Width           =   735
   End
   Begin Crystal.CrystalReport rptWork_In_Progress 
      Left            =   720
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Work In Progress Monitoring Report"
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer91 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   255
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCSMSWorkInProgressSublet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotype_Change()
    If cbotype = "As Of" Then
        frm1.Visible = False
        frm2.Visible = True
    Else
        frm2.Visible = False
        frm1.Visible = True
    End If
End Sub

Private Sub CBOtype_Click()
    If cbotype = "As Of" Then
        frm1.Visible = False
        frm2.Visible = True
    Else
        frm2.Visible = False
        frm1.Visible = True
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMainMenu, Me, 1
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear

    cboMonth.Text = MonthName(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    
    cbotype.AddItem "Monthly"
    cbotype.AddItem "As Of"
    cbotype.ListIndex = 0
    
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "WORKING IN PROGRESS SUBLET") = False Then Exit Sub
    Screen.MousePointer = 11

    Dim XXX As String
    Dim CrApp As CRAXDRT.Application
    Dim CrRep As CRAXDRT.Report
    
    XXX = MonthName(Month(DTPicker1.Value))
    
    If cbotype.Text = "As Of" Then
        Set CrApp = New CRAXDRT.Application
        Set CrRep = CrApp.OpenReport(CSMS_REPORT_PATH & "Work_In_Progress_Sublet_Asof.Rpt", 1)
        
        CrRep.DiscardSavedData
        CrRep.ParameterFields.GetItemByName("@COMPANYNAME").AddCurrentValue COMPANY_NAME
        CrRep.ParameterFields.GetItemByName("@COMPANYADDRESS").AddCurrentValue COMPANY_ADDRESS
        CrRep.ParameterFields.GetItemByName("@PRINTEDBY").AddCurrentValue LOGNAME
        CrRep.ParameterFields.GetItemByName("@J_DATE").AddCurrentValue (CDate(DTPicker1.Value))

    
    Else
    
        Set CrApp = New CRAXDRT.Application
        Set CrRep = CrApp.OpenReport(CSMS_REPORT_PATH & "Work_In_Progress_Sublet.Rpt", 1)
        
        CrRep.DiscardSavedData
        CrRep.ParameterFields.GetItemByName("@COMPANYNAME").AddCurrentValue COMPANY_NAME
        CrRep.ParameterFields.GetItemByName("@COMPANYADDRESS").AddCurrentValue COMPANY_ADDRESS
        CrRep.ParameterFields.GetItemByName("@PRINTEDBY").AddCurrentValue LOGNAME
        CrRep.ParameterFields.GetItemByName("@J_MONTH").AddCurrentValue What_month(cboMonth)
        CrRep.ParameterFields.GetItemByName("@J_YEAR").AddCurrentValue N2Str2IntZero(cboYear.Text)
   End If
   
    Me.WindowState = vbMaximized
    Me.BorderStyle = vbSizable
    
    With CRViewer91
        .ReportSource = CrRep
        .DisplayGroupTree = False
        .DisplayTabs = False
        .DisplayToolbar = True
        .Height = Me.Height - 800
        .Width = Me.Width
        .ZOrder 0
        
        .ViewReport
    End With
        
    Set CrRep = Nothing
    Set CrApp = Nothing
    
    Call NEW_LogAudit("V", "WORKING IN PROGRESS - SUBLET", "", "", "", cboMonth & " " & cboYear, "", "")
    
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


