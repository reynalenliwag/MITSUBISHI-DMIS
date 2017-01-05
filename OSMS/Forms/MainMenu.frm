VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSMS Main Menu"
   ClientHeight    =   5955
   ClientLeft      =   990
   ClientTop       =   1065
   ClientWidth     =   10170
   ForeColor       =   &H8000000F&
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10170
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      _Version        =   655364
      _ExtentX        =   18177
      _ExtentY        =   10557
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   3
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tbPageMainModules"
      Item(1).Caption =   "Inquiry"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tbPageFileMaintenance"
      Item(2).Caption =   "Reports"
      Item(2).Tooltip =   "Detail Report Regarding Payrolls"
      Item(2).ImageIndex=   1
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tbPageReport"
      Begin XtremeSuiteControls.TabControlPage tbPageReport 
         Height          =   5355
         Left            =   -69970
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   10245
         _Version        =   655364
         _ExtentX        =   18071
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton Command3 
            Height          =   585
            Left            =   300
            Picture         =   "MainMenu.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Height          =   585
            Left            =   300
            Picture         =   "MainMenu.frx":6CB4
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Height          =   585
            Left            =   300
            Picture         =   "MainMenu.frx":7116
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton Command16 
            Height          =   585
            Left            =   300
            Picture         =   "MainMenu.frx":7578
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Receipt by Supplier"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1110
            TabIndex        =   31
            Top             =   2550
            Width           =   4245
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Issuance by Department"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1110
            TabIndex        =   29
            Top             =   1830
            Width           =   4245
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Status Report "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1110
            TabIndex        =   27
            Top             =   1110
            Width           =   3225
         End
         Begin VB.Label Label62 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Inventory Rank"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1110
            TabIndex        =   25
            Top             =   450
            Width           =   3225
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageFileMaintenance 
         Height          =   5355
         Left            =   -69970
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   10245
         _Version        =   655364
         _ExtentX        =   18071
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton cmdFileCompanyProfile 
            Height          =   585
            Left            =   390
            Picture         =   "MainMenu.frx":79DA
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   300
            Width           =   615
         End
         Begin VB.CommandButton cmdFileBegBal 
            Height          =   585
            Left            =   390
            Picture         =   "MainMenu.frx":7EC9
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1020
            Width           =   615
         End
         Begin VB.CommandButton cmdFilePrevEmployer 
            Height          =   585
            Left            =   390
            Picture         =   "MainMenu.frx":81D3
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Issused"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1230
            TabIndex        =   12
            Top             =   1170
            Width           =   3225
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Inventory"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1230
            TabIndex        =   11
            Top             =   1950
            Width           =   3345
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Recieved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1230
            TabIndex        =   10
            Top             =   390
            Width           =   3225
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageMainModules 
         Height          =   5355
         Left            =   30
         TabIndex        =   1
         Top             =   600
         Width           =   10245
         _Version        =   655364
         _ExtentX        =   18071
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton Command4 
            Height          =   585
            Left            =   300
            Picture         =   "MainMenu.frx":84DD
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3960
            Width           =   615
         End
         Begin VB.CommandButton Command10 
            Height          =   585
            Left            =   330
            Picture         =   "MainMenu.frx":A84F
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "1102"
            Top             =   3240
            Width           =   615
         End
         Begin VB.CommandButton cmdRegProb 
            Height          =   585
            Left            =   330
            Picture         =   "MainMenu.frx":B491
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   255
            Width           =   615
         End
         Begin VB.CommandButton cmdContract 
            Height          =   585
            Left            =   330
            Picture         =   "MainMenu.frx":D803
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1005
            Width           =   615
         End
         Begin VB.CommandButton cmdAllowance 
            Height          =   585
            Left            =   330
            Picture         =   "MainMenu.frx":FB75
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1755
            Width           =   615
         End
         Begin VB.CommandButton cmdAPP 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   330
            Picture         =   "MainMenu.frx":11EE7
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2445
            Width           =   615
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receving Supply"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   960
            TabIndex        =   33
            Top             =   4125
            Width           =   2385
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Departments"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6540
            TabIndex        =   23
            Top             =   1800
            Visible         =   0   'False
            Width           =   3405
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Supply"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6540
            TabIndex        =   22
            Top             =   1320
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Signatories"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6540
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Suppliers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6540
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   3795
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1020
            TabIndex        =   19
            Top             =   2580
            Width           =   3915
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Reminders"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1080
            TabIndex        =   18
            Top             =   3300
            Width           =   2490
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receving Supply"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   990
            TabIndex        =   6
            Top             =   420
            Width           =   2385
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Issuance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   990
            TabIndex        =   5
            Top             =   1110
            Width           =   2355
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supply Inventory Adjustment"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   990
            TabIndex        =   4
            Top             =   1860
            Width           =   4215
         End
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAllowance_Click()
    frmOSMSTransactionInvAdjustment.Show
    frmOSMSTransactionInvAdjustment.ZOrder 0
End Sub

Private Sub cmdAPP_Click()
    frmOSMSFilesEmployee.Show
    frmOSMSFilesEmployee.ZOrder 0
End Sub

Private Sub cmdContract_Click()
    frmOSMSTransactionIssuance.Show
    frmOSMSTransactionIssuance.ZOrder 0
End Sub

Private Sub cmdFileBegBal_Click()
    frmOSMSInquiryIssued.Show
    frmOSMSInquiryIssued.ZOrder 0

End Sub

Private Sub cmdFileCompanyProfile_Click()


    frmOSMSInquiryReceiving.Show
    frmOSMSInquiryReceiving.ZOrder 0


End Sub

Private Sub cmdFilePrevEmployer_Click()
    frmOSMSInquirySupply.Show
    frmOSMSInquirySupply.ZOrder 0
End Sub

Private Sub cmdRegProb_Click()
    frmOSMSTransactionReceivingSupply.Show
    frmOSMSTransactionReceivingSupply.ZOrder 0

End Sub

Private Sub Command10_Click()
'Upating Code       :AXP-062620071225
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub Command2_Click()
    frmOSMSReportDepartment.Show
    frmOSMSReportDepartment.ZOrder 0
End Sub

Private Sub Command3_Click()
    frmOSMSReportSupplier.Show
    frmOSMSReportSupplier.ZOrder 0
End Sub

Private Sub Command4_Click()
    frmOSMS_PurchaseOrder.Show
End Sub

'    Case FILE_DEPARTMENT
'        Screen.MousePointer = 11
'        frmOSMSFilesDepartment.Show
'        frmOSMSFilesDepartment.ZOrder 0
'        Screen.MousePointer = 0
'    Case FILE_EMPLOYEE
'
'    Case FILE_SIGNATORIES
'        frmOSMSFilesSignatories.Show
'        frmOSMSFilesSignatories.ZOrder 0
'    Case FILE_SUPPLIER
'        frmOSMSFilesSupplier.Show
'        frmOSMSFilesSupplier.ZOrder 0
'    Case FILE_SUPPLY
'        frmOSMSFilesSupply.Show
'        frmOSMSFilesSupply.ZOrder 0
'    Case FILE_UNIT
'        frmOSMSFilesUnit.Show
'        frmOSMSFilesUnit.ZOrder 0
'
'



Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub
