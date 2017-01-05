VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicles Sales Monitoring"
   ClientHeight    =   8205
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   13680
   ClipControls    =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   912
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   7800
      Left            =   2220
      TabIndex        =   7
      Top             =   390
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   13758
      _StockProps     =   64
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   7
      Item(0).Caption =   "&Prospects"
      Item(0).Tooltip =   "Prospects"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "lstProspects"
      Item(0).Control(1)=   "picOptionProspects"
      Item(0).Control(2)=   "Picture2"
      Item(0).Control(3)=   "Combo2"
      Item(0).Control(4)=   "txtSearch_Prospect"
      Item(1).Caption =   "Quotations"
      Item(1).Tooltip =   "Quotations"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "Picture1"
      Item(1).Control(1)=   "lvQuotation"
      Item(1).Control(2)=   "txtSearch_Quotation"
      Item(2).Caption =   "&Sales Order"
      Item(2).Tooltip =   "Sales Order"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lstSalesOrder"
      Item(2).Control(1)=   "picOptSalesOrder"
      Item(3).Caption =   "&Loan Application"
      Item(3).Tooltip =   "Loan Application"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "lstIndividual"
      Item(3).Control(1)=   "picOptLoan(1)"
      Item(3).Control(2)=   "txtSearch_LoanApplication"
      Item(3).Control(3)=   "LABLOANNOTES"
      Item(4).Caption =   "&Vehicles"
      Item(4).Tooltip =   "Vehicles"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "lstVehicles"
      Item(4).Control(1)=   "picOptLoan(2)"
      Item(4).Control(2)=   "txtSearch_Vehicles"
      Item(5).Caption =   "LTO Monitoring"
      Item(5).ControlCount=   5
      Item(5).Control(0)=   "txtSearch_LTO"
      Item(5).Control(1)=   "ReportControl1"
      Item(5).Control(2)=   "Combo1"
      Item(5).Control(3)=   "Label1"
      Item(5).Control(4)=   "Command3"
      Item(6).Caption =   "Reminders && Tasks"
      Item(6).ControlCount=   10
      Item(6).Control(0)=   "Picture4"
      Item(6).Control(1)=   "lstReminders"
      Item(6).Control(2)=   "Text5"
      Item(6).Control(3)=   "Label10"
      Item(6).Control(4)=   "cboSAE"
      Item(6).Control(5)=   "Label9"
      Item(6).Control(6)=   "cboStatus"
      Item(6).Control(7)=   "Label7"
      Item(6).Control(8)=   "cboPriority"
      Item(6).Control(9)=   "Label8"
      Begin XtremeReportControl.ReportControl lstReminders 
         Height          =   6570
         Left            =   -69970
         TabIndex        =   62
         Top             =   1170
         Visible         =   0   'False
         Width           =   9495
         _Version        =   655364
         _ExtentX        =   16748
         _ExtentY        =   11589
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstSalesOrder 
         Height          =   6630
         Left            =   -69970
         TabIndex        =   8
         Top             =   1110
         Visible         =   0   'False
         Width           =   11400
         _Version        =   655364
         _ExtentX        =   20108
         _ExtentY        =   11695
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstVehicles 
         Height          =   6675
         Left            =   -69970
         TabIndex        =   28
         Top             =   1110
         Visible         =   0   'False
         Width           =   11370
         _Version        =   655364
         _ExtentX        =   20055
         _ExtentY        =   11774
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstProspects 
         Height          =   6705
         Left            =   30
         TabIndex        =   10
         Top             =   1050
         Width           =   8640
         _Version        =   655364
         _ExtentX        =   15240
         _ExtentY        =   11827
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lvQuotation 
         Height          =   6750
         Left            =   -69970
         TabIndex        =   42
         Top             =   1020
         Visible         =   0   'False
         Width           =   11400
         _Version        =   655364
         _ExtentX        =   20108
         _ExtentY        =   11906
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl ReportControl1 
         Height          =   6555
         Left            =   -69970
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   11370
         _Version        =   655364
         _ExtentX        =   20055
         _ExtentY        =   11562
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin XtremeReportControl.ReportControl lstIndividual 
         Height          =   5730
         Left            =   -69970
         TabIndex        =   9
         Top             =   1170
         Visible         =   0   'False
         Width           =   11415
         _Version        =   655364
         _ExtentX        =   20135
         _ExtentY        =   10107
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   6675
         Left            =   8700
         ScaleHeight     =   6645
         ScaleWidth      =   2715
         TabIndex        =   74
         Top             =   1050
         Width           =   2745
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3195
            Left            =   0
            ScaleHeight     =   3195
            ScaleWidth      =   4050
            TabIndex        =   75
            Top             =   3900
            Width           =   4050
            Begin VB.TextBox lblNotes 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   80
               Top             =   30
               Width           =   2655
            End
            Begin VB.TextBox txtCusPhone 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               HideSelection   =   0   'False
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   1530
               Width           =   1995
            End
            Begin VB.TextBox txtCusAdd 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   645
               HideSelection   =   0   'False
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   2070
               Width           =   2655
            End
            Begin VB.TextBox txtCusEmail 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               HideSelection   =   0   'False
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   1260
               Width           =   1995
            End
            Begin VB.TextBox txtCusName 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   465
               HideSelection   =   0   'False
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   780
               Width           =   2655
            End
            Begin VB.Label lblCustEmail 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Email :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   83
               Top             =   1260
               Width           =   675
            End
            Begin VB.Label lblCustPhone 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Phone:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   82
               Top             =   1530
               Width           =   675
            End
            Begin VB.Label lblCustAddress 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   30
               TabIndex        =   81
               Top             =   1800
               Width           =   2655
            End
         End
         Begin VB.Label lblAppt 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Appointment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   48
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblCalls 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Calls"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   105
            Top             =   1830
            Width           =   945
         End
         Begin VB.Label lblLetters 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Letters"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   104
            Top             =   2130
            Width           =   945
         End
         Begin VB.Label lblTest 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Test Drive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   103
            Top             =   630
            Width           =   945
         End
         Begin VB.Label lblVisits 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Visits"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   102
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label lblLoan 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Loan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   101
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label lblSalesOrder 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Sales order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   100
            Top             =   1230
            Width           =   945
         End
         Begin VB.Label lblLogSalesOrder 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   99
            Top             =   1230
            Width           =   1695
         End
         Begin VB.Label lblEmails 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   98
            Top             =   2430
            Width           =   945
         End
         Begin XtremeShortcutBar.ShortcutCaption captionInformation 
            Height          =   315
            Left            =   0
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   0
            Width           =   2715
            _Version        =   655364
            _ExtentX        =   4789
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Profile @"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Invoice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   96
            Top             =   3030
            Width           =   945
         End
         Begin VB.Label lblLogEmail 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   95
            Top             =   2430
            Width           =   1695
         End
         Begin VB.Label lblLogLoan 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   94
            Top             =   2730
            Width           =   1695
         End
         Begin VB.Label lblLogVisits 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   93
            Top             =   1530
            Width           =   1695
         End
         Begin VB.Label lblLogAppointment 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   92
            ToolTipText     =   "Last sales appointment made on and days elasped"
            Top             =   932
            Width           =   1695
         End
         Begin VB.Label lblLogTestDrive 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   91
            ToolTipText     =   " Test Drive Schedules On and Day Elasped"
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label lblLogCalls 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   90
            Top             =   1830
            Width           =   1695
         End
         Begin VB.Label lblLogLetters 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   89
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label lblInvoiceNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   88
            Top             =   3030
            Width           =   1695
         End
         Begin VB.Label lblQ 
            BackColor       =   &H00E0E0E0&
            Caption         =   " Quotation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   30
            TabIndex        =   87
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblLogQuote 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   990
            TabIndex        =   86
            ToolTipText     =   " Last Quotation Send "
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label lblAgeing 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   30
            TabIndex        =   85
            Top             =   3630
            Width           =   2640
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   270
            Left            =   30
            TabIndex        =   84
            Top             =   3330
            Width           =   2640
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   630
         Width           =   2235
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -69940
         ScaleHeight     =   525
         ScaleWidth      =   9465
         TabIndex        =   57
         Top             =   630
         Visible         =   0   'False
         Width           =   9465
         Begin VB.OptionButton optReminder 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Internal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   405
            Index           =   3
            Left            =   1260
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   60
            Width           =   1215
         End
         Begin VB.TextBox txtSearch_Activity 
            Height          =   375
            Left            =   6030
            TabIndex        =   71
            Top             =   90
            Width           =   3375
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H8000000D&
            Caption         =   "Prospect"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Customers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Index           =   1
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   60
            Width           =   1215
         End
         Begin VB.OptionButton optReminder 
            BackColor       =   &H0080FFFF&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   405
            Index           =   2
            Left            =   3780
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   60
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Filter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   5400
            TabIndex        =   72
            Top             =   180
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   330
         Left            =   -59800
         TabIndex        =   53
         Top             =   690
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "MainForm.frx":030A
         Left            =   -63640
         List            =   "MainForm.frx":0323
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   660
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.TextBox txtSearch_LTO 
         Height          =   375
         Left            =   -69850
         TabIndex        =   49
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtSearch_Quotation 
         Height          =   375
         Left            =   -69910
         TabIndex        =   47
         Top             =   630
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.PictureBox picOptSalesOrder 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   -69940
         ScaleHeight     =   540
         ScaleWidth      =   11400
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   11400
         Begin VB.OptionButton optSO 
            BackColor       =   &H00008000&
            Caption         =   "On Process"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "On Process"
            Top             =   60
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optSO 
            BackColor       =   &H00000080&
            Caption         =   "Cancelled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   6210
            Style           =   1  'Graphical
            TabIndex        =   20
            Tag             =   "Cancelled"
            Top             =   60
            Width           =   1035
         End
         Begin VB.TextBox txtSearch_SalesOrder 
            Height          =   405
            Left            =   0
            TabIndex        =   108
            Top             =   60
            Width           =   2925
         End
         Begin VB.ComboBox Combo3 
            Height          =   345
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   107
            Top             =   90
            Width           =   1995
         End
         Begin VB.CommandButton cmdPrintSO 
            Caption         =   "Print"
            Height          =   375
            Left            =   10290
            TabIndex        =   37
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optSO 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   9435
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "All"
            Top             =   60
            Width           =   855
         End
         Begin VB.OptionButton optSO 
            Caption         =   "Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   8325
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "All"
            Top             =   60
            Width           =   1095
         End
         Begin VB.OptionButton optSO 
            BackColor       =   &H8000000D&
            Caption         =   "Invoiced"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   7275
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "Invoiced"
            Top             =   60
            Width           =   1035
         End
      End
      Begin Crystal.CrystalReport rptMain 
         Left            =   10350
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.PictureBox picOptLoan 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   -66490
         ScaleHeight     =   435
         ScaleWidth      =   7875
         TabIndex        =   32
         Top             =   645
         Visible         =   0   'False
         Width           =   7875
         Begin VB.OptionButton optInventory 
            BackColor       =   &H000000C0&
            Caption         =   "Invoiced"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "S"
            ToolTipText     =   "View Invoiced Vehicles"
            Top             =   30
            Width           =   1140
         End
         Begin VB.CommandButton cmdPrintVehicles 
            Caption         =   "Print"
            Height          =   360
            Left            =   6630
            TabIndex        =   36
            ToolTipText     =   "Print Details"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00404040&
            Caption         =   "Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   4305
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "R"
            ToolTipText     =   "View Released Vehicles"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00008000&
            Caption         =   "On Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "O"
            ToolTipText     =   "View On Stock Vehicles"
            Top             =   30
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optInventory 
            BackColor       =   &H00000080&
            Caption         =   "Allocated"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   1935
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "A"
            ToolTipText     =   "View Allocated Vehicles"
            Top             =   30
            Width           =   1140
         End
      End
      Begin VB.TextBox txtSearch_Vehicles 
         Height          =   375
         Left            =   -69880
         TabIndex        =   29
         Top             =   690
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtSearch_Prospect 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   630
         Width           =   2250
      End
      Begin VB.TextBox txtSearch_LoanApplication 
         Height          =   405
         Left            =   -69880
         TabIndex        =   26
         Top             =   630
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.PictureBox picOptLoan 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   -64480
         ScaleHeight     =   435
         ScaleWidth      =   6195
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   6195
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00404040&
            Caption         =   "Cancelled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   4
            Left            =   4710
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "C"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H000000C0&
            Caption         =   "Disapproved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   2325
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "D"
            Top             =   30
            Width           =   1215
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00000080&
            Caption         =   "Pending"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   1
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "P"
            Top             =   30
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00800000&
            Caption         =   "On Process"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   1170
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "O"
            Top             =   30
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optLoan 
            BackColor       =   &H00008000&
            Caption         =   "Approved"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   3
            Left            =   3555
            Style           =   1  'Graphical
            TabIndex        =   12
            Tag             =   "A"
            Top             =   30
            Width           =   1140
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   -66190
         ScaleHeight     =   405
         ScaleWidth      =   5925
         TabIndex        =   43
         Top             =   630
         Visible         =   0   'False
         Width           =   5925
         Begin VB.OptionButton optQuote 
            BackColor       =   &H00C0C0C0&
            Caption         =   "On Sales Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   1
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "S"
            Top             =   0
            Width           =   1470
         End
         Begin VB.OptionButton optQuote 
            BackColor       =   &H0000C000&
            Caption         =   "On Process"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "P"
            Top             =   0
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton optQuote 
            BackColor       =   &H00C00000&
            Caption         =   "On Invoice"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Index           =   2
            Left            =   3105
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "I"
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.ComboBox cboPriority 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -60430
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2010
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -60430
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1410
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboSAE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -60430
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   810
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00400000&
         Height          =   3825
         Left            =   -60430
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   2700
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox picOptionProspects 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   4680
         ScaleHeight     =   405
         ScaleWidth      =   6675
         TabIndex        =   21
         Top             =   630
         Width           =   6675
         Begin VB.OptionButton optProspects 
            BackColor       =   &H008080FF&
            Caption         =   "Lost Sale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   5
            Left            =   4740
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   0
            Width           =   945
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Print"
            Height          =   375
            Left            =   5700
            TabIndex        =   110
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C0C0C0&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   3810
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   945
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C00000&
            Caption         =   "Follow Ups"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H0000C000&
            Caption         =   "Open"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   0
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inactive"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   1830
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   945
         End
         Begin VB.OptionButton optProspects 
            BackColor       =   &H000040C0&
            Caption         =   "Closed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   930
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.Label LABLOANNOTES 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   825
         Left            =   -69940
         TabIndex        =   109
         Top             =   6930
         Visible         =   0   'False
         Width           =   11355
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   315
         Left            =   -64270
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -60430
         TabIndex        =   70
         Top             =   1770
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -60430
         TabIndex        =   68
         Top             =   1170
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label9 
         Caption         =   "SAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   165
         Left            =   -60430
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label10 
         Caption         =   "Follow Up Note"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -60430
         TabIndex        =   64
         Top             =   2460
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboYear 
      Height          =   345
      ItemData        =   "MainForm.frx":03D4
      Left            =   75
      List            =   "MainForm.frx":03D6
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   90
      Width           =   2115
   End
   Begin VB.ListBox lstMonth 
      Appearance      =   0  'Flat
      Height          =   3285
      IntegralHeight  =   0   'False
      Left            =   60
      MouseIcon       =   "MainForm.frx":03D8
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   510
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   3735
      Width           =   2205
      Begin VB.CommandButton Command10 
         Caption         =   "Process Prospects Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         MouseIcon       =   "MainForm.frx":052A
         MousePointer    =   99  'Custom
         TabIndex        =   112
         ToolTipText     =   "View Prospects"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sales Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":067C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "View Sales Order"
         Top             =   915
         Width           =   1995
      End
      Begin VB.CommandButton cmddata 
         Caption         =   "Create Prospect data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   111
         Top             =   3345
         Width           =   1995
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Reminders/Task"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":07CE
         MousePointer    =   99  'Custom
         TabIndex        =   56
         ToolTipText     =   "View Customer"
         Top             =   2985
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Loan Application (C)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0920
         MousePointer    =   99  'Custom
         TabIndex        =   55
         ToolTipText     =   "View Customer"
         Top             =   2640
         Width           =   1995
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Loan Application (I)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0A72
         MousePointer    =   99  'Custom
         TabIndex        =   54
         ToolTipText     =   "View Customer"
         Top             =   2295
         Width           =   1995
      End
      Begin VB.CommandButton cmdSearchCustomer 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0BC4
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "View Customer"
         Top             =   1950
         Width           =   1995
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Sales Calculator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0D16
         MousePointer    =   99  'Custom
         TabIndex        =   1
         ToolTipText     =   "View Sales Calculator"
         Top             =   1605
         Width           =   1995
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0E68
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "View Inquiry"
         Top             =   1260
         Width           =   1995
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Prospects"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MouseIcon       =   "MainForm.frx":0FBA
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "View Prospects"
         Top             =   570
         Width           =   1995
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   3
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   2100
         _Version        =   655364
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   ":: OPTIONS  :::"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   64
      End
   End
   Begin VB.Label labVCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   2220
      TabIndex        =   106
      Top             =   0
      Width           =   11385
   End
   Begin VB.Menu mnuContextProspect 
      Caption         =   "CONTEXT PROSPECT"
      Begin VB.Menu mnuTestDrive 
         Caption         =   "Schedule Test Drive"
      End
      Begin VB.Menu mnuSalesAppointment 
         Caption         =   "Add Sales Appoinment"
      End
      Begin VB.Menu mnuSendQuotation 
         Caption         =   "Add Quotation"
      End
      Begin VB.Menu mnuLoanApplication 
         Caption         =   "Add Loan Application"
      End
      Begin VB.Menu mnuSalesOrder 
         Caption         =   "Make Sales Order"
      End
      Begin VB.Menu mnuShowGroup 
         Caption         =   "Show By Group "
      End
      Begin VB.Menu mspc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisits 
         Caption         =   "Log Visits"
      End
      Begin VB.Menu mnuCalls 
         Caption         =   "Log Calls"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Log Email"
      End
      Begin VB.Menu mnuLetter 
         Caption         =   "Log A Letter"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "View Prospect Inquiry"
      End
      Begin VB.Menu mnuUpdateProspectStatus 
         Caption         =   "Update Prospect Status"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PSTATUS                                                           As String
Dim SSTATUS                                                           As String
Dim lStatus                                                           As String
Dim VSTATUS                                                           As String
Dim QSTATUS                                                           As String
Dim RSTATUS                                                           As String
Dim Exist_SO                                                          As Boolean
Dim EXIST_LOAN                                                        As Boolean
Dim EXIST_INVOICE                                                     As Boolean
Dim CustomerCode                                                      As String
Dim WithEvents CustomerInformation                                    As frmAllCustomer
Attribute CustomerInformation.VB_VarHelpID = -1
Dim ReportFilter                                                      As String
Dim ProspID                                                           As Long
Dim ProspType                                                         As String
Dim RPTCAPTION                                                        As String

Sub FillAndConfigGrid()
    fillcbomoreyear cboYear
    With cboPriority
        .AddItem "Normal"
        .AddItem "High"
        .AddItem "Low"
        .AddItem "(ANY)"
    End With
    With cbostatus
        .AddItem "Not Started"
        .AddItem "In Progress"
        .AddItem "Completed"
        .AddItem "Waiting"
        .AddItem "Deferred"
        .AddItem "(ANY)"
    End With
    With lstMonth
        .AddItem ("January")
        .AddItem ("February")
        .AddItem ("March")
        .AddItem ("April")
        .AddItem ("May")
        .AddItem ("June")
        .AddItem ("July")
        .AddItem ("August")
        .AddItem ("September")
        .AddItem ("October")
        .AddItem ("November")
        .AddItem ("December")
        .AddItem ("All Months")
    End With


    ReportControlPaintManager lstVehicles

    ReportControlAddColumnHeader lstProspects, "Open-Date,Prospect Name,Model,Description,[C],SAE,LeadSource"
    ResizeColumnHeader lstProspects, "10,20,12,20,8,12,8"
    ReportControlPaintManager lstProspects

    ReportControlAddColumnHeader lvQuotation, "Date, ProspectName, Model , Type, Status"
    ResizeColumnHeader lvQuotation, "10,20,20,8,20"
    ReportControlPaintManager lvQuotation

    ReportControlAddColumnHeader lstSalesOrder, "Date,Customer Name,Model,Model Description,SAE, Status,SONO"
    ResizeColumnHeader lstSalesOrder, "8,18,8,20,20,10,10"
    ReportControlPaintManager lstSalesOrder

    ReportControlAddColumnHeader ReportControl1, "PULLOUTDATE,RELEASEDDATE,CS#,MODEL , COLOR, CUSTOMER, LTOSTATUS ,CSR, CSRDATE"
    ResizeColumnHeader ReportControl1, "8,8,8,20,10,20,8,8,8"

    ReportControlPaintManager ReportControl1

    ReportControlAddColumnHeader lstIndividual, "Date,Financing Company, Account Name,Model,SAE,Type, Status, DateApproved"
    ResizeColumnHeader lstIndividual, "8,18,18,18,10,0,8,8"
    ReportControlPaintManager lstIndividual
    lstIndividual.GroupsOrder.Add lstIndividual.Columns(5)
    lstIndividual.Columns(5).Visible = False

    ReportControlAddColumnHeader lstReminders, "Type, Date , Due By , Name,Reminder Type ,Subject,Priority, Status,SAE"
    ResizeColumnHeader lstReminders, "0,10,5,25,15,15,10,10,20"

    ReportControlPaintManager lstReminders
    lstReminders.GroupsOrder.Add lstReminders.Columns(0)

    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select name from smis_vw_srep ")
    SetComboWidth Combo2, 200
    SetComboWidth cboSAE, 200

    While Not temprs.EOF
        cboSAE.AddItem (temprs!Name) & ""
        Combo2.AddItem (temprs!Name) & ""
        Combo3.AddItem (temprs!Name) & ""
        temprs.MoveNext
    Wend
    cboSAE.AddItem "(Any)", 0
    Combo2.AddItem "(Any)", 0
    Combo3.AddItem "(Any)", 0

    lstReminders.Columns(0).Visible = False
    lstMonth.ListIndex = Month(LOGDATE) - 1

End Sub

Sub FillLoanApplication(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Dim SQL                                                           As String
    Set RsUploadData = New ADODB.Recordset

    SQL = "Select "
    SQL = SQL & "convert(varchar, DateApplied ,101),  " & vbCrLf
    SQL = SQL & "FINCOM , " & vbCrLf
    SQL = SQL & "isnull(Ind_Apl_LastName,'')  + ' . ' + isnull(Ind_Apl_FirstName, ''), " & vbCrLf
    SQL = SQL & "Ind_LoanApl_UnitModel , " & vbCrLf
    SQL = SQL & "Ind_LoanApl_SAE , " & vbCrLf
    SQL = SQL & "'Individual' ," & vbCrLf
    SQL = SQL & "Case LStatus  " & vbCrLf
    SQL = SQL & "WHEN 'O' THEN 'On Process' " & vbCrLf
    SQL = SQL & "WHEN 'P' THEN 'Pending' " & vbCrLf
    SQL = SQL & "WHEN 'D' THEN 'Disapproved' " & vbCrLf
    SQL = SQL & "WHEN 'C' THEN 'Cancelled' " & vbCrLf
    SQL = SQL & "WHEN 'A' THEN 'Approved' END as Status ," & vbCrLf
    SQL = SQL & "lastupdated,Apl_no, ID " & vbCrLf
    SQL = SQL & "from SMIS_LoanIndiv " & vbCrLf
    SQL = SQL & " WHERE " & xDate & vbCrLf
    SQL = SQL & "UNION " & vbCrLf
    SQL = SQL & "Select " & vbCrLf
    SQL = SQL & "convert(varchar, DateApplied ,101),  " & vbCrLf
    SQL = SQL & "FINCOM , " & vbCrLf
    SQL = SQL & "Busname, " & vbCrLf
    SQL = SQL & "unitmodel, " & vbCrLf
    SQL = SQL & "SAEName ," & vbCrLf
    SQL = SQL & "'Corporate' ," & vbCrLf
    SQL = SQL & "Case LStatus  " & vbCrLf
    SQL = SQL & "WHEN 'O' THEN 'On Process' " & vbCrLf
    SQL = SQL & "WHEN 'P' THEN 'Pending' " & vbCrLf
    SQL = SQL & "WHEN 'D' THEN 'Disapproved' " & vbCrLf
    SQL = SQL & "WHEN 'C' THEN 'Cancelled' " & vbCrLf
    SQL = SQL & "WHEN 'A' THEN 'Approved' END , " & vbCrLf
    SQL = SQL & "lastupdated,Aplcode , ID" & vbCrLf
    SQL = SQL & "from SMIS_LoanCORP " & vbCrLf
    SQL = SQL & " WHERE " & xDate & vbCrLf
    SQL = SQL & "Order By 1 DESC" & vbCrLf

    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstIndividual
End Sub

Sub FILLLTOSTATUS()
    Dim FILTER, SQL

    If UCase(Combo1.Text) <> "OTHERS" Then
        FILTER = " WHERE  SMIS_MrrInv.LTOStatus='" & Combo1.Text & "'"
    Else
        FILTER = ""

    End If

    SQL = " SELECT"
    SQL = SQL & " SMIS_MrrInv.PullOutDate,"
    SQL = SQL & " SMIS_SalesOrder.DateReleased,"
    SQL = SQL & "  SMIS_MrrInv.ignkey,"
    SQL = SQL & "  SMIS_MrrInv.DESCRIPT,"
    SQL = SQL & "  SMIS_MrrInv.color,"
    SQL = SQL & "  SMIS_SalesOrder.CustName,"
    SQL = SQL & "  isnull(SMIS_MrrInv.LTOStatus,'NO STATUS') AS ltostatus,"
    SQL = SQL & "  SMIS_MrrInv.CSR,"
    SQL = SQL & "  SMIS_MrrInv.CSRDATE"
    SQL = SQL & "  FROM SMIS_MrrInv INNER JOIN"
    SQL = SQL & "  SMIS_SalesOrder ON SMIS_MrrInv.ignkey = SMIS_SalesOrder.IGNKEY_NO  "


    If FILTER <> "" Then SQL = SQL & FILTER

    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    ReportControl1.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute(SQL & " ORDER BY SMIS_MrrInv.LTOStatus")

    flex_FillReportView RsUploadData, ReportControl1


End Sub

Sub FillProspect(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    lstProspects.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute("SELECT  LogInitialInquiry , AcctName , Model,Variant, Classification , SAE , LeadSource  , ProspectType, ProspectID ,CUSCDE,LOGQUOTE,LOGSO  FROM CRIS_Prospects Where " & xDate & PSTATUS & " order by LOGINITIALINQUIRY DESC")
    flex_FillReportView RsUploadData, lstProspects
End Sub

Sub FillQuotation(xDate As String)
    'UDPATING CODE      :   AXP-061007 750PM
    Dim SQL                                                           As String
    SQL = " SELECT CQ.QUOTATIONDATE, CRIS_PROSPECTS.ACCTNAME, CQ.MODELDESCRIPT, "
    SQL = SQL & " CASE ISNULL(CQ.OPT, 'B')  WHEN 'B' THEN 'CASH/FIN'"
    SQL = SQL & " WHEN 'F' THEN   'FINANCING'"
    SQL = SQL & " WHEN 'C' THEN 'CASH'"
    SQL = SQL & " END ,"
    SQL = SQL & " CASE ISNULL(CQ.STATUS,'P')  WHEN 'P'  THEN 'ON PROCESS'"
    SQL = SQL & " WHEN 'S' THEN 'ON SALES ORDER' "
    SQL = SQL & " WHEN 'L' THEN 'ON LOAN'"
    SQL = SQL & " WHEN 'I' THEN 'ON INVOICE' END, LOGID FROM  CRIS_QUOTATION CQ"
    SQL = SQL & " INNER JOIN CRIS_PROSPECTS ON CQ.PROSPECTID = CRIS_PROSPECTS.PROSPECTID "

    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    lvQuotation.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute(SQL & " WHERE" & xDate & QSTATUS & " order by CQ.QuotationDate DESC")
    flex_FillReportView RsUploadData, lvQuotation

End Sub

Sub FillReminder(xDate As String)
    Dim SQL                                                           As String
    SQL = " SELECT "
    SQL = SQL & " Case ENTITYTYPE"
    SQL = SQL & " WHEN 'P' THEN 'PROSPECT'"
    SQL = SQL & " WHEN 'C' THEN 'CUSTOMER'"
    SQL = SQL & " WHEN 'S' THEN 'INTERNAL'"
    SQL = SQL & " END ,"
    SQL = SQL & " CONVERT(VARCHAR,DATETIMEREMIND,101) AS DATE,"
    'SQL = SQL & "  CASE "
    'SQL = SQL & " WHEN nexttime> GETDATE() THEN  DATEDIFF(DAY, GETDATE(), nexttime) "
    'SQL = SQL & " ELSE DATEDIFF(DAY, nexttime,GETDATE()) END , "
    SQL = SQL & " DATEDIFF(DAY, nexttime,GETDATE()) , "
    SQL = SQL & " Case"
    SQL = SQL & " WHEN ENTITYTYPE ='P' THEN (SELECT ACCTNAME FROM CRIS_PROSPECTS WHERE CAST(PROSPECTID AS VARCHAR) =CAST(CSCDE AS VARCHAR))"
    SQL = SQL & " WHEN ENTITYTYPE ='C' THEN (SELECT ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE  =CSCDE)"
    SQL = SQL & " WHEN ENTITYTYPE ='S' AND LEFT(CSCDE,2)='PR' THEN (SELECT ACCTNAME FROM CRIS_PROSPECTS WHERE CAST(PROSPECTID AS VARCHAR) =CAST(RIGHT(CSCDE ,LEN(CSCDE)-3) AS VARCHAR))"
    SQL = SQL & " WHEN ENTITYTYPE ='S' AND LEFT(CSCDE,2)='CS' THEN (SELECT ACCTNAME FROM ALL_CUSTOMER WHERE CUSCDE  = RIGHT(CSCDE ,LEN(CSCDE)-3)) END ,"
    SQL = SQL & " REMINDERTYPE,SUBJECT,"
    SQL = SQL & " Case Priority"
    SQL = SQL & " WHEN 'H' THEN 'HIGH'"
    SQL = SQL & " WHEN 'L' THEN 'LOW'"
    SQL = SQL & " WHEN 'N' THEN 'NORMAL'"
    SQL = SQL & " END AS PRIORITY,"
    SQL = SQL & " Case status"
    SQL = SQL & " WHEN 'N' THEN 'NOT STARTED'"
    SQL = SQL & " WHEN 'I' THEN 'IN PROGRESS'"
    SQL = SQL & " WHEN 'C' THEN 'COMPLETED'"
    SQL = SQL & " WHEN 'W' THEN 'WAITING'"
    SQL = SQL & " WHEN 'D' THEN 'DEFERRED'"
    SQL = SQL & " Else 'N/A'"
    SQL = SQL & " END As status, NAME ,CRIS_REMINDERS.ID , followupnotes"
    SQL = SQL & " From CRIS_REMINDERS LEFT OUTER JOIN SMIS_VW_SREP ON SAECODE= USERCODE"
    SQL = SQL & " WHERE ENTITYTYPE <>('E')   " & xDate & RSTATUS
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset
    lstProspects.FilterText = vbNullString
    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstReminders
End Sub

Sub FillSalesOrder(xDate As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Set RsUploadData = New ADODB.Recordset

    Set RsUploadData = gconDMIS.Execute("SELECT * FROM SMIS_VW_INQSALESORDER  WHERE " & xDate & SSTATUS & " order by Deyt DESC")
    flex_FillReportView RsUploadData, lstSalesOrder
End Sub

Sub FillVehicles(xCondition As String, Optional ByVal xDateFilter As String)
    Dim RsUploadData                                                  As ADODB.Recordset
    Dim SQL                                                           As String
    DoEvents
    If xCondition = "A" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL,SO#,DATE#A , CUSTOMERNAME, DESCRIPTION, CSNO,COLOR,AGING,SAE,TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,10,30,30,10,10,8,15,8")
        SQL = "SELECT upper(SMIS_MrrInv.MODEL)MODEL , SMIS_SalesOrder.so_no, "
        SQL = SQL & "  CONVERT(VARCHAR, SMIS_SalesOrder.deyt,101),"
        SQL = SQL & " SMIS_SalesOrder.CustName,"
        SQL = SQL & " SMIS_MrrInv.DESCRIPT,"
        SQL = SQL & " SMIS_MrrInv.ignkey,"
        SQL = SQL & " SMIS_MrrInv.color,"
        SQL = SQL & " DATEDIFF(Day, SMIS_SalesOrder.Deyt, GETDATE()) As Aging,"
        SQL = SQL & " SMIS_SalesOrder.salesae ,"
        SQL = SQL & " SMIS_SalesOrder.TERM "
        SQL = SQL & " FROM SMIS_MrrInv INNER JOIN "
        SQL = SQL & " SMIS_SalesOrder ON SMIS_MrrInv.ignkey = SMIS_SalesOrder.IGNKEY_NO"
        SQL = SQL & " WHERE " & xDateFilter & " (isnull(SMIS_SalesOrder.STATUS,'')  <>'C') and  SMIS_SalesOrder.SOSTATUS ='P' and SMIS_SalesOrder.VI_NO is null  AND (SMIS_MrrInv.iSTATUS='A') "

    ElseIf xCondition = "S" Then

        ReportControlAddColumnHeader lstVehicles, "MODEL,DATE#I, CUSTOMER NAME ,DESCRIPTION, CSNO, COLOR,SAE, TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,30,30,10,10,15,8")
        SQL = "SELECT upper(SMIS_PURCHAGREE.MODEL)MODEL, CONVERT(VARCHAR, SMIS_PURCHAGREE.InvoicedDate,101),"
        SQL = SQL & " SMIS_PURCHAGREE.CustName,"
        SQL = SQL & " SMIS_PURCHAGREE.modeldescription,"
        SQL = SQL & " SMIS_PURCHAGREE.ignkey_no,"
        SQL = SQL & " SMIS_PURCHAGREE.color ,"
        SQL = SQL & " SMIS_PURCHAGREE.salesae ,"
        SQL = SQL & " SMIS_PURCHAGREE.TERM"
        SQL = SQL & " FROM SMIS_PURCHAGREE "
        SQL = SQL & " "
        SQL = SQL & " WHERE " & xDateFilter & "  SMIS_PURCHAGREE.VI_NO is Not Null  order by 1 "

    ElseIf xCondition = "O" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL, CSNO,DESCRIPTION, COLOR, AGE(REC), AGE(P/O)"
        Call ResizeColumnHeader(lstVehicles, "0,10,35,25,10,10")
        SQL = "SELECT  upper(SMIS_MrrInv.MODEL)MODEL, " _
            & " ignkey as CSNO,  " _
            & " DESCRIPT + ' ' + isnull(YEER ,'') as Descriptions," _
            & " Color, " _
            & " Datediff(day,DateReceived,getdate()) as Aging1, " _
            & " Datediff(day,PullOutDate,getdate()) as Aging2, " _
            & " ID FROM SMIS_MRRINV  WHERE RELEASED= 0 and status='P' order by 2 "

    ElseIf xCondition = "R" Then
        ReportControlAddColumnHeader lstVehicles, "MODEL,DATE REL, CUSTOMER NAME ,DESCRIPTION, CSNO, COLOR, TERM"
        Call ResizeColumnHeader(lstVehicles, "0,8,30,30,10,10,8")
        SQL = "SELECT upper(SMIS_PURCHAGREE.MODEL)MODEL,  CONVERT(VARCHAR, SMIS_PURCHAGREE.datereleased,101),"
        SQL = SQL & " SMIS_PURCHAGREE.CustName,"
        SQL = SQL & " SMIS_PURCHAGREE.modeldescription,"
        SQL = SQL & " SMIS_PURCHAGREE.ignkey_no +'/'+ SMIS_PURCHAGREE.vi_no  ,"
        SQL = SQL & " SMIS_PURCHAGREE.color ,"
        SQL = SQL & " SMIS_PURCHAGREE.TERM"
        SQL = SQL & " FROM SMIS_PURCHAGREE "
        SQL = SQL & " "
        SQL = SQL & " WHERE  " & xDateFilter & "   SMIS_PURCHAGREE.STATUS<>'C' order by 1 "
    End If
    lstVehicles.GroupsOrder.Add lstVehicles.Columns(0): lstVehicles.Columns(0).Visible = False
    Set RsUploadData = gconDMIS.Execute(SQL)
    flex_FillReportView RsUploadData, lstVehicles


End Sub

Sub ShowStatus(PROSPECTID)
    Dim temprs                                                        As ADODB.Recordset
    Dim xstatus                                                       As String

    Set temprs = gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)

    EXIST_INVOICE = False: EXIST_LOAN = False: Exist_SO = False: lblLogQuote = "": lblLogEmail = "": lblLogAppointment = "": lblLogTestDrive = "": lblLogCalls = ""
    lblLogLetters = "": lblLetters = "": lblLogLoan = "": lblLogVisits = "": lblLogSalesOrder = "": lblSTATUS = ""
    txtCusAdd = "": txtCusEmail = "": txtCusName = "": txtCusPhone = "": lblNotes = "": captionInformation.Caption = ""
    CustomerCode = "": captionInformation.Caption = ""
    If Not (temprs.EOF Or temprs.BOF) Then
        lblLogQuote = "x": lblLogQuote.ForeColor = &HC0&
        lblLogEmail = "x": lblLogEmail.ForeColor = &HC0&
        lblLogAppointment = "x": lblLogAppointment.ForeColor = &HC0&
        lblLogTestDrive = "x": lblLogTestDrive.ForeColor = &HC0&
        lblLogCalls = "x": lblLogCalls.ForeColor = &HC0&
        lblLogLetters = "x": lblLogLetters.ForeColor = &HC0&
        lblLogLoan = "x": lblLogLoan.ForeColor = &HC0&
        lblSTATUS = "x": lblSTATUS.BackColor = vbWhite
        lblLogVisits = "x": lblLogVisits.ForeColor = &HC0&

        lblInvoiceNo = "x": lblInvoiceNo.ForeColor = &HC0&
        lblLogSalesOrder.Caption = "x": lblLogSalesOrder.ForeColor = &HC0&
        lblNotes.Text = ""
        ProspType = Null2String(temprs!ProspectType)
        lblNotes.Text = Null2String(temprs!Notes)

        If IsNull(temprs!LogQuote) = False Then
            lblLogQuote = Chr(187) & FormatDateTime(temprs!LogQuote, vbShortDate): lblLogQuote.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogEmail) = False Then
            lblLogEmail = Chr(187) & FormatDateTime(temprs!LogEmail, vbShortDate): lblLogEmail.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogAppointment) = False Then
            lblLogAppointment = Chr(187) & FormatDateTime(temprs!LogAppointment, vbShortDate): lblLogAppointment.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogTestDrive) = False Then
            lblLogTestDrive = Chr(187) & FormatDateTime(temprs!LogTestDrive, vbShortDate) & "(" & DateDiff("d", temprs!LogTestDrive, LOGDATE) & ")": lblLogTestDrive.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogCall) = False Then
            lblLogCalls = Chr(187) & FormatDateTime(temprs!LogCall, vbShortDate): lblLogCalls.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogLetter) = False Then
            lblLogLetters = Chr(187) & FormatDateTime(temprs!LogLetter, vbShortDate): lblLogLetters.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogApplication) = False Then
            EXIST_LOAN = True
            lblLogLoan = Chr(187) & FormatDateTime(temprs!LogApplication, vbShortDate): lblLogLoan.ForeColor = &H8000&
        End If

        If IsNull(temprs!LogVisit) = False Then
            lblLogVisits = Chr(187) & FormatDateTime(temprs!LogVisit, vbShortDate): lblLogVisits.ForeColor = &H8000&
        End If

        If IsNull(temprs!INVOICENO) = False Then
            EXIST_INVOICE = True
            lblInvoiceNo = Chr(187) & temprs!INVOICENO: lblInvoiceNo.ForeColor = &H8000&
        End If



        If IsNull(temprs!LOGSO) = False Then
            Exist_SO = True
            lblLogSalesOrder = Chr(187) & FormatDateTime(temprs!LOGSO, vbShortDate) & "(" & Null2String(temprs!SO_NO) & ")": lblLogSalesOrder.ForeColor = &H8000&
            If IsDate(temprs!LOGCLOSINGDATE) = True Then
                lblAgeing = "Closed In: " & DateDiff("d", temprs!loginitialinquiry, temprs!LOGCLOSINGDATE) & " Days"
            End If

        Else
            lblAgeing = "Prospect Age: " & DateDiff("d", temprs!loginitialinquiry, LOGDATE) & " Days"
        End If
        xstatus = Null2String(temprs!STATUS)

        If xstatus = "O" Then
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        ElseIf xstatus = "C" And EXIST_INVOICE Then
            lblSTATUS = "CLOSED": lblSTATUS.BackColor = &H40C0&
        ElseIf xstatus = "C" And EXIST_INVOICE = False Then
            lblSTATUS = "INVOICING": lblSTATUS.BackColor = &H40C0&
        ElseIf xstatus = "I" Then
            lblSTATUS = "INACTIVE": lblSTATUS.BackColor = &HC0C0C0
        ElseIf xstatus = "L" Then
            lblSTATUS = "LOST SALES": lblSTATUS.BackColor = &HC0C0C0
        Else
            lblSTATUS = "OPEN": lblSTATUS.BackColor = &HC000&
        End If
        If IsNull(temprs!CUSCDE) = False Then
            CustomerCode = temprs!CUSCDE    ': ShowCustomerInfo CustomerCode
        End If
        If IsNull(temprs!EMAIL) = False Then
            txtCusEmail = temprs!EMAIL
        End If

        If IsNull(temprs!Telephone) = False Then
            txtCusPhone = temprs!Telephone
        End If

        If IsNull(temprs!Address) = False Then
            txtCusAdd = temprs!Address
        End If
    End If
    Set temprs = Nothing
End Sub

Private Sub cboPriority_Click()
    ShowData
End Sub

Private Sub cboSAE_Click()
    ShowData
End Sub

Private Sub cboStatus_Click()
    ShowData
End Sub

Private Sub cboYear_Click()
    ShowData
End Sub

Private Sub cmddata_Click()
    If Module_Access(LOGID, "PROSPECT DATA REPORT", "INQUIRY") = False Then Exit Sub
    frmFile_Prospectdata.Show 1
End Sub

Private Sub cmdPrintSO_Click()

    On Error GoTo Errorcode:

    If lstSalesOrder.Records.Count = 0 Then
        MsgSpeechBox "No Records To print"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Dim MANTH, YEER, REPORTSQL
    MANTH = Month(DateSerial(cboYear, (lstMonth.ListIndex + 1), 1))
    YEER = cboYear

    rptMain.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMain.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptMain.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"

    If optSO(0).Value = True Then
        REPORTSQL = " {SO.STATUS}='ON ORDER'"
        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If
    If optSO(1).Value = True Then
        REPORTSQL = " {SO.STATUS}='CANCELLED INVOICE' OR {SO.STATUS}='CANCELLED SO'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If
    If optSO(2).Value = True Then
        REPORTSQL = " {SO.STATUS}='INVOICED'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If
    If optSO(3).Value = True Then
        REPORTSQL = " {SO.STATUS}='RELEASED'"

        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " AND MONTH({SO.DEYT})=" & MANTH
        End If
        REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
    End If
    If optSO(4).Value = True Then
        If lstMonth.ListIndex <> 12 Then
            REPORTSQL = REPORTSQL & " MONTH({SO.DEYT})=" & MANTH
            REPORTSQL = REPORTSQL & " AND YEAR({SO.DEYT})=" & YEER
        Else
            REPORTSQL = REPORTSQL & " YEAR({SO.DEYT})=" & YEER
        End If
    End If
    PrintSQLReport rptMain, SMIS_REPORT_PATH & "SOListing.rpt", REPORTSQL, DMIS_REPORT_Connection, 1
    Screen.MousePointer = vbDefault
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrintVehicles_Click()
    On Error GoTo Errorcode:

    If lstVehicles.Records.Count = 0 Then
        MsgSpeechBox "No Records To print"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lstVehicles.PrintOptions.Header.Font.Size = 11
    lstVehicles.PrintOptions.Header.Font.Bold = True
    lstVehicles.PrintOptions.MarginLeft = 1.75
    lstVehicles.PrintOptions.MarginRight = 1.75

    If optInventory(0).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "ALLOCATED VEHICLES")
    ElseIf optInventory(1).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "INVOICED VEHICLES")
    ElseIf optInventory(2).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "RELEASED VEHICLES")
    ElseIf optInventory(3).Value = True Then
        lstVehicles.PrintOptions.Header.TextCenter = UCase(COMPANY_NAME & vbCrLf & COMPANY_ADDRESS & vbCrLf & "ON STOCK VEHICLES")
    End If
    lstVehicles.PrintPreview True
    '    If lQL = "D" Then
    '        PrintSQLReport rptMain, SMIS_REPORT_PATH & "VehilcesDemo.rpt", "", DMIS_REPORT_Connection, 1
    '    ElseIf lQL = "S" Then
    '        PrintSQLReport rptMain, SMIS_REPORT_PATH & "VehilcesSold.rpt", "", DMIS_REPORT_Connection, 1
    '    ElseIf lQL = "O" Then
    '        PrintSQLReport rptMain, SMIS_REPORT_PATH & "VehilcesOpen.rpt", "{SMIS_MRRINV.ISTATUS}='O'", DMIS_REPORT_Connection, 1
    '    ElseIf lQL = "T" Then
    '        PrintSQLReport rptMain, SMIS_REPORT_PATH & "VehilcesOpen.rpt", "{SMIS_MRRINV.ISTATUS}='T'", DMIS_REPORT_Connection, 1
    '    ElseIf lQL = "A" Then
    '        PrintSQLReport rptMain, SMIS_REPORT_PATH & "VSAListing-Allocated.rpt", "{SMIS_MRRINV.ISTATUS}='A'", DMIS_REPORT_Connection, 1
    '    Else
    '
    '    End If
    '    rptMain.PageZoom 75
    Screen.MousePointer = vbNormal
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSearchCustomer_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then: Exit Sub

    frmAllCustomer.Show
End Sub

Private Sub Combo1_Click()
    FILLLTOSTATUS
End Sub

Private Sub Combo2_Click()
    ShowData
End Sub

Private Sub Combo3_Click()
    ShowData
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "PROSPECT", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Prospects.Show
    If FormExist("frmSMIS_Files_Prospects") Then
        frmSMIS_Files_Prospects.ZOrder 0
    End If
End Sub

Private Sub Command10_Click()
    '    If Module_Access(LOGID, "PROSPECT STATUS", "PROCESS") = False Then Exit Sub
    'frmSMIS_Process_ProspectStatus.Show
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "SALES ORDER", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next

    frmSMIS_Trans_SalesOrder.Show


End Sub

Private Sub Command3_Click()
    With ReportControl1
        '.PrintOptions.BlackWhiteContrast = 0
        '.PrintOptions.BlackWhitePrinting = True
        .PrintOptions.Header.Font.Size = "12"
        .PrintOptions.Header.TextCenter = "LTO STATUS" & vbCrLf & COMPANY_NAME & vbCrLf & COMPANY_ADDRESS
        .PrintPreview True
    End With

End Sub

Private Sub Command4_Click()
    On Error Resume Next
    frmSMIS_Trans_ApplicationIndividual.Show
    If FormExist("frmSMIS_Trans_ApplicationIndividual") Then
        frmSMIS_Trans_ApplicationIndividual.ZOrder 0
    End If
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    frmSMIS_Trans_ApplicationCorporate.Show
    If FormExist("frmSMIS_Trans_ApplicationCorporate") Then
        frmSMIS_Trans_ApplicationCorporate.ZOrder 0
    End If
End Sub

Private Sub Command6_Click()
    frmSMIS_Inquiry_InquiryMain.optAdvSearch(0).Value = True
    frmSMIS_Inquiry_InquiryMain.Show
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "SALES CALCULATOR", "SYSTEM") = False Then Exit Sub
    frmSMIS_Mis_AOR.Show
End Sub

Private Sub Command8_Click()
    frmSMIS_Log_InternalReminder.Show
    frmSMIS_Log_InternalReminder.cmdAdd.Value = True
End Sub

Private Sub Command9_Click()
    If Module_Access(LOGID, "PROSPECT MASTER FILE", "REPORTS") = False Then Exit Sub
    rptMain.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMain.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptMain.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
    rptMain.Formulas(3) = "DRange = '" & " FOR THE  " & lstMonth.Text & cboYear & "'"
    rptMain.Formulas(4) = "Statusx = '" & RPTCAPTION & "'"
    PrintSQLReport rptMain, SMIS_REPORT_PATH & "Listing/ProspectListing.rpt", "(" & ReportFilter & ")", DMIS_Connection, 1

    NEW_LogAudit "V", "PROSPECT MASTER FILE", "", "", "", ReportFilter, "", ""
End Sub

Private Sub CustomerInformation_ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)
    Dim temprs                                                        As ADODB.Recordset
    On Error GoTo Errorcode:
    ShowStatus PROSPECTID
    gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
    Call frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID))
    Unload CustomerInformation
    On Error Resume Next
    frmSMIS_Trans_SalesOrder.Show
    If FormExist("frmSMIS_Trans_SalesOrder") Then
        On Error Resume Next
        frmSMIS_Trans_SalesOrder.SetFocus
        frmSMIS_Trans_SalesOrder.ZOrder 0
    End If
    Exit Sub
    If xGoingWhere = "Order" Then
        ShowStatus PROSPECTID
        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
        Call frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID))
        Unload CustomerInformation
        frmSMIS_Trans_SalesOrder.Show
    End If
    If xGoingWhere = "Prospects" Then
        ShowStatus PROSPECTID
        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
        Call Unload(CustomerInformation)
        Set CustomerInformation = Nothing
    End If
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillAndConfigGrid
    optProspects(0).Value = True
    optProspects_Click 0
    LOGSAE = ""
    MainForm.Enabled = True
    cboYear.Text = Format(LOGDATE, "yyyy")

End Sub

Private Sub lstIndividual_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    If UCase(Row.Record(5).Value) = "CORPORATE" Then
        frmSMIS_Trans_ApplicationCorporate.Show
        frmSMIS_Trans_ApplicationCorporate.SearchID (Row.Record(9).Value)
    Else

        frmSMIS_Trans_ApplicationIndividual.Show
        frmSMIS_Trans_ApplicationIndividual.SearchID (Row.Record(9).Value)

    End If

End Sub

Private Sub lstIndividual_SelectionChanged()
    If lstIndividual.SelectedRows(0).GroupRow = True Then Exit Sub
    If lstIndividual.SelectedRows.Count <= 0 Then: LABLOANNOTES = "": Exit Sub
    On Error GoTo ADDER:
    Dim temprs                                                        As ADODB.Recordset
    If UCase(lstIndividual.SelectedRows(0).Record(5).Value) = "CORPORATE" Then
        Set temprs = gconDMIS.Execute("select notes from smis_loancorp where id= " & (lstIndividual.SelectedRows(0).Record(9).Value))
    Else
        Set temprs = gconDMIS.Execute("select notes from smis_loanindiv where id= " & (lstIndividual.SelectedRows(0).Record(9).Value))
    End If
    If Not temprs.EOF Or Not temprs.BOF Then
        LABLOANNOTES = " " & Chr(187) & " " & Null2String(temprs!Notes)
    End If
    Exit Sub
ADDER:     Err.Clear
End Sub

Private Sub lstMonth_Click()
    ShowData
End Sub

Private Sub lstProspects_GotFocus()
    lstProspects_SelectionChanged
End Sub

''prospect grid
Private Sub lstProspects_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

    If Row.Record Is Nothing Then: Exit Sub
    'Set FormProspects = New frmSMIS_Files_Prospects
    frmSMIS_Files_Prospects.EditProspect (Row.Record(8).Value)
    frmSMIS_Files_Prospects.Show
End Sub

Private Sub lstProspects_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    lstProspects_SelectionChanged
            
    If Exist_SO = True Then
        mnuSalesOrder.Caption = "Open Sales Order"
    Else
        mnuSalesOrder.Caption = "Add New Sales Order"
    End If
    
    If EXIST_LOAN = True Then
        mnuLoanApplication.Caption = "Open Loan Application"
    Else
        mnuLoanApplication.Caption = "Add Loan Application"
    End If
    
    PopupMenu mnuContextProspect
End Sub

Private Sub lstProspects_SelectionChanged()
    On Error Resume Next
    If lstProspects.Records.Count = 0 Then
        ShowStatus 0
        Exit Sub
    End If
    ProspID = lstProspects.SelectedRows(0).Record(8).Value
    ShowStatus ProspID
End Sub

Private Sub lstReminders_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.GroupRow = True Then Exit Sub
    frmSMIS_Log_ReminderStatus.REMINDERID = Row.Record(9).Value
    frmSMIS_Log_ReminderStatus.Show
End Sub

Private Sub lstReminders_SelectionChanged()
    If lstReminders.SelectedRows(0).GroupRow = True Then Text5.Text = "": Exit Sub
    Text5.Text = lstReminders.SelectedRows(0).Record(10).Value & ""
End Sub

Private Sub lstSalesOrder_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then Exit Sub
    Load frmSMIS_Trans_SalesOrder
    Call frmSMIS_Trans_SalesOrder.SearchID(Row.Record(7).Value)
End Sub

Private Sub lvQuotation_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Load frmSMIS_Trans_Quotation
    Call frmSMIS_Trans_Quotation.SearchID(Row.Record(5).Value)
    Call frmSMIS_Trans_Quotation.Show
End Sub

Private Sub mnuCalls_Click()
    Call frmSMIS_Log_Call.AddCall(ProspID, vbNullString)
    frmSMIS_Log_Call.Show
    frmSMIS_Log_Call.ZOrder 0
End Sub

Private Sub mnuEmail_Click()
    Call frmSMIS_Log_Email.AddEmail(ProspID, vbNullString)
    frmSMIS_Log_Email.Show
    frmSMIS_Log_Email.ZOrder 0
End Sub

Private Sub mnuLetter_Click()
    Call frmSMIS_Log_Letter.AddLetter(ProspID, vbNullString)
    frmSMIS_Log_Letter.Show
    frmSMIS_Log_Letter.ZOrder 0
End Sub

Private Sub mnuLoanApplication_Click()


    Dim temprs                                                        As ADODB.Recordset
    Dim MyDecisionIs
    Dim rsExists                                                      As ADODB.Recordset

    On Error Resume Next

    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID= " & ProspID)

    If ProspType = "P" Then

        If EXIST_LOAN = False Then
LINDV:             frmSMIS_Trans_ApplicationIndividual.AddFromProspects ProspID
            frmSMIS_Trans_ApplicationIndividual.Show
            frmSMIS_Trans_ApplicationIndividual.ZOrder 0
        Else
            Set rsExists = gconDMIS.Execute("select ProspectID from SMIS_LoanINDIV where ProspectID=" & ProspID)
            If Not rsExists.EOF Or Not rsExists.BOF Then
                frmSMIS_Trans_ApplicationIndividual.Show
                frmSMIS_Trans_ApplicationIndividual.SearchID rsExists("ID").Value
                frmSMIS_Trans_ApplicationIndividual.ZOrder 0
            Else
                GoTo LINDV:
            End If
        End If

    Else
        If EXIST_LOAN = False Then
CORPDIV:             frmSMIS_Trans_ApplicationCorporate.AddFromProspects ProspID
            frmSMIS_Trans_ApplicationCorporate.Show
            frmSMIS_Trans_ApplicationCorporate.ZOrder 0
        Else
            Set rsExists = gconDMIS.Execute("select ProspectID from SMIS_LoanINDIV where ProspectID=" & ProspID)
            If Not rsExists.EOF Or Not rsExists.BOF Then
                frmSMIS_Trans_ApplicationCorporate.Show
                frmSMIS_Trans_ApplicationCorporate.SearchID rsExists("ID").Value
                frmSMIS_Trans_ApplicationCorporate.ZOrder 0
            Else
                GoTo CORPDIV:
            End If
        End If

    End If

End Sub

Private Sub mnuSalesAppointment_Click()
    frmSMIS_Log_SalesAppointment.AddSalesAppointment (ProspID)
    frmSMIS_Log_SalesAppointment.Show
    frmSMIS_Log_SalesAppointment.ZOrder 0
End Sub

Private Sub mnuSalesOrder_Click()
    Dim temprs                                                        As ADODB.Recordset
    Dim MyDecisionIs
    Dim rsExists                                                      As ADODB.Recordset

    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID= " & ProspID)
    If Null2String(temprs!CUSCDE) = "" Then
        If MsgBox("Current Prospect Has Not been Converted to Customer !" & vbCrLf & "Do you Like to Convert It to Customer?" _
         , vbYesNo Or vbExclamation Or vbDefaultButton1, App.TITLE) = vbYes Then
            Set CustomerInformation = New frmAllCustomer
            Load CustomerInformation
            Call CustomerInformation.AddCustomerFromProspect(temprs, "Order")
            CustomerInformation.Show 1
            Set CustomerInformation = Nothing
        End If
    Else
        '
        Set rsExists = gconDMIS.Execute("select ID from SMIS_SALESORDER WHERE sostatus<>'C' and CODE='" & CustomerCode & "'")
        If Not (rsExists.EOF Or rsExists.BOF) And Exist_SO = True Then
            MyDecisionIs = MsgBox(" There is Already Sales Order For this Prospect. Would You Like To Preview It?" & vbCrLf & _
                                " Click Yes To Open Existing Sales Order" & vbCrLf & _
                                " Click No To Add New Sales Order, Cancel To Abort", vbInformation + vbYesNoCancel)
            If MyDecisionIs = vbYes Then

                frmSMIS_Trans_SalesOrder.Show
                frmSMIS_Trans_SalesOrder.SearchID rsExists!ID
                frmSMIS_Trans_SalesOrder.ZOrder 0
            ElseIf MyDecisionIs = vbNo Then

                If frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(temprs) = True Then: frmSMIS_Trans_SalesOrder.Show: frmSMIS_Trans_SalesOrder.ZOrder 0
            Else
                Exit Sub
            End If
        Else
            If frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(temprs) = True Then: frmSMIS_Trans_SalesOrder.Show: frmSMIS_Trans_SalesOrder.ZOrder 0
        End If
    End If
End Sub

Private Sub mnuSendQuotation_Click()
    On Error Resume Next
    Call frmSMIS_Trans_Quotation.AddNewQuotation(ProspID)
    frmSMIS_Trans_Quotation.Show
    frmSMIS_Trans_Quotation.ZOrder 0
    frmSMIS_Trans_Quotation.cboModel_Change
End Sub

Private Sub mnuShowGroup_Click()
    '
    
            mnuShowGroup.Checked = Not mnuShowGroup.Checked
            If TabControl.Selected.Index = 0 And mnuShowGroup.Checked = True Then
                lstProspects.ShowGroupBox = True
            Else
                lstProspects.ShowGroupBox = False
            End If
        
    
    '
End Sub

Private Sub mnuTestDrive_Click()
    On Error Resume Next

    frmSMIS_Log_TestDriveAppointment.AddTestDriveAppointment (ProspID)
    frmSMIS_Log_TestDriveAppointment.Show
    frmSMIS_Log_TestDriveAppointment.ZOrder 0
End Sub

Private Sub mnuUpdateProspectStatus_Click()
    frmSMIS_Files_ProspectStatus.PROSPECTID = lstProspects.SelectedRows(0).Record(8).Value
    frmSMIS_Files_ProspectStatus.Show
End Sub

Private Sub mnuUpdateStaus_Click()
    'frmSMIS_Files_ProspectStatus.ProspectID = lstProspects.SelectedRows(0).Record(8).Value
    'frmSMIS_Files_ProspectStatus.Show
End Sub

Private Sub mnuViewLog_Click()
    frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG lstProspects.SelectedRows.Row(0).Record(8).Value, lstProspects.SelectedRows.Row(0).Record(1).Value
End Sub

Private Sub mnuVisits_Click()
    Call frmSMIS_Log_Visits.AddVisit(ProspID, vbNullString)
    frmSMIS_Log_Visits.Show
    frmSMIS_Log_Visits.ZOrder 0
End Sub

Private Sub optInventory_Click(Index As Integer)
    ShowData
End Sub

Private Sub optLoan_Click(Index As Integer)
    ShowData
End Sub

Private Sub optProspects_Click(Index As Integer)
    ShowData
End Sub

Private Sub optQuote_Click(Index As Integer)
    ShowData
End Sub

Private Sub optReminder_Click(Index As Integer)
    ShowData
End Sub

Private Sub optSO_Click(Index As Integer)
    ShowData
End Sub

Private Sub ReportControl1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error Resume Next
    frmSMIS_LTOStatus.UpdateStataus Row.Record(2).Value
    frmSMIS_LTOStatus.Show
End Sub

Private Sub SalesOrder_AE()
    ShowData
End Sub

Private Sub SalesOrder_Deleted()
    ShowData
End Sub

Private Sub ShowCustomerInfo(xxxcode)
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(xxxcode))
    If Not (temprs.EOF Or temprs.BOF) Then
        If temprs.Fields("CUSTYPE") = "P" Then
            txtCusAdd = Replace(Null2String(temprs("CUSTOMERADD")), Chr(10), "")
            txtCusEmail = Null2String(temprs("EMAIL"))
            txtCusName = Null2String(temprs("LASTNAME")) & IIf(IsNull(temprs("LASTNAME")), "", ",") & Null2String(temprs("FirstName")) & IIf(IsNull(temprs("MIDDLEINITIAL")), "", ".") & Null2String(temprs("MIDDLEINITIAL"))
            txtCusPhone = Null2String(temprs("HOMEPHONE"))
            captionInformation.Caption = Null2String(temprs("ACCTNAME"))
        Else
            txtCusAdd = Replace(Null2String(temprs("COMPANYADD")), Chr(10), "")
            txtCusEmail = Null2String(temprs("EMAIL"))
            txtCusName = Null2String(temprs("CUSCOMP"))
            txtCusPhone = Null2String(temprs("TELEPHONENO"))
            captionInformation.Caption = Null2String(temprs("ACCTNAME"))
        End If
    End If
    Set temprs = Nothing
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    DoEvents

    On Error Resume Next
    cboYear.Enabled = True
    lstMonth.Enabled = True
    If Item.Index = 0 Then
        txtSearch_Prospect.SetFocus
    ElseIf Item.Index = 1 Then
        txtSearch_Quotation.SetFocus
    ElseIf Item.Index = 2 Then
        txtSearch_SalesOrder.SetFocus
    ElseIf Item.Index = 3 Then
        txtSearch_LoanApplication.SetFocus
    ElseIf Item.Index = 4 Then
        'txtSearch_Vehicles.SetFocus
    ElseIf Item.Index = 5 Then

        txtSearch_LTO.SetFocus
        cboYear.Enabled = False
        lstMonth.Enabled = False
    ElseIf Item.Index = 6 Then
        txtSearch_Activity.SetFocus
    End If
    ShowData
End Sub

Private Sub Text5_Change()
    If Text5 = "" Then: Label10.Visible = False: Else Label10.Visible = True
End Sub

Private Sub txtSearch_Activity_Change()
    lstReminders.FilterText = txtSearch_Activity.Text
    lstReminders.Populate
End Sub

Private Sub txtSearch_LoanApplication_Change()
    lstIndividual.FilterText = txtSearch_LoanApplication
    lstIndividual.Populate
End Sub

Private Sub txtSearch_LTO_Change()
    ReportControl1.FilterText = txtSearch_LTO.Text
    ReportControl1.Populate
End Sub

Private Sub txtSearch_Prospect_Change()
    lstProspects.FilterText = txtSearch_Prospect.Text
    lstProspects.Populate
    ShowStatus 0
End Sub

Private Sub txtSearch_SalesOrder_Change()
    lstSalesOrder.FilterText = txtSearch_SalesOrder
    lstSalesOrder.Populate
End Sub

Private Sub txtSearch_Vehicles_Change()
    lstVehicles.FilterText = txtSearch_Vehicles.Text
    lstVehicles.Populate

End Sub

Public Sub ShowData()
    Dim Indx                                                          As Integer
    Dim Priority
    Dim xstatus
    Dim XSAE
    Dim LABDESC                                                       As String
    Select Case lstMonth.Text
        Case "January": Indx = 1
        Case "February": Indx = 2
        Case "March": Indx = 3
        Case "April": Indx = 4
        Case "May": Indx = 5
        Case "June": Indx = 6
        Case "July": Indx = 7
        Case "August": Indx = 8
        Case "September": Indx = 9
        Case "October": Indx = 10
        Case "November": Indx = 11
        Case "December": Indx = 12
        Case Else: Indx = -1
    End Select
    ReportFilter = ""
    Select Case TabControl.SelectedItem
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''PROSPECT''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 0

            If Left(Combo2, 1) <> "" And Left(Combo2, 1) <> "(" Then
                XSAE = " AND USERCODE='" & GetSAECode(Combo2) & "'"
                ReportFilter = " {p.USERCODE}='" & GetSAECode(Combo2) & "' AND "
            End If
            mnuCalls.Enabled = True
            mnuEmail.Enabled = True
            mnuLetter.Enabled = True
            mnuLoanApplication.Enabled = True
            mnuSalesAppointment.Enabled = True
            mnuSalesOrder.Enabled = True
            mnuSendQuotation.Enabled = True
            mnuTestDrive.Enabled = True
            mnuVisits.Enabled = True

            If optProspects(0).Value = True Then

                PSTATUS = " AND STATUS='O' "
                LABDESC = " SHOWING OPEN PROSPECT "
                RPTCAPTION = "OPEN"
                ReportFilter = ReportFilter & "{P.STATUS}='O' AND "

            ElseIf optProspects(1).Value = True Then

                PSTATUS = " AND STATUS='C' "
                ReportFilter = ReportFilter & "{P.STATUS}='C' AND "
                mnuLoanApplication.Enabled = False
                mnuSalesAppointment.Enabled = False
                mnuSalesOrder.Enabled = False
                mnuSendQuotation.Enabled = False
                mnuTestDrive.Enabled = False
                LABDESC = "SHOWING CLOSED PROSPECT "
                RPTCAPTION = "CLOSED"

            ElseIf optProspects(2).Value = True Then
                PSTATUS = " AND STATUS='I' "
                ReportFilter = ReportFilter & "{P.STATUS}='I' AND "
                mnuLoanApplication.Enabled = False
                mnuSalesAppointment.Enabled = False
                mnuSalesOrder.Enabled = False
                mnuSendQuotation.Enabled = False
                mnuTestDrive.Enabled = False
                LABDESC = "SHOWING INACTIVE PROSPECT "
                RPTCAPTION = "INACTIVE"
            ElseIf optProspects(3).Value = True Then
                PSTATUS = " AND LOGFOLLOWUPDATE IS NOT NULL AND LOGFOLLOWUPDATE >= CONVERT(varchar,GETDATE(),101) "
                LABDESC = "SHOWING PROSPECTS FOR FOLLOW UP "
                ReportFilter = ReportFilter & "ISNULL({P.LOGFOLLOWUPDATE})=FALSE AND {P.LOGFOLLOWUPDATE}>CURRENTDATE"
                RPTCAPTION = "FOLLOW UP"
            ElseIf optProspects(5).Value = True Then
                PSTATUS = " AND STATUS='L' "
                ReportFilter = ReportFilter & "{P.STATUS}='L' AND "
                mnuLoanApplication.Enabled = False
                mnuSalesAppointment.Enabled = False
                mnuSalesOrder.Enabled = False
                mnuSendQuotation.Enabled = False
                mnuTestDrive.Enabled = False
                LABDESC = "SHOWING LOST SALE PROSPECT "
                RPTCAPTION = "LOST SALE"
            Else
                PSTATUS = " "
                RPTCAPTION = ""
                LABDESC = " ALL PROSPECTS "

            End If

            If Indx = -1 Then
                FillProspect "( YEAR(LOGCLOSINGDATE)=" & cboYear.Text & " OR  YEAR(LOGINITIALINQUIRY)=" & cboYear.Text & ")" & XSAE
                labVCount = LABDESC & "(" & lstProspects.Records.Count & ")  For the " & cboYear
                ReportFilter = ReportFilter & " YEAR({P.LOGINITIALINQUIRY})=" & cboYear
            Else
                '''''*********
                FillProspect " (YEAR(LOGCLOSINGDATE)=" & cboYear.Text & "  AND MONTH(LOGCLOSINGDATE)  =" & Indx & " OR  YEAR(LOGINITIALINQUIRY)=" & cboYear.Text & "  AND MONTH(LOGINITIALINQUIRY)  =" & Indx & ")" & XSAE
                labVCount = LABDESC & "(" & lstProspects.Records.Count & ") For the  " & lstMonth & " " & cboYear
                ReportFilter = ReportFilter & " YEAR({P.LOGINITIALINQUIRY})=" & cboYear & " AND MONTH({P.LOGINITIALINQUIRY})= " & Indx
            End If

            If lstProspects.Records.Count > 0 Then
                lstProspects.Rows(0).Selected = True
                lstProspects.Rows(0).EnsureVisible
            End If
            lstProspects_SelectionChanged
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''QUOTATION''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 1
            If optQuote(0).Value = True Then
                QSTATUS = " AND (CQ.Status='P' OR CQ.Status IS NULL) "
                LABDESC = "SHOWING QUOTATION ON PROCESS"
            ElseIf optQuote(1).Value = True Then
                QSTATUS = " AND   CQ.STATUS='S' "
                LABDESC = "SHOWING QUOTATION ON SALES ORDER"
            ElseIf optQuote(2).Value = True Then
                QSTATUS = " AND   CQ.STATUS='I' "
                LABDESC = "SHOWING QUOTATION ON ON INVOICE"
            End If

            If Indx = -1 Then
                FillQuotation " YEAR(QuotationDate)=" & cboYear.Text
                labVCount = LABDESC & "(" & lvQuotation.Records.Count & ") For the " & cboYear
            Else
                FillQuotation " YEAR(QuotationDate)=" & cboYear.Text & "  AND MONTH(QuotationDate)  =" & Indx
                labVCount = LABDESC & "(" & lvQuotation.Records.Count & ") For the  " & lstMonth & " " & cboYear
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''SALES ORDER''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 2
            If Left(Combo3, 1) <> "" And Left(Combo3, 1) <> "(" Then
                XSAE = " AND USERCODE='" & GetSAECode(Combo3) & "'"
            End If
            If optSO(0).Value = True Then
                SSTATUS = " AND (Status='ON ORDER') "
                LABDESC = " SALES ORDER ON ORDER "
            ElseIf optSO(1).Value = True Then
                SSTATUS = "  AND  (STATUS='CANCELLED INVOICE' OR Status='CANCELLED SO')"
                LABDESC = " CANCELLED SALES ORDER/SALES INVOICE "
            ElseIf optSO(2).Value = True Then
                SSTATUS = " AND   STATUS='INVOICED' "
                LABDESC = " INVOICED SALES ORDER "
            ElseIf optSO(3).Value = True Then
                SSTATUS = " AND   STATUS='RELEASED' "
                LABDESC = " RELEASED SALES ORDER "
            Else
                SSTATUS = vbNullString
                LABDESC = " "
            End If

            If Indx = -1 Then
                FillSalesOrder " YEAR(DEYT)=" & cboYear.Text & XSAE
                labVCount = LABDESC & "(" & lstSalesOrder.Records.Count & ") Sales Order For the  " & cboYear
            Else
                FillSalesOrder " YEAR(DEYT)=" & cboYear.Text & "  and MOnth(deyt)=" & (Indx) & XSAE
                labVCount = LABDESC & "(" & lstSalesOrder.Records.Count & ") Sales Order For the " & lstMonth & " " & cboYear
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''LOAN APPLICATION'''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 3
            If optLoan(0).Value = True Then                  ' on process
                lStatus = " AND LSTATUS='O'"
                LABDESC = " LOAN APPLICATION - ON PROCESS"
                lstIndividual.Columns(6).Caption = " Last Updated"
            ElseIf optLoan(1).Value = True Then              ' pending
                lstIndividual.Columns(6).Caption = " Last Updated"
                lStatus = " AND LSTATUS='P'"
                LABDESC = " LOAN APPLICATION - PENDING"
            ElseIf optLoan(2).Value = True Then
                lstIndividual.Columns(6).Caption = "Date Disapproved"
                lStatus = " AND LSTATUS='D'"
                LABDESC = " LOAN APPLICATION - DISAPPROVED"
            ElseIf optLoan(4).Value = True Then
                lstIndividual.Columns(6).Caption = "Date Cancelled"
                lStatus = " AND LSTATUS='C'"
                LABDESC = " LOAN APPLICATION - CANCELLED"
            Else
                lstIndividual.Columns(6).Caption = "Date Approved"
                lStatus = " AND LSTATUS='A'"
                LABDESC = " LOAN APPLICATION - APPROVED"
            End If

            If Indx = -1 Then
                FillLoanApplication " YEAR(DateApplied)=" & cboYear.Text & lStatus
                labVCount = LABDESC & "(" & lstIndividual.Records.Count & ")  For the " & cboYear
            Else
                FillLoanApplication " YEAR(DateApplied)=" & cboYear.Text & " and MOnth(DateApplied)=" & (Indx) & lStatus
                labVCount = LABDESC & "(" & lstIndividual.Records.Count & ") For the " & lstMonth & " " & cboYear
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''' VEHICLE INQUIRY''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 4
            If optInventory(0).Value = True Then
                VSTATUS = "A"
                LABDESC = " ALLOCATED VEHICLES "
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_SALESORDER.Deyt)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_SALESORDER.Deyt)=" & cboYear & "  and MONTH(SMIS_SALESORDER.Deyt)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If

            ElseIf optInventory(1).Value = True Then
                VSTATUS = "S"
                LABDESC = " VEHICLES INVOICED"
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_PURCHAGREE.INVOICEDDATE)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") FOR  THE " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_PURCHAGREE.INVOICEDDATE)=" & cboYear & "  AND MONTH(SMIS_PURCHAGREE.INVOICEDDATE)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") FOR THE  " & lstMonth & " " & cboYear
                End If

            ElseIf optInventory(2).Value = True Then
                VSTATUS = "R"
                LABDESC = " TOTAL VEHICLES ON RELEASED "
                If Indx = -1 Then
                    FillVehicles VSTATUS, " YEAR(SMIS_PURCHAGREE.DateReleased)=" & cboYear & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, "YEAR(SMIS_PURCHAGREE.DateReleased)=" & cboYear & "  and MONTH(SMIS_PURCHAGREE.DateReleased)=" & Indx & " AND "
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If
            Else
                VSTATUS = "O"
                LABDESC = " TOTAL VEHICLES ON STOCK "
                If Indx = -1 Then
                    FillVehicles VSTATUS, ""
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For  the " & cboYear
                Else
                    FillVehicles VSTATUS, ""
                    labVCount = LABDESC & "(" & lstVehicles.Records.Count & ") For the  " & lstMonth & " " & cboYear
                End If
            End If


            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''' REMINDERS AND TASK'''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case 6
            Priority = ""
            xstatus = ""
            XSAE = ""
            If Left(cboPriority, 1) <> "" And Left(cboPriority, 1) <> "(" Then
                Priority = " AND Priority='" & Left(cboPriority, 1) & "'"
            End If
            If Left(cbostatus, 1) <> "" And Left(cbostatus, 1) <> "(" Then
                xstatus = " AND STATUS='" & Left(cbostatus, 1) & "'"
            End If

            If Left(cboSAE, 1) <> "" And Left(cboSAE, 1) <> "(" Then
                XSAE = " AND usercode='" & GetSAECode(cboSAE) & "'"
            End If

            If optReminder(0).Value = True Then
                RSTATUS = " AND ENTITYTYPE='P'" & Priority & xstatus & XSAE
            ElseIf optReminder(1).Value = True Then
                RSTATUS = " AND ENTITYTYPE='C'" & Priority & xstatus & XSAE
            ElseIf optReminder(3).Value = True Then
                RSTATUS = " AND ENTITYTYPE='S'" & Priority & xstatus & XSAE
            Else
                RSTATUS = "" & Priority & xstatus & XSAE

            End If

            If Indx = -1 Then
                FillReminder " and  YEAR(DATETIMEREMIND)>=" & cboYear
            Else
                FillReminder " and  YEAR(DATETIMEREMIND)=" & cboYear & "  and MONTH(DATETIMEREMIND)>=" & Indx
            End If
    End Select

End Sub

