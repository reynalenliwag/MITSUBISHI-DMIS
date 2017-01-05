VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_SalesOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Order"
   ClientHeight    =   9090
   ClientLeft      =   1125
   ClientTop       =   1200
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "SalesOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   13095
   Begin VB.PictureBox picsecurity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4005
      ScaleHeight     =   2145
      ScaleWidth      =   4455
      TabIndex        =   135
      Top             =   4020
      Visible         =   0   'False
      Width           =   4485
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3450
         MaskColor       =   &H00FFFFFF&
         Picture         =   "SalesOrder.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Cancel"
         Top             =   1350
         Width           =   885
      End
      Begin VB.TextBox txtUserPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1770
         PasswordChar    =   "*"
         TabIndex        =   141
         Top             =   900
         Width           =   2595
      End
      Begin VB.ComboBox CboAuthorized 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1770
         TabIndex        =   139
         Text            =   "Combo2"
         Top             =   540
         Width           =   2625
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   137
         Top             =   60
         Width           =   285
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2580
         MaskColor       =   &H00FFFFFF&
         Picture         =   "SalesOrder.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Log-In"
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   140
         Top             =   930
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Authorized Person:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   138
         Top             =   570
         Width           =   1635
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   -270
         TabIndex        =   136
         Top             =   -30
         Width           =   5175
         _Version        =   655364
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "::Transaction Security Access::"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
         GradientColorLight=   12632256
         GradientColorDark=   8421504
         ForeColor       =   0
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   0
      ScaleHeight     =   2550
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2520
         Left            =   11190
         ScaleHeight     =   2520
         ScaleWidth      =   2145
         TabIndex        =   150
         Top             =   0
         Width           =   2145
         Begin VB.TextBox txtInvoicedDate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   " "
            Top             =   2130
            Width           =   1755
         End
         Begin VB.TextBox txtTimeRelease 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   153
            Text            =   " "
            Top             =   1515
            Width           =   1755
         End
         Begin VB.TextBox txtPlaceRelease 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   152
            Text            =   " "
            Top             =   903
            Width           =   1755
         End
         Begin VB.TextBox txtDateRelease 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   " "
            Top             =   291
            Width           =   1755
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoiced Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   30
            TabIndex        =   158
            Top             =   1875
            Width           =   1755
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Release "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   60
            TabIndex        =   157
            Top             =   45
            Width           =   1755
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place of Release"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   30
            TabIndex        =   156
            Top             =   660
            Width           =   1755
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time of Release:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   30
            TabIndex        =   155
            Top             =   1275
            Width           =   1755
         End
      End
      Begin VB.Timer tmBlink 
         Interval        =   500
         Left            =   8730
         Top             =   540
      End
      Begin Crystal.CrystalReport rptReleased 
         Left            =   9510
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Units Released"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Frame fraCustInfo 
         BorderStyle     =   0  'None
         Caption         =   "Customer Information"
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
         Height          =   2700
         Left            =   60
         TabIndex        =   1
         Top             =   -120
         Width           =   11160
         Begin VB.TextBox txt_SONO 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   5850
            MaxLength       =   6
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   300
            Width           =   1275
         End
         Begin VB.CommandButton Command3 
            Caption         =   "::"
            Height          =   345
            Left            =   10740
            TabIndex        =   132
            Top             =   270
            Width           =   315
         End
         Begin VB.TextBox txtIssuedOn 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   8295
            TabIndex        =   30
            Text            =   " "
            Top             =   2295
            Width           =   2715
         End
         Begin VB.TextBox txtIssuedAt 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   4935
            TabIndex        =   29
            Text            =   " "
            Top             =   2295
            Width           =   3345
         End
         Begin VB.TextBox txtCTCNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1980
            TabIndex        =   28
            Text            =   " "
            Top             =   2295
            Width           =   2925
         End
         Begin VB.TextBox txtTIN 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   345
            TabIndex        =   27
            Text            =   " "
            Top             =   2295
            Width           =   1605
         End
         Begin VB.TextBox txtPosisyon 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   8295
            TabIndex        =   22
            Text            =   " "
            Top             =   1710
            Width           =   2715
         End
         Begin VB.TextBox txtPerson 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   4935
            TabIndex        =   21
            Text            =   " "
            Top             =   1710
            Width           =   3345
         End
         Begin VB.TextBox txtSpouse 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   1530
            TabIndex        =   20
            Text            =   " "
            Top             =   1710
            Width           =   3345
         End
         Begin VB.TextBox txtHomeTelNo 
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
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   8310
            TabIndex        =   10
            Text            =   " "
            Top             =   675
            Width           =   2715
         End
         Begin VB.TextBox txtCusName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   1560
            TabIndex        =   2
            Tag             =   "@R"
            ToolTipText     =   "Customer Name "
            Top             =   270
            Width           =   4215
         End
         Begin VB.TextBox txtHomeAdd 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1560
            TabIndex        =   8
            Top             =   690
            Width           =   5535
         End
         Begin VB.TextBox txtOfficeAdd 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   330
            Left            =   1560
            TabIndex        =   12
            Text            =   " "
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox txtOfficeTelNo 
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
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   8310
            TabIndex        =   14
            Text            =   " "
            Top             =   1065
            Width           =   2715
         End
         Begin VB.TextBox txtBirthDate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   360
            Left            =   345
            TabIndex        =   19
            Text            =   " "
            Top             =   1710
            Width           =   1125
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   10620
            Top             =   270
         End
         Begin MSComCtl2.DTPicker txtDeyt 
            Height          =   375
            Left            =   8310
            TabIndex        =   4
            Top             =   255
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   56492035
            CurrentDate     =   38941
         End
         Begin VB.TextBox txtSaveMe 
            Height          =   315
            Left            =   5880
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   270
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label LABALLOWREPRINT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   9420
            TabIndex        =   159
            Top             =   1290
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issued on"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   8325
            TabIndex        =   26
            Top             =   2085
            Width           =   825
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issued At"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   4980
            TabIndex        =   25
            Top             =   2085
            Width           =   795
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTC No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   1980
            TabIndex        =   24
            Top             =   2085
            Width           =   615
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   345
            TabIndex        =   23
            Top             =   2085
            Width           =   270
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   8310
            TabIndex        =   18
            Top             =   1485
            Width           =   690
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   4920
            TabIndex        =   17
            Top             =   1485
            Width           =   1320
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   1530
            TabIndex        =   16
            Top             =   1485
            Width           =   645
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   345
            TabIndex        =   15
            Top             =   1485
            Width           =   825
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   11
            Left            =   585
            TabIndex        =   3
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   7110
            TabIndex        =   6
            Top             =   330
            Width           =   1125
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Home Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   225
            TabIndex        =   7
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. No(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   1
            Left            =   7440
            TabIndex        =   9
            Top             =   743
            Width           =   795
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Office Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   2
            Left            =   210
            TabIndex        =   11
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. No(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   3
            Left            =   7440
            TabIndex        =   13
            Top             =   1133
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   2895
      ScaleHeight     =   4965
      ScaleWidth      =   6735
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   6765
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3420
         Left            =   150
         TabIndex        =   47
         Top             =   1080
         Width           =   6600
         _Version        =   655364
         _ExtentX        =   11642
         _ExtentY        =   6032
         _StockProps     =   64
         BorderStyle     =   4
         SkipGroupsFocus =   0   'False
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   885
         TabIndex        =   45
         Top             =   675
         Width           =   3915
      End
      Begin VB.OptionButton optList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Vehilce List"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2340
         TabIndex        =   44
         Top             =   360
         Width           =   1590
      End
      Begin VB.OptionButton optInventory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "From Vehicle Inventory"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   330
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6450
         TabIndex        =   42
         Top             =   15
         Width           =   285
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   5730
         TabIndex        =   41
         ToolTipText     =   "Cancel"
         Top             =   4500
         Width           =   825
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "Select "
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   48
         ToolTipText     =   "Select"
         Top             =   4500
         Width           =   825
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   -15
         TabIndex        =   40
         Top             =   0
         Width           =   6915
         _Version        =   655364
         _ExtentX        =   12197
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Preview Vehicles On Stock :::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   135
         TabIndex        =   46
         Top             =   750
         Width           =   2505
      End
   End
   Begin VB.PictureBox picMultipleInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   3180
      ScaleHeight     =   5055
      ScaleWidth      =   6075
      TabIndex        =   31
      Top             =   2565
      Visible         =   0   'False
      Width           =   6105
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   870
         TabIndex        =   34
         Text            =   "Combo1"
         Top             =   390
         Width           =   5085
      End
      Begin MSComctlLib.ListView lstMultipleInventory 
         Height          =   3585
         Left            =   60
         TabIndex        =   36
         Top             =   840
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6324
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CSNO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "MODEL DESCRIPTION"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COLOR"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdCancelMultiple 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   5160
         TabIndex        =   37
         ToolTipText     =   "Cancel"
         Top             =   4410
         Width           =   825
      End
      Begin VB.CommandButton cmdCloseMultiple 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5700
         TabIndex        =   33
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmdSelectMultiple 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4350
         TabIndex        =   38
         ToolTipText     =   "Select"
         Top             =   4410
         Width           =   825
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "FILTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   35
         Top             =   420
         Width           =   2505
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   -15
         TabIndex        =   32
         Top             =   0
         Width           =   6915
         _Version        =   655364
         _ExtentX        =   12197
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Select Vehicle Details"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picMiddles 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   13095
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2550
      Width           =   13095
      Begin VB.PictureBox picSalesOrder 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5685
         Left            =   0
         ScaleHeight     =   5685
         ScaleWidth      =   13275
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Width           =   13275
         Begin VB.Frame fraAdditionalInfo 
            Caption         =   "Additional Information"
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
            Height          =   5700
            Left            =   10560
            TabIndex        =   144
            Top             =   -30
            Width           =   2505
            Begin VB.TextBox txtRPPD 
               Alignment       =   1  'Right Justify
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
               Height          =   360
               Left            =   90
               TabIndex        =   147
               Text            =   " "
               Top             =   5190
               Width           =   2325
            End
            Begin VB.TextBox txtGMI 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00701E2A&
               Height          =   360
               Left            =   90
               TabIndex        =   146
               Text            =   " "
               Top             =   4575
               Width           =   2310
            End
            Begin VB.TextBox txtAdditionalInfo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   3510
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   145
               Top             =   240
               Width           =   2355
            End
            Begin VB.Label LABCUSTOMERCODE 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   465
               Left            =   90
               TabIndex        =   160
               Top             =   3840
               Width           =   2295
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GMI"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   11
               Left            =   90
               TabIndex        =   149
               ToolTipText     =   "Gross Monthly Interest"
               Top             =   4335
               Width           =   315
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RPPD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   10
               Left            =   90
               TabIndex        =   148
               ToolTipText     =   "NMI (Net Monthly Interest)"
               Top             =   4950
               Width           =   420
            End
         End
         Begin VB.ComboBox cboColor 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   1890
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "Color "
            Top             =   5220
            Width           =   2895
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Select Vehicles"
            Height          =   315
            Left            =   1920
            MouseIcon       =   "SalesOrder.frx":0EA3
            MousePointer    =   99  'Custom
            TabIndex        =   68
            ToolTipText     =   "Select Vehicle From Company Inventory Or Select From Model List"
            Top             =   2250
            Width           =   2835
         End
         Begin VB.Frame fraVehilcesInfo 
            Caption         =   "Vehicle Information"
            Enabled         =   0   'False
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
            Height          =   3570
            Left            =   30
            TabIndex        =   51
            Top             =   2070
            Width           =   4815
            Begin VB.TextBox txtModel 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1860
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Text            =   " "
               Top             =   510
               Width           =   2865
            End
            Begin VB.TextBox txtProdNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               TabIndex        =   60
               TabStop         =   0   'False
               Text            =   " "
               Top             =   1626
               Width           =   2865
            End
            Begin VB.TextBox txtConductionSticker 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               TabIndex        =   58
               TabStop         =   0   'False
               Text            =   " "
               Top             =   1254
               Width           =   2865
            End
            Begin VB.TextBox txtVinNumber 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               Locked          =   -1  'True
               TabIndex        =   66
               TabStop         =   0   'False
               Text            =   " "
               Top             =   2745
               Width           =   2865
            End
            Begin VB.TextBox txtEngineNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               Locked          =   -1  'True
               TabIndex        =   62
               TabStop         =   0   'False
               Text            =   " "
               Top             =   1998
               Width           =   2865
            End
            Begin VB.TextBox txtModelDescription 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Text            =   " "
               Top             =   882
               Width           =   2865
            End
            Begin VB.TextBox txtFrameNumber 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1845
               Locked          =   -1  'True
               TabIndex        =   64
               TabStop         =   0   'False
               Text            =   " "
               Top             =   2370
               Width           =   2865
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Model"
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
               Index           =   0
               Left            =   1245
               TabIndex        =   53
               ToolTipText     =   "xxx"
               Top             =   555
               Width           =   495
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Production No. "
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
               Left            =   480
               TabIndex        =   59
               Top             =   1605
               Width           =   1260
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CS Number"
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
               Left            =   765
               TabIndex        =   57
               Top             =   1260
               Width           =   975
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frame Number"
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
               Left            =   480
               TabIndex        =   63
               Top             =   2370
               Width           =   1260
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Engine No"
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
               Left            =   870
               TabIndex        =   61
               Top             =   1980
               Width           =   870
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
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
               Left            =   1290
               TabIndex        =   67
               Top             =   3120
               Width           =   450
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
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
               Left            =   795
               TabIndex        =   55
               Top             =   930
               Width           =   945
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vin Number"
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
               Index           =   0
               Left            =   765
               TabIndex        =   65
               Top             =   2745
               Width           =   975
            End
            Begin VB.Label lblVehicleStatus 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
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
               Height          =   360
               Left            =   105
               TabIndex        =   52
               Top             =   195
               Width           =   1860
            End
         End
         Begin VB.Frame fraInitialCashLayout 
            Caption         =   "Initial Cash Outlay"
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
            Height          =   5685
            Left            =   4920
            TabIndex        =   82
            Top             =   -30
            Width           =   5595
            Begin VB.TextBox txtOldCS 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   60
               TabIndex        =   161
               TabStop         =   0   'False
               Text            =   " "
               Top             =   3624
               Width           =   2865
            End
            Begin VB.TextBox txtCL_Discount 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   360
               Left            =   3000
               TabIndex        =   86
               Text            =   " "
               ToolTipText     =   "Encode Discount (if Applicable)"
               Top             =   526
               Width           =   2505
            End
            Begin VB.TextBox txtCL_Chattel 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   96
               Text            =   " "
               ToolTipText     =   "Chattel Mortgage"
               Top             =   2451
               Width           =   2505
            End
            Begin VB.TextBox txtCL_DownpaymentPert 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               MaxLength       =   5
               TabIndex        =   100
               TabStop         =   0   'False
               Text            =   " "
               ToolTipText     =   "Encode Downpayment Either By Percentage"
               Top             =   3233
               Width           =   585
            End
            Begin VB.TextBox txtCL_AORRate 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               MaxLength       =   10
               TabIndex        =   102
               Text            =   " "
               ToolTipText     =   "Encode Add on Rate For The Financing In Terms Of Percentage Value (10, 12,18.7)"
               Top             =   3624
               Width           =   1035
            End
            Begin VB.TextBox txtCL_BankTerms 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   104
               Text            =   " "
               ToolTipText     =   "Encode Total Bank Term (24 months, 36 months etc)"
               Top             =   4015
               Width           =   2505
            End
            Begin VB.TextBox txtCL_SalesPrice 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   360
               Left            =   3000
               TabIndex        =   84
               Tag             =   "@R"
               Text            =   " "
               Top             =   180
               Width           =   2505
            End
            Begin VB.TextBox txtCL_BalToFinanced 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   105
               TabStop         =   0   'False
               Text            =   " "
               ToolTipText     =   "Computed Balanced to be financed for the Transaction"
               Top             =   4406
               Width           =   2505
            End
            Begin VB.TextBox txtCL_NetMoAmort 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   108
               Text            =   " "
               ToolTipText     =   "Encode or Computed Net Monthly Amortization"
               Top             =   4797
               Width           =   2505
            End
            Begin VB.TextBox txtCL_LTORegFee 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   90
               Text            =   " "
               ToolTipText     =   "Land Transporation Registration Fees"
               Top             =   1278
               Width           =   2505
            End
            Begin VB.TextBox txtCL_DownPayment 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3600
               TabIndex        =   101
               Text            =   " "
               ToolTipText     =   "Encode Downpayment Either By Amount"
               Top             =   3240
               Width           =   1905
            End
            Begin VB.TextBox txtCL_Insurance 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   98
               Text            =   " "
               ToolTipText     =   "Insurance Amount"
               Top             =   2842
               Width           =   2505
            End
            Begin VB.TextBox txtCL_Freight 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   92
               Text            =   " "
               ToolTipText     =   "Additional Any Delivery Charges"
               Top             =   1669
               Width           =   2505
            End
            Begin VB.TextBox txtCL_Others 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               TabIndex        =   94
               Text            =   " "
               ToolTipText     =   "Encode Other Type of Input  Additional To Net Amount Due"
               Top             =   2070
               Width           =   2505
            End
            Begin VB.TextBox txtOthersDesc 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   375
               Left            =   450
               TabIndex        =   93
               Text            =   " "
               ToolTipText     =   "Other Type of Input  Additional To Net Amount Due"
               Top             =   2070
               Width           =   2535
            End
            Begin VB.TextBox txtCL_TotalDue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   111
               TabStop         =   0   'False
               Text            =   " "
               ToolTipText     =   "Computed Total Amount Due"
               Top             =   5190
               Width           =   2505
            End
            Begin VB.TextBox txtCL_NetSalesPrice 
               Alignment       =   1  'Right Justify
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
               ForeColor       =   &H00800000&
               Height          =   360
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   88
               Text            =   " "
               ToolTipText     =   "Computed Value Total Gross Amount - Discount"
               Top             =   902
               Width           =   2505
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DISCOUNT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1965
               TabIndex        =   85
               ToolTipText     =   "Encode Discount (if Applicable)"
               Top             =   570
               Width           =   945
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AOR (Rate)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   12
               Left            =   1980
               TabIndex        =   110
               ToolTipText     =   "Add on Rate For The Financing Option"
               Top             =   3720
               Width           =   930
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL AMOUNT DUE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   9
               Left            =   1155
               TabIndex        =   109
               ToolTipText     =   "Total Amount Due"
               Top             =   5280
               Width           =   1755
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DOWN PAYMENT (%)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   1140
               TabIndex        =   99
               ToolTipText     =   "Downpayment Either By Percentage or by Amount"
               Top             =   3330
               Width           =   1770
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BALANCE TO BE FINANCED"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   555
               TabIndex        =   107
               ToolTipText     =   "Computed Balanced to be financed for the Transaction"
               Top             =   4500
               Width           =   2355
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BANK TERM(s)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   1665
               TabIndex        =   103
               ToolTipText     =   "Total Bank Term (24 months, 36 months etc)"
               Top             =   4110
               Width           =   1245
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NET MONTHLY AMORTIZATION"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   315
               TabIndex        =   106
               ToolTipText     =   "Net Monthly Amortization"
               Top             =   4905
               Width           =   2595
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CHMO FEE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   1980
               TabIndex        =   95
               ToolTipText     =   "Chattel Mortgage"
               Top             =   2565
               Width           =   930
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SALES PRICE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   1740
               TabIndex        =   83
               ToolTipText     =   "Input Total Sales Price For The Vehicle"
               Top             =   225
               WhatsThisHelpID =   1
               Width           =   1170
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LTO REG. FEE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   6
               Left            =   1695
               TabIndex        =   89
               ToolTipText     =   "Land Transporation Registration Fees"
               Top             =   1395
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "INSURANCE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   7
               Left            =   1845
               TabIndex        =   97
               ToolTipText     =   "Insurance Amount"
               Top             =   2940
               Width           =   1065
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FREIGHT && HANDLING"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   8
               Left            =   990
               TabIndex        =   91
               ToolTipText     =   "Additional Any Delivery Charges"
               Top             =   1770
               Width           =   1920
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NET SALES PRICE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   1335
               TabIndex        =   87
               ToolTipText     =   "Computed Value Total Gross Amount - Discount"
               Top             =   990
               Width           =   1575
            End
         End
         Begin VB.Frame fraPurchInfo 
            Caption         =   "Purchase Information"
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
            Height          =   2115
            Left            =   30
            TabIndex        =   69
            Top             =   -30
            Width           =   4815
            Begin VB.ComboBox cboFinancingTerm 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               ItemData        =   "SalesOrder.frx":0FF5
               Left            =   2100
               List            =   "SalesOrder.frx":1005
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   510
               Width           =   2520
            End
            Begin VB.OptionButton opt1st 
               Caption         =   "1st"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   70
               ToolTipText     =   "Purchase Information(1st Purchase)"
               Top             =   240
               Width           =   675
            End
            Begin VB.OptionButton optRPL 
               Caption         =   "RPL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2280
               TabIndex        =   71
               Top             =   240
               Width           =   765
            End
            Begin VB.OptionButton optADDL 
               Caption         =   "ADDL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3090
               TabIndex        =   72
               Top             =   240
               Width           =   765
            End
            Begin VB.OptionButton optTRI 
               Caption         =   "TRI"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4020
               TabIndex        =   73
               Top             =   240
               Width           =   585
            End
            Begin VB.ComboBox cboSalesAE 
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
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   80
               Tag             =   "@R"
               ToolTipText     =   "Sales Agent"
               Top             =   1680
               Width           =   4575
            End
            Begin VB.ComboBox cboFinancingCo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   345
               ItemData        =   "SalesOrder.frx":102F
               Left            =   150
               List            =   "SalesOrder.frx":1031
               TabIndex        =   78
               Top             =   1080
               Width           =   4545
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Account Executive"
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
               Left            =   135
               TabIndex        =   79
               Top             =   1440
               Width           =   1980
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Financing Company"
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
               Left            =   165
               TabIndex        =   77
               Top             =   810
               Width           =   1650
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Term"
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
               Left            =   1575
               TabIndex        =   74
               Top             =   570
               Width           =   435
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Type"
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
               Index           =   3
               Left            =   255
               TabIndex        =   75
               Top             =   270
               Width           =   1230
            End
         End
      End
   End
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   13095
      TabIndex        =   113
      Top             =   8205
      Width           =   13095
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   3360
         ScaleHeight     =   915
         ScaleWidth      =   12870
         TabIndex        =   114
         Top             =   0
         Width           =   12870
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   8970
            MouseIcon       =   "SalesOrder.frx":1033
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":1185
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   8280
            MouseIcon       =   "SalesOrder.frx":14EB
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":163D
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Print this Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   7590
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "SalesOrder.frx":19A3
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":1AF5
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Cancel this Transaction"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6900
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "SalesOrder.frx":1E2F
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":1F81
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Post this Transaction"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdUnPost 
            Caption         =   "Unpost Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6210
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "SalesOrder.frx":22A6
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":23F8
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Unpost this Transaction"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   5520
            MouseIcon       =   "SalesOrder.frx":273D
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":288F
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   4830
            MouseIcon       =   "SalesOrder.frx":2BEB
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":2D3D
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Multiple SO"
            Height          =   795
            Left            =   4140
            MouseIcon       =   "SalesOrder.frx":3050
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":31A2
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Add Multiple Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   795
            Left            =   3450
            MouseIcon       =   "SalesOrder.frx":330C
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":345E
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Move to Last Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   795
            Left            =   2760
            MouseIcon       =   "SalesOrder.frx":37AE
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":3900
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Move to First Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   525
            Left            =   930
            MouseIcon       =   "SalesOrder.frx":3C5E
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":3DB0
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Delete Selected Record"
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Fin&d"
            Height          =   795
            Left            =   2070
            MouseIcon       =   "SalesOrder.frx":40DB
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":422D
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   1380
            MouseIcon       =   "SalesOrder.frx":4527
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":4679
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   690
            MouseIcon       =   "SalesOrder.frx":49D1
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":4B23
            Style           =   1  'Graphical
            TabIndex        =   116
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Refresh"
            Height          =   795
            Left            =   0
            MouseIcon       =   "SalesOrder.frx":4E82
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":4FD4
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   11565
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   129
         Top             =   30
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   780
            MouseIcon       =   "SalesOrder.frx":554F
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":56A1
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   90
            MouseIcon       =   "SalesOrder.frx":59DF
            MousePointer    =   99  'Custom
            Picture         =   "SalesOrder.frx":5B31
            Style           =   1  'Graphical
            TabIndex        =   131
            ToolTipText     =   "Save this Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
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
         Height          =   390
         Left            =   60
         TabIndex        =   128
         Top             =   420
         Width           =   3105
      End
      Begin VB.Label lblSalesStatus 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   30
         TabIndex        =   127
         Top             =   60
         Width           =   3105
      End
   End
   Begin VB.Label labID 
      Height          =   315
      Left            =   75
      TabIndex        =   112
      Top             =   9525
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmSMIS_Trans_SalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Private AddingSalesOrder                                              As Boolean
Dim rsS_Model                                                         As ADODB.Recordset
Dim ctl                                                               As Control
Dim xxSONO                                                            As String
Dim AcctName                                                          As String
Dim PROSPECTID                                                        As Long
Dim CUSCDE                                                            As String
Dim ProfileType                                                       As String
Dim rsSO                                                              As ADODB.Recordset
Dim vID
Dim VCODE
Dim VDescript
Dim ComputebyPert                                                     As Boolean
Dim rsCustomerInfo                                                    As Recordset
Dim MULTIPLESO                                                        As Boolean
Private WithEvents EntryPoint                                         As frmSMIS_Trans_SOEntryPoint
Attribute EntryPoint.VB_VarHelpID = -1
Dim AccessFlag                                                        As Boolean

Function AddNewSOFromProspect(oCusRs As ADODB.Recordset) As Boolean
    Dim temprs                                                        As ADODB.Recordset
    AddNewSOFromProspect = False
    CUSCDE = Null2String(oCusRs!CUSCDE)
    If (oCusRs.EOF Or oCusRs.BOF) Then Exit Function
    If Null2String(oCusRs!CUSCDE) <> "" Then
        AddingSalesOrder = True
        initMemvars
        Set temprs = gconDMIS.Execute("Select  * from ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(oCusRs!CUSCDE))
        If temprs.EOF Or temprs.BOF Then
            Call MsgBox("Prospect Information Altered In Customer Database .. " _
                      & vbCrLf & "Please Update Convert Prospect Information Again." _
                      & vbCrLf & " Sales Order Will Now Exit" _
               , vbQuestion, "Customer Code Error")
            AddNewSOFromProspect = False
            frmSMIS_Files_Prospects.EditProspect (PROSPECTID)
            frmSMIS_Files_Prospects.Show
            frmSMIS_Files_Prospects.Show
            Unload Me
            Exit Function
        End If
        If Null2String(temprs!CUSTYPE) = "P" Then
            txtCusName = Trim(Null2String(temprs!Firstname) & " " & Null2String(temprs!MiddleInitial) & " " & Null2String(temprs!lastname))
            txtSpouse = Null2String(temprs!Spouse)
        Else
            txtCusName = Trim(Null2String(temprs!Firstname) & " " & Null2String(temprs!MiddleInitial) & " " & Null2String(temprs!lastname))
            'txtCusName = Null2String(temprs!CUSCOMP)
            txtPerson = Null2String(temprs!Spouse)
        End If
        txtHomeAdd = Null2String(temprs!CUSTOMERADD)
        txtHomeTelNo = Null2String(temprs!HomePhone)
        txtOfficeAdd = Null2String(temprs!CompanyAdd)
        txtOfficeTelNo = Null2String(temprs!TelephoneNo)
        txtBirthDate = Null2String(temprs!BirthDate)
        txtPosisyon = Null2String(temprs!TITLE)
        txtTIN = Null2String(temprs!TIN)
        txtCTCNo = Null2String(temprs!Mobile)
        txtPerson = Null2String(temprs!Assistant)
        
        CUSCDE = Null2String(oCusRs!CUSCDE)
        PROSPECTID = N2Str2Zero(oCusRs!PROSPECTID)
        txtModelDescription = Null2String(oCusRs!Variant)
        cboSalesAE = Null2String(oCusRs!SAE)

        'UPDATED BY: JUN
        'DATE UPDATED: 07/28/2008
        'DESCRIPTION: FOR HAI IT WILL GENERATE A SERIES NO.
        If VALID_COMPANY_CODE_FORHAI = True Then
            GenerateHAISOseries
        Else
             txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
        End If
'        If COMPANY_CODE <> "HAI" Then
'            txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
'        Else
'            GenerateHAISOseries
'        End If

        LABCUSTOMERCODE = "CUSTOMER CODE:" & Null2String(oCusRs!CUSCDE)
        AddorEdit = "ADD"
        fraCustInfo.Enabled = True
        picSalesOrder.Enabled = True
        picSalesOrder.Enabled = True
        picAdds.Visible = False
        picSaves.Visible = True
        Set temprs = Nothing
    End If
End Function

Function AddNewSOfromQuotation(oCusRs As ADODB.Recordset) As Boolean
    Dim temprs                                                        As ADODB.Recordset
    Dim rsProspect                                                    As ADODB.Recordset
    AddNewSOfromQuotation = False
    Set rsProspect = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE ProspectID=" & oCusRs("PROSPECTID"))
    If (rsProspect.EOF Or rsProspect.BOF) Then Exit Function
    If Null2String(rsProspect!CUSCDE) <> "" Then
        AddingSalesOrder = True
        initMemvars
        Set temprs = gconDMIS.Execute("Select  * from ALL_CUSTOMER WHERE CUSCDE=" & N2Str2Null(rsProspect!CUSCDE))
        If temprs.EOF Or temprs.BOF Then
            Call MsgBox("Prospect Information Altered In Customer Database .. " _
                      & vbCrLf & "Please Update Convert Prospect Information Again." _
                      & vbCrLf & " Sales Order Will Now Exit" _
               , vbQuestion, "Customer Code Error")
            AddNewSOfromQuotation = False
            frmSMIS_Files_Prospects.EditProspect (PROSPECTID)
            frmSMIS_Files_Prospects.Show

            frmSMIS_Files_Prospects.Show
            Unload Me
            Exit Function
        End If
        txtCusName = Trim(Null2String(temprs!Firstname) & " " & Null2String(temprs!MiddleInitial) & " " & Null2String(temprs!lastname))
        txtHomeAdd = Null2String(temprs!CUSTOMERADD)
        txtHomeTelNo = Null2String(temprs!HomePhone)
        txtOfficeAdd = Null2String(temprs!CompanyAdd)
        txtOfficeTelNo = Null2String(temprs!TelephoneNo)
        txtBirthDate = Null2String(temprs!BirthDate)
        txtSpouse = Null2String(temprs!Spouse)
        txtPerson = Null2String(temprs!Assistant)
        txtPosisyon = Null2String(temprs!TITLE)
        CUSCDE = Null2String(temprs!CUSCDE)
        PROSPECTID = N2Str2Zero(oCusRs!PROSPECTID)
        txtPerson = Null2String(rsProspect!ContactPerson)
        txtModelDescription = Null2String(oCusRs!ModelDescript)
        cboSalesAE = Null2String(rsProspect!SAE)
        txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
        cboFinancingTerm.ListIndex = 1
        cboFinancingCo = Null2String(oCusRs!FINCOMPANY)
        txtCL_Chattel = FormatNumber(NumericVal(oCusRs!FinChattel))
        txtCL_DownPayment = FormatNumber(NumericVal(oCusRs!finDownpayment))
        txtCL_Insurance = FormatNumber(NumericVal(oCusRs!FinInsurance))
        txtCL_LTORegFee = FormatNumber(NumericVal(oCusRs!FinLTO))
        txtOthersDesc = Null2String(oCusRs!FinOtherDesc)
        txtCL_Others = FormatNumber(NumericVal(oCusRs!FinOthers))
        txtCL_NetMoAmort = FormatNumber(NumericVal(oCusRs!FinChattel))
        txtCL_NetSalesPrice = FormatNumber(NumericVal(oCusRs!CASHUNITPRICE))
        txtCL_SalesPrice = FormatNumber(NumericVal(oCusRs!FinUnitPrice))
        AddorEdit = "ADD"
        fraCustInfo.Enabled = True
        picSalesOrder.Enabled = True
        picSalesOrder.Enabled = True
        picAdds.Visible = False
        picSaves.Visible = True
        Set temprs = Nothing
    End If
End Function

Function GetModels(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select ID, Code, DESCRIPT from All_Model where ltrim(rtrim(DESCRIPT)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        vID = N2Str2Null(rsS_Model!ID)
        VCODE = N2Str2Null(rsS_Model!CODE)
        VDescript = N2Str2Null(rsS_Model!DESCRIPT)
        GetModels = N2Str2Null(rsS_Model!CODE)

    Else
        vID = "NULL"
        VCODE = "NULL"
        VDescript = "NULL"
        GetModels = "NULL"
    End If
    Set rsS_Model = Nothing
End Function

Function SetColor(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_CODE,COLOR_DESC from ALL_Color where ltrim(rtrim(COLOR_DESC)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetColor = N2Str2Null(rsS_Model!Color_code) Else SetColor = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetColorDesc(XXX As String) As String
    Dim rsColor                                                       As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    Set rsColor = gconDMIS.Execute("select color_desc,Color_code from ALL_Color where color_code = '" & ReplaceQuote(XXX) & "'")
    If Not (rsColor.EOF And rsColor.BOF) Then
        SetColorDesc = Null2String(rsColor!color_desc)
    End If
End Function

Function SetColorName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_CODE,COLOR_DESC from ALL_Color where ltrim(rtrim(COLOR_CODE)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetColorName = Null2String(rsS_Model!color_desc) Else SetColorName = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetFinancing(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,COMPANY from SMIS_FinCom where ltrim(rtrim(COMPANY)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetFinancing = N2Str2Null(rsS_Model!CODE) Else SetFinancing = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetFinancingName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,COMPANY from SMIS_FinCom where ltrim(rtrim(CODE)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetFinancingName = Null2String(rsS_Model!company) Else SetFinancingName = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetModelName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Code,DESCRIPT from All_Model where ltrim(rtrim(Code)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetModelName = Null2String(rsS_Model!DESCRIPT) Else SetModelName = "NULL"
    Set rsS_Model = Nothing
End Function

Private Function AORVALUE(Principal, AOR, TERM) As Double
    On Error Resume Next

    If AOR <= 0 Then: AORVALUE = 0: Exit Function
    If Principal <= 0 Then: AORVALUE = 0: Exit Function
    If TERM <= 0 Then: AORVALUE = 0: Exit Function
    Dim Interest                                                      As Double
    Interest = NumericVal(AOR)
    Interest = AOR / 1200
    '
    '        AORVALUE = FormatNumber((Principal * Interest / (1 - ((1 / (1 + Interest) ^ Term)))), 2)
    'AORVALUE = FormatNumber((Principal * AOR / 12), 2)

    'If AOR <= 0 Then Exit Function
    'If Principal <= 0 Then Exit Function
    'If Term <= 0 Then Exit Function
    'Dim Interest                        As Double

    '    AORVALUE = FormatNumber((txtCL_BalToFinanced * (1 + (AOR / 100))) / Term)

    'AORVALUE = FormatNumber((Principal * AOR / TERM), 2)
    If COMPANY_CODE = "HPI" Then
        'UPDATED BY: JUN
        'DATE UPDATED: 03192009
        AORVALUE = FormatNumber((Principal * (1 + (AOR / 100))) / TERM)
    Else
        AORVALUE = FormatNumber((Principal * AOR / TERM), 2)
    End If
End Function

Private Function Runvalidation(strcase As String) As Boolean
    Runvalidation = False
    Dim txt                                                           As Control
    For Each txt In Me.ControlS
        If (TypeOf txt Is TextBox Or TypeOf txt Is ComboBox) And txt.Tag = strcase Then
            If Trim(txt.Text) = vbNullString Then
                MessagePop RecSaveError, "Required Filed Missing", txt.ToolTipText & " is Required Field", 1000
                Call ColorIt(txt, Timer1)
                On Error Resume Next
                txt.SetFocus
                Exit Function
            End If
        End If
    Next
    Runvalidation = True
End Function

Sub AddNewSODirect(xCustRs As ADODB.Recordset)
    If Not (xCustRs.EOF Or xCustRs.BOF) Then
        initMemvars
        txtCusName = Trim(Null2String(xCustRs!Firstname) & " " & Null2String(xCustRs!MiddleInitial) & " " & Null2String(xCustRs!lastname))
        txtHomeAdd = Null2String(xCustRs!CUSTOMERADD)
        AcctName = Null2String(xCustRs!AcctName)
        CUSCDE = Null2String(xCustRs!CUSCDE)

        ProfileType = Null2String(xCustRs!CUSTYPE)
        txtOfficeAdd = Null2String(xCustRs!CUSTOMERADD)
        txtHomeTelNo = Null2String(xCustRs!HomePhone)
        txtOfficeTelNo = Null2String(xCustRs!TelephoneNo)
        txtBirthDate = Null2String(xCustRs!BirthDate)
        txtSpouse = Null2String(xCustRs!Spouse)
        txtPerson = Null2String(xCustRs!Assistant)
        txtPosisyon = Null2String(xCustRs!TITLE)
        txtTIN = Null2String(xCustRs!TIN)
        txtCTCNo = Null2String(xCustRs!Mobile)
        'txtCTCNo = ""
        txtIssuedAt = ""
        PROSPECTID = 0
        If VALID_COMPANY_CODE_FORHAI = True Then
            GenerateHAISOseries
        Else
             txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
        End If
'        If COMPANY_CODE <> "HAI" Then
'            txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
'        Else
'            GenerateHAISOseries
'        End If
        LABCUSTOMERCODE = "CUSTOMER CODE:" & Null2String(xCustRs!CUSCDE)
        AddorEdit = "ADD"
        picSalesOrder.Enabled = True
        fraCustInfo.Enabled = True
        picAdds.Visible = False
        picSaves.Visible = True

    End If
End Sub

Sub GenerateHAISOseries()
    'UPDATED BY: JUN
    'DATE UPDATE: 07/28/2008
    Dim rsSERIES_NO                                                   As ADODB.Recordset

    Set rsSERIES_NO = New ADODB.Recordset
    rsSERIES_NO.Open "select SO_NO from SMIS_Salesorder where len(SO_NO) = 6 order by SO_NO desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSERIES_NO.EOF And Not rsSERIES_NO.BOF Then
        txt_SONO.Text = Format(NumericVal(Mid$(rsSERIES_NO!SO_NO, 3, 6)) + 1, "000000")
    Else
        txt_SONO.Text = "000001"
    End If
End Sub

Sub AddNewSOFromApplication(oCusRs As ADODB.Recordset)
    If (oCusRs.EOF Or oCusRs.BOF) Then: Exit Sub
    initMemvars
    If oCusRs!ApplicationType = "I" Then
        txtCusName = Trim(Null2String(oCusRs!Ind_Apl_FirstName) & " " & Null2String(oCusRs!Ind_Apl_MidName) & " " & Null2String(oCusRs!Ind_Apl_LastName))
        txtHomeAdd = Null2String(oCusRs!Ind_Address)
        AcctName = Trim(Null2String(oCusRs!Ind_Apl_FirstName) & " " & Null2String(oCusRs!Ind_Apl_MidName) & " " & Null2String(oCusRs!Ind_Apl_LastName))
        CUSCDE = Null2String(oCusRs!AplCode)
        PROSPECTID = N2Str2Zero(oCusRs!PROSPECTID)
        txtPosisyon = Null2String(oCusRs!Ind_Apl_Position)
        txtOfficeAdd = Null2String(oCusRs!Ind_Apl_Address)
        txtHomeAdd = Null2String(oCusRs!Ind_Address)
        txtHomeTelNo = Null2String(oCusRs!Ind_TelNo)
        txtOfficeTelNo = Null2String(oCusRs!Ind_Apl_TelNo)
        txtBirthDate = Null2String(oCusRs!Ind_Apl_Birthday)
        txtSpouse = Null2String(oCusRs!Ind_Sps_LastName) & " " & Null2String(oCusRs!Ind_Sps_FirstName) & " " & Null2String(oCusRs!Ind_Sps_MidName)
        txtTIN = ""
        txtModelDescription = Null2String(oCusRs!Ind_LoanApl_UnitModel)
        txtCTCNo = ""
        txtIssuedAt = ""
        txtIssuedOn = ""
    Else
        txtCusName = Trim(Null2String(oCusRs!Busname))
        txtOfficeAdd = Null2String(oCusRs!OfficeAdd)
        txtPerson = Trim(Null2String(oCusRs!ODContactPerson))
        txtPosisyon = Null2String(oCusRs!ODDesignation)
        txtHomeTelNo = Null2String(oCusRs!ODTelNo)
        CUSCDE = Null2String(oCusRs!AplCode)
        PROSPECTID = N2Str2Zero(oCusRs!PROSPECTID)
        txtTIN = Null2String(oCusRs!TINNO)
        txtModelDescription = Null2String(oCusRs!UnitModel)
        txtCTCNo = Null2String(oCusRs!CCINo)
        txtIssuedAt = Null2String(oCusRs!PlaceOfIssue)
        txtIssuedOn = Null2String(oCusRs!DateofIssue)
        cboSalesAE = Null2String(oCusRs!SAENAME)
        txtCL_SalesPrice = FormatNumber(NumericVal(oCusRs!NetCostPrice))
        txtCL_BankTerms = FormatNumber(NumericVal(oCusRs!TERMS))
        txtCL_AORRate = FormatNumber(NumericVal(oCusRs!AOR))
        txtCL_BalToFinanced = FormatNumber(NumericVal(oCusRs!BalanceFianced))
        txtCL_DownPayment = FormatNumber(NumericVal(oCusRs!DownPayment))
        txtCL_NetMoAmort = FormatNumber(NumericVal(oCusRs!MonthlyAmortization))
        txtCL_NetSalesPrice = FormatNumber(NumericVal(oCusRs!NetCostPrice))
        cboFinancingTerm.ListIndex = 1
    End If
    txt_SONO = GenerateCode("SMIS_SalesOrder", "SO_NO", "000000")
    AddorEdit = "ADD"
    picSalesOrder.Enabled = True
    fraCustInfo.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
End Sub

Sub InitCbo()


    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_DESC from ALL_Color order by COLOR_DESC asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboColor.Clear
        Do While Not rsS_Model.EOF
            cboColor.AddItem Null2String(rsS_Model!color_desc)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COMPANY from SMIS_FinCom")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboFinancingCo.Clear
        Do While Not rsS_Model.EOF
            cboFinancingCo.AddItem Null2String(rsS_Model!company)
            rsS_Model.MoveNext
        Loop
    End If



    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select NAME from SMIS_vw_Srep order by NAME asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboSalesAE.Clear
        Do While Not rsS_Model.EOF
            cboSalesAE.AddItem Null2String(rsS_Model!Name)
            rsS_Model.MoveNext
        Loop
    End If
End Sub

Sub initMemvars()
    With Me
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = vbNullString
            End If
        Next ctl
    End With

    LABALLOWREPRINT = ""
    txtDeyt.Value = Format(LOGDATE, "MM/dd/yyyy")

    picViewVehicles.Visible = False

    txtCL_SalesPrice = "0.00"
    txtCL_Discount = "0.00"
    txtCL_DownPayment = "0.00"
    txtCL_BalToFinanced = "0.00"
    txtCL_NetSalesPrice = "0.00"
    txtCL_DownPayment = "0.00"
    txtCL_Insurance = "0.00"
    txtCL_LTORegFee = "0.00"
    txtCL_Freight = "0.00"
    txtCL_Others = "0.00"
    txtCL_TotalDue = "0.00"
    txtCL_Chattel = "0.00"
    txtOldCS = ""
    LABCUSTOMERCODE = ""
    txtGMI = "0.00"
    txtRPPD = "0.00"
    cboColor.Text = ""
    cboFinancingCo.Text = ""
    txtModelDescription.Text = ""
    cboSalesAE.Text = ""
    cboFinancingTerm.ListIndex = 0
    opt1st.Value = True
    lblSalesStatus = ""
    lblVehicleStatus = ""
    lblStatus = ""
End Sub

Sub rsRefresh()
    Set rsSO = New ADODB.Recordset
    If LOGSAE <> "" Then
        rsSO.Open "SELECT * FROM SMIS_SALESORDER   WHERE  USERCODE='" & LOGSAE & " ' ORDER BY ID DESC ", gconDMIS, adOpenKeyset, adLockReadOnly
    Else
        rsSO.Open "SELECT * FROM SMIS_SALESORDER  order by ID  DESC ", gconDMIS, adOpenKeyset, adLockReadOnly
    End If
End Sub

Sub SaveSalesOrder()
    Dim xxCustName                                                    As String
    Dim xxDeyt                                                        As String
    Dim xxHomeTelNo                                                   As String
    Dim xxHomeAddress                                                 As String
    Dim xxOfficeAdd                                                   As String
    Dim xxOfficeTelNo                                                 As String
    Dim xxBirthDate                                                   As String
    Dim xxSpouse                                                      As String
    Dim xxPerson                                                      As String
    Dim xxPosisyon                                                    As String
    Dim xxTIN                                                         As String
    Dim xxCTCNo                                                       As String
    Dim xxIssuedAt                                                    As String
    Dim xxIssuedOn                                                    As String
    Dim xxmodel                                                       As String
    Dim xxProdNo                                                      As String
    Dim xxConductionSticker                                           As String
    Dim xxEngineNo                                                    As String
    Dim xxFrameNo                                                     As String
    Dim xxColor                                                       As String
    Dim xxType                                                        As String
    Dim xxTerm                                                        As String
    Dim xxFinancingCo                                                 As String
    Dim xxSalesAE                                                     As String
    Dim xx_SalesPrice                                                 As String
    Dim xx_NetSalesPrice                                              As String
    Dim xx_DownPayment                                                As String
    Dim xx_BalToFinanced                                              As Double
    Dim xxAdditionalInfo                                              As String
    Dim xx_GMI                                                        As String
    Dim xx_RPPD                                                       As String
    Dim xx_NetMoAmort                                                 As String
    Dim xx_Insurance                                                  As String
    Dim xx_LTORegFee                                                  As String
    Dim xx_Freight                                                    As Double
    Dim xxVinNo                                                       As String
    Dim xxModelDescript                                               As String
    Dim xxOthersDesc                                                  As String
    Dim xxxOthers                                                     As String
    Dim rsIGNKEY                                                      As ADODB.Recordset
    Dim strIGNKEY                                                     As String
    ''''''''''''personal
    xxCustName = N2Str2Null(txtCusName)
    xxDeyt = N2Str2Null(txtDeyt)

    xxHomeTelNo = N2Str2Null(txtHomeTelNo)
    xxHomeAddress = N2Str2Null(txtHomeAdd)
    xxOfficeAdd = N2Str2Null(txtOfficeAdd)
    xxOfficeTelNo = N2Str2Null(txtOfficeTelNo)
    xxBirthDate = N2Str2Null(txtBirthDate)
    xxSpouse = N2Str2Null(txtSpouse)
    xxPerson = N2Str2Null(txtPerson)
    xxPosisyon = N2Str2Null(txtPosisyon)

    xxTIN = N2Str2Null(txtTIN)
    xxCTCNo = N2Str2Null(txtCTCNo)
    xxIssuedAt = N2Str2Null(txtIssuedAt)
    xxIssuedOn = N2Str2Null(txtIssuedOn)


    If opt1st.Value = True Then
        xxType = "'1st'"
    ElseIf optRPL.Value = True Then
        xxType = "'RPL'"
    ElseIf optADDL.Value = True Then
        xxType = "'ADDL'"
    ElseIf optTRI.Value = True Then
        xxType = "'TRI'"
    End If

    If Left(cboFinancingTerm.Text, 1) = "C" Then
        xxTerm = "'COD'"
    ElseIf Left(cboFinancingTerm.Text, 1) = "F" Then
        xxTerm = "'F'"
    ElseIf Left(cboFinancingTerm.Text, 1) = "B" Then
        xxTerm = "'BPO'"
    End If
    If (cboFinancingTerm.Text) = "COMPANY PO" Then
        xxTerm = "'CPO'"
    End If

    xxFinancingCo = N2Str2Null(cboFinancingCo)
    xxSalesAE = N2Str2Null(cboSalesAE)
    xx_SalesPrice = NumericVal(txtCL_SalesPrice)
    xx_NetSalesPrice = NumericVal(txtCL_NetSalesPrice)
    xx_DownPayment = NumericVal(txtCL_DownPayment)
    xx_BalToFinanced = NumericVal(txtCL_BalToFinanced)
    xxAdditionalInfo = N2Str2Null(txtAdditionalInfo)
    xx_GMI = NumericVal(txtGMI)
    xx_RPPD = NumericVal(txtRPPD)
    xx_NetMoAmort = NumericVal(txtCL_NetMoAmort)
    xx_Insurance = NumericVal(txtCL_Insurance)
    xx_LTORegFee = NumericVal(txtCL_LTORegFee)
    xx_Freight = NumericVal(txtCL_Freight)
    xxSONO = N2Str2Null(txt_SONO)
    xxOthersDesc = N2Str2Null(txtOthersDesc)
    xxxOthers = NumericVal(txtCL_Others)
    ''''''''''''AUTOS
    xxVinNo = N2Str2Null(txtVinNumber)
    xxModelDescript = N2Str2Null(txtModelDescription)
    xxmodel = N2Str2Null(txtModel)
    xxProdNo = N2Str2Null(txtProdNo)
    xxConductionSticker = N2Str2Null(txtConductionSticker)
    xxEngineNo = N2Str2Null(txtEngineNo)
    xxFrameNo = N2Str2Null(txtFrameNumber)
    xxVinNo = N2Str2Null(txtVinNumber)
    xxColor = N2Str2Null(cboColor)

    'RESET OLD CS NO TO INVENTORY
    'DUBIOS CODE
    If Not LTrim(RTrim(txtOldCS)) = "" Then
        gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET PROSPECTID=NULL,RELEASED=0, customercode=NULL,datereleased=null, invoiceddate=null,IStatus='O', WithProsBuyers='N'  WHERE IGNKEY=" & N2Str2Null(LTrim(RTrim(txtOldCS))))
    End If


    If AddorEdit = "ADD" Then

        'gconDMIS.Execute ("INSERT INTO SMIS_SALESORDER " & _
         "(DISCOUNT, CUSTNAME, PROSPECTID, SO_NO,CODE,DEYT,HOMETELNO,HOMEADDRESS,OFFICETELNO,OFFICEADD,BIRTHDATE,SPOUSE,PERSON,POSISYON,TIN,CTCNO," & _
         "ISSUEDAT,ISSUEDON,MODEL, MODELDESCRIPTION,PRODNO,IGNKEY_NO,ENGINENO,FRAMENO,COLOR,TYPE,TERM,FINANCINGCO,SALESAE,SALESPRICE,NETSALESPRICE," & _
         "DOWNPAYMENT,BALTOFINANCED,ADDITIONALINFO,GMI,RPPD,NETMOAMORT,INSURANCE,LTOREGFEE,FREIGHT, AOR,MONTHSAMORT, VINO, OTHERSDESC, OTHERS, TOTAL, CHMOFEE,DOWNPAYMENTRATE)" & _
         " VALUES (" & NumericVal(txtCL_Discount) & " ," & xxCustName & " , " & PROSPECTID & " , " & xxSONO & ",'" & CUSCDE & "', " & xxDeyt & ", " & xxHomeTelNo & ", " & xxHomeAddress & ", " & xxOfficeTelNo & ", " & xxOfficeAdd & ", " & xxBirthDate & ", " & xxSpouse & ", " & xxPerson & ", " & xxPosisyon & ", " & xxTIN & _
         "," & xxCTCNo & ", " & xxIssuedAt & ", " & xxIssuedOn & ", " & xxmodel & "," & xxModelDescript & ", " & xxProdNo & ", " & xxConductionSticker & ", " & xxEngineNo & ", " & xxFrameNo & ", " & xxColor & ", " & xxType & ", " & xxTerm & ", " & xxFinancingCo & ", " & xxSalesAE & _
         "," & xx_SalesPrice & ", " & xx_NetSalesPrice & ", " & xx_DownPayment & ", " & xx_BalToFinanced & ", " & xxAdditionalInfo & ", " & xx_GMI & _
         "," & xx_RPPD & ", " & xx_NetMoAmort & ", " & xx_Insurance & ", " & xx_LTORegFee & ", " & xx_Freight & "," & NumericVal(txtCL_AORRate) & "," & NumericVal(txtCL_BankTerms) & "," & xxVinNo & "," & xxOthersDesc & "," & xxxOthers & "," & NumericVal(txtCL_TotalDue) & "," & NumericVal(txtCL_Chattel) & "," & NumericVal(txtCL_DownpaymentPert) & " )")
        SQL_STATEMENT = ("INSERT INTO SMIS_SALESORDER " & _
                         "(DISCOUNT, CUSTNAME, PROSPECTID, SO_NO,CODE,DEYT,HOMETELNO,HOMEADDRESS,OFFICETELNO,OFFICEADD,BIRTHDATE,SPOUSE,PERSON,POSISYON,TIN,CTCNO," & _
                         "ISSUEDAT,ISSUEDON,MODEL, MODELDESCRIPTION,PRODNO,IGNKEY_NO,ENGINENO,FRAMENO,COLOR,TYPE,TERM,FINANCINGCO,SALESAE,SALESPRICE,NETSALESPRICE," & _
                         "DOWNPAYMENT,BALTOFINANCED,ADDITIONALINFO,GMI,RPPD,NETMOAMORT,INSURANCE,LTOREGFEE,FREIGHT, AOR,MONTHSAMORT, VINO, OTHERSDESC, OTHERS, TOTAL, CHMOFEE,DOWNPAYMENTRATE)" & _
                       " VALUES (" & NumericVal(txtCL_Discount) & " ," & xxCustName & " , " & PROSPECTID & " , " & xxSONO & ",'" & CUSCDE & "', " & xxDeyt & ", " & xxHomeTelNo & ", " & xxHomeAddress & ", " & xxOfficeTelNo & ", " & xxOfficeAdd & ", " & xxBirthDate & ", " & xxSpouse & ", " & xxPerson & ", " & xxPosisyon & ", " & xxTIN & _
                         "," & xxCTCNo & ", " & xxIssuedAt & ", " & xxIssuedOn & ", " & xxmodel & "," & xxModelDescript & ", " & xxProdNo & ", " & xxConductionSticker & ", " & xxEngineNo & ", " & xxFrameNo & ", " & xxColor & ", " & xxType & ", " & xxTerm & ", " & xxFinancingCo & ", " & xxSalesAE & _
                         "," & xx_SalesPrice & ", " & xx_NetSalesPrice & ", " & xx_DownPayment & ", " & xx_BalToFinanced & ", " & xxAdditionalInfo & ", " & xx_GMI & _
                         "," & xx_RPPD & ", " & xx_NetMoAmort & ", " & xx_Insurance & ", " & xx_LTORegFee & ", " & xx_Freight & "," & NumericVal(txtCL_AORRate) & "," & NumericVal(txtCL_BankTerms) & "," & xxVinNo & "," & xxOthersDesc & "," & xxxOthers & "," & NumericVal(txtCL_TotalDue) & "," & NumericVal(txtCL_Chattel) & "," & NumericVal(txtCL_DownpaymentPert) & " )")
        '******************
        'THIS IS THE NEW LOG AUDIT
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txt_SONO), "SO_No", "SMIS_SALESORDER"), "", "SO No:" & txt_SONO, "", ""
        '******************

        gconDMIS.Execute ("Update CRIS_Prospects Set LOGSO=getdate() , SO_NO=" & N2Str2Null(xxSONO) & " , STATUS='O' where ProspectID=" & PROSPECTID)
        MessagePop RecSave, "Sales Order Added", "New Sales Order Has Been Added", 1

        LogAudit "A", "SALES ORDER", txt_SONO & " " & cboSalesAE & " TERM:" & cboFinancingTerm
    Else
        Set rsIGNKEY = gconDMIS.Execute("SELECT IGNKEY_NO FROM SMIS_SALESORDER WHERE SO_NO='" & txt_SONO & "'")
        If Not (rsIGNKEY.EOF Or rsIGNKEY.BOF) Then
            strIGNKEY = Null2String(rsIGNKEY.Collect(0))

        End If
        Set rsIGNKEY = Nothing

        MessagePop RecSave, "Sales Order Updated", "Sales Order Has Been Updated", 1
        SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET" & _
                      " DEYT = " & xxDeyt & "," & _
                      " PROSPECTID = " & PROSPECTID & "," & _
                      " HOMEADDRESS= " & xxHomeAddress & "," & _
                      " HOMETELNO = " & xxHomeTelNo & "," & _
                      " OFFICEADD = " & xxOfficeAdd & "," & _
                      " OFFICETELNO = " & xxOfficeTelNo & "," & _
                      " BIRTHDATE = " & xxBirthDate & "," & _
                      " SPOUSE = " & xxSpouse & "," & _
                      " PERSON = " & xxPerson & "," & _
                      " POSISYON= " & xxPosisyon & "," & _
                      " CODE = '" & CUSCDE & "'" & _
                      " WHERE [ID] = " & labID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "EE", "SALES ORDER", SQL_STATEMENT, Null2String(labID), "", "SO No:" & txt_SONO, "", ""
        '*****Reset the Variable*****
        SQL_STATEMENT = ""
        '***************************
        SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET" & _
                      " TIN = " & xxTIN & "," & _
                      " CTCNO = " & xxCTCNo & "," & _
                      " ISSUEDAT = " & xxIssuedAt & "," & _
                      " ISSUEDON = " & xxIssuedOn & "," & _
                      " MODEL = " & xxmodel & "," & _
                      " MODELDESCRIPTION = " & xxModelDescript & "," & _
                      " PRODNO = " & xxProdNo & "," & _
                      " IGNKEY_NO = " & xxConductionSticker & "," & _
                      " ENGINENO = " & xxEngineNo & "," & _
                      " FRAMENO = " & xxFrameNo & "," & _
                      " VINO = " & xxVinNo & "," & _
                      " COLOR = " & xxColor & "," & _
                      " TYPE    = " & xxType & "," & _
                      " TERM = " & xxTerm & "" & _
                      " WHERE [ID] = " & labID

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "SALES ORDER", SQL_STATEMENT, Null2String(labID), "", "SO No:" & txt_SONO, "", ""

        '*****Reset the Variable*****
        SQL_STATEMENT = ""
        '***************************
        SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET" & _
                      " FINANCINGCO = " & xxFinancingCo & "," & _
                      " SALESAE = " & xxSalesAE & "," & _
                      " SALESPRICE = " & xx_SalesPrice & "," & _
                      " SO_NO= " & N2Str2Null(txt_SONO) & "," & _
                      " DISCOUNT= " & NumericVal(txtCL_Discount) & "," & _
                      " NETSALESPRICE = " & xx_NetSalesPrice & "," & _
                      " DOWNPAYMENT = " & xx_DownPayment & "," & _
                      " BALTOFINANCED = " & xx_BalToFinanced & "," & _
                      " ADDITIONALINFO = " & xxAdditionalInfo & "," & _
                      " GMI = " & xx_GMI & "," & _
                      " RPPD = " & xx_RPPD & "," & _
                      " MONTHSAMORT = " & NumericVal(txtCL_BankTerms) & "," & _
                      " AOR = " & NumericVal(txtCL_AORRate) & "," & _
                      " NETMOAMORT = " & xx_NetMoAmort & "," & _
                      " INSURANCE = " & xx_Insurance & "," & _
                      " LTOREGFEE = " & xx_LTORegFee & "," & _
                      " DOWNPAYMENTRATE = " & NumericVal(txtCL_DownpaymentPert) & "," & _
                      " CHMOFEE = " & NumericVal(txtCL_Chattel) & "," & _
                      " OTHERSDESC = " & xxOthersDesc & "," & _
                      " OTHERS = " & xxxOthers & "," & _
                      " TOTAL= " & NumericVal(txtCL_TotalDue) & "," & _
                      " FREIGHT = " & xx_Freight & "" & _
                      " WHERE [ID] = " & labID
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "SALES ORDER", SQL_STATEMENT, N2Str2Null(labID), "", "SO No:" & txt_SONO, "", ""
        LogAudit "E", "SALES ORDER", txt_SONO & " " & cboSalesAE & " TERM:" & cboFinancingTerm
        '*****Reset the Variable*****
        SQL_STATEMENT = ""
        '***************************

    End If

    'IF SALES ORDER HAS BEEN ALREADY BEEN INVOICED THEN ' DON'T CHANGE THE STATUS

    If AddorEdit = "ADD" And Len(txtConductionSticker.Text) > 0 Then
        'gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET " _
         & " ISTATUS='A'," _
         & " PROSPECTID=" & PROSPECTID & "," _
         & " PROSPECTCOUNTER=ISNULL(PROSPECTCOUNTER,0) + 1 ," _
         & " WITHPROSBUYERS='Y' ," _
         & " CUSTOMERCODE='" & CUSCDE & "'" _
         & " WHERE IGNKEY=" & N2Str2Null(txtConductionSticker))

        SQL_STATEMENT = ("UPDATE SMIS_MRRINV_TABLE SET " _
                       & " ISTATUS='A'," _
                       & " PROSPECTID=" & PROSPECTID & "," _
                       & " PROSPECTCOUNTER=ISNULL(PROSPECTCOUNTER,0) + 1 ," _
                       & " WITHPROSBUYERS='Y' ," _
                       & " CUSTOMERCODE='" & CUSCDE & "'" _
                       & " WHERE IGNKEY=" & N2Str2Null(txtConductionSticker))

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "EE", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtConductionSticker), "ignkey", "SMIS_MRRINV_TABLE"), "", "SO No:" & txt_SONO, "", ""

    ElseIf AddorEdit = "EDIT" And Len(txtConductionSticker.Text) > 0 Then
        If Null2String(rsSO!VI_NO) = "" Then
            If IsDate(rsSO!InvoicedDate) = False Then
                '        gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET " _
                         & " ISTATUS='A'," _
                         & " PROSPECTID=" & PROSPECTID & "," _
                         & " CUSTOMERCODE='" & rsSO!CODE & "'" _
                         & " WHERE UPPER(IGNKEY)=" & N2Str2Null(txtConductionSticker))
                SQL_STATEMENT = ("UPDATE SMIS_MRRINV_TABLE SET " _
                               & " ISTATUS='A'," _
                               & " PROSPECTID=" & PROSPECTID & "," _
                               & " CUSTOMERCODE='" & rsSO!CODE & "'" _
                               & " WHERE UPPER(IGNKEY)=" & N2Str2Null(txtConductionSticker))

                gconDMIS.Execute (SQL_STATEMENT)
                NEW_LogAudit "EE", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtConductionSticker), "ignkey", "SMIS_MRRINV_TABLE"), "", "SO No:" & txt_SONO, "", ""
            End If

        End If
    End If
    '*****Reset the Variable*****
    SQL_STATEMENT = ""
    '***************************
End Sub

Sub SearchID(SOID As Long)
    If SOID <> 0 Then
        rsSO.MoveFirst
        rsSO.Find ("ID=" & SOID)
        'If Null2String(rsSO!SOStatus) = "C" Then
        '    picAdds.Visible = True
        '    picSaves.Visible = False
        'ElseIf Null2String(rsSO!SOStatus) = "P" Then
        '    picAdds.Visible = True
        '    picSaves.Visible = False
        'Else
        '    picSalesOrder.Enabled = True
        '    fraCustInfo.Enabled = True
        '    picAdds.Visible = False
        '    picSaves.Visible = True
        'End If

        'If IsDate(rsSO!InvoicedDate) = True Then
        '    picAdds.Visible = True
        '    picSaves.Visible = False
        'End If
    End If
    StoreMemVars
End Sub

Sub SetModelNo(Kode As String)
    Dim rsMRRINV                                                      As ADODB.Recordset
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "select * from SMIS_MRRINV_TABLE WHERE prodno = '" & Kode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        'labModId.Caption = rsMRRINV!Id
        txtProdNo.Text = Null2String(rsMRRINV!prodno)
        txtEngineNo.Text = Null2String(rsMRRINV!EngineNo)
        txtVinNumber.Text = Null2String(rsMRRINV!SERIALNO)
        cboColor.Text = Null2String(Null2String(rsMRRINV!Color))
        txtModelDescription.Text = Null2String(rsMRRINV!DESCRIPT)
        txtConductionSticker.Text = Null2String(rsMRRINV!ignkey)
    Else
        txtProdNo.Text = vbNullString
        txtEngineNo.Text = vbNullString
        txtVinNumber.Text = vbNullString
        cboColor.Text = vbNullString
        txtModelDescription.Text = vbNullString
        txtConductionSticker.Text = vbNullString
    End If
End Sub

Sub StoreMemVars()
    If Not rsSO.EOF And Not rsSO.BOF Then
        labID = rsSO!ID
        LABALLOWREPRINT = Null2String(rsSO!PRINTED)
        txt_SONO = Null2String(rsSO!SO_NO)
        AcctName = Null2String(rsSO!CODE)
        txtCusName = Null2String(rsSO!CustName)
        txtDeyt = Null2String(rsSO!DEYT)
        txtHomeTelNo = Null2String(rsSO!HomeTelNo)

        txtHomeAdd = Null2String(rsSO!HomeAddress)
        txtOfficeAdd = Null2String(rsSO!OfficeAdd)
        txtOfficeTelNo = Null2String(rsSO!officetelno)
        txtBirthDate = Null2String(rsSO!BirthDate)
        txtSpouse = Null2String(rsSO!Spouse)
        txtPerson = Null2String(rsSO!Person)
        txtPosisyon = Null2String(rsSO!posisyon)
        txtTIN = Null2String(rsSO!TIN)
        txtCTCNo = Null2String(rsSO!CtcNo)
        txtIssuedAt = Null2String(rsSO!IssuedAt)
        txtIssuedOn = Null2String(rsSO!IssuedOn)
        PROSPECTID = Null2String(rsSO!PROSPECTID)
        txtModel = Null2String(rsSO!Model)
        LABCUSTOMERCODE = "CUSTOMER CODE:" & Null2String(rsSO!CODE)
        txtModelDescription = Null2String(rsSO!modeldescription)
        txtProdNo = Null2String(rsSO!prodno)
        txtConductionSticker = Null2String(rsSO!IGNKEY_NO)
        txtOldCS = Null2String(rsSO!IGNKEY_NO)
        txtEngineNo = Null2String(rsSO!EngineNo)
        txtVinNumber = Null2String(rsSO!VINO)
        cboColor.Text = Null2String(rsSO!Color)
        txtFrameNumber = Null2String(rsSO!frameno)
        If rsSO![Type] = "1st" Then
            opt1st.Value = True
        ElseIf rsSO![Type] = "RPL" Then
            optRPL.Value = True
        ElseIf rsSO![Type] = "ADDL" Then
            optADDL.Value = True
        ElseIf rsSO![Type] = "TRI" Then
            optTRI.Value = True
        End If
        If Null2String(rsSO!TERM) = "COD" Then
            cboFinancingTerm.ListIndex = 0
        ElseIf Null2String(rsSO!TERM) = "F" Then
            cboFinancingTerm.ListIndex = 1
        ElseIf Null2String(rsSO!TERM) = "BPO" Then
            cboFinancingTerm.ListIndex = 2
        ElseIf Null2String(rsSO!TERM) = "CPO" Then
            cboFinancingTerm.ListIndex = 3
        End If
        'BTT 01312008
        'If Null2String(rsSO!Term) = "CPO" Then
        '    cboFinancingTerm.ListIndex = 3
        'End If
        txtAdditionalInfo = Null2String(rsSO!ADDITIONALINFO)
        cboFinancingCo.Text = Null2String(rsSO!financingco)
        cboSalesAE.Text = Null2String(rsSO!salesae)
        txtCL_SalesPrice = FormatNumber(NumericVal(rsSO!SALESPRICE))
        txtCL_Discount = FormatNumber(NumericVal(rsSO!DISCOUNT))
        txtCL_NetSalesPrice = FormatNumber(NumericVal(rsSO!NETSALESPRICE))
        txtCL_DownPayment = FormatNumber(NumericVal(rsSO!DownPayment))
        txtCL_DownpaymentPert = FormatNumber(rsSO!DOWNPAYMENTRATE)
        txtCL_BalToFinanced = FormatNumber(NumericVal(rsSO!BALTOFINANCED))
        txtCL_NetMoAmort = FormatNumber(NumericVal(rsSO!NETMOAMORT))
        txtCL_AORRate = FormatNumber(NumericVal(rsSO!AOR))
        txtGMI = FormatNumber(NumericVal(rsSO!GMI))
        txtRPPD = FormatNumber(NumericVal(rsSO!RPPD))
        txtCL_Chattel = FormatNumber(rsSO!CHMOFEE)

        txtCL_TotalDue = FormatNumber(rsSO!Total)
        txtCL_Insurance = FormatNumber(NumericVal(rsSO!INSURANCE))
        txtCL_LTORegFee = FormatNumber(NumericVal(rsSO!LTOREGFEE))
        txtCL_Freight = FormatNumber(NumericVal(rsSO!FREIGHT))
        txtOthersDesc = Null2String(rsSO!OTHERSDESC)
        txtCL_Others = FormatNumber(NumericVal(rsSO!OTHERS))
        txtCL_BankTerms = Null2String(rsSO!MONTHSAMORT)
        CUSCDE = Null2String(rsSO!CODE)
        txtVinNumber = Null2String(rsSO!VINO)
        txtModel = Null2String(rsSO!Model)


        lblStatus = "": lblVehicleStatus = "": lblSalesStatus = ""




        If Null2String(rsSO!STATUS) = "C" Then
            lblStatus = "***CANCELLED INVOICED***"
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdPrint.Enabled = False
            lblVehicleStatus = ""
        Else
            If Null2String(rsSO!SOSTATUS) = "C" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = False
                cmdPost.Enabled = False
                lblSalesStatus = ""
                lblStatus = "***CANCELLED SALES ORDER***"
                lblSalesStatus = ""
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                lblVehicleStatus = ""
                cmdPrint.Enabled = False
            ElseIf Null2String(rsSO!SOSTATUS) = "P" Then
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = True
                cmdPost.Enabled = False
                lblStatus = "***POSTED***"
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdPrint.Enabled = True
            Else
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = True
                cmdEdit.Enabled = True
                cmdDelete.Enabled = False
                cmdPrint.Enabled = False
            End If

            If IsDate(rsSO!InvoicedDate) = True Then
                picSalesOrder.Enabled = False
                cmdEdit.Enabled = False
                cmdPrint.Enabled = True
                cmdUnPost.Enabled = False
                cmdPost.Enabled = False
            End If

        End If





        If txtConductionSticker.Text = "" Then
            lblVehicleStatus.Caption = "***NOT IN STOCK***"
        End If

        If Null2String(rsSO!SOSTATUS) <> "C" And Null2String(rsSO!STATUS) <> "C" Then

            If IsDate(rsSO!DateReleased) = True Then
                txtDateRelease = DateValue(Null2String(rsSO!DateReleased))
                txtTimeRelease = TimeValue(rsSO!DateReleased)
            Else
                txtDateRelease = ""
                txtTimeRelease = ""
            End If
            txtInvoicedDate = Null2String(rsSO!InvoicedDate)
            Dim grs                                                   As ADODB.Recordset
            Dim lStatus                                               As String
            Set grs = gconDMIS.Execute("select  IStatus  from SMIS_MRRINV_TABLE WHERE IGNKEY=" & N2Str2Null(txtConductionSticker))
            If Not (grs.EOF Or grs.BOF) Then
                lblVehicleStatus = ""
                lStatus = grs!ISTATUS
                If lStatus = "A" Then
                    lblVehicleStatus = "***ALLOCATED***"
                ElseIf lStatus = "D" Then
                    lblVehicleStatus = "***DEMO***"
                ElseIf lStatus = "T" Then
                    lblVehicleStatus = "***STOCK TRANSFERED***"
                ElseIf lStatus = "S" And txtInvoicedDate <> "" Then
                    lblVehicleStatus = "***SOLD***"
                ElseIf lStatus = "R" And txtDateRelease <> "" Then
                    lblVehicleStatus = "***RELEASED***"
                End If

            End If
            Set grs = Nothing
        Else
            lblVehicleStatus = ""
        End If

    Else
        If AddingSalesOrder = False Then
            ShowNoRecord
            Select Case MsgBox("There are no Sales Order. Do You Want To Add New Record", vbYesNo Or vbExclamation Or vbDefaultButton1, App.TITLE)
                Case vbYes
                    cmdAdd.Value = True
                Case vbNo
                    Unload Me
            End Select
        End If

    End If
End Sub

Sub UpdateLog()
    Dim SQL                                                           As String
    If PROSPECTID = 0 Then
        SQL = "INSERT INTO CRIS_PROSPECTS ( PROSCODE , "
        SQL = SQL & " PROSPECTTYPE , "
        SQL = SQL & " CUSCDE,"
        SQL = SQL & " SAE,"
        SQL = SQL & " ACCTNAME,"
        SQL = SQL & " TELEPHONE,"
        SQL = SQL & " ADDRESS,"
        SQL = SQL & " MODEL,"
        SQL = SQL & " VARIANT, "
        SQL = SQL & " COLOR,"
        SQL = SQL & " LEADSOURCE,"
        SQL = SQL & " CLASSIFICATION,"
        SQL = SQL & " LOGINITIALINQUIRY,"
        SQL = SQL & " SO_NO,"
        SQL = SQL & " HITCOUNTER) VALUES ( "
        SQL = SQL & N2Str2Null(GenerateCode("CRIS_PROSPECTS", "PROSCODE", "0000000000")) & ","
        SQL = SQL & N2Str2Null("P") & ","
        SQL = SQL & N2Str2Null(CUSCDE) & ","
        SQL = SQL & N2Str2Null(cboSalesAE) & ","
        SQL = SQL & N2Str2Null(txtCusName) & ","
        SQL = SQL & N2Str2Null(txtHomeTelNo) & ","
        SQL = SQL & N2Str2Null(txtHomeAdd) & ","

        SQL = SQL & N2Str2Null(txtModel) & ","
        SQL = SQL & N2Str2Null(txtModelDescription) & ","
        SQL = SQL & N2Str2Null(cboColor) & ","
        SQL = SQL & N2Str2Null("DIRECT SALES") & ","
        SQL = SQL & N2Str2Null("HOT") & ","
        SQL = SQL & N2Str2Null(txtDeyt) & ","
        SQL = SQL & N2Str2Null(txt_SONO) & ","
        SQL = SQL & "1)"
        gconDMIS.Execute SQL
    End If
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME" & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(DEYT) FROM smis_salesorder  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGSO=@DT, HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    gconDMIS.Execute (TSQL)
End Sub

Sub UpdateNetAmount()
    On Error Resume Next
    If AddorEdit = "" Then Exit Sub
    Dim A, C, D, E, F, G, Z1, z2, Z3, Z4, Z5, Z6
    A = NumericVal(txtCL_SalesPrice)
    D = NumericVal(txtCL_Discount)
    Z1 = NumericVal(txtCL_Insurance)
    z2 = NumericVal(txtCL_LTORegFee)
    Z3 = NumericVal(txtCL_Freight)
    Z4 = NumericVal(txtCL_Others)
    Z5 = NumericVal(txtCL_Chattel)
    Z6 = NumericVal(txtCL_DownPayment)
    G = FormatNumber(A - D)
    txtCL_NetSalesPrice = G

    If UCase(cboFinancingTerm) = "FINANCING" Or UCase(cboFinancingTerm = "BANK PO") Then
        txtCL_BalToFinanced = FormatNumber(A - Z6)
    End If

    If UCase(cboFinancingTerm) = "FINANCING" Or UCase(cboFinancingTerm = "BANK PO") Then
        If Z6 > 0 Then
            txtCL_TotalDue = FormatNumber((Z1 + z2 + Z3 + Z4 + Z5) + (Z6)) - D
        Else
            If NumericVal(txtCL_BalToFinanced) = 0 Then
                txtCL_TotalDue = FormatNumber(txtCL_BalToFinanced)
            Else
                txtCL_TotalDue = FormatNumber(G + Z1 + z2 + Z3 + Z4 + Z5 + Z6)
            End If
        End If
    Else
        txtCL_TotalDue = FormatNumber(G + Z1 + z2 + Z3 + Z4 + Z5 + Z6)
    End If

End Sub

Sub checktheCost()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT PURCHPRICE from SMIS_MRRINV_TABLE where IGNKEY = '" & txtConductionSticker & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If NumericVal(txtCL_SalesPrice.Text) < NumericVal(RS!PurchPrice) Then
            MsgBox "Contact Sales Administrator..Your Selling price is less than the Cost Price!", vbExclamation, "WARNING"
            txtCL_SalesPrice = "0.00"
            picsecurity.Visible = True
            loadAccesUsers
        End If
    End If
    Set RS = Nothing
End Sub

Sub loadAccesUsers()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT username from ALL_vw_RAMS_PAccess where loglevel= 'ADM'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    CboAuthorized.Clear

    Do While Not RS.EOF
        CboAuthorized.AddItem Null2String(RS!UserName)
        RS.MoveNext
        AccessFlag = False
    Loop
    Set RS = Nothing
End Sub

Private Sub cmdRefresh_Click()
    If Not (rsSO.EOF Or rsSO.BOF) Then
        rsRefresh
        rsSO.Find ("ID=" & labID)
        StoreMemVars
    End If
End Sub

Private Sub AORCAL_LineAOR(LOR_Fincom As Variant, LOR_Custinfo As Variant, LOR_Addinfo As Variant, LOR_Vehiclesinfo As Variant, LOR_Customerinfo As Variant, LOR_UnitPrice As Variant, LOR_LTO As Variant, LOR_Chattel As Variant, LOR_Insurance As Variant, LOR_TotalUnitCost As Variant, LOR_Discount As Variant, LOR_GrandTotal As Variant, LOR_DownPayment As Variant, LOR_BalToFinance As Variant, LOR_Term As Variant, LOR_MonthlyAmort As Variant, LOR_AOR As Variant, LOR_DownpaymentRate As Variant)
    txtCL_NetMoAmort = LOR_MonthlyAmort
    txtCL_BankTerms = LOR_Term
    cboFinancingCo.Text = LOR_Fincom
    txtCL_SalesPrice = LOR_UnitPrice
    txtCL_DownPayment = LOR_DownPayment
    txtCL_Insurance = LOR_Insurance
    txtCL_LTORegFee = LOR_LTO
    txtCL_BalToFinanced = LOR_BalToFinance
    txtCL_AORRate = LOR_AOR
    txtCL_Chattel = LOR_Chattel
    txtCL_DownpaymentPert = LOR_DownpaymentRate
    On Error Resume Next
    txtGMI.SetFocus
End Sub

Private Sub CboAuthorized_GotFocus()
    CboAuthorized.BackColor = &HFFFFC0
End Sub

Private Sub cboFinancingCo_GotFocus()
    VBComBoBoxDroppedDown cboFinancingCo
End Sub

Private Sub cboFinancingTerm_Click()
    If cboFinancingTerm.ListIndex = -1 Then Exit Sub

    If cboFinancingTerm.ListIndex = 0 Or cboFinancingTerm.ListIndex = 3 Then

        Call ShadeControl(cboFinancingCo, False, "")
        Call ShadeControl(txtCL_BankTerms, False, "0.00")
        Call ShadeControl(txtCL_NetMoAmort, False, "0.00")
        Call ShadeControl(txtCL_AORRate, False, "0.00")
        Call ShadeControl(txtCL_DownPayment, False, "0.00")
        Call ShadeControl(txtCL_Chattel, False, "0.00")
        Call ShadeControl(txtCL_DownpaymentPert, False, "0.00")
        Call ShadeControl(txtCL_BalToFinanced, False, "0.00")
        txtCL_Chattel = "0.00"
    Else
        Call ShadeControl(cboFinancingCo, True)
        Call ShadeControl(cboFinancingCo, True)
        Call ShadeControl(txtCL_BankTerms, True)
        Call ShadeControl(txtCL_NetMoAmort, True)
        Call ShadeControl(txtCL_AORRate, True)
        Call ShadeControl(txtCL_BalToFinanced, True)
        Call ShadeControl(txtCL_DownPayment, True)
        Call ShadeControl(txtCL_Chattel, True)
        Call ShadeControl(txtCL_DownpaymentPert, True)

    End If
    UpdateNetAmount

End Sub

Private Sub cboSalesAE_GotFocus()
    VBComBoBoxDroppedDown cboSalesAE
End Sub

Private Sub cboSalesAE_LostFocus()
    If cboSalesAE <> "" Then
        If SelectSAE(cboSalesAE, cboSalesAE) = False Then
            On Error Resume Next
            cboSalesAE = ""
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "SALES ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    MULTIPLESO = False

    Set EntryPoint = New frmSMIS_Trans_SOEntryPoint
    EntryPoint.Show vbModal
    Unload EntryPoint
    fraCustInfo.Enabled = True
    txtDeyt.Enabled = False
    Set EntryPoint = Nothing
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancel_Click()
    If AddingSalesOrder = True And rsSO.RecordCount = 0 Then
        Unload Me
        Exit Sub
    Else
        AddorEdit = ""
        picSalesOrder.Enabled = False
        fraCustInfo.Enabled = False
        picAdds.Visible = True
        picSaves.Visible = False
        initMemvars
        StoreMemVars
        On Error Resume Next

        picMultipleInventory.Visible = False
    End If
    AccessFlag = False
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "SALES ORDER") = False Then Exit Sub
    Dim rsIGNKEY                                                      As ADODB.Recordset
    Dim strIGNKEY                                                     As String

    On Error GoTo ErrorCode:
    If MsgBox("Do you Want to Cancel this Sales Order ", vbOKCancel + vbInformation, "Confirm Cancellation") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = False
    SQL_STATEMENT = ("UPDate SMIS_SalesOrder  Set SOSTATUS='C' Where ID=" & labID)
    '*********NEW LOG AUDIT
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "SALES ORDER", SQL_STATEMENT, Null2String(labID), "", "SO No:" & txt_SONO, "", ""
    '*********************


    gconDMIS.Execute ("UPDate CRIS_PROSPECTS  Set SO_NO=NULL,LogSO=NULL  Where ProspectID=" & PROSPECTID)
    Set rsIGNKEY = gconDMIS.Execute("SELECT IGNKEY_NO FROM SMIS_SALESORDER WHERE ID=" & labID)
    If Not (rsIGNKEY.EOF Or rsIGNKEY.BOF) Then
        strIGNKEY = Null2String(rsIGNKEY.Collect(0))
        If Not strIGNKEY = "" Then
            gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET PROSPECTID=NULL, VI_NO=NULL, CUSTOMERCODE=NULL,DATERELEASED=null, INVOICEDDATE=null,IStatus='O' , WithProsBuyers='N'  , RELEASED=0 WHERE IGNKEY=" & N2Str2Null(strIGNKEY))
        End If
    End If
    Set rsIGNKEY = Nothing
    LogAudit "C", "SALES ORDER", txt_SONO & " SC:" & cboSalesAE
    rsRefresh
    rsSO.Find ("ID=" & labID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Cancelled", "Record Sucessfully Cancelled", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancelMultiple_Click()
    MULTIPLESO = False
    ShowHidePictureBox2 picMultipleInventory, False
    cmdCancel.Value = True
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picViewVehicles, False, picBottoms
End Sub

Private Sub cmdCloseMultiple_Click()
    cmdCancelMultiple_Click
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SALES ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If IsDate(rsSO!InvoicedDate) = True Then
        MessagePop RecLocekd, "Cannot Delete Record", " Sales Invoice has already been issued for this Sales Order Cannot Delete this Record"
    Else
        If MsgBox(" Confirm Delete" & vbCrLf & " Do you want to Delete this Sales Order ", vbOKCancel + vbExclamation) = vbOK Then
            gconDMIS.Execute ("Delete from SMIS_SalesOrder where id=" & labID)

            UpdateLog
            If FormExist("MainForm") Then
                MainForm.ShowStatus PROSPECTID
            End If
            If Len(txtConductionSticker) > 0 Then
                gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET " _
                                & " STATUS='O'," _
                                & " PROSPECTID=NULL,  CUSTOMERCODE=NULL , Released=0, WithProsBuyers='N' WHERE ignkey=" & N2Str2Null(txtConductionSticker))


            End If
            rsRefresh
            initMemvars
            StoreMemVars
            If FormExist("MainForm") Then
                MainForm.ShowData
            End If
        End If
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "SALES ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "EDIT"
    If IsDate(rsSO!InvoicedDate) = True Then
        MsgBox "Vehicle has already been Invoiced .. " & vbCrLf & " Editing is Limited ", vbInformation
    End If
    txtDeyt.Enabled = False
    picSalesOrder.Enabled = True
    fraCustInfo.Enabled = True

    picAdds.Visible = False: picSaves.Visible = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

    frmSMIS_SearchVehicleSalesOrder.Show
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:
    If rsSO.BOF Then
        ShowFirstRecordMsg
    Else
        rsSO.MoveFirst
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:
    If rsSO.EOF Then
        ShowLastRecordMsg
    Else
        rsSO.MoveLast
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:
    rsSO.MoveNext
    If rsSO.EOF Then
        rsSO.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdok_Click()
    'BTT - 02/09/2008
    'This Routin validate the transaction
    Dim rsALL_vw_RAMS_PAccess                                         As New ADODB.Recordset
    On Error GoTo ErrorCode:
    If txtUserPass.Enabled = False Then Exit Sub
    Dim cnt, COUNTER                                                  As Integer
    With wizVar
        Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset
        rsALL_vw_RAMS_PAccess.Open "select * from ALL_vw_RAMS_PAccess where username = '" & CboAuthorized.Text & "'", gconACCESS, adOpenKeyset

        If Not rsALL_vw_RAMS_PAccess.EOF And Not rsALL_vw_RAMS_PAccess.BOF Then
            If txtUserPass.Text <> .DecryptAccess(rsALL_vw_RAMS_PAccess!userPass) Then
                MsgBox "Invalid Password..Please Check Your Pasword.", vbExclamation, "WARNING"
                txtUserPass.Text = ""
                txtUserPass.BackColor = vbRed
            Else
                MsgBox "Transaction Authorized..", vbInformation, "Confirm"
                AccessFlag = True
                picsecurity.Visible = False
            End If
        Else

            MsgBox "Invalid Password..Please Check Your Pasword.", vbExclamation, "WARNING"

        End If
    End With

    Exit Sub


ErrorCode:
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
    End
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "SALES ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Do you Want to Post this Transaction ", vbYesNo + vbInformation, "Confirm Posting") = vbNo Then Exit Sub
    cmdCancelCO.Enabled = False
    '   gconDMIS.Execute ("UPDate SMIS_SalesOrder  Set SOSTATUS='P' Where ID=" & labID)
    '   gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET ISTATUS='A' WHERE IGNKEY='" & txtConductionSticker & "'")

    SQL_STATEMENT = ("UPDate SMIS_SalesOrder  Set SOSTATUS='P' Where ID=" & labID)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "P", "SALES ORDER", SQL_STATEMENT, Null2String(labID), "", "SO No:" & txt_SONO, "", ""
    '********RESET THE VARIABLE
    SQL_STATEMENT = ""
    '*************************
    SQL_STATEMENT = ("UPDATE SMIS_MRRINV_TABLE SET ISTATUS='A' WHERE IGNKEY='" & txtConductionSticker & "'")
    NEW_LogAudit "E", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtConductionSticker), "ignkey", "SMIS_MRRINV_TABLE"), "", "SO No:" & txt_SONO, "", ""

    LogAudit "P", "SALES ORDER", txt_SONO & "- SC:" & cboSalesAE
    rsRefresh
    rsSO.Find ("ID=" & labID)
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Posted", "Record Sucessfully Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:
    rsSO.MovePrevious
    If rsSO.BOF Then
        rsSO.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "SALES ORDER") = False Then Exit Sub

    If LABALLOWREPRINT <> "" Then
        If AllowReprint("SALES ORDER") = False Then Exit Sub
    End If

    On Error GoTo ErrorCode:
    rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptReleased.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    LoadSignatories ("SALES ORDER")
    rptReleased.Formulas(3) = "APPROVEDBY = '" & ApprovedBy & "'"
    If Null2String(rsSO!SOSTATUS) = "C" Then
        rptReleased.Formulas(1) = "Status = 'CANCELLED'"
    Else
        rptReleased.Formulas(1) = "Status = ''"
    End If

    If rsSO!TERM = "F" Then
        rptReleased.WindowTitle = "Sales Order: Financing"
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "Purchase Agreement.rpt", "{SMIS_SALESORDER.SO_NO}= '" & Trim(txt_SONO) & "'", DMIS_REPORT_Connection, 1
        '****NEW LOG AUDIT
        NEW_LogAudit "V", "SALES ORDER", "", Null2String(labID), "", "SO No:" & txt_SONO, "", ""
        '****************
    Else
        rptReleased.WindowTitle = "Sales Order: Cash"
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "Purchase Agreement Cash.rpt", "{SMIS_SALESORDER.SO_NO}= '" & Trim(txt_SONO) & "'", DMIS_REPORT_Connection, 1
        '****NEW LOG AUDIT
        NEW_LogAudit "V", "SALES ORDER", "", Null2String(labID), "", "SO No:" & txt_SONO, "", ""
        '****************
    End If

    If rptReleased.RecordsPrinted = 1 Then
        gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET PRINTED=1 WHERE ID=" & labID)
        rsSO.Requery
        rsSO.Find ("ID=" & labID)
        StoreMemVars
    End If

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()

    If MULTIPLESO = True Then
        Dim COND
        Select Case Left(cboFinancingTerm.Text, 1)
            Case "C"

                If NumericVal(txtCL_SalesPrice) = 0 Then: COND = COND & "SALES PRICE." & vbCrLf
                If NumericVal(txtCL_NetSalesPrice) = 0 Then: COND = COND & "NET PRICE." & vbCrLf
                If NumericVal(txtCL_Insurance) = 0 Then: COND = COND & "INSURANCE." & vbCrLf
            Case "F", "B"
                If NumericVal(txtCL_AORRate) = 0 Then: COND = "AOR" & vbCrLf
                If NumericVal(txtCL_SalesPrice) = 0 Then: COND = COND & "SALES PRICE." & vbCrLf
                If NumericVal(txtCL_DownPayment) = 0 Then: COND = COND & "DOWN PAYMENT." & vbCrLf
                If NumericVal(txtCL_BalToFinanced) = 0 Then: COND = COND & "BALANCED TO BE FINANCED." & vbCrLf
                If NumericVal(txtCL_NetMoAmort) = 0 Then: COND = COND & "MONTHLY AMORTIZATION." & vbCrLf
                If NumericVal(txtCL_Insurance) = 0 Then: COND = COND & "INSURANCE." & vbCrLf
        End Select
        If Len(COND) <> 0 Then
            If MsgBox("FOLLOWING FILED DOESN'T HAVE VALUE " & vbCrLf & COND & "ARE YOU SURE?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If

        Listview_Loadval lstMultipleInventory.ListItems, gconDMIS.Execute("SELECT  IGNKEY, upper(DESCRIPT), upper(COLOR), ID  FROM SMIS_MRRINV WHERE RELEASED=0 order by descript")
        ShowHidePictureBox2 picMultipleInventory, True, picTops
        Exit Sub
    End If

    On Error GoTo ErrorCode:
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE SO_NO=" & N2Str2Null(txt_SONO)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Number Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsSO!SO_NO)) <> UCase(txt_SONO) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Already Exist"
            Exit Sub
        End If
    End If

    If txtModel = "" Then
        ShowIsRequiredMsg "Vehicle Model"
        On Error Resume Next
        Command1.SetFocus
        Exit Sub
    End If
    If Runvalidation("@R") = False Then: Exit Sub

    SaveSalesOrder

    rsRefresh
    rsSO.Find ("SO_NO='" & txt_SONO & "'")
    SQL_STATEMENT = "UPDATE SMIS_SALESORDER SET " & _
                  " usercode = '" & GetSAECode(cboSalesAE) & "' ," & _
                  " lastupdated   = '" & LOGDATE & "'" & _
                  " WHERE SO_NO = '" & txt_SONO & "'"

    gconDMIS.Execute (SQL_STATEMENT)
    '*******NEW LOG AUDIT
    NEW_LogAudit "EE", "SALES ORDER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txt_SONO), "SO_NO", "SMIS_SALESORDER"), "", "SO_No:" & txt_SONO, "", ""
    '********************
    UpdateLog
    On Error GoTo CANCELLINE:
CANCELLINE:
    cmdCancel.Value = True
    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If
    AccessFlag = False

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSelectMultiple_Click()
    picMultipleInventory.Visible = False
    If lstMultipleInventory.SelectedItem Is Nothing Then Exit Sub
    Dim xxCustName                                                    As String
    Dim xxDeyt                                                        As String
    Dim xxHomeTelNo                                                   As String
    Dim xxHomeAddress                                                 As String
    Dim xxOfficeAdd                                                   As String
    Dim xxOfficeTelNo                                                 As String
    Dim xxBirthDate                                                   As String
    Dim xxSpouse                                                      As String
    Dim xxPerson                                                      As String
    Dim xxPosisyon                                                    As String
    Dim xxTIN                                                         As String
    Dim xxCTCNo                                                       As String
    Dim xxIssuedAt                                                    As String
    Dim xxIssuedOn                                                    As String
    Dim xxmodel                                                       As String
    Dim xxProdNo                                                      As String
    Dim xxConductionSticker                                           As String
    Dim xxEngineNo                                                    As String
    Dim xxFrameNo                                                     As String
    Dim xxColor                                                       As String
    Dim xxType                                                        As String
    Dim xxTerm                                                        As String
    Dim xxFinancingCo                                                 As String
    Dim xxSalesAE                                                     As String
    Dim xx_SalesPrice                                                 As String
    Dim xx_NetSalesPrice                                              As String
    Dim xx_DownPayment                                                As String
    Dim xx_BalToFinanced                                              As Double
    Dim xxAdditionalInfo                                              As String
    Dim xx_GMI                                                        As String
    Dim xx_RPPD                                                       As String
    Dim xx_NetMoAmort                                                 As String
    Dim xx_Insurance                                                  As String
    Dim xx_LTORegFee                                                  As String
    Dim xx_Freight                                                    As Double
    Dim xxVinNo                                                       As String
    Dim xxModelDescript                                               As String
    Dim xxOthersDesc                                                  As String
    Dim xxxOthers                                                     As String
    Dim rsIGNKEY                                                      As ADODB.Recordset
    Dim strIGNKEY                                                     As String
    Dim Item                                                          As ListItem
    Dim rsMRRINV                                                      As ADODB.Recordset
    Dim RSPROSPECTID                                                  As ADODB.Recordset
    Dim SQL
    For Each Item In lstMultipleInventory.ListItems
        If Item.Checked = True Then

            xxCustName = N2Str2Null(txtCusName)
            xxDeyt = N2Str2Null(txtDeyt)
            xxHomeTelNo = N2Str2Null(txtHomeTelNo)
            xxHomeAddress = N2Str2Null(txtHomeAdd)
            xxOfficeAdd = N2Str2Null(txtOfficeAdd)
            xxOfficeTelNo = N2Str2Null(txtOfficeTelNo)
            xxBirthDate = N2Str2Null(txtBirthDate)
            xxSpouse = N2Str2Null(txtSpouse)
            xxPerson = N2Str2Null(txtPerson)
            xxPosisyon = N2Str2Null(txtPosisyon)
            xxTIN = N2Str2Null(txtTIN)
            xxCTCNo = N2Str2Null(txtCTCNo)
            xxIssuedAt = N2Str2Null(txtIssuedAt)
            xxIssuedOn = N2Str2Null(txtIssuedOn)


            If opt1st.Value = True Then
                xxType = "'1st'"
            ElseIf optRPL.Value = True Then
                xxType = "'RPL'"
            ElseIf optADDL.Value = True Then
                xxType = "'ADDL'"
            ElseIf optTRI.Value = True Then
                xxType = "'TRI'"
            End If

            If Left(cboFinancingTerm.Text, 1) = "C" Then
                xxTerm = "'COD'"
            ElseIf Left(cboFinancingTerm.Text, 1) = "F" Then
                xxTerm = "'F'"
            ElseIf Left(cboFinancingTerm.Text, 1) = "B" Then
                xxTerm = "'BPO'"
            End If

            xxFinancingCo = N2Str2Null(cboFinancingCo)
            xxSalesAE = N2Str2Null(cboSalesAE)
            xx_SalesPrice = NumericVal(txtCL_SalesPrice)
            xx_NetSalesPrice = NumericVal(txtCL_SalesPrice)
            xx_DownPayment = NumericVal(txtCL_DownPayment)
            xx_BalToFinanced = NumericVal(txtCL_BalToFinanced)
            xxAdditionalInfo = N2Str2Null(txtAdditionalInfo)
            xx_GMI = NumericVal(txtGMI)
            xx_RPPD = NumericVal(txtRPPD)
            xx_NetMoAmort = NumericVal(txtCL_NetMoAmort)
            xx_Insurance = NumericVal(txtCL_Insurance)
            xx_LTORegFee = NumericVal(txtCL_LTORegFee)
            xx_Freight = NumericVal(txtCL_Freight)
            xxSONO = N2Str2Null(GenerateCode("SMIS_SalesOrder", "SO_NO", "000000"))
            xxOthersDesc = N2Str2Null(txtOthersDesc)
            xxxOthers = FormatNumber(NumericVal(txtCL_Others))

            '
            '                        Dim lng                     As Integer
            '                        lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE SO_NO=" & N2Str2Null(txt_SONO)).Fields(0).Value
            '                        If AddorEdit = "ADD" Then
            '                            If lng >= 1 Then
            '                                MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Number Already Exist"
            '                                Exit Sub
            '                            End If
            '                        Else
            '                            If lng >= 1 And UCase(Null2String(rsSO!so_no)) <> UCase(txt_SONO) Then
            '                                MessagePop RecSaveWarning, "Duplicate Record", "Vehicle Sales Order Already Exist"
            '                                Exit Sub
            '                            End If
            '                        End If

            Set rsMRRINV = gconDMIS.Execute("Select * from SMIS_MRRINV WHERE ID  =" & Item.ListSubItems(3))
            If Not rsMRRINV.EOF Or Not rsMRRINV.BOF Then
                ''''''''''''AUTOS
                xxVinNo = N2Str2Null(rsMRRINV!VINO)
                xxModelDescript = N2Str2Null(rsMRRINV!DESCRIPT)
                xxmodel = N2Str2Null(rsMRRINV!Model)
                xxProdNo = N2Str2Null(rsMRRINV!prodno)
                xxConductionSticker = N2Str2Null(rsMRRINV!ignkey)
                xxEngineNo = N2Str2Null(rsMRRINV!EngineNo)
                xxFrameNo = N2Str2Null(rsMRRINV!frameno)
                xxColor = N2Str2Null(rsMRRINV!Color)
            End If
            gconDMIS.Execute ("INSERT INTO SMIS_SALESORDER " & _
                              "(DISCOUNT, CUSTNAME, PROSPECTID, SO_NO,CODE,DEYT,HOMETELNO,HOMEADDRESS,OFFICETELNO,OFFICEADD,BIRTHDATE,SPOUSE,PERSON,POSISYON,TIN,CTCNO," & _
                              "ISSUEDAT,ISSUEDON,MODEL, ModelDescription,PRODNO,IGNKEY_NO,ENGINENO,FRAMENO,COLOR,TYPE,TERM,FINANCINGCO,SALESAE,SALESPRICE,NETSALESPRICE," & _
                              "DOWNPAYMENT,BALTOFINANCED,ADDITIONALINFO,GMI,RPPD,NETMOAMORT,INSURANCE,LTOREGFEE,FREIGHT, AOR,MONTHSAMORT, Vino, OthersDesc, Others, Total, CHMOFEE,DOWNPAYMENTRATE,SOSTATUS)" & _
                            " VALUES (" & NumericVal(txtCL_Discount) & " ," & xxCustName & " , " & PROSPECTID & " , " & xxSONO & "," & N2Str2Null(CUSCDE) & ", " & xxDeyt & ", " & xxHomeTelNo & ", " & xxHomeAddress & ", " & xxOfficeTelNo & ", " & xxOfficeAdd & ", " & xxBirthDate & ", " & xxSpouse & ", " & xxPerson & ", " & xxPosisyon & ", " & xxTIN & _
                              "," & xxCTCNo & ", " & xxIssuedAt & ", " & xxIssuedOn & ", " & xxmodel & "," & xxModelDescript & ", " & xxProdNo & ", " & xxConductionSticker & ", " & xxEngineNo & ", " & xxFrameNo & ", " & xxColor & ", " & xxType & ", " & xxTerm & ", " & xxFinancingCo & ", " & xxSalesAE & _
                              "," & xx_SalesPrice & ", " & xx_NetSalesPrice & ", " & xx_DownPayment & ", " & xx_BalToFinanced & ", " & xxAdditionalInfo & ", " & xx_GMI & _
                              "," & xx_RPPD & ", " & xx_NetMoAmort & ", " & xx_Insurance & ", " & xx_LTORegFee & ", " & xx_Freight & "," & NumericVal(txtCL_AORRate) & "," & NumericVal(txtCL_BankTerms) & "," & xxVinNo & "," & xxOthersDesc & "," & xxxOthers & "," & NumericVal(txtCL_TotalDue) & "," & NumericVal(txtCL_Chattel) & "," & NumericVal(txtCL_DownpaymentPert) & " ,'')")

            gconDMIS.Execute ("Update CRIS_Prospects Set LOGSO=getdate() , SO_NO=" & N2Str2Null(xxSONO) & " , STATUS='O' where ProspectID=" & PROSPECTID)

            gconDMIS.Execute ("UPDATE SMIS_MRRINV SET " _
                            & " ISTATUS='A'," _
                            & " PROSPECTID=" & PROSPECTID & "," _
                            & " ProspectCounter=ISNULL(ProspectCounter,0) + 1 ," _
                            & " Released=0 ," _
                            & " WithProsBuyers='Y' ," _
                            & " CUSTOMERCODE=" & N2Str2Null(CUSCDE) _
                            & " WHERE ignkey=" & xxConductionSticker)

            If PROSPECTID = 0 Then
                SQL = "INSERT INTO CRIS_PROSPECTS ( PROSCODE , "
                SQL = SQL & " PROSPECTTYPE , "
                SQL = SQL & " CUSCDE,"
                SQL = SQL & " SAE,"
                SQL = SQL & " ACCTNAME,"
                SQL = SQL & " TELEPHONE,"
                SQL = SQL & " ADDRESS,"
                SQL = SQL & " MODEL,"
                SQL = SQL & " VARIANT, "
                SQL = SQL & " COLOR,"
                SQL = SQL & " LEADSOURCE,"
                SQL = SQL & " CLASSIFICATION,"
                SQL = SQL & " LOGINITIALINQUIRY,"
                SQL = SQL & " SO_NO,"
                SQL = SQL & " HITCOUNTER,STATUS) VALUES ( "
                SQL = SQL & N2Str2Null(GenerateCode("CRIS_PROSPECTS", "PROSCODE", "0000000000")) & ","
                SQL = SQL & N2Str2Null("P") & ","
                SQL = SQL & N2Str2Null(CUSCDE) & ","
                SQL = SQL & N2Str2Null(cboSalesAE) & ","
                SQL = SQL & N2Str2Null(txtCusName) & ","
                SQL = SQL & N2Str2Null(txtHomeTelNo) & ","
                SQL = SQL & N2Str2Null(txtHomeAdd) & ","

                SQL = SQL & xxmodel & ","
                SQL = SQL & xxModelDescript & ","
                SQL = SQL & xxColor & ","
                SQL = SQL & N2Str2Null("DIRECT SALES") & ","
                SQL = SQL & N2Str2Null("HOT") & ","
                SQL = SQL & N2Str2Null(txtDeyt) & ","
                SQL = SQL & N2Str2Null(xxSONO) & ","
                SQL = SQL & "1, 'O')"

                gconDMIS.Execute SQL

                Set RSPROSPECTID = gconDMIS.Execute("SELECT PROSPECTID FROM CRIS_PROSPECTS WHERE CUSCDE=" & N2Str2Null(CUSCDE))
                If Not RSPROSPECTID.EOF Or RSPROSPECTID.BOF Then
                    PROSPECTID = N2Str2IntZero(RSPROSPECTID(0).Value)
                    Set RSPROSPECTID = Nothing

                End If
            End If
            Dim TSQL                                                  As String
            TSQL = " DECLARE @DT DATETIME" & vbCrLf
            TSQL = TSQL & " SELECT @DT=MAX(DEYT) FROM smis_salesorder  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
            TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
            TSQL = TSQL & " BEGIN " & vbCrLf
            TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGSO=@DT, HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
            TSQL = TSQL & " End " & vbCrLf
            gconDMIS.Execute (TSQL)


        End If
    Next
    MULTIPLESO = False
    rsRefresh
    cmdCancel.Value = True
    ShowHidePictureBox2 picMultipleInventory, False
    Command1.Enabled = True

End Sub

Private Sub cmdSelectViewVehicles_Click()
    If lvViewVehicles.SelectedRows(0).GroupRow = True Then: Exit Sub



    Dim rsCus                                                         As ADODB.Recordset
    Dim CUSTNAMEX
    If optInventory.Value = True Then
        With lvViewVehicles.SelectedRows.Row(0)

            If Null2String(.Record(10).Value) = "A" Then
                Set rsCus = gconDMIS.Execute("SELECT count(*) FROM SMIS_SALESORDER WHERE IGNKEY_NO='" & lvViewVehicles.SelectedRows.Row(0).Record(2).Value & "' AND SO_NO<>'" & txt_SONO & "' AND ISNULL(STATUS,'') <>'C' ")
                If Not rsCus(0).Value = 0 Then
                    Screen.MousePointer = 0
                    MsgBox "Current Vehicle Has Been Allocated to Customer. " & vbCrLf & CUSTNAMEX & vbCrLf & "Please Cancel Sales Order For " & vbCrLf & txtCusName & "! in order to Re-allocated.", vbCritical
                    Exit Sub
                End If
            End If


            txtModel = Null2String(.Record(0).Value)
            txtModelDescription = Null2String(.Record(1).Value)
            txtConductionSticker.Text = Null2String(.Record(2).Value)
            txtProdNo.Text = Null2String(.Record(3).Value)
            txtEngineNo = Null2String(.Record(4).Value)
            txtFrameNumber = Null2String(.Record(5).Value)
            txtVinNumber = Null2String(.Record(6).Value)
            cboColor = Null2String(.Record(7).Value)
            txtCL_LTORegFee = NumericVal(.Record(8).Value)
            txtCL_Freight = NumericVal(.Record(9).Value)
            Dim rsPrice As ADODB.Recordset
            'UNITCOST IS SRP ITS A TYPO ERROR  IN DATABASE
            Set rsPrice = gconDMIS.Execute("SELECT UNITCOST FROM ALL_MODEL WHERE DESCRIPT=" & N2Str2Null(txtModelDescription))
            If Not (rsPrice.EOF Or rsPrice.BOF) Then
            txtCL_SalesPrice = FormatNumber(N2Str2IntZero(rsPrice!unitcost))
            
            End If

        End With
    Else
        With lvViewVehicles.SelectedRows.Row(0)
            txtModel = Null2String(.Record(0).Value)
            txtModelDescription = Null2String(.Record(1).Value)
            cboColor.Enabled = True
            txtConductionSticker.Text = vbNullString
            txtProdNo.Text = vbNullString
            txtEngineNo = vbNullString
            txtFrameNumber = vbNullString
            txtVinNumber = vbNullString
        End With
    End If

    ShowHidePictureBox2 picViewVehicles, False, picBottoms
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "SALES ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Do you Want to Unpost this Sales Order ", vbYesNo + vbInformation, "Confirm un-posting") = vbNo Then Exit Sub
    cmdCancelCO.Enabled = True

    SQL_STATEMENT = ("UPDate SMIS_SalesOrder  Set SOSTATUS='U' Where ID=" & labID)

    '*****NEW LOG AUDIT
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "U", "SALES ORDER", SQL_STATEMENT, Null2String(labID), "", "SO No:" & txt_SONO, "", ""
    '*****************

    rsRefresh
    rsSO.Find ("ID=" & labID)
    LogAudit "U", "SALES ORDER", txt_SONO & "- SC:" & cboSalesAE
    StoreMemVars
    MessagePop RecSaveOk, "Transaction Unposted", "Record Sucessfully Un-Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    If optInventory.Value = True Then
        optInventory_Click
    Else
        optList_Click
    End If
    If MULTIPLESO = True Then
        optInventory.Enabled = False
        optList.Value = True
    Else
        optInventory.Enabled = True
    End If
    ShowHidePictureBox2 picViewVehicles, True, picBottoms
    On Error Resume Next
    txtFilterViewVehicles.SetFocus
End Sub

Private Sub Command2_Click()
    cmdAdd.Value = True
    MULTIPLESO = True
    fraVehilcesInfo.Enabled = False

End Sub

Private Sub Command3_Click()
    'If AddorEdit = "EDIT" Then
    If Function_Access(LOGID, "ACESS_SYSTEM", "SALES ORDER") = False Then Exit Sub
    txtDeyt.Enabled = True: txtDeyt.SetFocus
    'End If
End Sub

Private Sub Command4_Click()
    picsecurity.Visible = False
    txtUserPass.Text = ""
End Sub

Private Sub Command5_Click()
    picsecurity.Visible = False
    txtUserPass.Text = ""
End Sub

Private Sub EntryPoint_NothingSelected()
    AddorEdit = ""
    picSalesOrder.Enabled = False
    fraCustInfo.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    If (rsSO.EOF Or rsSO.BOF) Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then

    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES ORDER)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labID), "SALES ORDER")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    picSalesOrder.Enabled = False
    fraCustInfo.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    picsecurity.Visible = False
    initMemvars
    InitCbo

     If LTrim(RTrim(LOGCODE)) = "NET" Then
        txtOldCS.Visible = True
    Else
        txtOldCS.Visible = False
    End If
    
    rsRefresh
    
    StoreMemVars
    
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AddorEdit = vbNullString
    AddingSalesOrder = False
    Set rsS_Model = Nothing
    Set ctl = Nothing
    xxSONO = vbNullString
    AcctName = vbNullString
    PROSPECTID = 0
    CUSCDE = vbNullString
    ProfileType = vbNullString
    vID = 0
    VCODE = vbNullString
    VDescript = vbNullString
End Sub

Private Sub lstMultipleInventory_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        cmdSelectMultiple.Enabled = True
        Exit Sub
    End If
    For Each Item In lstMultipleInventory.ListItems
        If Item.Checked = True Then
            cmdSelectMultiple.Enabled = True
            Exit Sub
        End If
    Next
    cmdSelectMultiple.Enabled = False
End Sub

Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If cmdSelectViewVehicles.Enabled = True Then
            cmdSelectViewVehicles_Click
        End If
    End If
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdSelectViewVehicles_Click
End Sub

Private Sub lvViewVehicles_SelectionChanged()
    cmdSelectViewVehicles.Enabled = True
End Sub

Private Sub optInventory_Click()
    If optInventory.Value = True Then
        ReportControlAddColumnHeader lvViewVehicles, "MODEL, DESCRIPTION, CS #"
        ReportControlPaintManager lvViewVehicles
        'lvViewVehicles.GroupsOrder.Add lvViewVehicles.Columns(0)
        lvViewVehicles.Columns(0).Visible = False
        flex_FillReportView gconDMIS.Execute("select  MODEL, Descript, ignkey, prodno , EngineNo ,FrameNo, Vino,color , lto , pullout , ISTATUS ,CUSTOMERCODE from SMIS_MRRINV_TABLE where STATUS='P' and Released=0 AND (ISTATUS='O' OR ISTATUS='A') ORDER BY MODEL"), lvViewVehicles
        cap3.Caption = "Search for Vehicles In Stock"
    End If

End Sub

Private Sub optList_Click()
    If optList.Value = True Then
        ReportControlAddColumnHeader lvViewVehicles, "Model, Descript"
        ReportControlPaintManager lvViewVehicles
        lvViewVehicles.GroupsOrder.Add lvViewVehicles.Columns(0)
        lvViewVehicles.Columns(0).Visible = False
        flex_FillReportView gconDMIS.Execute("select Model, Descript from ALL_MODEL WHERE LEN(CODE)>0  ORDER BY CODE  "), lvViewVehicles
        cap3.Caption = "Search for Vehicles From List"
    End If
End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Timer1_Timer()
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If
        End If
    Next
    Timer1.Enabled = False
End Sub

Private Sub tmBlink_Timer()
    If lblStatus.Caption <> "" Then
        If lblStatus.Visible = True Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If
    End If
    If lblVehicleStatus.Caption <> "" Then
        If lblVehicleStatus.Visible = True Then
            lblVehicleStatus.Visible = False
        Else
            lblVehicleStatus.Visible = True
        End If
    End If
    If lblSalesStatus.Caption <> "" Then
        If lblSalesStatus.Visible = True Then
            lblSalesStatus.Visible = False
        Else
            lblSalesStatus.Visible = True
        End If
    End If
End Sub

Private Sub txt_SONO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txt_SONO_LostFocus()
    txt_SONO = Format(txt_SONO, "000000")
End Sub

Private Sub txtAdditionalInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtBirthDate_Validate(Cancel As Boolean)
    If IsDate(txtBirthDate) = False Then: txtBirthDate = vbNullString
End Sub

Private Sub txtCL_AORRate_Change()
    If AddorEdit = "ADD" Then
        txtCL_NetMoAmort = AORVALUE(NumericVal(txtCL_BalToFinanced), NumericVal(txtCL_AORRate), NumericVal(txtCL_BankTerms))
    End If
End Sub

Private Sub txtCL_AORRate_GotFocus()
    If NumericVal(txtCL_AORRate.Text) <= 0 Then txtCL_AORRate = ""
End Sub

Private Sub txtCL_BalToFinanced_Change()
    If AddorEdit = "ADD" Then
        txtCL_NetMoAmort = AORVALUE(NumericVal(txtCL_BalToFinanced), NumericVal(txtCL_AORRate), NumericVal(txtCL_BankTerms))
    End If
    UpdateNetAmount
End Sub

Private Sub txtCL_BankTerms_Change()
    If AddorEdit = "ADD" Then
        txtCL_NetMoAmort = AORVALUE(NumericVal(txtCL_BalToFinanced), NumericVal(txtCL_AORRate), NumericVal(txtCL_BankTerms))
    End If
End Sub

Private Sub txtCL_BankTerms_GotFocus()
    If NumericVal(txtCL_BankTerms.Text) <= 0 Then txtCL_BankTerms = ""
End Sub

Private Sub txtCL_BankTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Chattel_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_Chattel_GotFocus()
    If NumericVal(txtCL_Chattel.Text) <= 0 Then txtCL_Chattel = ""
End Sub

Private Sub txtCL_Chattel_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Chattel_LostFocus()
    If NumericVal(txtCL_Chattel) <= 0 Then txtCL_Chattel = "0.00"
    txtCL_Chattel = FormatNumber(NumericVal(txtCL_Chattel))
End Sub

Private Sub txtCL_Discount_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_Discount_GotFocus()
    If NumericVal(txtCL_Discount.Text) <= 0 Then txtCL_Discount = ""
End Sub

Private Sub txtCL_Discount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Discount_LostFocus()
    If NumericVal(txtCL_Discount) <= 0 Then txtCL_Discount = "0.00"
    txtCL_Discount = FormatNumber(NumericVal(txtCL_Discount))
End Sub

Private Sub txtCL_DownPayment_change()
    On Error GoTo ADDER:
    If ComputebyPert = False And AddorEdit <> "" Then
        txtCL_DownpaymentPert = (NumericVal(txtCL_DownPayment) / NumericVal(txtCL_NetSalesPrice)) * 100
    End If
    UpdateNetAmount
    Exit Sub
ADDER:
    If Err.Number = 6 Then
        Err.Clear
        Exit Sub
    End If

End Sub

Private Sub txtCL_DownPayment_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_DownPayment_LostFocus()
    txtCL_DownPayment = FormatNumber(NumericVal(txtCL_DownPayment))
End Sub

Private Sub txtCL_DownPayment_Validate(Cancel As Boolean)
    If IsNumeric(txtCL_DownPayment) = True Then
        txtCL_DownPayment = FormatNumber(txtCL_DownPayment.Text)
        UpdateNetAmount
    End If
End Sub

Private Sub txtCL_DownpaymentPert_Change()
    If Not txtCL_DownpaymentPert = "" Then
        If NumericVal(txtCL_DownpaymentPert) > 100 Then
            txtCL_DownpaymentPert = 100
        ElseIf NumericVal(txtCL_DownpaymentPert) <= 0 Then
            txtCL_DownpaymentPert = 0
        End If
    End If
    If ComputebyPert = True And AddorEdit <> "" Then
        txtCL_DownPayment = FormatNumber(NumericVal(txtCL_NetSalesPrice) * (NumericVal(txtCL_DownpaymentPert) / 100))
    End If
    UpdateNetAmount
End Sub

Private Sub txtCL_DownpaymentPert_GotFocus()
    If NumericVal(txtCL_DownpaymentPert.Text) <= 0 Then txtCL_DownpaymentPert = ""
    ComputebyPert = True
End Sub

Private Sub txtCL_DownpaymentPert_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_DownpaymentPert_LostFocus()
    ComputebyPert = False
    txtCL_DownPayment = FormatNumber(NumericVal(txtCL_DownPayment))
End Sub

Private Sub txtCL_Freight_change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_Freight_GotFocus()
    If NumericVal(txtCL_Freight.Text) <= 0 Then txtCL_Freight = ""
End Sub

Private Sub txtCL_Freight_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Freight_LostfOCUS()
    If NumericVal(txtCL_Freight) <= 0 Then txtCL_Freight = "0.00"
    txtCL_Freight = FormatNumber(NumericVal(txtCL_Freight))
End Sub

Private Sub txtCL_Insurance_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_Insurance_GotFocus()

    If NumericVal(txtCL_Insurance.Text) <= 0 Then txtCL_Insurance = ""
End Sub

Private Sub txtCL_Insurance_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Insurance_LostFocus()
    If NumericVal(txtCL_Insurance) <= 0 Then txtCL_Insurance = "0.00"
    txtCL_Insurance = FormatNumber(NumericVal(txtCL_Insurance))
End Sub

Private Sub txtCL_LTORegFee_change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_LTORegFee_GotFocus()
    If NumericVal(txtCL_LTORegFee.Text) <= 0 Then txtCL_LTORegFee = ""
End Sub

Private Sub txtCL_LTORegFee_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_LTORegFee_LostFocus()
    If NumericVal(txtCL_LTORegFee) <= 0 Then txtCL_Insurance = "0.00"
    txtCL_LTORegFee = FormatNumber(NumericVal(txtCL_LTORegFee))
End Sub

Private Sub txtCL_NetMoAmort_Change()
    If AddorEdit = "ADD" Then
        txtGMI = NumericVal(txtCL_NetMoAmort) * 1.03
        txtRPPD = NumericVal(txtGMI) - NumericVal(txtCL_NetMoAmort)
    End If
End Sub

Private Sub txtCL_NetMoAmort_GotFocus()
    If NumericVal(txtCL_NetMoAmort.Text) <= 0 Then txtCL_NetMoAmort = ""
End Sub

Private Sub txtCL_NetMoAmort_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_NetSalesPrice_change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_NetSalesPrice_GotFocus()
    If NumericVal(txtCL_NetSalesPrice.Text) <= 0 Then txtCL_NetSalesPrice = ""
End Sub

Private Sub txtCL_NetSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_NetSalesPrice_LostFocus()
    txtCL_NetSalesPrice = FormatNumber(NumericVal(txtCL_NetSalesPrice), 2, vbTrue, vbTrue)
    UpdateNetAmount
End Sub

Private Sub txtCL_Others_change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_Others_GotFocus()
    If NumericVal(txtCL_Others.Text) <= 0 Then txtCL_Others = ""
End Sub

Private Sub txtCL_Others_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_Others_LostFocus()
    If NumericVal(txtCL_Others) <= 0 Then txtCL_Others = "0.00"
    txtCL_Others = FormatNumber(NumericVal(txtCL_Others), 2, vbTrue, vbTrue)
End Sub

Private Sub txtCL_SalesPrice_change()
    If AddorEdit = "" Then Exit Sub
    UpdateNetAmount
End Sub

Private Sub txtCL_SalesPrice_GotFocus()
    If NumericVal(txtCL_SalesPrice.Text) <= 0 Then txtCL_SalesPrice = ""
End Sub

Private Sub txtCL_SalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCL_SalesPrice_LostFocus()
    If NumericVal(txtCL_SalesPrice) <= 0 Then txtCL_SalesPrice = "0.00"
    txtCL_SalesPrice = FormatNumber(NumericVal(txtCL_SalesPrice))
    UpdateNetAmount
    If AccessFlag = True Then Exit Sub
    checktheCost

End Sub

Private Sub txtConductionSticker_Change()
    If txtConductionSticker.Text = "" Then
        cboColor.Enabled = True
    Else
        cboColor.Enabled = False
    End If
End Sub

Private Sub txtCTCNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDownpayment_GotFocus()
    If NumericVal(txtCL_DownPayment.Text) <= 0 Then txtCL_DownPayment = ""
End Sub

Private Sub txtFilterViewVehicles_Change()
    lvViewVehicles.FilterText = txtFilterViewVehicles.Text
    lvViewVehicles.Populate
    cmdSelectViewVehicles.Enabled = IIf(lvViewVehicles.Rows.Count = 0, False, True)
End Sub

Private Sub txtGMI_GotFocus()
    If NumericVal(txtGMI.Text) <= 0 Then txtGMI = ""
End Sub

Private Sub txtGMI_LostFocus()
    txtGMI = FormatNumber(NumericVal(txtGMI))
End Sub

Private Sub txtIssuedAt_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtIssuedOn_LostFocus()
    If Trim(txtIssuedOn) <> "" Then
        If IsDate(txtIssuedOn) = False Then
            txtIssuedOn = ""
        End If
    End If
End Sub

Private Sub txtRPPD_GotFocus()
    If NumericVal(txtRPPD.Text) <= 0 Then txtRPPD = ""
End Sub

Private Sub txtRPPD_LostFocus()
    txtRPPD = FormatNumber(NumericVal(txtRPPD))
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtUserPass_GotFocus()
    txtUserPass.BackColor = &HFFFFC0
End Sub

