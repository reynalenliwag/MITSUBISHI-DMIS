VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmSMIS_Mis_AOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Calculator"
   ClientHeight    =   8415
   ClientLeft      =   2145
   ClientTop       =   315
   ClientWidth     =   11745
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AOR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11745
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLow 
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   0
      ScaleHeight     =   5865
      ScaleWidth      =   11745
      TabIndex        =   15
      Top             =   2550
      Width           =   11745
      Begin VB.CommandButton Command1 
         Caption         =   "Print AOR Computation Sheet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8370
         TabIndex        =   56
         ToolTipText     =   "Print AOR Computation"
         Top             =   5340
         Width           =   2550
      End
      Begin VB.CommandButton Command3 
         Height          =   585
         Left            =   10980
         MouseIcon       =   "AOR.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "AOR.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Exit"
         Top             =   5250
         Width           =   675
      End
      Begin VB.Frame Frame1 
         Height          =   5895
         Left            =   60
         TabIndex        =   17
         Top             =   -90
         Width           =   4170
         Begin VB.CommandButton Command4 
            Caption         =   "Compute"
            Height          =   465
            Left            =   2430
            TabIndex        =   42
            Top             =   4890
            Width           =   1695
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   32
            Top             =   2810
            Width           =   2415
         End
         Begin VB.TextBox txtDiscountPert 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   3225
            Width           =   540
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
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
            Left            =   2205
            MaxLength       =   10
            TabIndex        =   34
            Top             =   3225
            Width           =   1890
         End
         Begin VB.TextBox txtDownpaymentPert 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
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
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1170
            Width           =   540
         End
         Begin VB.ComboBox cboFinCom 
            Height          =   345
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   375
            Width           =   3990
         End
         Begin VB.TextBox txtUnitPrice 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   21
            Top             =   750
            Width           =   2445
         End
         Begin VB.TextBox txtDownPayment 
            Alignment       =   1  'Right Justify
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
            Left            =   2250
            TabIndex        =   24
            Top             =   1170
            Width           =   1860
         End
         Begin VB.TextBox txtLTO 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   26
            Top             =   1574
            Width           =   2415
         End
         Begin VB.TextBox txtChattelMorgage 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1986
            Width           =   2415
         End
         Begin VB.TextBox txtInsurance 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   30
            Top             =   2398
            Width           =   2415
         End
         Begin VB.TextBox txtNetAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3630
            Width           =   2415
         End
         Begin VB.TextBox txtAmountDue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   4080
            Width           =   2415
         End
         Begin VB.TextBox txtFinancingBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   405
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   5400
            Width           =   4035
         End
         Begin VB.OptionButton optRural 
            Caption         =   "OMA"
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
            Left            =   705
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   4905
            Width           =   675
         End
         Begin VB.OptionButton optUrban 
            Caption         =   "AOR"
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
            Left            =   1530
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   4920
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Other(s)"
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
            Left            =   105
            TabIndex        =   31
            Top             =   2880
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Financing Company"
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
            Left            =   75
            TabIndex        =   18
            Top             =   135
            Width           =   1650
         End
         Begin VB.Label lblPrincipal 
            AutoSize        =   -1  'True
            Caption         =   "Unit Price"
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
            Left            =   105
            TabIndex        =   20
            Top             =   825
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "Less: Downpayment"
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
            Height          =   435
            Left            =   105
            TabIndex        =   22
            Top             =   1140
            Width           =   1500
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "LTO Registration"
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
            Left            =   105
            TabIndex        =   25
            Top             =   1620
            Width           =   1425
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Chattel Mortgage"
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
            Left            =   105
            TabIndex        =   27
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Less :Discount"
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
            Height          =   225
            Left            =   105
            TabIndex        =   35
            Top             =   3375
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Net Amount"
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
            Left            =   105
            TabIndex        =   36
            ToolTipText     =   "Unit Price Discount"
            Top             =   3780
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Amount To Be Financed"
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
            Left            =   105
            TabIndex        =   44
            Top             =   5130
            Width           =   1995
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount Due"
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
            Left            =   105
            TabIndex        =   38
            ToolTipText     =   "Total Downpayment  + LTO + Chattel + Insurance + Others - Discount"
            Top             =   4200
            Width           =   1500
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Insurance"
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
            Left            =   105
            TabIndex        =   29
            Top             =   2490
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Financing Computation Rate Option"
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
            Left            =   105
            TabIndex        =   40
            Top             =   4635
            Width           =   2970
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5670
         Left            =   4170
         ScaleHeight     =   5670
         ScaleWidth      =   7905
         TabIndex        =   16
         Top             =   -420
         Width           =   7905
         Begin FlexCell.Grid Grid1 
            Height          =   5265
            Left            =   120
            TabIndex        =   59
            Top             =   450
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   9287
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12632256
            Rows            =   30
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   585
         Left            =   4320
         TabIndex        =   55
         Top             =   5250
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   1032
         Picture         =   "AOR.frx":07C2
         ForeColor       =   32768
         BorderStyle     =   2
         BarPicture      =   "AOR.frx":07DE
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Segments        =   -1  'True
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox picSchedule 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5850
         Left            =   4260
         ScaleHeight     =   5820
         ScaleWidth      =   7425
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   7455
         Begin MSFlexGridLib.MSFlexGrid gridSchedule 
            Height          =   4800
            Left            =   0
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   0
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   8467
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColorBkg    =   -2147483628
            Redraw          =   -1  'True
            FocusRect       =   2
            HighLight       =   2
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "Term No|Monthly Amortization  |Interest                 |Principle                         |Balance                   "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Back"
            Height          =   405
            Left            =   5910
            TabIndex        =   50
            Top             =   5250
            Width           =   1455
         End
         Begin VB.TextBox txtTotalInterest 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   4965
            Width           =   2190
         End
         Begin VB.TextBox txtTotalPayment 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   5370
            Width           =   2190
         End
         Begin VB.CommandButton cmdPrintSched 
            Caption         =   "Print Schedule"
            Height          =   405
            Left            =   4470
            TabIndex        =   51
            Top             =   5250
            Width           =   1455
         End
         Begin VB.Label labPayments 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Interest"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   75
            TabIndex        =   49
            Top             =   5010
            Width           =   1515
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Payment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   75
            TabIndex        =   53
            Top             =   5445
            Width           =   1515
         End
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   0
      ScaleHeight     =   2580
      ScaleWidth      =   11745
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      Begin VB.OptionButton optFormula2 
         Caption         =   "Formula 2 : (Principal * Interest / (1 - ((1 / (1 + Interest) ^ TERM))))"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4290
         TabIndex        =   58
         Top             =   2310
         Width           =   7095
      End
      Begin VB.OptionButton optFormula1 
         Caption         =   "Formula 1 : (Amount To Be Financed  * (1 + (AOR / 100))) / TERM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4290
         TabIndex        =   57
         Top             =   2070
         Value           =   -1  'True
         Width           =   7095
      End
      Begin VB.TextBox txtNotes 
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
         Height          =   510
         Left            =   4245
         TabIndex        =   13
         Top             =   1470
         Width           =   7305
      End
      Begin VB.TextBox txtSAEPhone 
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
         Left            =   60
         TabIndex        =   14
         Top             =   2100
         Width           =   4095
      End
      Begin VB.ComboBox cboSAEName 
         Height          =   345
         Left            =   60
         TabIndex        =   11
         Top             =   1500
         Width           =   4110
      End
      Begin VB.TextBox txtContact 
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
         Left            =   4245
         TabIndex        =   8
         Top             =   885
         Width           =   7305
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   4245
         TabIndex        =   4
         Top             =   250
         Width           =   7305
      End
      Begin VB.TextBox txtCustomer 
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
         Left            =   60
         TabIndex        =   3
         Top             =   250
         Width           =   4110
      End
      Begin VB.ComboBox cboVDetails 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   885
         Width           =   4110
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes "
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
         Left            =   4245
         TabIndex        =   10
         Top             =   1230
         Width           =   2520
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SAE's Cell Phone Number"
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
         Left            =   60
         TabIndex        =   12
         Top             =   1860
         Width           =   2160
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Account Executive"
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
         Left            =   60
         TabIndex        =   9
         Top             =   1260
         Width           =   2520
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Details"
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
         Left            =   60
         TabIndex        =   6
         Top             =   675
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   4245
         TabIndex        =   1
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Details"
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
         Left            =   4245
         TabIndex        =   5
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   15
         Width           =   1860
      End
   End
   Begin Crystal.CrystalReport rptAOR 
      Left            =   255
      Top             =   4275
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmSMIS_Mis_AOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Dontshowselectbutton                                           As Boolean
Event LineAOR( _
      LOR_Fincom As Variant, _
      LOR_Custinfo As Variant, _
      LOR_Addinfo As Variant, _
      LOR_Vehiclesinfo As Variant, _
      LOR_Customerinfo As Variant, _
      LOR_UnitPrice As Variant, _
      LOR_LTO As Variant, _
      LOR_Chattel As Variant, _
      LOR_Insurance As Variant, _
      LOR_TotalUnitCost As Variant, _
      LOR_Discount As Variant, _
      LOR_GrandTotal As Variant, _
      LOR_DownPayment As Variant, _
      LOR_BalToFinance As Variant, _
      LOR_Term As Variant, _
      LOR_MonthlyAmort As Variant, _
      LOR_AOR As Variant, _
      LOR_DownpaymentRate As Variant)
Dim RSPercentages                                                     As ADODB.Recordset
Dim ComputebyPert                                                     As Boolean

Public Function MyRound(payment As Currency)
    payment = payment + 0.005
    payment = payment * 100
    payment = Int(payment)
    payment = payment / 100
    MyRound = payment
End Function

Sub initGrid()

    With Grid1
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .BackColorFixed = RGB(90, 158, 214)
        .BackColorFixedSel = RGB(110, 180, 230)
        .BackColorBkg = RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cols = 7
        .Cell(0, 0).Text = ""
        .Column(0).Width = 0

        .Cell(0, 1).Text = "Select"
        .Column(1).CellType = cellCheckBox
        .Column(1).Width = 50
        .Column(1).Locked = False

        .Cell(0, 2).Text = "Terms"
        .Column(2).Width = 50
        .Column(2).Alignment = cellCenterCenter
        .Column(2).Locked = True

        .Cell(0, 3).Text = "Amortization"
        .Column(3).Width = 120
        .Column(3).Alignment = cellRightGeneral
        .Column(3).Locked = True

        .Cell(0, 4).Text = "AOR"
        .Column(4).Alignment = cellCenterCenter
        .Column(4).Width = 50
        .Column(4).Locked = True

        .Cell(0, 5).Text = "Interest"
        .Column(5).Width = 100
        .Column(5).Alignment = cellRightGeneral
        .Column(5).Locked = True

        .Cell(0, 6).Text = "Options"
        .Column(6).Locked = True
        .Column(6).Alignment = cellCenterCenter

    End With
End Sub

Sub PrintSchedule()
    Dim i                                                             As Long
    Dim vtxtMonths                                                    As Long
    Dim vtxtMonthlyAmoritization                                      As Currency
    Dim vtxtInterest                                                  As Double
    Dim vtxtPrincipal                                                 As Currency
    Dim vtxtBalance                                                   As Currency
    rptAOR.Reset
    progCPB.Max = gridSchedule.Rows - 1

    gconDMIS.Execute ("delete from ALL_AOR WHERE LOGNAME IS NULL OR LogName=" & N2Str2Null(LOGNAME))


    For i = 1 To gridSchedule.Rows - 1
        vtxtMonths = CInt(gridSchedule.TextMatrix(i, 0))     ' Months
        vtxtMonthlyAmoritization = CCur(gridSchedule.TextMatrix(i, 1))    ' MonthlyAmoritization
        vtxtInterest = CDbl(gridSchedule.TextMatrix(i, 2))   ' Interest
        vtxtPrincipal = CCur(gridSchedule.TextMatrix(i, 3))  ' Principal
        vtxtBalance = CCur(gridSchedule.TextMatrix(i, 4))    ' Balance
        gconDMIS.Execute ("Insert Into ALL_AOR (LogName,termno, payment,monthlyinterest, principal  ,balance) " _
                        & " values(" & N2Str2Null(LOGNAME) & "," & vtxtMonths & "," & vtxtMonthlyAmoritization & "," & vtxtInterest & "," & vtxtPrincipal & " ," & vtxtBalance & " )")
        progCPB.Value = progCPB.Value + 1
    Next


    With rptAOR
        .Formulas(0) = "VehicleName='" & cboVDetails & "'"
        .Formulas(1) = "finco='" & cboFinCom & "'"
        .Formulas(2) = "netsaleprice='" & txtUnitPrice & "'"
        .Formulas(3) = "downpayment ='" & txtDownPayment & "'"
        .Formulas(4) = "downpaymentpercent='" & NumericVal(txtDownpaymentPert) & "%'"
        .Formulas(5) = "baltofin='" & txtFinancingBalance & "'"
        .Formulas(6) = "AOR ='" & Grid1.Cell(Grid1.ActiveCell.Row, 4).DoubleValue & "'"
        .Formulas(7) = "NoTerms = '" & Grid1.Cell(Grid1.ActiveCell.Row, 2).SingleValue & " Months '"
        .Formulas(8) = "companyname= '" & COMPANY_NAME & "'"
        .Formulas(9) = "NoTerms = '" & COMPANY_ADDRESS & "'"
        .WindowTitle = "AOR Schedule"
    End With

    PrintSQLReport rptAOR, SMIS_REPORT_PATH & "AORSchedule.rpt", "{ALL_AOR.LogName}='" & LOGNAME & "'", DMIS_REPORT_Connection, 1


End Sub

Sub ProcessSchedule()
    Dim termno                                                        As Integer
    Dim payment                                                       As Double
    Dim Principal                                                     As Currency
    Dim TERM                                                          As Integer
    Dim Interest                                                      As Double
    Dim MonthlyInterest                                               As Double


    On Error GoTo Errorcode:

    Principal = CCur(txtFinancingBalance.Text)
    TERM = Grid1.Cell(Grid1.ActiveCell.Row, 2).SingleValue
    payment = Grid1.Cell(Grid1.ActiveCell.Row, 3).DoubleValue
    Interest = Grid1.Cell(Grid1.ActiveCell.Row, 4).DoubleValue

    Interest = Interest / 1200
    txtTotalInterest = FormatNumber(Grid1.Cell(Grid1.ActiveCell.Row, 3).DoubleValue)
    txtTotalPayment = FormatNumber(NumericVal(Grid1.Cell(Grid1.ActiveCell.Row, 3).DoubleValue) + Principal)
    gridSchedule.Rows = 1
    progCPB.Max = TERM
    For termno = 1 To TERM
        DoEvents

        MonthlyInterest = MyRound(Interest * Principal)

        Principal = MyRound(Principal - (payment - MonthlyInterest))
        If Principal < 0 Then

        End If
        If termno = TERM And Principal <> 0 Then
            payment = payment + Principal
            Principal = 0
        End If
        gridSchedule.AddItem termno & Chr(9) & FormatNumber(payment) & Chr(9) & FormatNumber(MonthlyInterest) & Chr(9) & FormatNumber((payment - MonthlyInterest)) & Chr(9) & FormatNumber(Principal)
        progCPB.Text = " Generating Schedules for ::" & termno & " Month"
        progCPB.Value = progCPB.Value + 1
    Next






    Exit Sub
Errorcode:
    ShowVBError

End Sub

Sub ShowOnlyCalculator()
    picTop.Visible = False
    picLow.Visible = True
    Me.Height = picLow.Height + 400
End Sub

Sub ShowonlyComputation()
    Dontshowselectbutton = True
    txtUnitPrice = "0.00"
    txtLTO = "0.00"
    txtChattelMorgage = "0.00"
    txtInsurance = "0.00"
    txtOthers = "0.00"
    txtDiscount = "0.00"
    txtDiscountPert = "0.00"
    txtDownPayment = "0.00"
    txtDownpaymentPert = "0.00"
End Sub

Sub UpdateUnitCost()
    Dim myVal                                                         As Double
    Grid1.Rows = 1
    txtNetAmount = FormatNumber(NumericVal(txtUnitPrice) - NumericVal(txtDiscount) - NumericVal(txtDownPayment))
    txtAmountDue = FormatNumber(NumericVal(txtDownPayment) + NumericVal(txtLTO) + NumericVal(txtChattelMorgage) + NumericVal(txtInsurance) + NumericVal(txtOthers))

    txtFinancingBalance = FormatNumber(NumericVal(txtUnitPrice) - NumericVal(txtDownPayment))

    Command4.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub cboFinCom_Click()
    If cboFinCom.ListIndex = -1 Then: Exit Sub
    If NumericVal(txtUnitPrice) > 0 Then
        UpdateUnitCost
    End If
End Sub

Private Sub cboVDetails_Click()
    If cboVDetails.ListIndex <> -1 Then
        Dim rsTemp                                                    As ADODB.Recordset
        Set rsTemp = gconDMIS.Execute("select * from all_model where id=" & cboVDetails.ItemData(cboVDetails.ListIndex))
        If Not rsTemp.EOF Or rsTemp.BOF Then
            txtUnitPrice = FormatNumber(NumericVal(rsTemp!unitcost))
            txtLTO = FormatNumber(NumericVal(rsTemp!LTO))

        Else
            txtUnitPrice = "0.00"
            txtLTO = "0.00"
        End If

    End If
End Sub

Private Sub cmdPrintSched_Click()
    Screen.MousePointer = vbHourglass
    progCPB.Visible = True: progCPB.Value = 0

    PrintSchedule
    progCPB.Visible = False: progCPB.Value = 0
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSchedule_Click()
    On Error GoTo Errorcode:
    Screen.MousePointer = vbHourglass
    progCPB.Visible = True: progCPB.Value = 0
    optFormula1.Enabled = False
    optFormula2.Enabled = False
    ProcessSchedule
    progCPB.Visible = False: progCPB.Value = 0
    picSchedule.Visible = True: picSchedule.ZOrder 0
    Screen.MousePointer = vbNormal
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    On Error GoTo Errorcode:

    Dim Valuex                                                        As String
    Dim printing                                                      As Boolean
    If LTrim(RTrim(txtCustomer)) = "" Then
        Valuex = "Customer Information"
    End If

    If LTrim(RTrim(cboSAEName)) = "" Then
        Valuex = Valuex & ", Sales Agent Information"
    End If

    If LTrim(RTrim(txtSAEPhone)) = "" Then
        Valuex = Valuex & " , Sales Agent Cell Phone"
    End If
    If Valuex <> "" Then
        If MsgBox("Following Information are Missing Are you Sure You Want To Print The Computation Sheet" & vbCrLf & Valuex, vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    If Grid1.Rows = 1 Then: Exit Sub
    Screen.MousePointer = vbHourglass
    progCPB.Visible = True: progCPB.Value = 0
    rptAOR.Reset
    With rptAOR
        .Formulas(0) = "CustomerName='" & txtCustomer & "'"
        .Formulas(1) = "address='" & txtAddress & "'"
        .Formulas(2) = "unit='" & cboVDetails & "'"
        .Formulas(3) = "unitprice='" & txtUnitPrice & "'"
        .Formulas(4) = "DownPayment='" & txtDownPayment & "'"
        .Formulas(5) = "baltofin='" & txtFinancingBalance & "'"
        .Formulas(6) = "downpaymentperct='" & txtDownpaymentPert & "'"
        .Formulas(7) = "finreq='Financial Requirements'"
        .Formulas(8) = "LTOReg='" & txtLTO & "'"
        .Formulas(9) = "Chattel='" & txtChattelMorgage & "'"
        .Formulas(10) = "Insurance='" & txtInsurance & "'"
        .Formulas(11) = "Total='" & txtNetAmount & "'"
        .Formulas(12) = "discount='" & txtDiscount & "'"
        .Formulas(13) = "GrandTotal='" & txtAmountDue & "'"
        .Formulas(15) = "Cell'" & txtContact & "'"
        .Formulas(16) = "CompanyName = '" & COMPANY_NAME & "'"
        .Formulas(17) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        .Formulas(18) = "finreq= '" & txtNotes & "'"
        .Formulas(19) = "sae= '" & cboSAEName & "'"
        .Formulas(20) = "Cell= '" & txtSAEPhone & "'"

    End With
    Dim i                                                             As Long
    gconDMIS.Execute ("delete from ALL_AORComputation")
    Dim grows                                                         As Long
    grows = Grid1.Rows - 1
    progCPB.Max = grows

    For i = 1 To grows
        progCPB.Value = progCPB.Value + i
        If Grid1.Cell(i, 1).BooleanValue = True Then
            gconDMIS.Execute ("INSERT INTO ALL_AORComputation (Term, MonthlyAMort,AOR) VALUES (" & Grid1.Cell(i, 2).SingleValue & "," & Grid1.Cell(i, 3).DoubleValue & ", " & Grid1.Cell(i, 4).DoubleValue & " )")
        End If
        progCPB.Text = "Printing Computation Sheet " & CInt((i / grows) * 100) & "%"
    Next
    PrintSQLReport rptAOR, SMIS_REPORT_PATH & "AOR.rpt", "", DMIS_REPORT_Connection, 1
    progCPB.Visible = False
    Screen.MousePointer = vbNormal





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command2_Click()
    picSchedule.Visible = False
    picSchedule.ZOrder 1
    Command4.Enabled = True
    optFormula1.Enabled = True
    optFormula2.Enabled = True
End Sub

Private Sub Command3_Click()
    gconDMIS.Execute ("delete from ALL_AOR WHERE LogName=" & N2Str2Null(LOGNAME))
    Set RSPercentages = Nothing
    Unload Me

End Sub

Private Sub Command4_Click()
    If NumericVal(txtFinancingBalance) > 0 Then
        ProcessAmort
        Command4.Enabled = False
        Command1.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If picTop.Visible = True Then
        On Error Resume Next
        txtCustomer.SetFocus
    Else
        On Error Resume Next
        cboFinCom.SetFocus
        SendKeys "{F4}"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initGrid
    FillCombo "Select ID, Company  from SMIS_FINCOM ", 0, 1, cboFinCom
    FillCombo "select descript ,id from ALL_MODEL", 1, 0, cboVDetails
    Set RSPercentages = New ADODB.Recordset

    Call RSPercentages.Open("SELECT * FROM SMIS_FINCOM_RATE ORDER BY TERM", gconDMIS, adOpenKeyset, adLockReadOnly)
    FillCombo "SELECT [NAME] FROM SMIS_vw_Srep ", -1, 0, cboSAEName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dontshowselectbutton = False

End Sub

Private Sub Grid1_DblClick()
    If Grid1.Selection.FirstCol = Grid1.Cols - 1 Then
        cmdSchedule_Click
        Command4.Enabled = False
    End If
End Sub

Private Sub optFormula1_Click()
    Command4_Click
End Sub

Private Sub optFormula2_Click()
    Command4_Click
End Sub

Private Sub optRural_Click()
    Grid1.Cell(0, 4).Text = "OMA"
    ProcessAmort
End Sub

Private Sub optUrban_Click()
    Grid1.Cell(0, 4).Text = "AOR"
    ProcessAmort
End Sub

Private Sub ProcessAmort()
    On Error Resume Next
    Dim payment                                                       As Long
    Dim Total_Interest                                                As Double
    Dim Principal                                                     As Currency
    Dim TERM                                                          As Integer
    Dim Interest                                                      As Double
    Dim AOR                                                           As Double
    Dim i                                                             As Long

    Principal = NumericVal(txtFinancingBalance)
    Grid1.Rows = 1
    While Not RSPercentages.EOF
        If optRural.Value = True Then
            AOR = NumericVal(RSPercentages!RPerct)
        Else
            AOR = NumericVal(RSPercentages!UPerct)
        End If
        Interest = AOR / 1200
        TERM = NumericVal(RSPercentages!TERM)
        If optFormula2.Value = True Then
            payment = (Principal * Interest / (1 - ((1 / (1 + Interest) ^ TERM))))
            Total_Interest = (TERM * payment) - Principal
            Grid1.Column(6).Width = 70
            Grid1.AddItem "0" & Chr(9) & TERM & Chr(9) & FormatNumber(payment) & Chr(9) & AOR & Chr(9) & FormatNumber(Total_Interest) & Chr(9) & "SCHEDULE", False
        Else
            payment = (Principal * (1 + (AOR / 100))) / TERM
            Total_Interest = (TERM * payment) - Principal
            Grid1.Column(6).Width = 0
            Grid1.AddItem "0" & Chr(9) & TERM & Chr(9) & FormatNumber(payment) & Chr(9) & AOR & Chr(9) & FormatNumber(Total_Interest), False
        End If
        RSPercentages.MoveNext

    Wend
    Grid1.Refresh
    If Not RSPercentages.BOF Then: RSPercentages.MoveFirst

End Sub

Private Sub txtChattelMorgage_Change()
    UpdateUnitCost
End Sub

Private Sub txtChattelMorgage_GotFocus()
    If NumericVal(txtChattelMorgage.Text) <= 0 Then txtChattelMorgage = ""

End Sub

Private Sub txtChattelMorgage_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtChattelMorgage_LostFocus()
    If NumericVal(txtChattelMorgage.Text) <= 0 Then txtChattelMorgage = "0.00"
    txtChattelMorgage = FormatNumber(txtChattelMorgage)
End Sub

'''''''''DISCOUNT AND '''''''''DOWNPAYMENTS
Private Sub txtDiscount_Change()
    On Error GoTo ADDER:
    If ComputebyPert = True Then Exit Sub
    txtDiscountPert = (NumericVal(txtDiscount) / NumericVal(txtUnitPrice)) * 100
    UpdateUnitCost
    Exit Sub
ADDER:
    Err.Clear
End Sub

Private Sub txtDiscount_GotFocus()
    If NumericVal(txtDiscount.Text) <= 0 Then txtDiscount = ""

End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtDiscountPert.ZOrder 0
        On Error Resume Next
        txtDiscountPert.SetFocus
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDiscount_LostFocus()
    If NumericVal(txtDiscount.Text) <= 0 Then txtDiscount = "0.00"
    txtDiscount = FormatNumber(txtDiscount)
End Sub

Private Sub txtDiscountPert_Change()
    On Error GoTo ADDER:
    If ComputebyPert = False Then: Exit Sub
    txtDiscount = FormatNumber(NumericVal(txtUnitPrice) * (NumericVal(txtDiscountPert) / 100))
    'txtDownpaymentPert = FormatNumber((NumericVal(txtDownPayment) / NumericVal(txtAmountDue)) * 100)
    UpdateUnitCost
    Exit Sub
ADDER:
    Err.Clear
End Sub

Private Sub txtDiscountPert_GotFocus()
    If NumericVal(txtDiscountPert.Text) <= 0 Then txtDiscountPert = ""
    ComputebyPert = True
End Sub

Private Sub txtDiscountPert_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDiscountPert_LostFocus()
    If NumericVal(txtDiscountPert.Text) <= 0 Then txtDiscountPert = "0.00"
    txtDiscountPert = FormatNumber(txtDiscountPert)
    ComputebyPert = False
End Sub

Private Sub txtDownPayment_change()
    On Error GoTo ADDER:
    If ComputebyPert = True Then Exit Sub
    If NumericVal(txtAmountDue) <> 0 Then
        txtDownpaymentPert = (NumericVal(txtDownPayment) / NumericVal(txtUnitPrice)) * 100
    End If
    UpdateUnitCost
    Exit Sub
ADDER:
    Err.Clear
End Sub

Private Sub txtDownpayment_GotFocus()
    If NumericVal(txtDownPayment.Text) <= 0 Then txtDownPayment = ""
End Sub

Private Sub txtDownpayment_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDownPayment_LostFocus()
    If NumericVal(txtDownPayment.Text) <= 0 Then txtDownPayment = "0.00"
    txtDownPayment = FormatNumber(txtDownPayment)
    ComputebyPert = False
End Sub

Private Sub txtDownpaymentPert_Change()
    If ComputebyPert = False Then: Exit Sub
    txtDownPayment = FormatNumber(NumericVal(txtUnitPrice) * (NumericVal(txtDownpaymentPert) / 100))

    UpdateUnitCost
End Sub

Private Sub txtDownpaymentPert_GotFocus()
    ComputebyPert = True
    If NumericVal(txtDownpaymentPert.Text) <= 0 Then txtDownpaymentPert = ""
End Sub

Private Sub txtDownpaymentPert_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDownpaymentPert_LostFocus()
    If NumericVal(txtDownpaymentPert.Text) <= 0 Then txtDownpaymentPert = "0.00"
    txtDownpaymentPert = FormatNumber(txtDownpaymentPert)
    ComputebyPert = False
End Sub

'''''''''DISCOUNT AND '''''''''DOWNPAYMENTS
Private Sub txtInsurance_Change()
    UpdateUnitCost
End Sub

Private Sub txtInsurance_GotFocus()
    If NumericVal(txtInsurance.Text) <= 0 Then txtInsurance = ""
End Sub

Private Sub txtInsurance_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtInsurance_LostFocus()
    If NumericVal(txtInsurance.Text) <= 0 Then txtInsurance = "0.00"
    txtInsurance = FormatNumber(txtInsurance)
End Sub

Private Sub txtLto_Change()
    UpdateUnitCost
End Sub

Private Sub txtLto_GotFocus()
    If NumericVal(txtLTO.Text) <= 0 Then txtLTO = ""

End Sub

Private Sub txtLto_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLto_LostFocus()
    If NumericVal(txtLTO.Text) <= 0 Then txtLTO = "0.00"
    txtLTO = FormatNumber(txtLTO)
End Sub

Private Sub txtOthers_Change()
    UpdateUnitCost
End Sub

Private Sub txtOthers_GotFocus()
    If NumericVal(txtOthers.Text) <= 0 Then txtOthers = ""

End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtOthers_LostFocus()
    If NumericVal(txtOthers.Text) <= 0 Then txtOthers = "0.00"
    txtOthers = FormatNumber(txtOthers)
End Sub

Private Sub txtUnitPrice_Change()
    UpdateUnitCost
End Sub

Private Sub txtUnitPrice_GotFocus()
    If NumericVal(txtUnitPrice.Text) <= 0 Then txtUnitPrice = ""

End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtUnitPrice_LostFocus()
    If NumericVal(txtUnitPrice.Text) <= 0 Then txtUnitPrice = "0.00"
    txtUnitPrice = FormatNumber(txtUnitPrice)
End Sub

