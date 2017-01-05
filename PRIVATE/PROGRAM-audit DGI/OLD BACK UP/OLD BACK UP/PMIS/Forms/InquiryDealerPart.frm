VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmPMISTrans_DealerPartInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dealer Part  - Inquiry"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InquiryDealerPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   10560
   Begin VB.PictureBox picInquiryPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4740
      ScaleHeight     =   2145
      ScaleWidth      =   4500
      TabIndex        =   36
      Top             =   2460
      Visible         =   0   'False
      Width           =   4530
      Begin VB.CommandButton cmdDelPriceInquiry 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3570
         MouseIcon       =   "InquiryDealerPart.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Delete Entry"
         Top             =   1230
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelInquiry_Price 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2850
         MouseIcon       =   "InquiryDealerPart.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Cancel Entry"
         Top             =   1230
         Width           =   735
      End
      Begin VB.CommandButton cmdOkInquiry_Price 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2130
         MouseIcon       =   "InquiryDealerPart.frx":11D7
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":1329
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Save Entry"
         Top             =   1230
         Width           =   735
      End
      Begin VB.TextBox txtPrice_SSO_SRP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7590
         TabIndex        =   52
         Top             =   2250
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_SSO_DP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6540
         TabIndex        =   50
         Top             =   2250
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_SAO_SRP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7590
         TabIndex        =   49
         Top             =   1860
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_SAO_DP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6540
         TabIndex        =   48
         Top             =   1860
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_Reg_SRP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7590
         TabIndex        =   45
         Top             =   1470
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_Reg_DP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6420
         TabIndex        =   44
         Top             =   1440
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtPrice_PartName 
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
         Height          =   345
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   765
         Width           =   2805
      End
      Begin VB.ComboBox cboPrice_PartNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   1530
         TabIndex        =   39
         Top             =   390
         Width           =   2835
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Special Air Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   4650
         TabIndex        =   47
         Top             =   1860
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Special Sea Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   4530
         TabIndex        =   51
         Top             =   2250
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Regular "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   16
         Left            =   5640
         TabIndex        =   46
         Top             =   1530
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   7650
         TabIndex        =   43
         Top             =   1230
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   6600
         TabIndex        =   42
         Top             =   1260
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1605
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   345
         Left            =   0
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Width           =   4485
         _Version        =   655364
         _ExtentX        =   7911
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "PRICE INQUIRY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picDPI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   5250
      ScaleHeight     =   3285
      ScaleWidth      =   3495
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   3525
      Begin VB.OptionButton optDPIInq 
         BackColor       =   &H00FFFF00&
         Caption         =   "PARTS ESTIMATE TIME OF ARRIVAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "ETA"
         Top             =   2040
         Width           =   3225
      End
      Begin VB.OptionButton optDPIInq 
         BackColor       =   &H0000C000&
         Caption         =   "PRICE INQUIRY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "PRCS"
         Top             =   1545
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker dtDateDPI 
         Height          =   345
         Left            =   90
         TabIndex        =   30
         Top             =   660
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
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
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39202
         MinDate         =   36526
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2460
         MouseIcon       =   "InquiryDealerPart.frx":1679
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   2550
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1620
         MouseIcon       =   "InquiryDealerPart.frx":17CB
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   2550
         Width           =   855
      End
      Begin VB.OptionButton optDPIInq 
         BackColor       =   &H0000FFFF&
         Caption         =   "TECHNICAL INQUIRY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "TECH"
         Top             =   1050
         Width           =   3225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE: (MM /DD / YYYY)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   480
         TabIndex        =   29
         Top             =   390
         Width           =   2805
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   3645
         _Version        =   655364
         _ExtentX        =   6429
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Select Your Date && DPI Type"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2490
      ScaleHeight     =   285
      ScaleWidth      =   8715
      TabIndex        =   67
      Top             =   6000
      Width           =   8715
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5940
         TabIndex        =   72
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4080
         TabIndex        =   71
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2550
         TabIndex        =   69
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1290
         TabIndex        =   70
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   30
         TabIndex        =   68
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7050
      Left            =   0
      ScaleHeight     =   7050
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      Begin VB.Frame Frame2 
         Height          =   7035
         Left            =   30
         TabIndex        =   1
         Top             =   -30
         Width           =   2355
         Begin VB.OptionButton optDate 
            Caption         =   "DA&TE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   660
            Width           =   1845
         End
         Begin VB.OptionButton optDPI 
            Caption         =   "&DPI No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   390
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.TextBox textSearch 
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   90
            MaxLength       =   35
            TabIndex        =   7
            Text            =   "TEXT"
            Top             =   960
            Width           =   2205
         End
         Begin MSComctlLib.ListView lstDPIList 
            Height          =   5565
            Left            =   60
            TabIndex        =   8
            Top             =   1350
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   9816
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "InquiryDealerPart.frx":191D
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DPINO"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.OptionButton optVIN 
            Caption         =   "&VIN NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   210
            TabIndex        =   4
            Top             =   2160
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.OptionButton optCatalgoue 
            Caption         =   "CATALO&GUE NO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   5
            Top             =   2460
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label22 
            Caption         =   "Search by:"
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
            Left            =   90
            TabIndex        =   2
            Top             =   150
            Width           =   1455
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3015
      Left            =   2490
      TabIndex        =   53
      Top             =   3000
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColor       =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483633
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   $"InquiryDealerPart.frx":1A7F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "InquiryDealerPart.frx":1B4B
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2820
      ScaleHeight     =   915
      ScaleWidth      =   7740
      TabIndex        =   73
      Top             =   6240
      Width           =   7740
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6900
         MouseIcon       =   "InquiryDealerPart.frx":1E65
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":1FB7
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6180
         MouseIcon       =   "InquiryDealerPart.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
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
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "InquiryDealerPart.frx":27D5
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":2927
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Cancel this Transaction"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4740
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "InquiryDealerPart.frx":2C61
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":2DB3
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Post this Transaction"
         Top             =   60
         Width           =   735
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
         Left            =   4020
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "InquiryDealerPart.frx":30D8
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":322A
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Unpost this Transaction"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3300
         MouseIcon       =   "InquiryDealerPart.frx":356F
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":36C1
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2580
         MouseIcon       =   "InquiryDealerPart.frx":3A1D
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":3B6F
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1860
         MouseIcon       =   "InquiryDealerPart.frx":3E82
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":3FD4
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1140
         MouseIcon       =   "InquiryDealerPart.frx":42CE
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":4420
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   420
         MouseIcon       =   "InquiryDealerPart.frx":4778
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":48CA
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
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
      Left            =   8940
      ScaleHeight     =   885
      ScaleWidth      =   2130
      TabIndex        =   84
      Top             =   6240
      Width           =   2130
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   780
         MouseIcon       =   "InquiryDealerPart.frx":4C29
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":4D7B
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         MouseIcon       =   "InquiryDealerPart.frx":50B9
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":520B
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2985
      Left            =   2430
      ScaleHeight     =   2985
      ScaleWidth      =   8385
      TabIndex        =   90
      Top             =   0
      Width           =   8385
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4860
         TabIndex        =   100
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtVinNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         TabIndex        =   98
         Top             =   1320
         Width           =   3885
      End
      Begin VB.TextBox txtCatalgoueNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         TabIndex        =   97
         Top             =   1950
         Width           =   3885
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2220
         Top             =   1350
      End
      Begin VB.CheckBox chkEnclose 
         Caption         =   "ENCLOSED ATTACHMENT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4860
         TabIndex        =   95
         Top             =   1200
         Width           =   2715
      End
      Begin VB.TextBox txtDPINo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4890
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   420
         Width           =   3045
      End
      Begin VB.TextBox txtDateDPI2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   840
         Width           =   3075
      End
      Begin VB.TextBox txtNotedBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   2550
         Width           =   1905
      End
      Begin VB.TextBox txtReqBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   2550
         Width           =   1935
      End
      Begin VB.CommandButton cmdDPINo 
         Caption         =   "ASSIGN DPI NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4890
         MouseIcon       =   "InquiryDealerPart.frx":555B
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   30
         Width           =   3075
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5070
         TabIndex        =   101
         Top             =   1770
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtDPIDetailID 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   6210
         TabIndex        =   102
         Text            =   "0"
         Top             =   1740
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtCompanyName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   99
         Top             =   300
         Width           =   3915
      End
      Begin VB.Label lblVin 
         Caption         =   "Vehicle Identitfication Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   111
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "Catalogue Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   110
         Top             =   1710
         Width           =   2295
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4890
         TabIndex        =   109
         Top             =   2700
         Width           =   3255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "SUBJECT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4860
         TabIndex        =   108
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "DPI No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4215
         TabIndex        =   107
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   4410
         TabIndex        =   106
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Noted By"
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
         Height          =   225
         Left            =   2070
         TabIndex        =   105
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Request By"
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
         Height          =   225
         Left            =   60
         TabIndex        =   104
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Dealer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   103
         Top             =   60
         Width           =   2715
      End
   End
   Begin VB.PictureBox picInquiryEstimateArrival 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   3870
      ScaleHeight     =   3405
      ScaleWidth      =   5430
      TabIndex        =   9
      Top             =   1590
      Visible         =   0   'False
      Width           =   5460
      Begin VB.ComboBox cboETA_OrderNo 
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
         Left            =   180
         TabIndex        =   112
         Text            =   "Combo1"
         Top             =   720
         Width           =   2265
      End
      Begin VB.CommandButton cmdDel_ETA 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4410
         MouseIcon       =   "InquiryDealerPart.frx":56AD
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":57FF
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Delete Entry"
         Top             =   2430
         Width           =   735
      End
      Begin VB.ComboBox cboPartsEstimate_Status 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   1800
         TabIndex        =   24
         Top             =   4740
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.TextBox txtETA_QTY 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4350
         TabIndex        =   19
         Top             =   1950
         Width           =   795
      End
      Begin VB.TextBox txtETA_PARTNAME 
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
         Left            =   180
         TabIndex        =   14
         Top             =   1980
         Width           =   4095
      End
      Begin VB.ComboBox cboETA_PNO 
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
         IntegralHeight  =   0   'False
         Left            =   180
         TabIndex        =   11
         Top             =   1320
         Width           =   2265
      End
      Begin MSComCtl2.DTPicker txtETA_DateOrd 
         Height          =   360
         Left            =   2520
         TabIndex        =   17
         Top             =   720
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
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
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39202
         MinDate         =   36526
      End
      Begin MSComCtl2.DTPicker dtPartsEstimate_ETA 
         Height          =   360
         Left            =   1800
         TabIndex        =   26
         Top             =   5130
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
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
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39202
         MinDate         =   36526
      End
      Begin VB.CommandButton cmdCancel_ETA 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3690
         MouseIcon       =   "InquiryDealerPart.frx":5B2A
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":5C7C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel Entry"
         Top             =   2430
         Width           =   735
      End
      Begin VB.CommandButton cmdSave_ETA 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2970
         MouseIcon       =   "InquiryDealerPart.frx":5FBA
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":610C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate Time of Arrival"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   13
         Left            =   150
         TabIndex        =   25
         Top             =   5100
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   210
         TabIndex        =   23
         Top             =   4710
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   4170
         TabIndex        =   18
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Ordered"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2550
         TabIndex        =   16
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   180
         TabIndex        =   15
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   12
         Top             =   1680
         Width           =   1140
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   5475
         _Version        =   655364
         _ExtentX        =   9657
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "PARTS ESTIMATE TIME OF ARRIVAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picInquiryTechincal 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   4680
      ScaleHeight     =   2205
      ScaleWidth      =   4500
      TabIndex        =   54
      Top             =   1890
      Visible         =   0   'False
      Width           =   4530
      Begin VB.CommandButton cmdDel_Tech 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3540
         MouseIcon       =   "InquiryDealerPart.frx":645C
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":65AE
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Delete Entry"
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox txtTech_PartName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   59
         Top             =   810
         Width           =   2595
      End
      Begin VB.ComboBox cboTech_PartNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   1680
         TabIndex        =   57
         Top             =   435
         Width           =   2595
      End
      Begin VB.TextBox txtTech_PNC 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   66
         Top             =   2580
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txtTech_GroupNumber 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   63
         Top             =   2160
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CommandButton cmdCancel_Tech 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2820
         MouseIcon       =   "InquiryDealerPart.frx":68D9
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":6A2B
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Cancel Entry"
         Top             =   1260
         Width           =   735
      End
      Begin VB.CommandButton cmdSave_Tech 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2100
         MouseIcon       =   "InquiryDealerPart.frx":6D69
         MousePointer    =   99  'Custom
         Picture         =   "InquiryDealerPart.frx":6EBB
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Save Entry"
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   58
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "PNC Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   65
         Top             =   2670
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   2190
         Visible         =   0   'False
         Width           =   1605
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   315
         Left            =   0
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   0
         Width           =   4665
         _Version        =   655364
         _ExtentX        =   8229
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "TECHNICAL INQUIRY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPMISTrans_DealerPartInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsDPIR                                                            As ADODB.Recordset
Dim ADDOREDIT                                                         As String
Dim dpiInqType                                                        As String
Dim dpiSeq                                                            As String
Dim olddpiInqType                                                     As String
Dim RSPARTS                                                           As ADODB.Recordset
Dim ReqBy                                                             As String
Dim CompName                                                          As String
Dim NotedBy                                                           As String

Function GenerateDPISEQ() As String

    Dim rsID                                                          As ADODB.Recordset
    Dim SQL                                                           As String
    Dim TEMPRS                                                        As ADODB.Recordset

    If ADDOREDIT = "EDIT" Then

        Set TEMPRS = gconDMIS.Execute("Select DPI_INQType ,Dpi_date, DPI_SEQNO from PMIS_DPIHeader Where ID=" & txtID)

        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then

            If dpiInqType = TEMPRS!DPI_INQTYPE And Month(dtDateDPI) = Month(TEMPRS!DPI_DATE) Then
                olddpiInqType = TEMPRS!DPI_INQTYPE
                GenerateDPISEQ = TEMPRS!DPI_SEQNO
            Else
                SQL = "SELECT ISNULL(MAX(DPI_SEQNO),0) AS MAXVAL  FROM PMIS_DPIHeader WHERE  " & vbCrLf & _
                    " YEAR(DPI_DATE)    =   YEAR    ('" & dtDateDPI & "' ) AND " & vbCrLf & _
                    " MONTH(DPI_DATE)   =   MONTH   ('" & dtDateDPI & "')  AND " & vbCrLf & _
                    " DPI_INQTYPE       =   '" & dpiInqType & "'"

                Set rsID = gconDMIS.Execute(SQL)
                If rsID.Fields(0).Value = 0 Then
                    GenerateDPISEQ = Format(1, "000")
                Else
                    GenerateDPISEQ = Format(Val(N2Str2Zero(rsID![MAXVAL])) + 1, "000")
                End If
            End If
            Set rsID = Nothing
        End If
    Else
        If IsDate(dtDateDPI) = True Then
            SQL = "SELECT ISNULL(MAX(DPI_SEQNO),0) AS MAXVAL  FROM PMIS_DPIHeader WHERE  " & vbCrLf & _
                " YEAR(DPI_DATE)    =   YEAR    ('" & dtDateDPI & "' ) AND " & vbCrLf & _
                " MONTH(DPI_DATE)   =   MONTH   ('" & dtDateDPI & "')  AND " & vbCrLf & _
                " DPI_INQTYPE       =   '" & dpiInqType & "'"

            Set rsID = gconDMIS.Execute(SQL)
            If rsID.Fields(0).Value = 0 Then
                GenerateDPISEQ = Format(1, "000")
            Else
                GenerateDPISEQ = Format(Val(N2Str2Zero(rsID![MAXVAL])) + 1, "000")

            End If
            Set rsID = Nothing
        End If
    End If
End Function

Function GetStatus() As String
    'when 'FO' then 'FOR ORDERING'
    'when 'BO' then 'BACK ORDER STAGE'
    'when 'AS' then 'ALLOCATION STAGE'
    'when 'KS' then 'PICKING STAGE'
    'when 'PS' then 'PACKING STAGE'
    'when 'SS' then 'SHIPPING STAGE'

    Select Case cboPartsEstimate_Status
        Case "FOR ORDERING"
            GetStatus = "FO"
        Case "BACK ORDER STAGE"
            GetStatus = "BO"
        Case "ALLOCATION STAGE"
            GetStatus = "AS"
        Case "PICKING STAGE"
            GetStatus = "KS"
        Case "PACKING STAGE"
            GetStatus = "PS"
        Case "SHIPPING STAGE"
            GetStatus = "SS"
    End Select
End Function

Function ItemExists(StringToFind As String, ColumnToLook As Integer) As Integer
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, ColumnToLook) = StringToFind Then
            ItemExists = ItemExists + 1
            Exit For
        End If
    Next
End Function

Function SelectCombo(c As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If c.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                                                             As Long
    Dim ItemDataX                                                     As Long
    If ByItemData = False Then
        For i = 0 To c.ListCount - 1
            If UCase(c.List(i)) = UCase(Trim(STR)) Then
                SelectCombo = i
                Exit Function
            End If
        Next
    Else
        If STR = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If

        ItemDataX = CLng(STR)

        For i = 0 To c.ListCount - 1
            If c.ItemData(i) = STR Then
                SelectCombo = i
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function

Function SetStatus(XString) As String
    'when 'FO' then 'FOR ORDERING'
    'when 'BO' then 'BACK ORDER STAGE'
    'when 'AS' then 'ALLOCATION STAGE'
    'when 'KS' then 'PICKING STAGE'
    'when 'PS' then 'PACKING STAGE'
    'when 'SS' then 'SHIPPING STAGE'

    Select Case XString
        Case "FO"
            SetStatus = "FOR ORDERING"
        Case "BO"
            SetStatus = "BACK ORDER STAGE"
        Case "AS"
            SetStatus = "ALLOCATION STAGE"
        Case "KS"
            SetStatus = "PICKING STAGE"
        Case "PS"
            SetStatus = "PACKING STAGE"
        Case "SS"
            SetStatus = "SHIPPING STAGE"
    End Select
End Function

Sub AddDetails()
    ADDOREDIT = "ADD"
    txtDPIDetailID = 0
    cmdDel_Tech.Enabled = False
    cmdDelPriceInquiry.Enabled = False
    cmdDel_ETA.Enabled = False
    Select Case dpiInqType
        Case "TECH"
            ShowHidePictureBox picInquiryTechincal.hwnd, True
            cboTech_PartNo.Enabled = True
            txtTech_PartName.Enabled = True
            cboTech_PartNo = vbNullString
            txtTech_GroupNumber = vbNullString
            txtTech_PartName = vbNullString
            txtTech_PNC = vbNullString
            ShortcutCaption4.Caption = "::TECHNICAL INQUIRY:: ADD MODE"
            On Error Resume Next
            cboTech_PartNo.SetFocus
        Case "PRCS"
            ShowHidePictureBox picInquiryPrice.hwnd, True
            cboPrice_PartNo.Enabled = True
            txtPrice_PartName.Enabled = False
            'txtPrice_PartName = ""
            txtPrice_Reg_DP = 0
            txtPrice_Reg_SRP = 0
            txtPrice_SAO_DP = 0
            txtPrice_SAO_SRP = 0
            txtPrice_SSO_DP = 0
            txtPrice_SSO_SRP = 0
            On Error Resume Next
            cboPrice_PartNo.SetFocus
        Case "ETA"
            cboETA_PNO = vbNullString
            cboETA_OrderNo = vbNullString
            txtETA_PARTNAME = vbNullString
            txtETA_QTY = 1
            ShowHidePictureBox picInquiryEstimateArrival.hwnd, True
            On Error Resume Next
            cboETA_PNO.SetFocus
    End Select
End Sub

Sub EditDetails()
    cmdDel_Tech.Enabled = True
    cmdDelPriceInquiry.Enabled = True
    cmdDel_ETA.Enabled = True
    ADDOREDIT = "EDIT"

    Select Case dpiInqType
        Case "TECH"
            txtDPIDetailID = Grid1.TextMatrix(Grid1.Row, 2)
            ShowHidePictureBox picInquiryTechincal.hwnd, True
            cboTech_PartNo.ListIndex = SelectCombo(cboTech_PartNo, Grid1.TextMatrix(Grid1.Row, 0))
            txtTech_PartName.Text = Grid1.TextMatrix(Grid1.Row, 1)
            'txtTech_GroupNumber = Grid1.TextMatrix(Grid1.Row, 2)
            'txtTech_PNC = Grid1.TextMatrix(Grid1.Row, 3)
            ShortcutCaption4.Caption = "EDIT TECHNICAL INQUIRY"
            On Error Resume Next
            cboTech_PartNo.SetFocus
        Case "PRCS"
            txtDPIDetailID = Grid1.TextMatrix(Grid1.Row, 2)
            cboPrice_PartNo.ListIndex = SelectCombo(cboPrice_PartNo, Grid1.TextMatrix(Grid1.Row, 0))
            txtPrice_PartName.Text = Grid1.TextMatrix(Grid1.Row, 1)
            'txtPrice_Reg_DP = Grid1.TextMatrix(Grid1.Row, 2)
            'txtPrice_Reg_SRP = Grid1.TextMatrix(Grid1.Row, 3)
            'txtPrice_SAO_DP = Grid1.TextMatrix(Grid1.Row, 4)
            'txtPrice_SAO_SRP = Grid1.TextMatrix(Grid1.Row, 5)
            'txtPrice_SSO_DP = Grid1.TextMatrix(Grid1.Row, 6)
            'txtPrice_SSO_SRP = Grid1.TextMatrix(Grid1.Row, 7)
            ShortcutCaption3.Caption = "EDIT PRICE INQUIRY "
            ShowHidePictureBox picInquiryPrice.hwnd, True
            On Error Resume Next
            cboPrice_PartNo.SetFocus
        Case "ETA"
            txtDPIDetailID = Grid1.TextMatrix(Grid1.Row, 5)
            cboETA_PNO.ListIndex = SelectCombo(cboETA_PNO, Grid1.TextMatrix(Grid1.Row, 0))
            txtETA_PARTNAME = Grid1.TextMatrix(Grid1.Row, 1)
            cboETA_OrderNo = Grid1.TextMatrix(Grid1.Row, 2)
            If IsDate(Grid1.TextMatrix(Grid1.Row, 3)) = True Then
                txtETA_DateOrd = Grid1.TextMatrix(Grid1.Row, 3)
            End If
            txtETA_QTY = Grid1.TextMatrix(Grid1.Row, 4)
            ' cboPartsEstimate_Status.ListIndex = SelectCombo(cboPartsEstimate_Status, UCase(Grid1.TextMatrix(Grid1.Row, 5)))

            'If Grid1.TextMatrix(Grid1.Row, 6) <> "" Then
            '    dtPartsEstimate_ETA = Grid1.TextMatrix(Grid1.Row, 6)
            'End If
            ShortcutCaption2.Caption = "EDIT PARTS ESTIMATE TIME OF ARRIVAL"
            ShowHidePictureBox picInquiryEstimateArrival.hwnd, True

            On Error Resume Next
            cboETA_PNO.SetFocus
    End Select

End Sub

Sub FillGrid()

    Dim recRs                                                         As ADODB.Recordset
    Set recRs = gconDMIS.Execute("Select * from PMIS_DPIDetails WHERE DPINO=" & N2Str2Null(txtDPINo))
    Grid1.Rows = 1

    If recRs.EOF Or recRs.BOF Then
        
    Else

        While Not recRs.EOF
            With Grid1
                Select Case dpiInqType
                    Case "TECH"
                        Grid1.AddItem recRs!PART_NO & Chr(9) _
                                    & recRs!PART_NAME & Chr(9) _
                                    & recRs!ID


                    Case "PRCS"
                        Grid1.AddItem recRs!PART_NO & Chr(9) _
                                    & recRs!PART_NAME & Chr(9) _
                                    & recRs!ID

                    Case "ETA"
                        Grid1.AddItem recRs!PART_NO & Chr(9) _
                                    & recRs!PART_NAME & Chr(9) _
                                    & recRs!ORDER_NO & Chr(9) _
                                    & recRs!DATE_ORDERED & Chr(9) _
                                    & recRs!QTY & Chr(9) _
                                    & recRs!ID
                End Select
            End With
            recRs.MoveNext
        Wend
    End If
End Sub

Sub FillParts()
    Dim PART_NO                                                       As String
    Dim PART_ID                                                       As Long

    Dim SQL                                                           As String

    SQL = "Select ID, STOCKDESC , STOCKNO from PMIS_StockMas WHERE TYPE='P' and NON_HARI='N' ORDER BY STOCKNO  "
    Set RSPARTS = New ADODB.Recordset
    Call RSPARTS.Open(SQL, gconDMIS, adOpenKeyset, adLockReadOnly)

    While Not RSPARTS.EOF
        PART_NO = Null2String(RSPARTS!STOCKNO)
        PART_ID = RSPARTS!ID
        cboTech_PartNo.AddItem PART_NO
        cboTech_PartNo.ItemData(cboTech_PartNo.NewIndex) = PART_ID

        cboPrice_PartNo.AddItem PART_NO
        cboPrice_PartNo.ItemData(cboPrice_PartNo.NewIndex) = PART_ID



        RSPARTS.MoveNext
    Wend
    cboETA_OrderNo.Clear
    Combo_Loadval cboETA_OrderNo, gconDMIS.Execute("SELECT DISTINCT DON FROM PMIS_PO_HIST WHERE DON IS NOT NULL UNION SELECT DISTINCT DON FROM PMIS_PO_HD WHERE DON IS NOT NULL")


    'cboETA_PNO.AddItem PART_NO
    '    cboETA_PNO.ItemData(cboETA_PNO.NewIndex) = PART_ID
End Sub

Sub FillSearchGrid()
    Dim TEMPRS                                                        As ADODB.Recordset

    lstDPIList.Enabled = False

    If optVIN.Value = True Then
        Set TEMPRS = gconDMIS.Execute("Select DPI_Date, DPINO,ID from PMIS_DPIHeader WHERE VINNO Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
    ElseIf optCatalgoue.Value = True Then
        Set TEMPRS = gconDMIS.Execute("Select DPI_Date, DPINO,ID from PMIS_DPIHeader WHERE CATALOGUENO Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
    ElseIf optDate.Value = True Then
        Set TEMPRS = gconDMIS.Execute("Select DPI_Date, DPINO,ID from PMIS_DPIHeader WHERE DPI_Date Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
    ElseIf optDPI.Value = True Then
        Set TEMPRS = gconDMIS.Execute("Select DPI_Date, DPINO,ID from PMIS_DPIHeader WHERE DPINO Like " & N2Str2Null(ReplaceQuote(textSearch & "%")))
    End If

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        Listview_Loadval lstDPIList.ListItems, TEMPRS
        lstDPIList.Enabled = True
    End If
    Set TEMPRS = Nothing
End Sub

Sub SetCompany()
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select CompanyName, CompanyAddress,PreparedBy,ApprovedBy from ALL_PRofile Where ModuleName='PMIS'")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        CompName = Null2String(TEMPRS!CompanyName)
        NotedBy = Null2String(TEMPRS!PreparedBy)
        ReqBy = Null2String(TEMPRS!ApprovedBy)
    End If

End Sub

Sub InitGrid()
    Select Case dpiInqType
        Case "TECH", "PRCS"
            With Grid1

                .Rows = 1
                .Cols = 3
                .FormatString = "PART NUMBER" & Chr(9) & "PART NAME"
                .ColWidth(0) = .Width * 0.28
                .ColWidth(1) = .Width * 0.7
                .ColAlignment(0) = 0
            End With

        Case "ETA"
            With Grid1
                .Rows = 1
                .Cols = 6
                .FormatString = "PART NUMBER" & Chr(9) & "PART NAME" & Chr(9) & "ORDER NUMBER" & Chr(9) & "DATE ORDERED" & Chr(9) & "QUANTITY"
                .ColWidth(0) = .Width * 0.18
                .ColWidth(1) = .Width * 0.35
                .ColWidth(2) = .Width * 0.18
                .ColWidth(3) = .Width * 0.18
                .ColWidth(4) = .Width * 0.12
                .ColWidth(5) = .Width * 0.1

                .ColAlignment(0) = 0
                .ColAlignment(4) = 3
            End With
    End Select
    Grid1.ColWidth(Grid1.Cols - 1) = 0
End Sub

Sub InitMemVars()
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            cntrl.Text = vbNullString
        End If
    Next
    txtDPIDetailID = 0
    txtID = 0
    lblStatus = ""
    txtQty = 1
    txtCompanyName = CompName
    txtReqBy = ReqBy
    txtNotedBy = NotedBy
    cleargrid Grid1
    txtDateDPI2 = dtDateDPI

End Sub

Sub rsRefresh()
    Set RsDPIR = New ADODB.Recordset
    Call RsDPIR.Open("SELECT  * FROM PMIS_DPIHeader", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub SetPartsLines(PartIDNo As Variant, ForCombo As Boolean)
    Dim TEMPRS                                                        As ADODB.Recordset
    If ForCombo = False Then
        Set TEMPRS = gconDMIS.Execute("SELECT  SRP, STOCKNO , STOCKDESC FROM PMIS_STOCKMAS WHERE ID=" & PartIDNo)
        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            txtUnitAmount = FormatNumber(NumericVal(TEMPRS!SRP))
            txtPartNo = Null2String(TEMPRS!STOCKNO)
            txtPART_NUMBER = Null2String(TEMPRS!STOCKNO)
            txtPartDescription = Null2String(TEMPRS!STOCKDESC)
        End If
    Else
        Set TEMPRS = gconDMIS.Execute("SELECT  SRP,  STOCKDESC , STOCKNO FROM PMIS_STOCKMAS WHERE STOCKNO=" & N2Str2Null(PartIDNo))
        If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
            txtUnitAmount = FormatNumber(NumericVal(TEMPRS!SRP))
            cboPART_NAME.Text = Null2String(TEMPRS!STOCKDESC)
            txtPART_NUMBER = Null2String(TEMPRS!STOCKNO)
            txtPartDescription = Null2String(TEMPRS!STOCKDESC)
        End If
    End If
End Sub

Sub ShowHidePictureBox(hwnd As Long, State As Boolean)
    Dim cntl                                                          As Control
    For Each cntl In Me.ControlS
        If TypeOf cntl Is PictureBox Then
            If Not cntl.Container.hwnd = hwnd Then
                If cntl.hwnd = hwnd Then
                    cntl.Enabled = State
                    cntl.Visible = State
                    If State = True Then
                        cntl.ZOrder 0
                    Else
                        cntl.ZOrder 1
                    End If
                Else

                    cntl.Enabled = Not (State)
                    If State = True Then
                    Else
                    End If
                End If
            End If
        End If
    Next

End Sub

Sub StoreMemvars()
    If Not (RsDPIR.EOF Or RsDPIR.BOF) Then
        cmdEdit.Enabled = True
        txtDateDPI2 = Null2String(RsDPIR!DPI_DATE)
        txtDPINo = Null2String(RsDPIR!DPIno)
        txtVinNo = Null2String(RsDPIR!VINNO)
        txtCatalgoueNo = Null2String(RsDPIR!CATALOGUENO)
        txtSubject = Null2String(RsDPIR!Subject)
        txtID = Null2String(RsDPIR!ID)
        If Null2String(RsDPIR!Enclose) = True Then
            chkEnclose = 1
        End If
        dpiInqType = Null2String(RsDPIR!DPI_INQTYPE)
        dpiSeq = Null2String(RsDPIR!DPI_SEQNO)

        txtNotedBy = Null2String(RsDPIR!NotedBy)
        txtReqBy = Null2String(RsDPIR!ReqBy)
        If Null2String(RsDPIR!STATUS) = "C" Then
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            lblStatus = "***Cancelled***"
            cmdEdit.Enabled = False
        ElseIf Null2String(RsDPIR!STATUS) = "P" Then
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = True
            cmdPost.Enabled = False
            lblStatus = "***Posted ***"
            cmdEdit.Enabled = False
            cmdPrint.Enabled = True
        Else
            cmdCancelCO.Enabled = True
            cmdUnPost.Enabled = False
            cmdPost.Enabled = True
            lblStatus = ""
            cmdEdit.Enabled = True
            cmdPrint.Enabled = False
        End If
        InitGrid
        FillGrid
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

Private Sub cboETA_OrderNo_Change()

    Dim PARTDET                                                       As ADODB.Recordset
    Dim RSPOHIST                                                      As ADODB.Recordset
    Set RSPOHIST = gconDMIS.Execute("SELECT PONO ,PODATE FROM PMIS_PO_HIST WHERE DON='" & Repleys(cboETA_OrderNo) & "' UNION SELECT PONO,PODATE FROM PMIS_PO_HD WHERE DON='" & Repleys(cboETA_OrderNo) & "'")

    If Not (RSPOHIST.EOF Or RSPOHIST.BOF) Then
        If IsDate(RSPOHIST!PODATE) = True Then
            txtETA_DateOrd.Value = RSPOHIST!PODATE
        End If
        Set PARTDET = gconDMIS.Execute("SELECT STOCK_SUP FROM PMIS_ALLDAYTRAN WHERE trantype='PO' AND TYPE='P' AND STATUS='P' AND TRANNO =" & N2Str2Null(RSPOHIST!PONO))
        Combo_Loadval cboETA_PNO, PARTDET
    End If


End Sub

Private Sub cboETA_OrderNo_Click()
    cboETA_OrderNo_Change
End Sub

Private Sub cboETA_PNO_Change()

    RSPARTS.MoveFirst
    RSPARTS.Find ("STOCKNO=" & N2Str2Null(Repleys(cboETA_PNO.Text)))
    If Not RSPARTS.EOF Or RSPARTS.BOF Then
        txtETA_PARTNAME = RSPARTS!STOCKDESC
    End If
End Sub

Private Sub cboETA_PNO_Click()
    cboETA_PNO_Change
End Sub

Private Sub cboPrice_PartNo_Change()
    If cboPrice_PartNo.ListIndex = -1 Then Exit Sub
    If RSPARTS.RecordCount > 0 Then
        RSPARTS.MoveFirst
        RSPARTS.Find ("ID=" & cboPrice_PartNo.ItemData(cboPrice_PartNo.ListIndex))
        If Not RSPARTS.EOF Or RSPARTS.BOF Then
            txtPrice_PartName = RSPARTS!STOCKDESC
        End If
    End If
End Sub

Private Sub cboPrice_PartNo_Click()

    cboPrice_PartNo_Change
End Sub

Private Sub cboTech_PartNo_Change()

    If cboTech_PartNo.ListIndex = -1 Then Exit Sub

    RSPARTS.MoveFirst

    RSPARTS.Find ("ID=" & cboTech_PartNo.ItemData(cboTech_PartNo.ListIndex))

    If Not RSPARTS.EOF Or RSPARTS.BOF Then
        txtTech_PartName = RSPARTS!STOCKDESC
    End If
End Sub

Private Sub cboTech_PartNo_Click()
    cboTech_PartNo_Change
End Sub

Private Sub cmdAdd_Click()
    InitMemVars
    ADDOREDIT = "ADD"
    picAdds.Visible = False
    picSaves.Visible = True
    picTop.Enabled = True
    dtDateDPI = LOGDATE
    txtID = 0

End Sub

Private Sub cmdCancel_Click()
    ADDOREDIT = ""
    picAdds.Visible = True
    picSaves.Visible = False
    picTop.Enabled = False
    Grid1.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdCancelCO_Click()
    On Error GoTo Errorcode:
    If MsgBox("Do you Want to Cancel this Transaction ", vbOKCancel, "Confirm Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = True
    gconDMIS.Execute ("UPDate PMIS_DPIHeader Set Status='C'  Where ID=" & txtID)
    LogAudit "C", "DPI", txtDPINo
    rsRefresh
    RsDPIR.Find ("ID=" & txtID)
    StoreMemvars
    MessagePop RecSaveOk, "Transaction Cancelled", "Transaction Sucessfully Cancelled"
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_ETA_Click()
    ShowHidePictureBox picInquiryEstimateArrival.hwnd, False
End Sub

Private Sub cmdCancelInquiry_Price_Click()
    ShowHidePictureBox picInquiryPrice.hwnd, False
End Sub

Private Sub cmdCancel_Tech_Click()
    ShowHidePictureBox picInquiryTechincal.hwnd, False
End Sub

Private Sub cmdDel_ETA_Click()

    Form_KeyDown 116, 1

End Sub

Private Sub cmdDelPriceInquiry_Click()
    Form_KeyDown 116, 1
End Sub

Private Sub cmdDel_Tech_Click()
    Form_KeyDown 116, 1
End Sub

Private Sub cmdEdit_Click()
    If NumericVal(txtID) <> 0 Then
        ADDOREDIT = "EDIT"
        picAdds.Visible = False
        picSaves.Visible = True
        picTop.Enabled = True
        txtDPINo.Locked = False
        Grid1.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    RsDPIR.MoveNext

    If RsDPIR.EOF Then
        RsDPIR.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars



End Sub

Private Sub cmdOkInquiry_Price_Click()
    If txtPrice_PartName = "" Then
        On Error Resume Next
        txtPrice_PartName.SetFocus
        Exit Sub
    ElseIf cboPrice_PartNo = "" Then
        On Error Resume Next

        cboPrice_PartNo.SetFocus
        Exit Sub
    End If


    Dim ijx                                                           As Integer

    ijx = ItemExists(cboPrice_PartNo, 0)


    If (ijx > 0 And txtDPIDetailID = 0) Or (ijx >= 1 And Grid1.Text <> cboPrice_PartNo And txtDPIDetailID <> 0) Then
        MsgBox "Part Number Already used in this Transaction.", vbInformation, "Duplicate Entry"
        txtPrice_PartName = vbNullString
        cboPrice_PartNo.SetFocus
        Exit Sub
    ElseIf ijx = 0 And txtDPIDetailID = 0 Then
        SQL = "INSERT INTO PMIS_DPIDetails (DPINO, PART_NO,PART_NAME,REG_DP,REG_SRP,SP_DP,SP_SRP,SSP_DP,SSP_SRP) VALUES ("
        SQL = SQL & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboPrice_PartNo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtPrice_PartName) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_Reg_DP) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_Reg_SRP) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_SAO_DP) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_SAO_SRP) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_SSO_DP) & "," & vbCrLf
        SQL = SQL & NumericVal(txtPrice_SSO_SRP) & ")"
        gconDMIS.Execute SQL

    ElseIf (ijx = 1 And Grid1.Text = cboPrice_PartNo) Or (ijx = 0 And Grid1.Text <> cboPrice_PartNo) And txtDPIDetailID <> 0 Then

        SQL = "Update PMIS_DPIDetails SET " & vbCrLf
        SQL = SQL & "DPINO=" & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & "PART_NO=" & N2Str2Null(cboPrice_PartNo) & "," & vbCrLf
        SQL = SQL & "PART_NAME=" & N2Str2Null(txtPrice_PartName) & "," & vbCrLf
        SQL = SQL & "REG_DP=" & NumericVal(txtPrice_Reg_DP) & "," & vbCrLf
        SQL = SQL & "REG_SRP=" & NumericVal(txtPrice_Reg_SRP) & "," & vbCrLf
        SQL = SQL & "SP_DP=" & NumericVal(txtPrice_SAO_DP) & "," & vbCrLf
        SQL = SQL & "SP_SRP=" & NumericVal(txtPrice_SAO_SRP) & "," & vbCrLf
        SQL = SQL & "SSP_DP=" & NumericVal(txtPrice_SSO_DP) & "," & vbCrLf
        SQL = SQL & "SSP_SRP=" & NumericVal(txtPrice_SSO_SRP) & vbCrLf
        SQL = SQL & "Where ID= " & txtDPIDetailID
        gconDMIS.Execute SQL
        ShowHidePictureBox picInquiryPrice.hwnd, False


    Else
        MsgBox "Else Statements"
    End If



    '''''''''''


    txtPrice_PartName = vbNullString
    cboPrice_PartNo = vbNullString
    txtPrice_Reg_DP = FormatNumber(0)
    txtPrice_Reg_SRP = FormatNumber(0)
    txtPrice_SAO_DP = FormatNumber(0)
    txtPrice_SAO_SRP = FormatNumber(0)
    txtPrice_SSO_DP = FormatNumber(0)
    txtPrice_SSO_SRP = FormatNumber(0)

    InitGrid
    FillGrid

End Sub

Private Sub cmdPost_Click()
    On Error GoTo Errorcode:
    If MsgBox("Do you want to Post this Transaction?", vbOKCancel + vbInformation, "Confirm Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = False
    gconDMIS.Execute ("UPDATE PMIS_DPIHeader Set Status='P'  Where ID=" & txtID)
    rsRefresh
    LogAudit "P", "DPI", txtDPINo
    RsDPIR.Find ("ID=" & txtID)
    StoreMemvars
    MessagePop RecSaveOk, "Transaction Posted", "Transaction Sucessfully Posted"
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrevious_Click()

    RsDPIR.MovePrevious

    If RsDPIR.BOF Then
        RsDPIR.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()

    If Len(Dir(App.Path & "\DPI.xlt")) <= 0 Then
        If EXTRACT_FILES(103, "DPI.xlt") = False Then
            MsgBox "Please Put DPI.xlt on " & vbCrLf & App.Path, vbInformation
            Exit Sub
        End If
    End If

    Screen.MousePointer = 11
    Dim vDPINO                                         As String
    Dim vDPIDate                                       As String
    Dim vDPIGroupNo                                    As String
    Dim vDPIPartNo                                     As String
    Dim vDPIPartName                                   As String
    Dim vDPIPNC                                        As String
    Dim vDPIOrderNo                                    As String
    Dim vDPIDateOrdered                                As String
    Dim vDPIQty                                        As String
    Dim vDPIVINNO                                      As String
    Dim vDPICATALOGNUMBER                              As String
    Dim vDPISUBJECT                                    As String
    Dim vDPIReqBy                                      As String
    Dim vDPINotedBy                                    As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\DPI.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    Dim VIN_1, VIN_2, VIN_3, VIN_4, VIN_5, VIN_6, VIN_7, VIN_8, VIN_9, VIN_10, VIN_11, VIN_12, VIN_13 As String
    Dim CAT_1, CAT_2, CAT_3, CAT_4, CAT_5, CAT_6, CAT_7, CAT_8, CAT_9, CAT_10 As String
    Dim rsDPI                                          As ADODB.Recordset
    Dim rsDPIDetails                                   As ADODB.Recordset
    Set rsDPI = New ADODB.Recordset
    Set rsDPI = gconDMIS.Execute("Select * from PMIS_DPIHeader where DPINo= '" & txtDPINo & "'")

    If Not rsDPI.EOF And Not rsDPI.BOF Then

        vDPINO = Null2String(rsDPI!DPIno)
        vDPIDate = Null2String(rsDPI!DPI_DATE)
        vDPIVINNO = Null2String(rsDPI!VINNO)
        vDPICATALOGNUMBER = Null2String(rsDPI!CATALOGUENO)
        vDPISUBJECT = Null2String(rsDPI!Subject)
        vDPIReqBy = Null2String(rsDPI!ReqBy)
        vDPINotedBy = Null2String(rsDPI!NotedBy)

        'Get VIN Number per Letter
        VIN_1 = Mid(vDPIVINNO, 1, 1): VIN_2 = Mid(vDPIVINNO, 2, 1): VIN_3 = Mid(vDPIVINNO, 3, 1)
        VIN_4 = Mid(vDPIVINNO, 4, 1): VIN_5 = Mid(vDPIVINNO, 5, 1): VIN_6 = Mid(vDPIVINNO, 6, 1)
        VIN_7 = Mid(vDPIVINNO, 7, 1): VIN_8 = Mid(vDPIVINNO, 8, 1): VIN_9 = Mid(vDPIVINNO, 9, 1)
        VIN_10 = Mid(vDPIVINNO, 10, 1): VIN_11 = Mid(vDPIVINNO, 11, 1): VIN_12 = Mid(vDPIVINNO, 12, 1)
        VIN_13 = Mid(vDPIVINNO, 13, 1): VIN_14 = Mid(vDPIVINNO, 14, 1): VIN_15 = Mid(vDPIVINNO, 15, 1)
        VIN_16 = Mid(vDPIVINNO, 16, 1): VIN_17 = Mid(vDPIVINNO, 17, 1)

        'Get CATALOGUE Number per Letter
        CAT_1 = Mid(vDPICATALOGNUMBER, 1, 1): CAT_2 = Mid(vDPICATALOGNUMBER, 2, 1): CAT_3 = Mid(vDPICATALOGNUMBER, 3, 1)
        CAT_4 = Mid(vDPICATALOGNUMBER, 4, 1): CAT_5 = Mid(vDPICATALOGNUMBER, 5, 1): CAT_6 = Mid(vDPICATALOGNUMBER, 6, 1)
        CAT_7 = Mid(vDPICATALOGNUMBER, 7, 1): CAT_8 = Mid(vDPICATALOGNUMBER, 8, 1): CAT_9 = Mid(vDPICATALOGNUMBER, 9, 1)
        CAT_10 = Mid(vDPICATALOGNUMBER, 10, 1)

        'Header
        xlSheet.Shapes("RECTANGLE 7").TextFrame.Characters(1, 255).Text = "DEALER NAME: " & txtCompanyName
        xlSheet.Cells(2, "EB") = vDPINO
        xlSheet.Cells(4, "EB") = vDPIDate
        'VIN NO
        xlSheet.Cells(10, "P") = VIN_1
        xlSheet.Cells(10, "S") = VIN_2
        xlSheet.Cells(10, "V") = VIN_3
        xlSheet.Cells(10, "Y") = VIN_4
        xlSheet.Cells(10, "AB") = VIN_5
        xlSheet.Cells(10, "AE") = VIN_6
        xlSheet.Cells(10, "AH") = VIN_7
        xlSheet.Cells(10, "AK") = VIN_8
        xlSheet.Cells(10, "AN") = VIN_9
        xlSheet.Cells(10, "AQ") = VIN_10
        xlSheet.Cells(10, "AT") = VIN_11
        xlSheet.Cells(10, "AW") = VIN_12
        xlSheet.Cells(10, "AZ") = VIN_13
        xlSheet.Cells(10, "BC") = VIN_14
        xlSheet.Cells(10, "BF") = VIN_15
        xlSheet.Cells(10, "BI") = VIN_16
        xlSheet.Cells(10, "BL") = VIN_17

        'CATALOGUE
        xlSheet.Cells(10, "BO") = CAT_1
        xlSheet.Cells(10, "BR") = CAT_2
        xlSheet.Cells(10, "BU") = CAT_3
        xlSheet.Cells(10, "BX") = CAT_4
        xlSheet.Cells(10, "CA") = CAT_5
        xlSheet.Cells(10, "CD") = CAT_6
        xlSheet.Cells(10, "CG") = CAT_7
        xlSheet.Cells(10, "CJ") = CAT_8
        xlSheet.Cells(10, "CM") = CAT_9
        xlSheet.Cells(10, "CP") = CAT_10

        xlSheet.Cells(10, "CS") = vDPISUBJECT
        'Footer
        xlSheet.Cells(41, "EP") = vDPIReqBy
        xlSheet.Cells(45, "EP") = vDPINotedBy


        '=====================================================
        'DPI Details
        Dim Row_Number                                 As Integer
        Dim PART_1, PART_2, PART_3, PART_4, PART_5, PART_6, PART_7, PART_8, PART_9, PART_10, PART_11, PART_12, PART_13 As String
        Set rsDPIDetails = New ADODB.Recordset
        Set rsDPIDetails = gconDMIS.Execute("Select * from PMIS_DPIDetails where DPINo= '" & txtDPINo & "'")
        If Not rsDPIDetails.EOF And Not rsDPIDetails.BOF Then
            'Details
            Row_Number = 14
            Do While Not rsDPIDetails.EOF
                vDPIGroupNo = Null2String(rsDPIDetails!GP_NO)
                vDPIPNC = Null2String(rsDPIDetails!PNC)
                vDPIPartNo = Null2String(rsDPIDetails!PART_NO)

                'Get PART Number per Letter
                PART_1 = Mid(vDPIPartNo, 1, 1): PART_2 = Mid(vDPIPartNo, 2, 1): PART_3 = Mid(vDPIPartNo, 3, 1)
                PART_4 = Mid(vDPIPartNo, 4, 1): PART_5 = Mid(vDPIPartNo, 5, 1): PART_6 = Mid(vDPIPartNo, 6, 1)
                PART_7 = Mid(vDPIPartNo, 7, 1): PART_8 = Mid(vDPIPartNo, 8, 1): PART_9 = Mid(vDPIPartNo, 9, 1)
                PART_10 = Mid(vDPIPartNo, 10, 1): PART_11 = Mid(vDPIPartNo, 11, 1): PART_12 = Mid(vDPIPartNo, 12, 1)
                PART_13 = Mid(vDPIPartNo, 13, 1)

                xlSheet.Cells(Row_Number, "T") = PART_1
                xlSheet.Cells(Row_Number, "V") = PART_2
                xlSheet.Cells(Row_Number, "X") = PART_3
                xlSheet.Cells(Row_Number, "Z") = PART_4
                xlSheet.Cells(Row_Number, "AB") = PART_5
                xlSheet.Cells(Row_Number, "AD") = PART_6
                xlSheet.Cells(Row_Number, "AF") = PART_7
                xlSheet.Cells(Row_Number, "AH") = PART_8
                xlSheet.Cells(Row_Number, "AJ") = PART_9
                xlSheet.Cells(Row_Number, "AL") = PART_10
                xlSheet.Cells(Row_Number, "AN") = PART_11
                xlSheet.Cells(Row_Number, "AP") = PART_12
                xlSheet.Cells(Row_Number, "AR") = PART_13


                vDPIPartName = Null2String(rsDPIDetails!PART_NAME)
                'vDPIRegular_DP = Null2String(rsDPIDetails!REG_DP)
                'vDPIRegular_SRP = Null2String(rsDPIDetails!REG_SRP)
                'vDPISpecialAirOrder_DP = Null2String(rsDPIDetails!SP_DP)
                'vDPISpecialAirOrder_SRP = Null2String(rsDPIDetails!SP_SRP)
                'vDPISpecialSeaOrder_DP = Null2String(rsDPIDetails!SSP_DP)
                'vDPISpecialSeaOrder_SRP = Null2String(rsDPIDetails!SSP_SRP)
                vDPIOrderNo = Null2String(rsDPIDetails!ORDER_NO)
                vDPIDateOrdered = Null2String(rsDPIDetails!DATE_ORDERED)
                vDPIQty = Null2String(rsDPIDetails!QTY)
                'vDPIStatus = Null2String(rsDPIDetails!Status)
                'vDPIEstTime_Arrival = Null2String(rsDPIDetails!ESTARRIVAL)

                Select Case dpiInqType
                    Case "PRCS"
                        xlSheet.Cells(Row_Number, "A") = vDPIGroupNo
                        xlSheet.Cells(Row_Number, "K") = vDPIPNC
                        xlSheet.Cells(Row_Number, "AT") = vDPIPartName
                    Case "TECH"
                        xlSheet.Cells(Row_Number, "A") = vDPIGroupNo
                        xlSheet.Cells(Row_Number, "K") = vDPIPNC
                        xlSheet.Cells(Row_Number, "AT") = vDPIPartName
                        'xlSheet.Cells(Row_Number, "BZ") = vDPIRegular_DP
                        'xlSheet.Cells(Row_Number, "CJ") = vDPIRegular_SRP
                        'xlSheet.Cells(Row_Number, "CT") = vDPISpecialAirOrder_DP
                        'xlSheet.Cells(Row_Number, "DD") = vDPISpecialAirOrder_SRP
                        'xlSheet.Cells(Row_Number, "DN") = vDPISpecialSeaOrder_DP
                        'xlSheet.Cells(Row_Number, "DX") = vDPISpecialSeaOrder_SRP
                    Case "ETA"
                        xlSheet.Cells(Row_Number, "EH") = vDPIOrderNo
                        xlSheet.Cells(Row_Number, "ES") = vDPIDateOrdered
                        xlSheet.Cells(Row_Number, "FC") = vDPIQty
                        'xlSheet.Cells(Row_Number, "FH") = vDPIStatus
                        'xlSheet.Cells(Row_Number, "FW") = vDPIEstTime_Arrival
                End Select
                Row_Number = Row_Number + 1
                rsDPIDetails.MoveNext
            Loop
        End If
        xlApp.Visible = True
        Set xlApp = Nothing
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdUnPost_Click()
    On Error GoTo Errorcode:
    If MsgBox("Do you Want to Unpost this Transaction ", vbOKCancel + vbInformation, "Confirm Un-Posting") = vbCancel Then Exit Sub
    cmdCancelCO.Enabled = True
    gconDMIS.Execute ("UPDate PMIS_DPIHeader Set Status='U' Where ID=" & txtID)
    LogAudit "U", "DPI", txtDPINo
    rsRefresh
    RsDPIR.Find ("ID=" & txtID)
    StoreMemvars
    MessagePop RecSaveOk, "Transaction Unposted", "Transaction Sucessfully Unposted"
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Command1_Click()

    If olddpiInqType <> dpiInqType And ADDOREDIT = "EDIT" Then
        cmdSave.Value = True
        ShowHidePictureBox picDPI.hwnd, False
    Else
        ShowHidePictureBox picDPI.hwnd, False
    End If
    picSaves.Visible = True
    picAdds.Visible = False
End Sub

Private Sub cmdDPINo_Click()
    ShowHidePictureBox picDPI.hwnd, True
    If txtID = 0 Then
        ADDOREDIT = "ADD"
        dtDateDPI = LOGDATE
    Else
        ADDOREDIT = "EDIT"
        If IsDate(RsDPIR!DPI_DATE) = True Then
            dtDateDPI = RsDPIR!DPI_DATE
        End If

    End If
End Sub

Private Sub Command3_Click()
    ShowHidePictureBox picDPI.hwnd, False
    cmdCancel.Value = True
End Sub

Private Sub cmdSave_Click()

    On Error GoTo Errorcode:

    Dim TEMPRS                                                        As ADODB.Recordset
    Dim vDPI_DATE, vDPI_INQTYPE, vDPI_SEQNO, vVINNO, vCATALOGUENO, vSubject, vEnclose, vReqBy, vNotedBy, vDPINO

    vDPI_DATE = N2Str2Null(txtDateDPI2)
    vDPI_INQTYPE = N2Str2Null(dpiInqType)
    vDPI_SEQNO = N2Str2Null(dpiSeq)
    vVINNO = N2Str2Null(txtVinNo)
    vCATALOGUENO = N2Str2Null(txtCatalgoueNo)
    vSubject = N2Str2Null(txtSubject)
    vReqBy = N2Str2Null(txtReqBy)
    vNotedBy = N2Str2Null(txtNotedBy)
    vDPINO = N2Str2Null(txtDPINo)
    If LTrim(RTrim(txtDPINo)) = "" Then
        ShowIsRequiredMsg "DPI No"
        cmdDPINo.SetFocus
        Exit Sub
    End If
    If chkEnclose.Value = 1 Then
        vEnclose = 1
    Else
        vEnclose = 0
    End If

    If ADDOREDIT = "ADD" Then
        SQL = "INSERT INTO PMIS_DPIHEADER("
        SQL = SQL & "DPINo,DPI_DATE,DPI_INQTYPE, DPI_SEQNO ,VINNO , CATALOGUENO, SUBJECT , ENCLOSE, ReqBy, NotedBy)"
        SQL = SQL & "VALUES( " & N2Str2Null(txtDPINo) & "," & vDPI_DATE & "," & vDPI_INQTYPE & ", " & vDPI_SEQNO & " ," & vVINNO & " , " & vCATALOGUENO & ", " & vSubject & " , " & vEnclose & " , " & vReqBy & " , " & vNotedBy & ") " & vbCrLf
        SQL = SQL & " SELECT @@IDENTITY"
        LogAudit "A", "DPI"
    Else

        SQL = "UPDATE PMIS_DPIHEADER SET "
        SQL = SQL & " DPINo=" & N2Str2Null(txtDPINo) & ", "
        SQL = SQL & " DPI_DATE=" & vDPI_DATE & ", "
        SQL = SQL & " DPI_INQTYPE=" & vDPI_INQTYPE & " , "
        SQL = SQL & " DPI_SEQNO=" & vDPI_SEQNO & ", "
        SQL = SQL & " VINNO=" & vVINNO & " , "
        SQL = SQL & " CATALOGUENO=" & vCATALOGUENO & " , "
        SQL = SQL & " SUBJECT=" & vSubject & ", "
        SQL = SQL & " NotedBy= " & vNotedBy & " , "
        SQL = SQL & " Reqby= " & vReqBy & " , "
        SQL = SQL & " ENCLOSE=" & vEnclose
        SQL = SQL & " where  ID= " & txtID

        LogAudit "E", "DPI", txtDPINo
        If LTrim(RTrim(txtDPINo)) <> Null2String(RsDPIR!DPIno) Then
            gconDMIS.Execute ("update PMIS_DPIDetails  set dpino='" & txtDPINo & "' where dpino='" & Null2String(RsDPIR!DPIno) & "'")
        End If
    End If



    If ADDOREDIT = "ADD" And txtID = 0 Then
        AddDetails
    End If
    Set TEMPRS = gconDMIS.Execute(SQL)
    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        txtID = TEMPRS.Collect(0)
    End If
    FillSearchGrid
    picAdds.Visible = True
    picSaves.Visible = False
    rsRefresh
    RsDPIR.Find ("ID=" & txtID)
    cmdCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSave_Tech_Click()

    'If txtTech_GroupNumber = "" Then
    'On Error Resume Next
    'txtTech_GroupNumber.SetFocus
    '   Exit Sub
    'ElseIf txtTech_PNC = "" Then
    '    On Error Resume Next
    '    txtTech_PNC.SetFocus
    '    Exit Sub
    'ElseIf txtTech_PartName = "" Then
    '   On Error Resume Next
    '   txtTech_PartName.SetFocus
    '   Exit Sub
'    If LTrim(RTrim(cboTech_PartNo)) = "" Then
'        On Error Resume Next
'        cboTech_PartNo.SetFocus
'        Exit Sub
'    End If

    If txtTech_PartName = "" Then
      On Error Resume Next
      txtTech_PartName.SetFocus
      Exit Sub
    End If
    
    
    Dim ijx          As Integer

    ijx = ItemExists(cboTech_PartNo, 0)


    If (ijx > 0 And txtDPIDetailID = 0) Or (ijx >= 1 And Grid1.Text <> cboTech_PartNo And txtDPIDetailID <> 0) Then
        MsgBox "Part Number Already Used In This Transaction", vbInformation, "Duplicate Entry"
        txtTech_PartName = vbNullString
        cboTech_PartNo.SetFocus
        Exit Sub
    ElseIf ijx = 0 And txtDPIDetailID = 0 Then
        SQL = "INSERT INTO PMIS_DPIDetails (DPINO, PART_NO,PART_NAME,GP_NO,PNC) VALUES ("
        SQL = SQL & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboTech_PartNo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtTech_PartName) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtTech_GroupNumber) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtTech_PNC) & ")"
        gconDMIS.Execute SQL
    ElseIf (ijx = 1 And Grid1.Text = cboTech_PartNo) Or (ijx = 0 And Grid1.Text <> cboTech_PartNo) And txtDPIDetailID <> 0 Then
        SQL = "Update PMIS_DPIDetails SET " & vbCrLf
        SQL = SQL & "DPINO=" & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & "PART_NO=" & N2Str2Null(cboTech_PartNo) & "," & vbCrLf
        SQL = SQL & "PART_NAME=" & N2Str2Null(txtTech_PartName) & "," & vbCrLf
        SQL = SQL & "GP_NO=" & N2Str2Null(txtTech_GroupNumber) & "," & vbCrLf
        SQL = SQL & "PNC=" & N2Str2Null(txtTech_PNC) & vbCrLf
        SQL = SQL & "Where ID= " & txtDPIDetailID
        gconDMIS.Execute SQL
        ShowHidePictureBox picInquiryTechincal.hwnd, False
    Else
        MsgBox "Else Statements"
    End If
    txtTech_GroupNumber = vbNullString
    txtTech_PNC = vbNullString
    txtTech_PartName = vbNullString
    cboTech_PartNo = vbNullString
    InitGrid
    FillGrid
End Sub

Private Sub cmdSave_ETA_Click()

    If cboETA_PNO.ListIndex = -1 Then
        On Error Resume Next
        cboETA_PNO.SetFocus
        Exit Sub
        'ElseIf txtETA_PARTNAME.Text = "" Then
        '    On Error Resume Next
        '    txtETA_PARTNAME.SetFocus
        '    Exit Sub

    ElseIf txtETA_QTY.Text = "" Then
        On Error Resume Next
        txtETA_QTY.SetFocus
        Exit Sub
        'ElseIf cboPartsEstimate_Status.ListIndex = -1 Then
        '    On Error Resume Next
        '    cboPartsEstimate_Status.SetFocus
        '    Exit Sub
    End If


    Dim ijx                                                           As Integer

    ijx = ItemExists(cboETA_PNO, 0)


    If (ijx > 0 And txtDPIDetailID = 0) Or (ijx >= 1 And Grid1.Text <> cboETA_PNO And txtDPIDetailID <> 0) Then
        MsgBox "Part Number Already Used In This Transaction", vbInformation, "Duplicate Entry"
        txtETA_PARTNAME = vbNullString
        cboETA_PNO.SetFocus
        Exit Sub
    ElseIf ijx = 0 And txtDPIDetailID = 0 Then
        SQL = "INSERT INTO PMIS_DPIDetails (DPINO, PART_NO,PART_NAME,ORDER_NO,DATE_ORDERED,QTY,Status,ESTARRIVAL) VALUES ("
        SQL = SQL & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboETA_PNO) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtETA_PARTNAME) & "," & vbCrLf
        SQL = SQL & N2Str2Null(cboETA_OrderNo) & "," & vbCrLf
        SQL = SQL & N2Str2Null(txtETA_DateOrd) & "," & vbCrLf
        SQL = SQL & NumericVal(txtETA_QTY) & "," & vbCrLf
        SQL = SQL & N2Str2Null(GetStatus) & "," & vbCrLf
        SQL = SQL & N2Str2Null(dtPartsEstimate_ETA) & ")"
        gconDMIS.Execute SQL

    ElseIf (ijx = 1 And Grid1.Text = cboETA_PNO) Or (ijx = 0 And Grid1.Text <> cboETA_PNO) And txtDPIDetailID <> 0 Then

        SQL = "Update PMIS_DPIDetails SET " & vbCrLf
        SQL = SQL & " DPINO=" & N2Str2Null(txtDPINo) & "," & vbCrLf
        SQL = SQL & " PART_NO=" & N2Str2Null(cboETA_PNO) & "," & vbCrLf
        SQL = SQL & " PART_NAME=" & N2Str2Null(txtETA_PARTNAME) & "," & vbCrLf
        SQL = SQL & " ORDER_NO=" & N2Str2Null(cboETA_OrderNo) & "," & vbCrLf
        SQL = SQL & " DATE_ORDERED=" & N2Str2Null(txtETA_DateOrd) & "," & vbCrLf
        SQL = SQL & " QTY=" & NumericVal(txtETA_QTY) & "," & vbCrLf
        SQL = SQL & " Status=" & N2Str2Null(GetStatus) & "," & vbCrLf
        SQL = SQL & " ESTARRIVAL=" & N2Str2Null(dtPartsEstimate_ETA) & vbCrLf
        SQL = SQL & " Where ID= " & txtDPIDetailID
        gconDMIS.Execute SQL
        ShowHidePictureBox picInquiryEstimateArrival.hwnd, False


    Else
        MsgBox "Else Statements"
    End If

    txtETA_PARTNAME = vbNullString
    cboETA_PNO = vbNullString
    cboETA_OrderNo = vbNullString
    txtETA_DateOrd.Value = FormatDateTime(Now, vbShortDate)
    txtETA_QTY = 0
    dtPartsEstimate_ETA.Value = FormatDateTime(Now, vbShortDate)

    InitGrid
    FillGrid

End Sub

Private Sub dtDateDPI_Change()
    Dim i                                                             As Integer

    txtDateDPI2 = dtDateDPI

    For i = 0 To 2
        If optDPIInq(i).Value = True Then
            optDPIInq_Click (i)
            Exit For
        End If
    Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ADDER:
    If KeyCode = vbKeyF3 Then
        If picAdds.Visible = True Then
            If Null2String(RsDPIR!STATUS) = "P" Then
                MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
            ElseIf Null2String(RsDPIR!STATUS) = "C" Then
                MsgSpeechBox "Transaction are already Cancelled and cannot be Change."
            Else
                AddDetails
            End If
        End If
    ElseIf KeyCode = vbKeyF4 Then
        If picAdds.Visible = True Then
            If Null2String(RsDPIR!STATUS) <> "P" And Null2String(RsDPIR!STATUS) <> "C" Then
                EditDetails
            End If
        End If
    ElseIf KeyCode = vbKeyF5 Then

        If Null2String(RsDPIR!STATUS) <> "P" And Null2String(RsDPIR!STATUS) <> "C" Then
            Grid1_DblClick
            If MsgBox("Confirm selected Record?", vbOKCancel + vbExclamation, "DPI") = vbOK Then

                txtDPIDetailID = Grid1.TextMatrix(Grid1.Row, Grid1.Cols - 1)

                gconDMIS.Execute ("DELETE From PMIS_DPIDETAILS WHERE ID=" & txtDPIDetailID)

                cleargrid Grid1

                FillGrid



            End If
            Select Case dpiInqType
                Case "TECH"
                    ShowHidePictureBox picInquiryTechincal.hwnd, False
                Case "PRCS"
                    ShowHidePictureBox picInquiryPrice.hwnd, False
                Case "ETA"
                    ShowHidePictureBox picInquiryEstimateArrival.hwnd, False
            End Select
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Select Case dpiInqType
            Case "TECH"
                ShowHidePictureBox picInquiryTechincal.hwnd, False
            Case "PRCS"
                ShowHidePictureBox picInquiryPrice.hwnd, False
            Case "ETA"
                ShowHidePictureBox picInquiryEstimateArrival.hwnd, False
        End Select
    Else
        MoveKeyPress KeyCode
    End If
    Exit Sub
ADDER:
    err.Clear
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillParts
    With cboPartsEstimate_Status
        .AddItem "FOR ORDERING"
        .AddItem "BACK ORDER STAGE"
        .AddItem "ALLOCATION STAGE"
        .AddItem "PICKING STAGE"
        .AddItem "PACKING STAGE"
        .AddItem "SHIPPING STAGE"
    End With
    picAdds.Visible = True
    picSaves.Visible = False
    picTop.Enabled = False
    SetCompany
    rsRefresh
    InitMemVars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Grid1_DblClick()
    If Grid1.Text = "" Or LCase(Grid1.Text) = "no entry" Then
        If MsgBox("There is Nothing To Edit" & vbCrLf, vbOKCancel + vbInformation) = vbOK Then
            Exit Sub
        End If

        Exit Sub
    End If

    If Null2String(RsDPIR!STATUS) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf Null2String(RsDPIR!STATUS) = "C" Then
        MsgSpeechBox "Transactions Are Already Cancelled and cannot be Change"
    Else
        EditDetails
    End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Grid1_DblClick
    End If
End Sub

Private Sub Label19_Click()
    If LTrim(RTrim(LOGCODE)) = "NET" Then

    End If

End Sub

Private Sub lstDPIList_DblClick()
    If lstDPIList.SelectedItem Is Nothing Then: Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub lstDPIList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'On Error Resume Next
    RsDPIR.MoveFirst
    RsDPIR.Bookmark = rsFind(RsDPIR.Clone, "DPINO", Item.ListSubItems(1).Text).Bookmark
    StoreMemvars
    Exit Sub
ADDER:
    err.Clear

    'rsDPI.Requery
    'rsDPI.MoveFirst
    'rsDPI.Find "id=" & ITEM.ListSubItems(1).Text
    'StoreMemvars
    'Exit Sub



End Sub

Private Sub optCatalgoue_Click()
    textSearch_Change
End Sub

Private Sub optDate_Click()
    textSearch_Change
End Sub

Private Sub optDPI_Click()
    textSearch_Change
End Sub

Private Sub optDPIInq_Click(Index As Integer)
    If optDPIInq(Index).Value = True Then

        If ADDOREDIT = "ADD" Then
            InitMemVars
            InitGrid
            FillGrid
        End If

        dpiInqType = optDPIInq(Index).Tag
        dpiSeq = GenerateDPISEQ
        txtDPINo = optDPIInq(Index).Tag & "-" & Right(Year(dtDateDPI), 2) & "-" & Format(Month(dtDateDPI), "00") & "-" & dpiSeq
        lbldpiInqType = optDPIInq(Index).Caption
        txtDateDPI2 = dtDateDPI
    End If
End Sub

Private Sub optVIN_Click()
    textSearch_Change
End Sub

Private Sub textSearch_Change()
    FillSearchGrid
End Sub

Private Sub Timer1_Timer()
    If lblStatus.Caption <> "" Then
        If lblStatus.Visible = True Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If
    End If
End Sub




