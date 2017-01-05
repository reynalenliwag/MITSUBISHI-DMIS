VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISCheckEncashment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Encashment Data Entry"
   ClientHeight    =   6210
   ClientLeft      =   885
   ClientTop       =   930
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CheckEncashment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11160
   Begin VB.PictureBox picCheckEncashment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   1208
      ScaleHeight     =   3225
      ScaleWidth      =   9345
      TabIndex        =   26
      Top             =   1523
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton cmdDeleteInCash 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   1950
         MouseIcon       =   "CheckEncashment.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2340
         Width           =   705
      End
      Begin VB.ComboBox cboTseklase 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1890
         Width           =   4515
      End
      Begin VB.TextBox txtCheckNumber 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5010
         TabIndex        =   11
         Top             =   1470
         Width           =   1455
      End
      Begin VB.TextBox txtTimeInCash 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1020
         Width           =   1185
      End
      Begin VB.TextBox txtChkAmount 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7770
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1020
         Width           =   1455
      End
      Begin VB.ComboBox cboBankCode 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1020
         Width           =   4545
      End
      Begin MSComCtl2.DTPicker dtpInCashDate 
         Height          =   405
         Left            =   1950
         TabIndex        =   6
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50331649
         CurrentDate     =   38216
      End
      Begin VB.CommandButton cmdCancelInCash 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   8520
         MouseIcon       =   "CheckEncashment.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2340
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveInCash 
         Caption         =   "&Save"
         Height          =   795
         Left            =   7830
         MouseIcon       =   "CheckEncashment.frx":11D7
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":1329
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2340
         Width           =   705
      End
      Begin VB.TextBox txtCheckDte 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1950
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   345
         Left            =   0
         TabIndex        =   45
         Top             =   -30
         Width           =   9345
         _Version        =   655364
         _ExtentX        =   16484
         _ExtentY        =   609
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
         ForeColor       =   16777215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Encashment :"
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
         Left            =   105
         TabIndex        =   34
         Top             =   570
         Width           =   1740
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number  :"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Type  :"
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
         Left            =   735
         TabIndex        =   32
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date  :"
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
         Left            =   780
         TabIndex        =   31
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank  :"
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
         Left            =   1305
         TabIndex        =   30
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   6510
         TabIndex        =   29
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
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
         Left            =   7980
         TabIndex        =   28
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label labBankDepoID 
         BackColor       =   &H000000FF&
         Caption         =   "Label14"
         Height          =   165
         Left            =   510
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2100
      ScaleHeight     =   525
      ScaleWidth      =   8985
      TabIndex        =   18
      Top             =   60
      Width           =   9015
      Begin VB.TextBox txtInCashDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   60
         Width           =   1875
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   6720
         Top             =   30
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Encashment Date  :"
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
         Height          =   240
         Left            =   825
         TabIndex        =   19
         Top             =   90
         Width           =   1830
      End
   End
   Begin VB.Frame fraDetails 
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
      Height          =   6165
      Left            =   30
      TabIndex        =   15
      Top             =   -30
      Width           =   2055
      Begin VB.TextBox textSearch 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   30
         MaxLength       =   35
         TabIndex        =   16
         Top             =   150
         Width           =   1965
      End
      Begin MSComctlLib.ListView lstInCash 
         Height          =   5535
         Left            =   30
         TabIndex        =   17
         Top             =   570
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   9763
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
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CheckEncashment.frx":1679
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "In Cash Date"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture5 
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
      Height          =   870
      Left            =   6900
      ScaleHeight     =   870
      ScaleWidth      =   4245
      TabIndex        =   35
      Top             =   5340
      Width           =   4245
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   3510
         MouseIcon       =   "CheckEncashment.frx":17DB
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":192D
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "CheckEncashment.frx":1C93
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":1DE5
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   795
         Left            =   2130
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CheckEncashment.frx":214B
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":229D
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Post this Transaction"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   1440
         MouseIcon       =   "CheckEncashment.frx":25C2
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":2714
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   750
         MouseIcon       =   "CheckEncashment.frx":2A70
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":2BC2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   60
         MouseIcon       =   "CheckEncashment.frx":2ED5
         MousePointer    =   99  'Custom
         Picture         =   "CheckEncashment.frx":3027
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   2100
      ScaleHeight     =   3255
      ScaleWidth      =   8985
      TabIndex        =   20
      Top             =   630
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid grdInCash 
         Height          =   3105
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   5477
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   " Code           |   Bank Name                                            |    Time            | Check Amount   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CheckEncashment.frx":3321
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2100
      ScaleHeight     =   1245
      ScaleWidth      =   8985
      TabIndex        =   21
      Top             =   3930
      Width           =   9015
      Begin VB.TextBox txtTotalChkAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   390
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   60
         Width           =   1815
      End
      Begin VB.TextBox txtChkDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1500
         TabIndex        =   2
         Top             =   60
         Width           =   1485
      End
      Begin VB.TextBox txtChkNumber 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1500
         TabIndex        =   3
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox txtTseklase 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1500
         TabIndex        =   4
         Top             =   840
         Width           =   5175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount  :"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date  :"
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
         Left            =   390
         TabIndex        =   24
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number  :"
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
         Left            =   90
         TabIndex        =   23
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Type  :"
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
         Left            =   345
         TabIndex        =   22
         Top             =   870
         Width           =   1110
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2910
      TabIndex        =   14
      Top             =   1650
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCMISCheckEncashment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsINCASH                                                        As ADODB.Recordset
Dim TOTAL_CHECK_ENCASHMENT                                          As Double
Dim Previous_EncashAmount                                           As Double
Dim AddorEdit                                                       As String

Function SetBankCode(XXX As Variant)
    Dim rsBankName                                                  As New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("SELECT CODE FROM CMIS_SBOOK WHERE Book = 'B' AND DESCNAME = " & N2Str2Null(XXX))
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankCode = rsBankName!Code
    End If
End Function

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                  As New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'B' AND Code = '" & XXX & "'")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankName = rsBankName!DESCNAME
    End If
End Function

Function SetCheckClass(XXX As Variant)
    Dim rsSBOOK                                             As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE Book = 'F' AND CODE = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClass = rsSBOOK!DESCNAME
    End If
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                             As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT CODE FROM CMIS_SBOOK WHERE Book = 'F' AND DESCNAME = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
End Function

Sub StoreInCashEntry(XXX As Variant)
    Dim rsINCASH2                                           As New ADODB.Recordset
    Set rsINCASH2 = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE ID = " & XXX)
    If Not rsINCASH2.EOF And Not rsINCASH2.BOF Then
        labBankDepoID.Caption = rsINCASH2!Id
        dtpInCashDate = Null2String(rsINCASH2!incashdate)
        cboBankCode = SetBankName(Null2String(rsINCASH2!bankcode))
        txtTimeInCash.Text = Null2String(rsINCASH2!timeincash)
        txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsINCASH2!CHKAMOUNT))
        txtCheckDte.Text = Null2String(rsINCASH2!CHKDATE)
        txtCheckNumber.Text = Null2String(rsINCASH2!CHKNUMBER)
        cboTseklase.Text = SetCheckClass(Null2String(rsINCASH2!Tseklase))
        Previous_EncashAmount = N2Str2Zero(rsINCASH2!CHKAMOUNT)
    End If
End Sub

Sub rsRefresh()
    Set rsINCASH = New ADODB.Recordset
    Set rsINCASH = gconDMIS.Execute("SELECT DISTINCT INCASHDATE FROM CMIS_InCash WHERE month(incashdate) = " & PERIODMONTH & " AND year(incashdate) = " & PERIODYEAR & " ORDER BY INCASHDATE ASC")
End Sub

Sub InitCheckEncashmentMemVars()
    dtpInCashDate = LOGDATE
    cboBankCode.ListIndex = -1
    txtTimeInCash.Text = ""
    txtChkAmount.Text = "0.00"
    txtCheckDte.Text = LOGDATE
    txtCheckNumber.Text = ""
    cboTseklase.ListIndex = -1
End Sub

Sub InitCbo()
    Dim rsBANK                                              As New ADODB.Recordset
    Set rsBANK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE BOOK = 'B' ORDER BY DESCNAME ASC")
    If Not rsBANK.EOF And Not rsBANK.BOF Then
        Combo_Loadval cboBankCode, rsBANK
    End If
    Set rsBANK = Nothing
    
    Dim rsSBOOK                                                     As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE Book = 'F' ORDER BY DESCNAME ASC")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboTseklase, rsSBOOK
    End If
    Set rsSBOOK = Nothing
End Sub

Sub StoreMemVars()
    If Not (rsINCASH.EOF And rsINCASH.BOF) Then
        txtInCashDate.Text = Null2Date(rsINCASH!incashdate)
        Call StoreDetails
    Else
        Call ShowNoRecord
        Call cmdAdd_Click
        txtInCashDate.Text = LOGDATE
    End If
End Sub

Sub StoreDetails()
    Dim rsINCASHDet                                                 As ADODB.Recordset
    Dim i                                                           As Long
    
    TOTAL_CHECK_ENCASHMENT = 0
    InitGrid
    i = 0
    
    Set rsINCASHDet = New ADODB.Recordset
    Set rsINCASHDet = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE INCASHDATE = '" & txtInCashDate.Text & "' ORDER BY ID ASC")
    If Not rsINCASHDet.EOF And Not rsINCASHDet.BOF Then
        rsINCASHDet.MoveFirst
        Do While Not rsINCASHDet.EOF
            i = i + 1
            grdInCash.AddItem Null2String(rsINCASHDet!bankcode) & _
                Chr(9) & SetBankName(Null2String(rsINCASHDet!bankcode)) & _
                Chr(9) & Null2String(rsINCASHDet!timeincash) & _
                Chr(9) & ToDoubleNumber(N2Str2Zero(rsINCASHDet!CHKAMOUNT)) & _
                Chr(9) & rsINCASHDet!Id
                
            If i = 1 Then grdInCash.RemoveItem 1
            TOTAL_CHECK_ENCASHMENT = TOTAL_CHECK_ENCASHMENT + N2Str2Zero(rsINCASHDet!CHKAMOUNT)
            rsINCASHDet.MoveNext
        Loop
    End If
    txtTotalChkAmount.Text = ToDoubleNumber(TOTAL_CHECK_ENCASHMENT)
End Sub

Sub InitGrid()
    cleargrid grdInCash
    grdInCash.FormatString = " Code           |   Bank Name                                            |    Time            | Check Amount   "
    grdInCash.ColWidth(4) = 1
End Sub

Sub FillGrid()

    lstInCash.Sorted = False
    lstInCash.ListItems.Clear
    lstInCash.Enabled = False
    
    Dim rsINCASH2                                                   As ADODB.Recordset
    Set rsINCASH2 = New ADODB.Recordset
    Set rsINCASH2 = gconDMIS.Execute("SELECT DISTINCT INCASHDATE FROM CMIS_InCash ORDER BY INCASHDATE DESC")
    If Not (rsINCASH2.EOF And rsINCASH2.BOF) Then
        lstInCash.Enabled = True
        Listview_Loadval Me.lstInCash.ListItems, rsINCASH2
        lstInCash.Refresh
        lstInCash.Enabled = True
    Else
        lstInCash.Enabled = False
    End If
    Set rsINCASH2 = Nothing
End Sub

Sub FillSearchGrid(XXX As Variant)
    
    lstInCash.Sorted = False
    lstInCash.ListItems.Clear
    lstInCash.Enabled = False
    XXX = Repleys(LTrim(RTrim(XXX)))
    
    Dim rsINCASH2                                                   As New ADODB.Recordset
    Set rsINCASH2 = gconDMIS.Execute("SELECT DISTINCT INCASHDATE FROM CMIS_InCash WHERE INCASHDATE LIKE '" & XXX & "%' ORDER BY INCASHDATE DESC")
    If Not (rsINCASH2.EOF And rsINCASH2.BOF) Then
        lstInCash.Enabled = True
        Listview_Loadval Me.lstInCash.ListItems, rsINCASH2
        lstInCash.Refresh
        lstInCash.Enabled = True
    Else
        lstInCash.Enabled = False
    End If
    Set rsINCASH2 = Nothing
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "TRANSACTION CHECK ENCASHMENT") = False Then Exit Sub
    
    AddorEdit = "ADD"
    Call EnablePictures(False)
    cmdDeleteInCash.Visible = False
    picCheckEncashment.Visible = True
    picCheckEncashment.ZOrder 0
    InitCheckEncashmentMemVars
    lstInCash.Enabled = False
    textSearch.Enabled = False
End Sub

Private Sub cmdCancelInCash_Click()
    AddorEdit = ""
    picCheckEncashment.Visible = False
    picCheckEncashment.ZOrder 1
    lstInCash.Enabled = True
    Call EnablePictures(True)
    Call StoreMemVars
End Sub

Private Sub cmdDeleteInCash_Click()
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "DELETE FROM CMIS_InCash WHERE ID = " & labBankDepoID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "TRANSACTION CHECK ENCASHMENT", SQL_STATEMENT, labBankDepoID, "", Null2String(txtCheckNumber), "", ""

        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                          " CASH = (CASH + " & NumericVal(txtChkAmount.Text) & ")," & _
                          " INCASHMENT = (INCASHMENT - " & NumericVal(txtChkAmount.Text) & ")," & _
                          " [CHECK] = [CHECK] - " & NumericVal(txtChkAmount.Text) & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        
        grdInCash.Rows = 1
        Call rsRefresh
        Call textSearch_Change
        Call EnablePictures(True)
        Call cmdCancelInCash_Click
        
        If Not rsINCASH.EOF And Not rsINCASH.BOF Then rsINCASH.MoveLast
        LogAudit "X", "CHECK ENCASHMENT", txtInCashDate
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Sub EnablePictures(COND As Boolean)
    Picture1.Enabled = COND
    Picture2.Enabled = COND
    Picture4.Enabled = COND
    Picture5.Enabled = COND
End Sub

Private Sub cmdSaveInCash_Click()
    On Error GoTo ErrorCode:
    Dim vInCashDate                                                 As String
    Dim vBankCode                                                   As String
    Dim vChkNumber                                                  As String
    Dim vChkDate                                                    As String
    Dim vChkAmount                                                  As Double
    Dim vTimeInCash                                                 As String
    Dim vTseklase                                                   As String

    vInCashDate = N2Str2Null(dtpInCashDate)
    vBankCode = N2Str2Null(SetBankCode(cboBankCode.Text))
    vChkNumber = N2Str2Null(txtCheckNumber.Text)
    vChkDate = N2Str2Null(txtCheckDte.Text)
    vChkAmount = NumericVal(txtChkAmount.Text)
    vTimeInCash = N2Str2Null(txtTimeInCash.Text)
    vTseklase = N2Str2Null(SetCheckClassCode(cboTseklase.Text))

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO CMIS_InCash " & _
                        "(InCashDate,BankCode,ChkNumber,ChkDate,ChkAmount,TimeInCash,Tseklase,datecreate,timecreate)" & _
                        " VALUES (" & vInCashDate & _
                        ", " & vBankCode & _
                        ", " & vChkNumber & _
                        ", " & vChkDate & _
                        ", " & vChkAmount & _
                        ", " & vTimeInCash & _
                        ", " & vTseklase & _
                        ", '" & LOGDATE & _
                        "', '" & Time & "')"
        
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "TRANSACTION CHECK ENCASHMENT", SQL_STATEMENT, labBankDepoID, "", Null2String(txtCheckNumber), vTseklase, ""
        
        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                          " CASH = CASH - " & vChkAmount & "," & _
                          " INCASHMENT = INCASHMENT + " & vChkAmount & "," & _
                          " [CHECK] = [CHECK] + " & vChkAmount & _
                          " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                
        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE CMIS_InCash SET " & _
                        " InCashDate = " & vInCashDate & "," & _
                        " BankCode = " & vBankCode & "," & _
                        " ChkNumber = " & vChkNumber & "," & _
                        " ChkDate = " & vChkDate & "," & _
                        " ChkAmount = " & vChkAmount & "," & _
                        " TimeInCash = " & vTimeInCash & _
                        " WHERE ID = " & labBankDepoID.Caption

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "TRANSACTION CHECK ENCASHMENT", SQL_STATEMENT, labBankDepoID, "", Null2String(txtCheckNumber), vTseklase, ""
        
        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                          " CASH = (CASH - " & vChkAmount & ") + " & Previous_EncashAmount & _
                          ", INCASHMENT = (INCASHMENT + " & vChkAmount & ") - " & Previous_EncashAmount & _
                          ", [CHECK] = [CHECK] + " & vChkAmount - Previous_EncashAmount & _
                          " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                
        Call ShowSuccessFullyUpdated
    End If

    Call rsRefresh
    On Error Resume Next
    rsINCASH.Find "InCashDate = '" & vInCashDate & "'"
    Call textSearch_Change
    Call cmdCancelInCash_Click
    
    If Not rsINCASH.EOF And Not rsINCASH.BOF Then rsINCASH.MoveLast
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyEscape
            cmdCancelInCash_Click
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    
    Dim rsProfile                                                   As New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("SELECT * FROM ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    
    Set rsProfile = Nothing
    Call CenterMe(frmMain, Me, 1)
    Call textSearch_Change
    Call InitCbo
    Call rsRefresh
    
    If Not rsINCASH.EOF And Not rsINCASH.BOF Then rsINCASH.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub grdInCash_DblClick()
    grdInCash.Col = 4
    If grdInCash.Text <> "" Then
        AddorEdit = "EDIT"
        picCheckEncashment.Visible = True
        picCheckEncashment.ZOrder 0
        cmdDeleteInCash.Visible = True
        Call StoreInCashEntry(grdInCash.Text)
    End If
End Sub

'SEARCH MODULE
Private Sub lstINCASH_GotFocus()
    On Error Resume Next
    
    txtInCashDate.Text = lstInCash.SelectedItem
    'rsINCASH.Bookmark = rsFind(rsINCASH.Clone, "INCASHDATE", lstINCASH.SelectedItem).Bookmark
    StoreDetails
End Sub

Private Sub lstINCASH_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtInCashDate.Text = lstInCash.SelectedItem
    'rsINCASH.Bookmark = rsFind(rsINCASH.Clone, "INCASHDATE", lstINCASH.SelectedItem).Bookmark
    StoreDetails
End Sub

Private Sub lstINCASH_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstInCash
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstINCASH_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstINCASH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstInCash.ListItems.Count > 0 And lstInCash.Enabled = True Then: lstInCash.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    If AddorEdit = "ADD" Then
        txtTimeInCash.Text = Time
        DoEvents
    End If
End Sub

