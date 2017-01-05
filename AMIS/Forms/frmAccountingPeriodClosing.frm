VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccountingPeriodClosing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounting Period"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmAccountingPeriodClosing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   9600
   Begin VB.PictureBox picAcctYear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2700
      ScaleHeight     =   1185
      ScaleWidth      =   3885
      TabIndex        =   16
      Top             =   2130
      Visible         =   0   'False
      Width           =   3915
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2760
         TabIndex        =   37
         Top             =   600
         Width           =   1005
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Proceed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1770
         TabIndex        =   36
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   90
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   345
         Left            =   1710
         TabIndex        =   17
         Top             =   1290
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   98369537
         CurrentDate     =   40114
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   345
         Left            =   3660
         TabIndex        =   18
         Top             =   1290
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   98369537
         CurrentDate     =   40114
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select Year to close:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   150
         Width           =   2595
      End
      Begin VB.Label lblAccountingYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   3450
         TabIndex        =   19
         Top             =   1320
         Width           =   165
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   465
         Left            =   -60
         TabIndex        =   38
         Top             =   0
         Width           =   4005
         _Version        =   655364
         _ExtentX        =   7064
         _ExtentY        =   820
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12582912
         GradientColorDark=   16777215
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picOpenClose 
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4755
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   30
      Width           =   9525
      Begin VB.PictureBox picControl 
         Height          =   4695
         Left            =   6540
         ScaleHeight     =   4635
         ScaleWidth      =   2835
         TabIndex        =   3
         Top             =   30
         Width           =   2895
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   30
            Picture         =   "frmAccountingPeriodClosing.frx":09AA
            TabIndex        =   9
            Top             =   4200
            Width           =   2775
         End
         Begin VB.CommandButton cmdYearEnd 
            Caption         =   "&Year End Process"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   30
            Picture         =   "frmAccountingPeriodClosing.frx":1A2C
            TabIndex        =   10
            Top             =   3810
            Width           =   2775
         End
         Begin VB.CommandButton cmdClosePeriod 
            Caption         =   "&Month End Process"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3420
            Width           =   2775
         End
         Begin VB.CommandButton cmdSet 
            Caption         =   "&Set Journal Period"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   30
            Picture         =   "frmAccountingPeriodClosing.frx":2AAE
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2490
            Width           =   2775
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   5160
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton Command 
            BackColor       =   &H00C00000&
            Enabled         =   0   'False
            Height          =   225
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3210
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "GJ:"
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
            Height          =   210
            Index           =   5
            Left            =   300
            TabIndex        =   34
            Top             =   2070
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DRJ:"
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
            Height          =   210
            Index           =   4
            Left            =   300
            TabIndex        =   33
            Top             =   1770
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CRJ:"
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
            Height          =   210
            Index           =   3
            Left            =   300
            TabIndex        =   32
            Top             =   1470
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SJ:"
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
            Height          =   210
            Index           =   2
            Left            =   300
            TabIndex        =   31
            Top             =   1170
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CDJ:"
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
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   30
            Top             =   870
            Width           =   465
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "APJ:"
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
            Height          =   210
            Index           =   0
            Left            =   300
            TabIndex        =   29
            Top             =   570
            Width           =   465
         End
         Begin VB.Label lblGJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   28
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label lblDRJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   27
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label lblCRJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   26
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Label lblSJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   25
            Top             =   1140
            Width           =   1905
         End
         Begin VB.Label lblCDJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label lblAPJPeriod 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   840
            TabIndex        =   8
            Top             =   540
            Width           =   1905
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Journal Period"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   300
            TabIndex        =   7
            Top             =   150
            Width           =   2250
         End
         Begin VB.Shape Shape 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            Height          =   405
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   2775
         End
         Begin VB.Shape Shape 
            BackStyle       =   1  'Opaque
            Height          =   2055
            Index           =   1
            Left            =   30
            Top             =   420
            Width           =   2775
         End
      End
      Begin VB.ComboBox cboBookType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5640
         Visible         =   0   'False
         Width           =   3345
      End
      Begin FlexCell.Grid gridAccounting 
         Height          =   4725
         Left            =   0
         TabIndex        =   35
         Top             =   30
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   8334
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         DefaultRowHeight=   24
         GridColor       =   12632256
         ReadOnlyFocusRect=   0
         Rows            =   30
         MultiSelect     =   0   'False
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Journal Type"
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
         Left            =   150
         TabIndex        =   2
         Top             =   5700
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.PictureBox picStatus 
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4755
      ScaleWidth      =   6705
      TabIndex        =   12
      Top             =   30
      Visible         =   0   'False
      Width           =   6765
      Begin VB.CommandButton cmdClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   23
         Top             =   60
         Width           =   285
      End
      Begin MSComctlLib.ListView lvStatus 
         Height          =   4275
         Left            =   60
         TabIndex        =   13
         Top             =   420
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "JDate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Voucher No"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "J. Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Debit"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Credit"
            Object.Width           =   2558
         EndProperty
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List of Unposted Transactions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   150
         TabIndex        =   15
         Top             =   90
         Width           =   3210
      End
      Begin VB.Label lblJType 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1050
         TabIndex        =   14
         Top             =   5280
         Width           =   4815
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   315
         Index           =   2
         Left            =   60
         Top             =   60
         Width           =   6585
      End
   End
End
Attribute VB_Name = "frmAccountingPeriodClosing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xAccountingPeriod                                  As String
Dim xVOUCHERNO                                         As String
Dim xJType                                             As String
Dim xPeriodFrom                                        As Date
Dim xAccountingMonth                                   As Date
Dim xAccountingMonth2                                  As Date
Dim xBackMonth                                         As Date
Dim xNextMonth                                         As Date
Dim Current_Cash_GrossSales                            As Double
Dim Current_Charge_GrossSales                          As Double
Dim Current_Cash_SalesDiscountsAndReturns              As Double
Dim Current_Charge_SalesDiscountsAndReturns            As Double
Dim Current_Cash_CostOfSales                           As Double
Dim Current_Charge_CostOfSales                         As Double
Dim Current_LessSellingExpense                         As Double
Dim Current_LessAdminExpense                           As Double
Dim Current_LessOtherExpense                           As Double
Dim Current_AddOtherIncome                             As Double
Dim xCurrentPeriod                                     As Boolean
Dim rsJournal_Det                                      As ADODB.Recordset

Dim rsUSP_AR                                           As ADODB.Recordset
Dim rsUSP_AP                                           As ADODB.Recordset
Dim CMD_RRL                                            As ADODB.Command

Private Sub cboBookType_Click()
    initGrid
    FillGrid
End Sub

Private Sub cmdCancel_Click()
    picOpenClose.Enabled = True
    picAcctYear.Visible = False
    picAcctYear.ZOrder 1
End Sub

Private Sub cmdClose_Click()
    picStatus.Visible = False
    picStatus.ZOrder 1
End Sub

Private Sub cmdClosePeriod_Click()
    If Module_Access(LOGID, "MONTH END PROCESS", "SYSTEM") = False Then Exit Sub
    Screen.MousePointer = 11
    
    Dim iRow                                           As Long
    'gridAccounting.Selection.FirstCol = gridAccounting.Cols - 1 And
    If gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "Open" Then
        If CheckOpenBook(xBackMonth) = "Open" Then
            MsgBox "The selected Accounting Period cannot be close because the previous Period is not closed." & _
                   vbCrLf & "You must close the previous Period before you may close this Period." & _
                   vbCrLf & "Close Accounting Period " & Format(xBackMonth, "mmmm yyyy"), vbInformation, "System Message"
            Screen.MousePointer = 0
        Else
            If MsgBox("Are you sure you want to CLOSE these books?", vbQuestion + vbYesNo) = vbNo Then
                Screen.MousePointer = 0
                Exit Sub
            Else
            
                 'RRL 7/27/2015
                 'DESCRIPTION: TO CHECK WHETHER THE AR SCHEDULES ARE TALLY.
'                 If CheckDisbalanced_AR = False Then
'                     MsgBox "Accounts receivables are not balanced. Can not process month-end.", vbCritical, "Message"
'                     Exit Sub
'                 End If
'
'                 If CheckDisbalanced_AP = False Then
'                     MsgBox "Accounts Payables are not balanced. Can not process month-end.", vbCritical, "Message"
'                     Exit Sub
'                 End If
                 'RRL 7/27/2015
                
                If CheckUnposted = False Then
                    CheckTrialStatus
                    If CDate(Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "mm-yyyy")) < CDate(Format(LOGDATE, "mm-yyyy")) Then
                        gconDMIS.Execute ("UPDATE AMIS_AccountingPeriod SET STATUS=1,CurrPeriod=0 where AcctMonth='" & Format(xAccountingMonth, "mm/dd/yyyy") & "'")
                        'gconDMIS.Execute ("UPDATE AMIS_AccountingPeriod SET CurrPeriod=1 where Status=0 and AcctMonth='" & Format(xNextMonth, "mm/dd/yyyy") & "'")
                        FillGrid
                        ListCurrentPeriod
                        'BalanceForwarding
                        Screen.MousePointer = 0
                    Else
                        MsgBox "Selected month is NOT YET FINISH...", vbExclamation, "System Message"
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    ElseIf gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "Closed" Then
        '        MsgBox "Cannot process month end when period is Closed.", vbInformation, "Select Period"
        '        Screen.MousePointer = 0
        '        Exit Sub
        If CheckClosedBook(xNextMonth) = "1" Then
            MsgBox "The selected Accounting Period cannot be open because the suceeding Period is not opened." & _
                   vbCrLf & "You must open the succeeding Period before you may open this Period." & _
                   vbCrLf & "Open Accounting Period " & Format(xNextMonth, "mmmm yyyy"), vbInformation, "System Message"
            Screen.MousePointer = 0
        Else
            If MsgBox("Are you sure you want to OPEN these books?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            gconDMIS.Execute ("UPDATE AMIS_AccountingPeriod SET STATUS=0 where AcctMonth='" & Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "mm/dd/yyyy") & "'")
            ListCurrentPeriod
            cmdClosePeriod.BackColor = &H8000000F
            FillGrid
            Screen.MousePointer = 0
        End If
    Else
        MsgBox "Kindly select Period to close...", vbInformation, "Select Period"
        Screen.MousePointer = 0
    End If

End Sub

Function CheckDisbalanced_AR()
Set rsUSP_AR = Nothing
Set CMD_RRL = New ADODB.Command
    
With CMD_RRL
    .ActiveConnection = gconDMIS
    .CommandType = adCmdStoredProc
    .CommandText = "XSP_TALLY"
    .CommandTimeout = 1000
    Set rsUSP_AR = .Execute
End With

If Not rsUSP_AR.EOF And Not rsUSP_AR.BOF Then
    rsUSP_AR.MoveFirst
    Do While Not rsUSP_AR.EOF
        If rsUSP_AR!xREMARKS = "NOT BALANCED" Then
            CheckDisbalanced_AR = False
            Exit Function
        End If
        rsUSP_AR.MoveNext
    Loop
End If
    CheckDisbalanced_AR = True
    Set rsUSP_AR = Nothing
    Exit Function
End Function

Function CheckDisbalanced_AP()
Set rsUSP_AP = Nothing
Set CMD_RRL = New ADODB.Command
    
With CMD_RRL
    .ActiveConnection = gconDMIS
    .CommandType = adCmdStoredProc
    .CommandText = "XSP_TALLY2"
    .CommandTimeout = 1000
    Set rsUSP_AP = .Execute
End With

If Not rsUSP_AP.EOF And Not rsUSP_AP.BOF Then
    rsUSP_AP.MoveFirst
    Do While Not rsUSP_AP.EOF
        If rsUSP_AP!xREMARKS = "NOT BALANCED" Then
            CheckDisbalanced_AP = False
            Exit Function
        End If
        rsUSP_AP.MoveNext
    Loop
End If
    CheckDisbalanced_AP = True
    Set rsUSP_AP = Nothing
    Exit Function
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
If cboYear.Text = "" Then
    MsgBox "Please select year to close.", vbInformation, "Closing"
    cboYear.SetFocus
    Exit Sub
Else
    If Module_Access(LOGID, "YEAR END PROCESS", "SYSTEM") = False Then Exit Sub
    If MsgBox("Are you sure you want to process Year End Closing?", vbQuestion + vbYesNo) = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    Else
        If CheckCloseBook = True Then
            Screen.MousePointer = 11
            Dim str_MSG                                        As String
            str_MSG = "Error During @ACL09182716350" & vbCrLf
            str_MSG = str_MSG & "Imported Data Will Now Roll back." & vbCrLf
            str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
            str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
            str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
            str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
            str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

            gconDMIS.BeginTrans
            If ClosingEntries = False Then
                str_MSG = Replace(str_MSG, "@ACL09182716350", "Year End Closing")
                MsgBox str_MSG, vbCritical, "Year End Closing Error "
                gconDMIS.RollbackTrans
                Screen.MousePointer = 0
                Exit Sub
            End If
            gconDMIS.CommitTrans
        Else
            MsgBox "All books must be CLOSED before processing Year End Closing", vbInformation, "Accounting Period"
        End If
    End If
End If
End Sub

Private Sub cmdSet_Click()
    If xAccountingPeriod = " " Then
        MsgBox "Kindly select Accounging Period to be set.", vbInformation, "Accounting Period"
    Else
        '        If gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "" Then
        If gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "Open" Then
            If MsgBox("Set Current Accounting Period to " & xAccountingPeriod & "?", vbQuestion + vbYesNo, "Change Accounting Period") = vbYes Then
                Call ClearPreviousPeriod(xJType)
                gconDMIS.Execute ("UPDATE AMIS_AccountingPeriod SET CurrPeriod=0 where CurrPeriod=1 and Status=0 and JType='" & xJType & "'")
                gconDMIS.Execute ("UPDATE AMIS_AccountingPeriod SET CurrPeriod=1 where JType='" & xJType & "' and Status=0 and AcctMonth ='" & Format(xAccountingMonth, "mm/dd/yyyy") & "'")
                ListCurrentPeriod
                xAccountingPeriod = " "
            Else
                xAccountingPeriod = " "
            End If
        Else
            MsgBox "Accounting period selected is closed.", vbInformation, "Accounting Period"
            Exit Sub
        End If
        '        Else
        '            MsgBox "Please select journal type from the list.", vbInformation, "Accounting Period"
        '            cboBookType.SetFocus
        '            Exit Sub
        '        End If
    End If
End Sub

Function ClosingEntries() As Boolean
On Error GoTo ErrorCode
Dim cnt                                    As Integer
            Dim J_ACCT_CODE, J_ACCT_NAME               As String
            Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET As Double
            Dim J_STATUS, J_JITEMNO                    As String

            Dim TOTAL_DEBIT, TOTAL_CREDIT              As Double
            Dim J_JDATE As String, J_VOUCHERNO As String, J_JTYPE As String
            Dim J_JNO As String, J_REMARKS             As String
            Dim rsJournal_HD                           As ADODB.Recordset
            Dim TOTAL_DEBIT_BALANCE, TOTAL_CREDIT_BALANCE As Double
            Dim DEBIT_BALANCE, CREDIT_BALANCE          As Double

            Set rsJournal_HD = New ADODB.Recordset
            rsJournal_HD.Open "select SUM(DEBIT) AS DEBIT_TOTAL, SUM(CREDIT) AS CREDIT_TOTAL, ACCT_CODE from AMIS_Journal_Det where LEFT(ACCT_CODE,1) > 3 AND jtype NOT IN ('CCM','CLO') and Status = 'P' AND YEAR(jdate) = '" & cboYear.Text & "' group by ACCT_CODE order by ACCT_CODE asc", gconDMIS, adOpenDynamic
            If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
                rsJournal_HD.MoveFirst
                TOTAL_DEBIT_BALANCE = 0: TOTAL_CREDIT_BALANCE = 0
                Screen.MousePointer = 11
                gconDMIS.Execute ("Delete from AMIS_Journal_HD Where Jtype = 'CLO' and ISNULL(STATUS,'N') = 'N' AND YEAR(jdate) = '" & cboYear.Text & "'")
                gconDMIS.Execute ("Delete from AMIS_Journal_Det Where Jtype = 'CLO' and ISNULL(STATUS,'N') = 'N' AND YEAR(jdate) = '" & cboYear.Text & "'")

                Dim rsJournal_HDDup                    As ADODB.Recordset
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
                Set rsJournal_HDDup = Nothing
                J_JDATE = N2Str2Null("12/31/" & NumericVal(cboYear.Text))
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                cnt = 0
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                 " (Jno,jdate,voucherno,jtype,remarks)" & _
                                 " values (" & J_JNO & "," & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', 'CLOSING ENTRIES (AUTOMATED BY SYSTEM)')"

                Do While Not rsJournal_HD.EOF
                    If NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) Then
                        DEBIT_BALANCE = NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL))
                        CREDIT_BALANCE = 0
                    Else
                        If NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) > NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)) Then
                            CREDIT_BALANCE = NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)) - NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL))
                            DEBIT_BALANCE = 0
                        Else
                            CREDIT_BALANCE = 0: DEBIT_BALANCE = 0
                        End If
                    End If
                    If DEBIT_BALANCE > 0 Or CREDIT_BALANCE > 0 Then
                        TOTAL_DEBIT_BALANCE = TOTAL_DEBIT_BALANCE + DEBIT_BALANCE
                        TOTAL_CREDIT_BALANCE = TOTAL_CREDIT_BALANCE + CREDIT_BALANCE
                        cnt = cnt + 1
    
                        'gconDMIS.Execute "update AMIS_ChartAccount Set" & _
                         " Debit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!DEBIT_TOTAL)), 2) & "," & _
                         " Credit_Total = " & Round(NumericVal(Null2String(rsJournal_HD!CREDIT_TOTAL)), 2) & "," & _
                         " DebitBalance = " & DEBIT_BALANCE & "," & _
                         " CreditBalance = " & CREDIT_BALANCE & _
                         " Where AcctCode = '" & Null2String(rsJournal_HD!Acct_Code) & "'"
    
                        J_JITEMNO = "'" & Format(cnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(rsJournal_HD!ACCT_CODE)
                        J_ACCT_NAME = N2Str2Null(Setacctname(Null2String(rsJournal_HD!ACCT_CODE)))
                        J_DEBIT = Round(CREDIT_BALANCE, 2)
                        J_CREDIT = Round(DEBIT_BALANCE, 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If

                    rsJournal_HD.MoveNext
                Loop
                If TOTAL_DEBIT_BALANCE - TOTAL_CREDIT_BALANCE > 0 Then
                    cnt = cnt + 1
                    J_JITEMNO = "'" & Format(cnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("REU"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("REU")))
                    J_DEBIT = Round(TOTAL_DEBIT_BALANCE - TOTAL_CREDIT_BALANCE, 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                Else
                    cnt = cnt + 1
                    J_JITEMNO = "'" & Format(cnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("REU"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("REU")))
                    J_DEBIT = 0
                    J_CREDIT = Round(TOTAL_CREDIT_BALANCE - TOTAL_DEBIT_BALANCE, 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", 'CLO', " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                gconDMIS.Execute "Update AMIS_AccountingPeriod Set ActivePeriod=0 where Status=1"
                cmdCancel_Click
                MsgBox "Closing Entries for Accounting Year = " & Format(dtTo.Value, "mm") & "/" & Format(dtTo.Value, "dd") & "/" & Format(dtTo.Value, "yyyy") & " Successfully Created!", vbInformation, "Done"
                FillGrid
                Screen.MousePointer = 0
            Else
                gconDMIS.Execute "Update AMIS_AccountingPeriod Set ActivePeriod=0 where Status=1"
                FillGrid
                Screen.MousePointer = 0
            End If
            
    ClosingEntries = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & Err.Number & vbCrLf & "Error Description :" & Err.DESCRIPTION
    ClosingEntries = False
End Function

Private Sub cmdview_Click()
'FillGrid
    ListCurrentPeriod
End Sub

Private Sub cmdYearEnd_Click()
    picOpenClose.Enabled = False
    picAcctYear.Visible = True
    picAcctYear.ZOrder 0
    InitCombo
End Sub

Private Sub dtFrom_Change()
    dtFrom.Value = firstDay(dtFrom.Value)
    dtTo.Value = lastDay(DateAdd("m", 11, dtFrom.Value))
End Sub

Private Sub dtTo_Change()
    dtTo.Value = lastDay(DateAdd("m", 11, dtFrom.Value))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF12
        cmdClosePeriod.Caption = "&Open Journal Period"
        cmdClosePeriod.BackColor = &HC0E0FF
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]" '"." & App.Revision & "]"
    'SetOpenClose
    CenterMe frmMain, Me, 1
    If SetActivePeriod = True Then
        dtFrom.Value = firstDay(xPeriodFrom)
        dtTo.Value = lastDay(DateAdd("m", 11, xPeriodFrom))
    End If

    If CheckAcctYear = False Then
        NewAccountingPeriod
    End If
    initGrid
    InitCombo
    ListCurrentPeriod
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xCurrentPeriod = False
End Sub

Private Sub gridAccounting_Click()
    On Error Resume Next
    If gridAccounting.Rows <= 2 Then
        Exit Sub
    Else
        If gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "Open" Then
            cmdClosePeriod.Caption = "&Month End Process"
            cmdClosePeriod.BackColor = &H8000000F
            '        ElseIf gridAccounting.Cell(gridAccounting.ActiveCell.Row, 5).Text = "Closed" Then
            '            cmdClosePeriod.Caption = "&Open Journal Period"
            '            cmdClosePeriod.BackColor = &HC0E0FF
        End If
        xAccountingPeriod = Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "mmmm") & " " & Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "yyyy")
        xAccountingMonth = Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "mm/dd/yyyy")
        xAccountingMonth2 = lastDay(Format(gridAccounting.Cell(gridAccounting.ActiveCell.Row, 2).Text, "mm/dd/yyyy"))
        xBackMonth = xAccountingMonth - 1
        xNextMonth = xAccountingMonth + Day(lastDay(xAccountingMonth))
        Dim xRow, xColumn                              As Integer
        xRow = NumericVal(gridAccounting.ActiveCell.Row) - NumericVal(gridAccounting.ActiveCell.Row)
        xColumn = NumericVal(gridAccounting.ActiveCell.Col)
        xJType = gridAccounting.Cell(xRow, xColumn).Text
        If gridAccounting.Cell(xRow, xColumn).Text = "APJ" Then
            lblAPJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbRed
            lblCDJPeriod.ForeColor = vbBlue
            lblSJPeriod.ForeColor = vbBlue
            lblCRJPeriod.ForeColor = vbBlue
            lblDRJPeriod.ForeColor = vbBlue
            lblGJPeriod.ForeColor = vbBlue
        ElseIf gridAccounting.Cell(xRow, xColumn).Text = "CDJ" Then
            lblCDJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbBlue
            lblCDJPeriod.ForeColor = vbRed
            lblSJPeriod.ForeColor = vbBlue
            lblCRJPeriod.ForeColor = vbBlue
            lblDRJPeriod.ForeColor = vbBlue
            lblGJPeriod.ForeColor = vbBlue
        ElseIf gridAccounting.Cell(xRow, xColumn).Text = "SJ" Then
            lblSJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbBlue
            lblCDJPeriod.ForeColor = vbBlue
            lblSJPeriod.ForeColor = vbRed
            lblCRJPeriod.ForeColor = vbBlue
            lblDRJPeriod.ForeColor = vbBlue
            lblGJPeriod.ForeColor = vbBlue
        ElseIf gridAccounting.Cell(xRow, xColumn).Text = "CRJ" Then
            lblCRJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbBlue
            lblCDJPeriod.ForeColor = vbBlue
            lblSJPeriod.ForeColor = vbBlue
            lblCRJPeriod.ForeColor = vbRed
            lblDRJPeriod.ForeColor = vbBlue
            lblGJPeriod.ForeColor = vbBlue
        ElseIf gridAccounting.Cell(xRow, xColumn).Text = "DRJ" Then
            lblDRJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbBlue
            lblCDJPeriod.ForeColor = vbBlue
            lblSJPeriod.ForeColor = vbBlue
            lblCRJPeriod.ForeColor = vbBlue
            lblDRJPeriod.ForeColor = vbRed
            lblGJPeriod.ForeColor = vbBlue
        ElseIf gridAccounting.Cell(xRow, xColumn).Text = "GJ" Then
            lblGJPeriod.Caption = xAccountingPeriod
            lblAPJPeriod.ForeColor = vbBlue
            lblCDJPeriod.ForeColor = vbBlue
            lblSJPeriod.ForeColor = vbBlue
            lblCRJPeriod.ForeColor = vbBlue
            lblDRJPeriod.ForeColor = vbBlue
            lblGJPeriod.ForeColor = vbRed
        End If
    End If
    With gridAccounting
        .BackColorFixedSel = vbRed
    End With
End Sub

Sub initGrid()
    With gridAccounting
        .Cols = 9: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True
        .Cell(0, 0).Text = ""
        .Column(0).Width = 0
        .SelectionMode = cellSelectionFree

        .DrawMode = cellOwnerDraw

        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Period"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "APJ"
        .Cell(0, 4).Text = "CDJ"
        .Cell(0, 5).Text = "SJ"
        .Cell(0, 6).Text = "CRJ"
        .Cell(0, 7).Text = "DRJ"
        .Cell(0, 8).Text = "GJ"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox
        .Column(3).CellType = cellTextBox
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox

        '.Column(6).Mask = cellLetter

        '.ComboBox(6).AddItem ("Open")
        '.ComboBox(6).AddItem ("Close")

        .Column(1).Width = 50: .Column(1).Locked = True: .Column(1).Alignment = cellCenterGeneral
        .Column(2).Width = 75: .Column(2).Locked = True: .Column(2).Alignment = cellLeftGeneral
        .Column(3).Width = 48: .Column(3).Locked = True: .Column(3).Alignment = cellCenterGeneral
        .Column(4).Width = 48: .Column(4).Locked = True: .Column(4).Alignment = cellCenterGeneral
        .Column(5).Width = 48: .Column(5).Locked = True: .Column(5).Alignment = cellCenterGeneral
        .Column(6).Width = 48: .Column(6).Locked = True: .Column(6).Alignment = cellCenterGeneral
        .Column(7).Width = 48: .Column(7).Locked = True: .Column(7).Alignment = cellCenterGeneral
        .Column(8).Width = 48: .Column(8).Locked = True: .Column(8).Alignment = cellCenterGeneral
        .AllowUserSort = False
        .RowHeight(0) = 20
        .Range(1, 5, .Rows - 1, 5).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Sub InitCombo()
    cboBookType.Clear
    cboYear.Clear
    Dim rsBooks                                        As ADODB.Recordset
    Set rsBooks = New ADODB.Recordset
    rsBooks.Open "Select * from AMIS_Books", gconDMIS, adOpenForwardOnly
    If Not rsBooks.EOF And Not rsBooks.BOF Then
        Do While Not rsBooks.EOF
            cboBookType.AddItem Null2String(rsBooks!JTYPE)
            rsBooks.MoveNext
        Loop
    End If
    cboBookType.ListIndex = 0
    xCurrentPeriod = False
    
    cboYear.Clear
    Dim rsYear As ADODB.Recordset
    Set rsYear = New ADODB.Recordset
    rsYear.Open "SELECT DISTINCT YEAR(ACCTMONTH) AS ACCTYEAR FROM AMIS_ACCOUNTINGPERIOD WHERE ActivePeriod=1", gconDMIS, adOpenForwardOnly
    If Not rsYear.EOF And Not rsYear.BOF Then
        Do While Not rsYear.EOF
            cboYear.AddItem Null2String(rsYear!ACCTYEAR)
            rsYear.MoveNext
        Loop
    End If
    Set rsBooks = Nothing
    Set rsYear = Nothing
End Sub

Sub FillGrid()
    On Error GoTo ErrorCode
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Dim Qtr                                            As String
    Dim iRow                                           As Long
    Dim xAPJStatus                                     As String
    Dim xCDJStatus                                     As String
    Dim xSJStatus                                      As String
    Dim xCRJStatus                                     As String
    Dim xDRJStatus                                     As String
    Dim xGJStatus                                      As String
    gridAccounting.Rows = 1
    gridAccounting.AutoRedraw = False
    Set rsAccountingPeriod = New ADODB.Recordset
    If xCurrentPeriod = True Then
        rsAccountingPeriod.Open "SELECT ACCTYEAR,ACCTMONTH,APJ,CDJ,SJ,CRJ,DRJ,GJ FROM (SELECT YEAR(ACCTMONTH)AS ACCTYEAR, ACCTMONTH,SUM(CASE WHEN JTYPE='APJ' THEN STATUS END) AS APJ, SUM(CASE WHEN JTYPE='CDJ' THEN STATUS END) AS CDJ, " & _
                                "SUM(CASE WHEN JTYPE='SJ' THEN STATUS END) AS SJ,SUM(CASE WHEN JTYPE='CRJ' THEN STATUS END) AS CRJ,SUM(CASE WHEN JTYPE='DRJ' THEN STATUS END) AS DRJ, " & _
                                "SUM(CASE WHEN JTYPE='GJ' THEN STATUS END) AS GJ,ACTIVEPERIOD FROM AMIS_ACCOUNTINGPERIOD GROUP BY ACCTMONTH,ACTIVEPERIOD) X " & _
                                "WHERE ACCTYEAR = '" & Format(LOGDATE, "yyyy") & "' AND ACTIVEPERIOD=1 order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    Else
        rsAccountingPeriod.Open "SELECT ACCTYEAR,ACCTMONTH,APJ,CDJ,SJ,CRJ,DRJ,GJ FROM (SELECT YEAR(ACCTMONTH)AS ACCTYEAR, ACCTMONTH,SUM(CASE WHEN JTYPE='APJ' THEN STATUS END) AS APJ, SUM(CASE WHEN JTYPE='CDJ' THEN STATUS END) AS CDJ, " & _
                                "SUM(CASE WHEN JTYPE='SJ' THEN STATUS END) AS SJ,SUM(CASE WHEN JTYPE='CRJ' THEN STATUS END) AS CRJ,SUM(CASE WHEN JTYPE='DRJ' THEN STATUS END) AS DRJ, " & _
                                "SUM(CASE WHEN JTYPE='GJ' THEN STATUS END) AS GJ,ACTIVEPERIOD FROM AMIS_ACCOUNTINGPERIOD GROUP BY ACCTMONTH,ACTIVEPERIOD) X " & _
                                "WHERE ACTIVEPERIOD=1 order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    End If
    ' where ActivePeriod=1
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        Do While Not rsAccountingPeriod.EOF
            'If Format(rsAccountingPeriod!AcctMonth, "yyyy-mm") <= Format(LOGDATE, "yyyy-mm") Then
            iRow = iRow + 1
            If rsAccountingPeriod!APJ = 0 And rsAccountingPeriod!CDJ = 0 And rsAccountingPeriod!SJ = 0 And rsAccountingPeriod!CRJ = 0 And rsAccountingPeriod!DRJ = 0 And rsAccountingPeriod!GJ = 0 Then
                xAPJStatus = "Open"
                xCDJStatus = "Open"
                xSJStatus = "Open"
                xCRJStatus = "Open"
                xDRJStatus = "Open"
                xGJStatus = "Open"
            Else
                xAPJStatus = "Closed"
                xCDJStatus = "Closed"
                xSJStatus = "Closed"
                xCRJStatus = "Closed"
                xDRJStatus = "Closed"
                xGJStatus = "Closed"
            End If
            gridAccounting.AddItem iRow & vbTab & Null2String(rsAccountingPeriod!ACCTYEAR) + "-" + Format(rsAccountingPeriod!AcctMonth, "mmm") & vbTab & _
                                   xAPJStatus & vbTab & xCDJStatus & vbTab & xSJStatus & vbTab & xCRJStatus & vbTab & xDRJStatus & vbTab & xGJStatus
            rsAccountingPeriod.MoveNext
        Loop
        gridAccounting.AutoRedraw = True
        gridAccounting.Refresh
    Else
        MsgBox "No record to view...", vbInformation, "Message"
    End If
    Set rsAccountingPeriod = Nothing
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function SetBookType(xBookType As String)
    Dim rsBookType                                     As ADODB.Recordset
    Set rsBookType = New ADODB.Recordset
    rsBookType.Open "Select * from AMIS_Books where JType = " & N2Str2Null(xBookType), gconDMIS, adOpenForwardOnly
    If Not rsBookType.EOF And Not rsBookType.BOF Then
        SetBookType = N2Str2Null(rsBookType!Code)
    End If
End Function

Function CheckOpenBook(nBackMonth As Date) As String
    Dim xSTATUS                                        As String
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "select Distinct AcctMonth,Status,ActivePeriod from AMIS_AccountingPeriod where AcctMonth = '" & firstDay(nBackMonth) & "' order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        If Null2String(rsAccountingPeriod!Status) = 0 Then
            xSTATUS = "Open"
        Else
            xSTATUS = "Closed"
        End If
        CheckOpenBook = xSTATUS
    End If
    Set rsAccountingPeriod = Nothing
End Function

Function CheckClosedBook(nNextMonth As Date) As String
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "select Distinct AcctMonth,Status,ActivePeriod from AMIS_AccountingPeriod where AcctMonth = '" & firstDay(nNextMonth) & "' order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        CheckClosedBook = rsAccountingPeriod!Status
    End If
    Set rsAccountingPeriod = Nothing
End Function

Function GetIDNextMonth(nNextMonth As Date)
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "select JType,AcctMonth,Status,ID from AMIS_AccountingPeriod where AcctMonth = '" & firstDay(nNextMonth) & "' order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        Do While Not rsAccountingPeriod.EOF
            GetIDNextMonth = rsAccountingPeriod!ID
            rsAccountingPeriod.MoveNext
        Loop
    End If
    Set rsAccountingPeriod = Nothing
End Function

Function LastPeriod(xBookType As String, xLogDate As String) As Boolean
    Dim rsLastPeriod                                   As ADODB.Recordset
    Set rsLastPeriod = New ADODB.Recordset
    rsLastPeriod.Open "Select * from AMIS_AccountingPeriod where JType =" & SetBookType(xBookType) & " and AcctYear = '" & Format(xLogDate, "yyyy") & "' order by JType,AcctMonth ASC", gconDMIS, adOpenForwardOnly
    If Not rsLastPeriod.EOF And Not rsLastPeriod.BOF Then
        Do While Not rsLastPeriod.EOF
            If Format(lastDay(rsLastPeriod!AcctMonth), "yyyy-mm-dd") <> Format(dtTo.Value, "yyyy-mm-dd") Then
                LastPeriod = True
            End If
            rsLastPeriod.MoveNext
        Loop
    End If
End Function

Sub ListCurrentPeriod()
    On Error GoTo ErrorCode
    Dim GridNo, GridNo2, xRow                          As Integer
    Dim xYearMonth                                     As String
    Dim xList                                          As ListItem
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "select JType,AcctMonth from AMIS_AccountingPeriod where CurrPeriod = 1 order by AcctMonth ASC", gconDMIS, adOpenForwardOnly
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        '        lblAPJPeriod.Caption = Null2String(Format(rsAccountingPeriod!ACCTMONTH, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!ACCTMONTH, "yyyy"))
        '    Else
        '        lblAPJPeriod.Caption = ""
        '    End If
        'InitGrid
        With gridAccounting
            For GridNo = 1 To .Rows - 1

                Do While Not rsAccountingPeriod.EOF
                    If rsAccountingPeriod!JTYPE = "APJ" Then
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        .Cell(xRow, 3).BackColor = vbYellow
                        lblAPJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    ElseIf rsAccountingPeriod!JTYPE = "CDJ" Then
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        .Cell(xRow, 4).BackColor = vbYellow
                        lblCDJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    ElseIf rsAccountingPeriod!JTYPE = "SJ" Then
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        .Cell(xRow, 5).BackColor = vbYellow
                        lblSJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    ElseIf rsAccountingPeriod!JTYPE = "CRJ" Then
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        .Cell(xRow, 6).BackColor = vbYellow
                        lblCRJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    ElseIf rsAccountingPeriod!JTYPE = "DRJ" Then
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        .Cell(xRow, 7).BackColor = vbYellow
                        lblDRJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    ElseIf rsAccountingPeriod!JTYPE = "GJ" Then
                        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
                        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
                        For GridNo2 = 1 To .Rows - 1
                            If .Cell(GridNo2, 2).Text = xYearMonth Then
                                xRow = GridNo2
                            End If
                        Next
                        .Cell(xRow, 8).BackColor = vbYellow
                        lblGJPeriod.Caption = Null2String(Format(rsAccountingPeriod!AcctMonth, "mmmm")) & " " & Null2String(Format(rsAccountingPeriod!AcctMonth, "yyyy"))
                    End If
                    rsAccountingPeriod.MoveNext
                Loop
            Next
        End With
    End If


    Set rsAccountingPeriod = Nothing
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function CheckUnposted() As Boolean
    On Error Resume Next
    Dim xList                                          As ListItem
    Dim rsCheckStatus                                  As ADODB.Recordset
    Set rsCheckStatus = New ADODB.Recordset
    rsCheckStatus.Open "SELECT VoucherNo,JType,JDate,Sum(Debit) as Debit,Sum(Credit) as Credit FROM AMIS_Journal_Det where Status <> 'P' and Status <> 'C' and JType IN ('APJ','CDJ','SJ','CRJ','DRJ','GJ') and JDate >= '" & firstDay(xAccountingMonth) & "' and JDate <= '" & lastDay(xAccountingMonth) & "' Group by VoucherNo,JType,JDate Order by JType,JDate", gconDMIS, adOpenForwardOnly
    lvStatus.ListItems.Clear
    If Not rsCheckStatus.EOF And Not rsCheckStatus.BOF Then
        CheckUnposted = True
        MsgBox "Unable to close this period..." & vbCrLf & "See details", vbExclamation, "Checking transactions"
        'Me.Width = 9375
        'picControl.Left = 6180
        'picOpenClose.Width = 9165
        Me.Width = 9960
        picOpenClose.Width = 9765
        picControl.Left = 6780
        picStatus.Visible = True
        picStatus.ZOrder 0
        Do While Not rsCheckStatus.EOF
            DoEvents
            Set xList = lvStatus.ListItems.Add(, , Format(Null2Date(rsCheckStatus!JDATE), "mm/dd/yyyy"))
            xList.SubItems(1) = Null2String(rsCheckStatus!VOUCHERNO)
            xList.SubItems(2) = Null2String(rsCheckStatus!JTYPE)
            xList.SubItems(3) = ToDoubleNumber(NumericVal(rsCheckStatus!Debit))
            xList.SubItems(4) = ToDoubleNumber(NumericVal(rsCheckStatus!Credit))
            '                        lblVoucherNo.Caption = Null2String(rsCheckStatus!VOUCHERNO)
            rsCheckStatus.MoveNext
        Loop
        '        picCheck.Visible = False
        '        picStatus.Visible = True
        '        picNext.Visible = True
        Screen.MousePointer = 0
    Else
        '        picStatus.Visible = False
        '        picControl.Enabled = True
        '        picNext.Visible = False
        '        picCheck.Visible = False
    End If
    Set rsCheckStatus = Nothing
End Function

Sub CheckTrialStatus()
    Dim xList                                          As ListItem
    Dim rsTrialBalance                                 As ADODB.Recordset
    Set rsTrialBalance = New ADODB.Recordset
    Set rsTrialBalance = gconDMIS.Execute("SELECT * FROM (SELECT VOUCHERNO,JTYPE,JDate,SUM(DEBIT) AS DEBIT, SUM(CREDIT) AS CREDIT FROM AMIS_JOURNAL_DET " & _
                                          " WHERE STATUS = 'P' and JDate <= '" & lastDay(xAccountingMonth) & "' GROUP BY VOUCHERNO,JTYPE,JDate) A WHERE DEBIT <> CREDIT")
    If Not rsTrialBalance.EOF And Not rsTrialBalance.BOF Then
        Do While Not rsTrialBalance.EOF
            DoEvents
            '            Set xList = lvStatus.ListItems.Add(, , Format(Null2Date(rsTrialBalance!JDate), "mm/dd/yyyy"))
            '            xList.SubItems(1) = Null2String(rsTrialBalance!VOUCHERNO)
            '            xList.SubItems(2) = Null2String(rsTrialBalance!jtype)
            '            xList.SubItems(3) = NumericVal(rsTrialBalance!DEBIT)
            '            xList.SubItems(4) = NumericVal(rsTrialBalance!CREDIT)
            '            lblVoucherNo.Caption = Null2String(rsTrialBalance!VOUCHERNO)
            rsTrialBalance.MoveNext
        Loop
    Else
        MsgBox "Trial Balance is balance", vbInformation, "Trial Balance Status"
    End If
    Set rsTrialBalance = Nothing
End Sub

Function GetVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'CLO' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function Setacctname(VVV As String) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
    Set rsChartAccount2 = Nothing
End Function

Function SetActivePeriod() As Boolean
    On Error Resume Next
    Dim rsActivePeriod                                 As ADODB.Recordset
    Set rsActivePeriod = New ADODB.Recordset
    rsActivePeriod.Open "Select * from AMIS_AccountingPeriod where ActivePeriod = 1 ", gconDMIS, adOpenKeyset
    If Not rsActivePeriod.EOF And Not rsActivePeriod.BOF Then
        rsActivePeriod.MoveFirst
        xPeriodFrom = rsActivePeriod!AcctMonth
        SetActivePeriod = True
    End If
    Set rsActivePeriod = Nothing
End Function

Function CheckCloseBook() As Boolean
    Dim rsCheckCloseBook                               As ADODB.Recordset
    Set rsCheckCloseBook = New ADODB.Recordset
    rsCheckCloseBook.Open "SELECT DISTINCT ACCTYEAR FROM (SELECT YEAR(ACCTMONTH) AS ACCTYEAR,STATUS FROM AMIS_ACCOUNTINGPERIOD WHERE Status=0) X  WHERE AcctYear = '" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    If Not rsCheckCloseBook.EOF And Not rsCheckCloseBook.BOF Then
        CheckCloseBook = False
    Else
        CheckCloseBook = True
    End If
    Set rsCheckCloseBook = Nothing
End Function

Function CheckAcctYear() As Boolean
    On Error Resume Next
    Dim rsCheckAcctYear                                As ADODB.Recordset
    Set rsCheckAcctYear = New ADODB.Recordset
    rsCheckAcctYear.Open "SELECT ACCTYEAR FROM (SELECT DISTINCT YEAR(ACCTMONTH) AS ACCTYEAR FROM AMIS_ACCOUNTINGPERIOD) X WHERE ACCTYEAR <= '" & Format(LOGDATE, "yyyy") & "'", gconDMIS, adOpenKeyset
    If Not rsCheckAcctYear.EOF And Not rsCheckAcctYear.BOF Then
        CheckAcctYear = False
    End If
    Set rsCheckAcctYear = Nothing
End Function
Sub BalanceForwarding()
'================ CURRENT ================
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                         " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "')" & _
                                         " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '40')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '40'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & xAccountingMonth & "' AND AMIS_Journal_Det.jdate <= '" & xAccountingMonth2 & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        Current_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
    End If
End Sub


Private Sub gridAccounting_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
'    Dim rc As RECT
'    Dim lngLeft As Long
'
'    If Row = 0 Then
'        If Col = 1 Then
'            lngLeft = Left + (Right - Left - GetTextWidth(hDC, gridAccounting.Cell(Row, Col).Text) - 18) / 2
'            SetRect rc, lngLeft + 18, Top, Right, Bottom
'            DrawIconEx hDC, lngLeft, Top + (Bottom - Top - 16) / 2, imgNote.Picture, 16, 16, 0, 0, DI_NORMAL
'            DrawText hDC, gridAccounting.Cell(Row, Col).Text, -1, rc, DT_SINGLELINE Or DT_VCENTER
'            Handled = True
'        ElseIf Col = 6 Then
'            lngLeft = Left + (Right - Left - GetTextWidth(hDC, gridAccounting.Cell(Row, Col).Text) - 18) / 2
'            SetRect rc, lngLeft + 18, Top, Right, Bottom
'            DrawIconEx hDC, lngLeft, Top + (Bottom - Top - 16) / 2, imgEarth.Picture, 16, 16, 0, 0, DI_NORMAL
'            DrawText hDC, gridAccounting.Cell(Row, Col).Text, -1, rc, DT_SINGLELINE Or DT_VCENTER
'            Handled = True
'        End If
End Sub

Private Sub Label_Click(Index As Integer)
    xCurrentPeriod = True
    FillGrid
    ListCurrentPeriod
End Sub

Private Sub lvStatus_DblClick()
    Dim rsCheckStatus                                  As ADODB.Recordset
    Set rsCheckStatus = New ADODB.Recordset
    Dim xAMIS_JOURNAL                                  As String
    'rsCheckStatus.Open "SELECT VoucherNo,JType,JDate,Sum(Debit) as Debit,Sum(Credit) as Credit FROM AMIS_Journal_Det where Status <> 'P' and Status <> 'C' and JType <> 'BOB' and JDate <= '" & lastDay(xAccountingMonth) & "' and VoucherNo ='" & lvStatus.SelectedItem.SubItems(1) & "' AND JTYPE= '" & lvStatus.SelectedItem.SubItems(2) & "' Group by VoucherNo,JType,JDate", gconDMIS, adOpenForwardOnly
    rsCheckStatus.Open "SELECT VoucherNo,JType FROM AMIS_Journal_Det where Status <> 'P' and Status <> 'C' and JType <> 'BOB' and JDate <= '" & lastDay(xAccountingMonth) & "' and VoucherNo ='" & lvStatus.SelectedItem.SubItems(1) & "' AND JTYPE= '" & lvStatus.SelectedItem.SubItems(2) & "'", gconDMIS, adOpenForwardOnly
    If Not rsCheckStatus.EOF And Not rsCheckStatus.BOF Then
        xVOUCHERNO = rsCheckStatus!VOUCHERNO
        JOURNALTYPE = rsCheckStatus!JTYPE
        If JOURNALTYPE = "APJ" Then
            Call frmAMISJournalEntry_APJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_APJ.Show
            frmAMISJournalEntry_APJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_APJ.ZOrder 0
        ElseIf JOURNALTYPE = "CDJ" Then
            Call frmAMISJournalEntry_CDJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_CDJ.Show
            frmAMISJournalEntry_CDJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_CDJ.ZOrder 0
        ElseIf JOURNALTYPE = "SJ" Then
            Call frmAMISJournalEntry_SJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_SJ.Show
            frmAMISJournalEntry_SJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_SJ.ZOrder 0
        ElseIf JOURNALTYPE = "CRJ" Then
            Call frmAMISJournalEntry_CRJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_CRJ.Show
            frmAMISJournalEntry_CRJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_CRJ.ZOrder 0
        ElseIf JOURNALTYPE = "DRJ" Then
            Call frmAMISJournalEntry_DRJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_DRJ.Show
            frmAMISJournalEntry_DRJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_DRJ.ZOrder 0
        ElseIf JOURNALTYPE = "GJ" Then
            Call frmAMISJournalEntry_GJ.LOADJOURNAL(JOURNALTYPE)
            frmAMISJournalEntry_GJ.Show
            frmAMISJournalEntry_GJ.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISJournalEntry_GJ.ZOrder 0
        ElseIf JOURNALTYPE = "COB" Then
            Call frmAMISCustomerAROpening.LOADJOURNAL(JOURNALTYPE)
            frmAMISCustomerAROpening.Show
            frmAMISCustomerAROpening.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISCustomerAROpening.ZOrder 0
        ElseIf JOURNALTYPE = "VPJ" Then
            Call frmAMISVendorAPOpening.LOADJOURNAL(JOURNALTYPE)
            frmAMISVendorAPOpening.Show
            frmAMISVendorAPOpening.SearchVoucherNo Trim(xVOUCHERNO)
            frmAMISVendorAPOpening.ZOrder 0
        End If
        '        frmAMISJournalEntry.Show
        '        frmAMISJournalEntry.SearchVoucherNo Trim(xVOUCHERNO)
        '        frmAMISJournalEntry.ZOrder 0
    End If
    Set rsCheckStatus = Nothing
End Sub

Sub NewAccountingPeriod()
    On Error GoTo ErrorCode
    Dim xAcctMonth                                     As Date
    Dim xDate                                          As Date
    Dim xMonth                                         As Integer
    Dim xDay                                           As Integer
    Dim xYear                                          As Integer
    Dim iMonth, iMonth2                                As Integer
    Dim rsBooks                                        As ADODB.Recordset
    Dim rsAccountingPeriod                             As ADODB.Recordset
    xDate = firstDay(Format(LOGDATE, "mm/dd/yyyy"))
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "SELECT STARTMONTH FROM (SELECT MIN(ACCTMONTH) AS STARTMONTH FROM AMIS_ACCOUNTINGPERIOD) X  WHERE STARTMONTH IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        If Format(LOGDATE, "mm/dd/yyyy") < CDate(rsAccountingPeriod!StartMonth) Then
            '            MsgBox "Entry not permitted...", vbExclamation, "System Message"
            Exit Sub
        End If
    End If
    For iMonth = 0 To 11
        xAcctMonth = DateAdd("m", iMonth, xDate)
        Set rsAccountingPeriod = New ADODB.Recordset
        rsAccountingPeriod.Open "SELECT DISTINCT AcctMonth from AMIS_AccountingPeriod where AcctMonth ='" & xAcctMonth & "'", gconDMIS, adOpenForwardOnly
        '        If rsAccountingPeriod.RecordCount = 0 Then Exit Sub
        If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
            '                    MsgBox "Accounting Period already exist...", vbInformation, "Accounting Year"
            Exit Sub
        Else
            'gconDMIS.Execute "Update AMIS_AccountingPeriod Set ActivePeriod=0"
            If iMonth = 11 Then
                '                        If MsgBox("Are you sure you want to set this accounting period?", vbQuestion + vbYesNo, "Accounting Year") = vbYes Then
                xMonth = 1
                xDay = 1
                xYear = Format(xDate, "yyyy")
                xDate = xMonth & "/" & xDay & "/" & xYear
                Set rsBooks = New ADODB.Recordset
                rsBooks.Open "Select Code from AMIS_Books", gconDMIS, adOpenForwardOnly
                If Not rsBooks.EOF And Not rsBooks.BOF Then
                    Do While Not rsBooks.EOF
                        For iMonth2 = 0 To 11
                            xAcctMonth = DateAdd("m", iMonth2, xDate)
                            gconDMIS.Execute "Insert into AMIS_AccountingPeriod (JType,AcctMonth,Status,CurrPeriod,ActivePeriod) values (" & N2Str2Null(rsBooks!Code) & ",'" & CDate(xAcctMonth) & "',0,0,1)"
                        Next iMonth2
                        rsBooks.MoveNext
                    Loop
                End If
                '                                    MsgBox "New Accounting Period Save!", vbInformation, "Accounting Year"
                '                        Else
                '                            Exit Sub
                '                        End If
            End If
        End If
    Next iMonth
    Set rsBooks = Nothing
    Set rsAccountingPeriod = Nothing
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function ClearPreviousPeriod(XXX As String) As Integer
    Dim xYearMonth                                     As String
    Dim GridNo                                         As Integer
    Dim rsAccountingPeriod                             As ADODB.Recordset
    Set rsAccountingPeriod = New ADODB.Recordset
    Dim xRow, xColor                                   As Integer
    rsAccountingPeriod.Open "Select * from AMIS_ACCOUNTINGPERIOD where CurrPeriod=1 and JType='" & XXX & "'", gconDMIS, adOpenKeyset
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        xYearMonth = NumericVal(Format(rsAccountingPeriod!AcctMonth, "yyyy")) & "-" & Null2String(Format(rsAccountingPeriod!AcctMonth, "mmm"))
        With gridAccounting
            For GridNo = 1 To .Rows - 1
                If .Cell(GridNo, 2).Text = xYearMonth Then
                    xRow = GridNo
                End If
            Next
        End With
        'xRow = NumericVal(Format(rsAccountingPeriod!AcctMonth, "mm"))
        xColor = xRow Mod 2
        With gridAccounting
            If xColor = 0 Then
                .Cell(xRow, .ActiveCell.Col).BackColor = RGB(239, 243, 255)
            Else
                .Cell(xRow, .ActiveCell.Col).BackColor = RGB(231, 235, 247)
            End If
        End With
    End If
End Function


Function ReturnAccountCode(XXX As String, Optional YYY As String)
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If YYY = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "' AND TRANTYPE3 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function


