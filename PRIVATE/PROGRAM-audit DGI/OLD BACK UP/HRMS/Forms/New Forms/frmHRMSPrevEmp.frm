VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSPrevEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Previous Employer"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   Icon            =   "frmHRMSPrevEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11970
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   60
      MaxLength       =   35
      TabIndex        =   45
      Top             =   150
      Width           =   2475
   End
   Begin VB.Frame fmeInfo 
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
      Height          =   5955
      Left            =   2610
      TabIndex        =   13
      Top             =   60
      Width           =   9315
      Begin VB.TextBox txtTaxWithheld 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6930
         TabIndex        =   47
         Top             =   5490
         Width           =   2115
      End
      Begin VB.TextBox txtTotalTaxable 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2370
         TabIndex        =   42
         Top             =   4410
         Width           =   2115
      End
      Begin VB.TextBox txtTaxSalaries 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2370
         TabIndex        =   40
         Top             =   3870
         Width           =   2115
      End
      Begin VB.TextBox txtTax13thMonth 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   2370
         TabIndex        =   38
         Top             =   3390
         Width           =   2115
      End
      Begin VB.TextBox txtNONTaxPremium 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6930
         TabIndex        =   35
         Top             =   4440
         Width           =   2115
      End
      Begin VB.TextBox txtNONTaxSalaries 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6930
         TabIndex        =   33
         Top             =   3900
         Width           =   2115
      End
      Begin VB.TextBox txtNONTax13thMonth 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   6930
         TabIndex        =   31
         Top             =   3330
         Width           =   2115
      End
      Begin VB.TextBox txtCompName 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2400
         TabIndex        =   28
         Top             =   1170
         Width           =   5925
      End
      Begin VB.TextBox txtCompAdd 
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
         Height          =   330
         Left            =   2400
         TabIndex        =   27
         Top             =   1530
         Width           =   5925
      End
      Begin VB.TextBox txtTin 
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
         Height          =   330
         Left            =   2400
         TabIndex        =   26
         Top             =   1920
         Width           =   2115
      End
      Begin VB.TextBox txtZip 
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
         Height          =   330
         Left            =   6240
         TabIndex        =   25
         Top             =   1950
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   2430
         TabIndex        =   14
         Top             =   2400
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         Format          =   40173569
         CurrentDate     =   39352
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   6300
         TabIndex        =   15
         Top             =   2430
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         Format          =   40173569
         CurrentDate     =   39352
      End
      Begin VB.Label lblYear 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   7650
         TabIndex        =   49
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Withheld"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   4920
         TabIndex        =   48
         Top             =   5550
         Width           =   1455
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   315
         Left            =   0
         TabIndex        =   46
         Top             =   5100
         Width           =   9285
         _Version        =   655364
         _ExtentX        =   16378
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "TAX WITHHELD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   7
         Left            =   330
         TabIndex        =   43
         Top             =   4410
         Width           =   1905
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salaries and Other Compensation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Index           =   6
         Left            =   330
         TabIndex        =   41
         Top             =   3780
         Width           =   1905
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "13th Month and Other Benefits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Index           =   4
         Left            =   330
         TabIndex        =   39
         Top             =   3270
         Width           =   1905
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   0
         TabIndex        =   37
         Top             =   720
         Width           =   9285
         _Version        =   655364
         _ExtentX        =   16378
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "PREVIOUS EMPLOYER"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS, GSIS, PAGIBIG and Union Dues"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   720
         Index           =   15
         Left            =   4890
         TabIndex        =   36
         Top             =   4350
         Width           =   1905
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salaries and Other Compensation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Index           =   14
         Left            =   4890
         TabIndex        =   34
         Top             =   3810
         Width           =   1905
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "13th Month and Other Benefits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Index           =   13
         Left            =   4860
         TabIndex        =   32
         Top             =   3270
         Width           =   1905
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   570
         TabIndex        =   30
         Top             =   1650
         Width           =   1755
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   810
         TabIndex        =   29
         Top             =   1260
         Width           =   1515
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   4680
         TabIndex        =   24
         Top             =   2850
         Width           =   5865
         _Version        =   655364
         _ExtentX        =   10345
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "NON TAXABLE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   -30
         TabIndex        =   23
         Top             =   2850
         Width           =   4725
         _Version        =   655364
         _ExtentX        =   8334
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "TAXABLE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   1320
         TabIndex        =   22
         Top             =   2430
         Width           =   1020
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   5430
         TabIndex        =   21
         Top             =   2460
         Width           =   795
      End
      Begin VB.Label lblEmpName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1890
         TabIndex        =   20
         Top             =   270
         Width           =   4035
      End
      Begin VB.Label lblEmpNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6120
         TabIndex        =   19
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tin no."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   1620
         TabIndex        =   18
         Top             =   1980
         Width           =   690
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   5250
         TabIndex        =   17
         Top             =   2010
         Width           =   870
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   6795
      Left            =   30
      ScaleHeight     =   6795
      ScaleWidth      =   2595
      TabIndex        =   11
      Top             =   150
      Width           =   2595
      Begin MSComctlLib.ListView lsvEmp 
         Height          =   6345
         Left            =   60
         TabIndex        =   12
         Top             =   390
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   11192
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
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmHRMSPrevEmp.frx":058A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
         Picture         =   "frmHRMSPrevEmp.frx":06EC
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   60
      Picture         =   "frmHRMSPrevEmp.frx":14459
      ScaleHeight     =   5145
      ScaleWidth      =   2475
      TabIndex        =   10
      Top             =   660
      Width           =   2505
   End
   Begin Crystal.CrystalReport rptPE 
      Left            =   2670
      Top             =   6180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6090
      ScaleHeight     =   855
      ScaleWidth      =   5820
      TabIndex        =   0
      Top             =   6210
      Width           =   5820
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
         Left            =   5040
         MouseIcon       =   "frmHRMSPrevEmp.frx":281B6
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":28308
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
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
         Left            =   4350
         MouseIcon       =   "frmHRMSPrevEmp.frx":2866E
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":287C0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Height          =   795
         Left            =   3660
         MouseIcon       =   "frmHRMSPrevEmp.frx":28B26
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":28C78
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
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
         Left            =   2970
         MouseIcon       =   "frmHRMSPrevEmp.frx":28FA3
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":290F5
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Height          =   795
         Left            =   2280
         MouseIcon       =   "frmHRMSPrevEmp.frx":29451
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":295A3
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Left            =   1590
         MouseIcon       =   "frmHRMSPrevEmp.frx":298B6
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":29A08
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   10260
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   7
      Top             =   6210
      Width           =   1440
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
         Left            =   30
         MouseIcon       =   "frmHRMSPrevEmp.frx":29D02
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":29E54
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
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
         Left            =   720
         MouseIcon       =   "frmHRMSPrevEmp.frx":2A1A4
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSPrevEmp.frx":2A2F6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label lblID 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmHRMSPrevEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                                         As ADODB.Recordset
Dim rsPREV                                                            As New ADODB.Recordset
Dim ADD_EDIT                                                          As String

Function FindEmpName(VEMPNO As String)
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select LastName,FirstName From HRMS_EmpInfo WHere EmpNo = '" & VEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindEmpName = Null2String(RSTMP!lastname) & "," & Null2String(RSTMP!FIRSTNAME)
    Else
        FindEmpName = ""
    End If
    Set RSTMP = Nothing
End Function

Sub EnabledAllFrame(COND As Boolean)
    lsvEmp.Enabled = COND
    picSearch.Enabled = COND
    fmeInfo.Enabled = Not COND
    Picture1.Visible = COND
    Picture2.Visible = Not COND
End Sub

Sub StoreMemVars()
    Dim FULLNAME                                                      As String

    If Not (rsPREV.EOF And rsPREV.BOF) Then
        lblID.Caption = rsPREV!ID
        
        FULLNAME = FindEmpName(rsPREV!EMPNO)

        lblEmpNo.Caption = Null2String(rsPREV!EMPNO)
        lblEmpName.Caption = FULLNAME

        txtCompAdd.Text = Null2String(rsPREV!COMPADD)
        txtTin.Text = Null2String(rsPREV!Tin)
        txtZip.Text = Null2String(rsPREV!ZipCode)
        dtpFrom.Value = Null2String(rsPREV!FromDate)
        dtpTo.Value = Null2String(rsPREV!ToDate)
        txtNONTax13thMonth = N2Str2Zero(rsPREV!NONTax13thMonth)
        txtNONTaxPremium = N2Str2Zero(rsPREV!NONTaxPremium)
        txtNONTaxSalaries = N2Str2Zero(rsPREV!NONTaxSalaries)
        txtTax13thMonth = N2Str2Zero(rsPREV!Tax13thMonth)
        txtTaxSalaries = N2Str2Zero(rsPREV!TaxSalaries)
        txtTaxWithheld = N2Str2Zero(rsPREV!TaxWithheld)
    Else
        InitMemvars
    End If
End Sub

Sub rsRefresh_REC()
    Set rsPREV = New ADODB.Recordset
    rsPREV.Open "select * from HRMS_PrevEmp Where Empno = '" & lblEmpNo.Caption & _
                "' Order by ID asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    StoreMemVars
End Sub

Sub rsrefresh()
    Set rsEmpInfo = New ADODB.Recordset
    rsEmpInfo.Open "select * from HRMS_EmpInfo WHERE EMPLEVEL = 'E' AND PREVIOUSCOMPANY IS NOT NULL AND YEAR(DATEHIRED) = '" & GetSavedYear() & "' order by lastname,firstname,middlename asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsvEmp.Sorted = False
    lsvEmp.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo WHERE PREVIOUSCOMPANY IS NOT NULL AND YEAR(DATEHIRED) = '" & GetSavedYear() & "'order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsvEmp.ListItems, rsEMPINFO2
        lsvEmp.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)

    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsvEmp.Sorted = False
    lsvEmp.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo  where PREVIOUSCOMPANY IS NOT NULL AND YEAR(DATEHIRED) = '" & GetSavedYear() & "' and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsvEmp.ListItems, rsEMPINFO2
        lsvEmp.Refresh
    End If
End Sub

Private Sub cmdCancel_Click()
    EnabledAllFrame True
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "EMPLOYEE PREVIOUS EMPLOYER") = False Then Exit Sub

    If txtCompName.Text <> "" Then
        EnabledAllFrame False
        txtCompAdd.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0

    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Not lblEmpNo.Caption = "" Then
        If Function_Access(LOGID, "Acess_Print", "EMPLOYEE PREVIOUS EMPLOYER") = False Then Exit Sub
        Screen.MousePointer = 11
        rptPE.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
        rptPE.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
        rptPE.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
        rptPE.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
        PrintSQLReport rptPE, HRMS_REPORT_PATH & "Previous Employer.rpt", "{HRMS_PrevEmp.EmpNo} = '" & lblEmpNo.Caption & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0

        LogAudit "V", "PRINT PREVIOUS EMPLOYER RECORD", lblEmpNo.Caption
    Else
        MsgBox "Choose a Employee To Print", vbInformation, "Print Previous Employer"
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    Dim COMPNAME                                                        As String
    Dim COMPADD                                                         As String
    Dim TIN_NO                                                          As String
    Dim ZIP_CODE                                                        As String
    
    Dim VtxtNONTax13thMonth                                             As String
    Dim VtxtNONTaxSalaries                                              As String
    Dim VtxtNONTaxPremium                                               As String
    Dim VtxtTax13thMonth                                                As String
    Dim VtxtTaxSalaries                                                 As String
    Dim VtxtTaxWithheld                                                 As String
    Dim VtxtTotalTaxable                                                As String
    
    If txtCompName.Text = "" Then
        ShowIsRequiredMsg "Company Name must not Blank"
        txtCompName.SetFocus
        Exit Sub
    End If
    If txtCompAdd.Text = "" Then
        ShowIsRequiredMsg "Company Address must not Blank"
        txtCompAdd.SetFocus
        Exit Sub
    End If

    COMPNAME = N2Str2Null(txtCompName.Text)
    COMPADD = N2Str2Null(txtCompAdd.Text)
    TIN_NO = N2Str2Null(txtTin.Text)
    ZIP_CODE = N2Str2Null(txtZip.Text)
    
    VtxtNONTax13thMonth = N2Str2Null(txtNONTax13thMonth)
    VtxtNONTaxSalaries = N2Str2Null(txtNONTaxSalaries)
    VtxtNONTaxPremium = N2Str2Null(txtNONTaxPremium)
    VtxtTax13thMonth = N2Str2Null(txtTax13thMonth)
    VtxtTaxSalaries = N2Str2Null(txtTaxSalaries)
    VtxtTaxWithheld = N2Str2Null(txtTaxWithheld)
    VtxtTotalTaxable = N2Str2Null(txtTotalTaxable)
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PREVEMP WHERE EMPNO = '" & lblEmpNo.Caption & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        gconDMIS.Execute ("Update HRMS_PrevEmp set Compname = " & COMPNAME & _
                              ",CompAdd = " & COMPADD & _
                              ",Tin = " & TIN_NO & _
                              ",ZipCode = " & ZIP_CODE & _
                              ",NONTax13thMonth = " & VtxtNONTax13thMonth & _
                              ",NONTaxSalaries = " & VtxtNONTaxSalaries & _
                              ",NONTaxPremium = " & VtxtNONTaxPremium & _
                              ",Tax13thMonth = " & VtxtTax13thMonth & _
                              ",TaxSalaries = " & VtxtTaxSalaries & _
                              ",TaxWithheld = " & VtxtTaxWithheld & _
                              ",TotalTaxable = " & VtxtTotalTaxable & _
                              ",FromDate= '" & dtpFrom.Value & _
                              "',Todate = '" & dtpTo.Value & _
                              "',PREVEMPYEAR = '" & lblYear.Caption & _
                              "' Where Empno = '" & lblEmpNo.Caption & _
                              "' And ID = '" & lblID.Caption & "'")
    Else
        gconDMIS.Execute ("Insert Into HRMS_PrevEmp (EmpLevel,Empno,Compname,CompAdd,Tin,ZipCode,NONTax13thMonth,NONTaxSalaries,NONTaxPremium,Tax13thMonth,TaxSalaries,TaxWithheld, Totaltaxable ,FromDate,Todate, PREVEMPYEAR)Values('" & GetEmpLevel(lblEmpNo.Caption) & _
                          "','" & lblEmpNo.Caption & _
                          "'," & COMPNAME & _
                          "," & COMPADD & _
                          "," & TIN_NO & _
                          "," & ZIP_CODE & _
                          "," & VtxtNONTax13thMonth & _
                          "," & VtxtNONTaxSalaries & _
                          "," & VtxtNONTaxPremium & _
                          "," & VtxtTax13thMonth & _
                          "," & VtxtTaxSalaries & _
                          "," & VtxtTaxWithheld & _
                          "," & VtxtTotalTaxable & _
                          ",'" & dtpFrom.Value & _
                          "','" & dtpTo.Value & _
                          "','" & lblYear.Caption & "')")
    End If
    LogAudit "E", "UPDATE PREVIOUS EMPLOYER RECORD", lblEmpNo.Caption
    ShowSuccessFullyUpdated
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'DrawXPCtl Me
    rsrefresh
    lblYear.Caption = GetSavedYear()
    txtSearch.Text = "aa": txtSearch.Text = ""
End Sub

Private Sub lsvEMP_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lsvEmp_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim INDEX                                                         As Double
    
    If Not lsvEmp.ListItems.count = 0 Then
        With lsvEmp
            INDEX = .SelectedItem.INDEX
            lblEmpNo.Caption = .ListItems(INDEX).SubItems(1)
            lblEmpName.Caption = .ListItems(INDEX).Text
            Dim rsTemp As ADODB.Recordset
            Set rsTemp = New ADODB.Recordset
            Set rsTemp = gconDMIS.Execute("SELECT PREVIOUSCOMPANY FROM HRMS_EMPINFO WHERE EMPNO ='" & lblEmpNo.Caption & "'")
            If Not rsTemp.EOF And Not rsTemp.BOF Then
                If Null2String(rsTemp!PreviousCompany) <> "" Then
                    txtCompName = Null2String(rsTemp!PreviousCompany)
                Else
                    txtCompName = ""
                End If
            End If
            Set rsTemp = Nothing
            rsRefresh_REC
        End With
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then FillGrid Else FillSearchGrid txtSearch.Text
End Sub

Sub InitMemvars()
    txtCompAdd.Text = ""
    txtTin.Text = ""
    txtZip.Text = ""
    txtNONTax13thMonth = 0
    txtNONTaxSalaries = 0
    txtNONTaxPremium = 0
    txtTax13thMonth = 0
    txtTaxSalaries = 0
    txtTaxWithheld = 0
    txtTotalTaxable = 0
End Sub

Function GetEmpLevel(EMPNO As String) As String
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT EMPLEVEL FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    GetEmpLevel = "E"
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetEmpLevel = Null2String(rsTemp!EMPLEVEL)
    End If
    Set rsTemp = Nothing
End Function

Function GetSavedYear() As Integer
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT PERIODYEAR FROM HRMS_PAYROLLSETUP")
    
    GetSavedYear = YEAR(Date)
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetSavedYear = N2Str2Zero(rsTemp!PERIODYEAR)
        If N2Str2Zero(rsTemp!PERIODYEAR) = 0 Then
            GetSavedYear = YEAR(Date)
        End If
    End If
End Function
