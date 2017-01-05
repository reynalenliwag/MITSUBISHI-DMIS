VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISAC_CheckPrevBal 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Check Previous Balance [Requested, Receipts, Issuance]"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FF8080&
   Icon            =   "AC_CheckPrevBal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   9990
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
      Left            =   9120
      MouseIcon       =   "AC_CheckPrevBal.frx":01CA
      MousePointer    =   99  'Custom
      Picture         =   "AC_CheckPrevBal.frx":031C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit Window"
      Top             =   6420
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8400
      MouseIcon       =   "AC_CheckPrevBal.frx":0682
      MousePointer    =   99  'Custom
      Picture         =   "AC_CheckPrevBal.frx":07D4
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Print Report"
      Top             =   6420
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "[ Inventory Balances ]"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   4605
      Left            =   90
      TabIndex        =   19
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtTotalPO 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtTP 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   1770
         Width           =   1455
      End
      Begin VB.TextBox txtLMOH 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtOH 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtLMMAC 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtMAC 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1350
         Width           =   1455
      End
      Begin VB.TextBox txtTR 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtTI 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2490
         Width           =   1455
      End
      Begin VB.TextBox txtTotalISS 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtTotalRR 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3270
         Width           =   1455
      End
      Begin VB.TextBox txtLastY_OH 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3660
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3990
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Total Purchase Orders"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   62
         Top             =   2910
         Width           =   3105
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD Purchase Orders"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   60
         Top             =   1770
         Width           =   3105
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Month On-Hand"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   28
         Top             =   330
         Width           =   3105
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "This Month On-Hand"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   660
         Width           =   3105
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Month Moving Ave. Cost"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   26
         Top             =   1020
         Width           =   3105
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "This Month Moving Ave. Cost"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   25
         Top             =   1350
         Width           =   3105
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD Accessories Receipts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   24
         Top             =   2160
         Width           =   3105
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD Accessories Issuance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   23
         Top             =   2490
         Width           =   3105
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Total Accessories Issuance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   22
         Top             =   3600
         Width           =   3195
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Total Accessories Receipts"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   21
         Top             =   3270
         Width           =   3195
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Year On-Hand"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   300
         TabIndex        =   20
         Top             =   3990
         Width           =   3105
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Dealer <--> Distributor"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   5670
      TabIndex        =   10
      Top             =   4170
      Width           =   4215
      Begin VB.TextBox txtDUnserved 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2280
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDFillRate 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2280
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1290
         Width           =   1455
      End
      Begin VB.TextBox txtDONRequest 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox txtDServed 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   315
         Left            =   3840
         TabIndex        =   38
         Top             =   1290
         Width           =   1755
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD FillRate"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   510
         TabIndex        =   18
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD UnServed"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   510
         TabIndex        =   17
         Top             =   990
         Width           =   1755
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD Ordered"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   510
         TabIndex        =   16
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MTD Served"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   510
         TabIndex        =   15
         Top             =   630
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Customer <--> Dealer"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   4065
      Left            =   5670
      TabIndex        =   9
      Top             =   60
      Width           =   4215
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         Caption         =   "Over the Counter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   90
         TabIndex        =   49
         Top             =   2160
         Width           =   4035
         Begin VB.TextBox txtServed 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   630
            Width           =   1455
         End
         Begin VB.TextBox txtONRequest 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   52
            Text            =   "Text1"
            Top             =   270
            Width           =   1455
         End
         Begin VB.TextBox txtFillRate 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   1350
            Width           =   1455
         End
         Begin VB.TextBox txtUnserved 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   990
            Width           =   1455
         End
         Begin VB.Label Label24 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   315
            Left            =   3750
            TabIndex        =   58
            Top             =   1410
            Width           =   1755
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD Served"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   57
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label22 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD Requested"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   56
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label21 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD UnServed"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   55
            Top             =   990
            Width           =   1755
         End
         Begin VB.Label Label20 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD FillRate"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   54
            Top             =   1350
            Width           =   1755
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Caption         =   "Workshop Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   90
         TabIndex        =   39
         Top             =   300
         Width           =   4035
         Begin VB.TextBox txtS_Unserved 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   990
            Width           =   1455
         End
         Begin VB.TextBox txtS_FillRate 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   1350
            Width           =   1455
         End
         Begin VB.TextBox txtS_ONRequest 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtS_Served 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2190
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   630
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD FillRate"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   48
            Top             =   1350
            Width           =   1755
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD UnServed"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   47
            Top             =   990
            Width           =   1755
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD Requested"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   46
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MTD Served"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   420
            TabIndex        =   45
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label19 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   315
            Left            =   3750
            TabIndex        =   44
            Top             =   1410
            Width           =   1755
         End
      End
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
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
      Left            =   7680
      MouseIcon       =   "AC_CheckPrevBal.frx":0C73
      MousePointer    =   99  'Custom
      Picture         =   "AC_CheckPrevBal.frx":0DC5
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Process Checking of Accessories Previous Balance"
      Top             =   6420
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   60
      ScaleHeight     =   1155
      ScaleWidth      =   9825
      TabIndex        =   0
      Top             =   5730
      Width           =   9825
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   7395
         TabIndex        =   1
         Top             =   750
         Width           =   7395
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   8235
         TabIndex        =   3
         Top             =   660
         Width           =   8235
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
            Top             =   0
            Width           =   7485
            _ExtentX        =   13203
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "AC_CheckPrevBal.frx":1060
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   556
         Picture         =   "AC_CheckPrevBal.frx":107C
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "AC_CheckPrevBal.frx":1098
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   30
         Width           =   9705
      End
   End
   Begin Crystal.CrystalReport rptPrintStkStat 
      Left            =   0
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Stock Status Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frmPMISAC_CheckPrevBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CheckBalance()
    Dim vLYOH, vLMOH, vOH, vLMMAC                                     As Double
    Dim vMAC, vTP, vTR, vTI, vTotalPO, vTotalRR, vTotalISS            As Double
    Dim vREQUESTED, vSERVED, vUNSERVED, vFILLRATE                     As Double
    Dim vS_REQUESTED, vS_SERVED, vS_UNSERVED, vS_FILLRATE             As Double
    Dim vORDERED, vDSERVED, vDUNSERVED, vDFILLRATE                    As Double

    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    Set rsPartMas = gconDMIS.Execute("Select SUM(ORDERED) TOTAL_ORDERED, SUM(SERVED) AS TOTAL_SERVED, SUM(UNSERVED) AS TOTAL_UNSERVED, SUM(FILLRATE) AS TOTAL_FILLRATE, SUM(S_ONREQUEST) TOTAL_S_REQUESTED, SUM(S_REQSERVED) AS TOTAL_S_REQSERVED, SUM(S_REQUNSERVED) AS TOTAL_S_REQUNSERVED, SUM(S_REQFILLRATE) AS TOTAL_S_REQFILLRATE, SUM(ONREQUEST) TOTAL_REQUESTED, SUM(REQSERVED) AS TOTAL_REQSERVED, SUM(REQUNSERVED) AS TOTAL_REQUNSERVED, SUM(REQFILLRATE) AS TOTAL_REQFILLRATE, SUM(lasty_oh) AS TOTAL_LASTY_OH, SUM(lastm_oh) AS TOTAL_LASTM_OH, ROUND(SUM(ROUND(lastm_mac,2) * ROUND(lastm_oh,2)),2) TOTAL_LASTM_MAC_ONHAND, SUM(onhand) AS TOTAL_ONHAND, ROUND(SUM(ROUND(mac,2) * ROUND(onhand,2)),2) TOTAL_MAC_ONHAND, SUM(tpoqty) AS TOTAL_TPOQTY,SUM(trecqty) AS TOTAL_TRECQTY,SUM(tissqty) AS TOTAL_TISSQTY,SUM(purchases) TOTAL_PURCHASES,SUM(receipts) TOTAL_RECEIPTS,SUM(issuances) AS TOTAL_ISSUANCES from PMIS_Accessories")
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        vLMOH = vLMOH + N2Str2IntZero(rsPartMas!TOTAL_LASTM_OH)
        vLYOH = vLYOH + N2Str2IntZero(rsPartMas!TOTAL_LASTY_OH)
        vOH = vOH + N2Str2IntZero(rsPartMas!total_onhand)
        vLMMAC = N2Str2Zero(rsPartMas!TOTAL_LASTM_MAC_ONHAND)
        vMAC = N2Str2Zero(rsPartMas!TOTAL_MAC_ONHAND)
        vTP = vTR + N2Str2IntZero(rsPartMas!TOTAL_tpoqty)
        vTR = vTR + N2Str2IntZero(rsPartMas!TOTAL_trecqty)
        vTI = vTI + N2Str2IntZero(rsPartMas!TOTAL_tissqty)
        vS_REQUESTED = vS_REQUESTED + N2Str2IntZero(rsPartMas!TOTAL_S_REQUESTED)
        vS_SERVED = vS_SERVED + N2Str2IntZero(rsPartMas!TOTAL_S_REQSERVED)
        vS_UNSERVED = vS_UNSERVED + N2Str2IntZero(rsPartMas!TOTAL_S_REQUNSERVED)
        vS_FILLRATE = vS_FILLRATE + N2Str2IntZero(rsPartMas!TOTAL_S_REQFILLRATE)

        vREQUESTED = vREQUESTED + N2Str2IntZero(rsPartMas!TOTAL_REQUESTED)
        vSERVED = vSERVED + N2Str2IntZero(rsPartMas!TOTAL_REQSERVED)
        vUNSERVED = vUNSERVED + N2Str2IntZero(rsPartMas!TOTAL_REQUNSERVED)
        vFILLRATE = vFILLRATE + N2Str2IntZero(rsPartMas!TOTAL_REQFILLRATE)

        vORDERED = vORDERED + N2Str2IntZero(rsPartMas!TOTAL_ORDERED)
        vDSERVED = vDSERVED + N2Str2IntZero(rsPartMas!TOTAL_SERVED)
        vDUNSERVED = vDUNSERVED + N2Str2IntZero(rsPartMas!TOTAL_UNSERVED)
        vDFILLRATE = vDFILLRATE + N2Str2IntZero(rsPartMas!TOTAL_FILLRATE)
        vTotalPO = vTotalPO + N2Str2IntZero(rsPartMas!TOTAL_purchases)
        vTotalRR = vTotalRR + N2Str2IntZero(rsPartMas!TOTAL_receipts)
        vTotalISS = vTotalISS + N2Str2IntZero(rsPartMas!TOTAL_issuances)
        DoEvents
        txtLMOH.Text = Format(vLMOH, DIGIT_FORMAT)
        txtOH.Text = Format(vOH, DIGIT_FORMAT)
        txtLMMAC.Text = Format(vLMMAC, MAXIMUM_DIGIT)
        txtMAC.Text = Format(vMAC, MAXIMUM_DIGIT)
        txtTP.Text = Format(vTP, DIGIT_FORMAT)
        txtTR.Text = Format(vTR, DIGIT_FORMAT)
        txtTI.Text = Format(vTI, DIGIT_FORMAT)

        txtS_ONRequest.Text = Format(vS_REQUESTED, DIGIT_FORMAT)
        txtS_Served.Text = Format(vS_SERVED, DIGIT_FORMAT)
        txtS_Unserved.Text = Format(vS_REQUESTED - vS_SERVED, DIGIT_FORMAT)

        If vS_REQUESTED > 0 Then
            txtS_FillRate.Text = Format((vS_SERVED / vS_REQUESTED) * 100, "##0")
        Else
            txtS_FillRate.Text = Format(0, "##0")
        End If

        txtONRequest.Text = Format(vREQUESTED, DIGIT_FORMAT)
        txtServed.Text = Format(vSERVED, DIGIT_FORMAT)
        txtUnserved.Text = Format(vREQUESTED - vSERVED, DIGIT_FORMAT)

        If vREQUESTED > 0 Then
            txtFillRate.Text = Format((vSERVED / vREQUESTED) * 100, "##0")
        Else
            txtFillRate.Text = Format(0, "##0")
        End If

        txtDONRequest.Text = Format(vORDERED, DIGIT_FORMAT)
        txtDServed.Text = Format(vDSERVED, DIGIT_FORMAT)
        txtDUnserved.Text = Format(vORDERED - vDSERVED, DIGIT_FORMAT)
        If vORDERED > 0 Then
            txtDFillRate.Text = Format((vDSERVED / vORDERED) * 100, "##0")
        Else
            txtDFillRate.Text = Format(0, "##0")
        End If

        txtTotalPO.Text = Format(vTotalPO, DIGIT_FORMAT)
        txtTotalRR.Text = Format(vTotalRR, DIGIT_FORMAT)
        txtTotalISS.Text = Format(vTotalISS, DIGIT_FORMAT)
        txtLastY_OH.Text = Format(vLYOH, DIGIT_FORMAT)
        progCPB.Value = 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        labProcessing.Caption = ""
        DoEvents
    Else
        MsgSpeechBox "Error Opening Part Master File"
    End If
    Set rsPartMas = Nothing
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "ACCESSORIES CHECK PREVIOUS BALANCE") = False Then Exit Sub
    CheckBalance
    
    Call NEW_LogAudit("I", "ACCESSORIES CHECK PREVIOUS BALANCE", "", "", "Accessories", "", "", "")
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptPrintStkStat.ReportTitle = "STOCK STATUS REPORT (ACCESSORIES)"
    PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "stock2.rpt", "{STKSTAT.TYPE} = 'A'", DMIS_REPORT_Connection, 1
    
    Call NEW_LogAudit("V", "ACCESSORIES CHECK PREVIOUS BALANCE", "", "", "", "", "", "")
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCESSORIES CHECK PREVIOUS BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACCESSORIES CHECK PREVIOUS BALANCE", "PRINTING")
    
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtLMOH.Text = 0: txtOH.Text = 0
    txtLMMAC.Text = 0: txtMAC.Text = 0
    txtTP.Text = 0: txtTR.Text = 0: txtTI.Text = 0
    txtS_ONRequest.Text = 0: txtS_Served.Text = 0
    txtS_Unserved.Text = 0: txtS_FillRate.Text = 0
    txtONRequest.Text = 0: txtServed.Text = 0
    txtUnserved.Text = 0: txtFillRate.Text = 0
    txtDONRequest.Text = 0: txtDServed.Text = 0
    txtDUnserved.Text = 0: txtDFillRate.Text = 0
    txtTotalPO.Text = 0: txtTotalRR.Text = 0: txtTotalISS.Text = 0
    txtLastY_OH.Text = 0
    CheckBalance
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

