VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Files_jobRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Request Form"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMIS_Files_jobRequest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   8565
   Begin VB.PictureBox picJobDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   3360
      ScaleHeight     =   4185
      ScaleWidth      =   4035
      TabIndex        =   43
      Top             =   1290
      Visible         =   0   'False
      Width           =   4065
      Begin VB.CommandButton cmdClose1 
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
         Left            =   3660
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.TextBox txtSCName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   2130
         Width           =   3735
      End
      Begin VB.CheckBox chkIsFree 
         Caption         =   "FREE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   54
         Top             =   2580
         Width           =   795
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2730
         Width           =   2175
      End
      Begin VB.ComboBox cboJobsCode 
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
         Left            =   120
         TabIndex        =   49
         Text            =   "Combo1"
         Top             =   570
         Width           =   3345
      End
      Begin VB.CommandButton cmdAddjob 
         Height          =   345
         Left            =   3480
         Picture         =   "frmSMIS_Files_jobRequest.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   540
         Width           =   405
      End
      Begin VB.TextBox txtJobDescription 
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
         Left            =   120
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1110
         Width           =   3705
      End
      Begin VB.CommandButton cmdClose2 
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
         Height          =   765
         Left            =   3180
         MouseIcon       =   "frmSMIS_Files_jobRequest.frx":0A34
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Files_jobRequest.frx":0B86
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Exit Entry"
         Top             =   3270
         Width           =   675
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
         Height          =   765
         Left            =   2520
         MouseIcon       =   "frmSMIS_Files_jobRequest.frx":0EEC
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Files_jobRequest.frx":103E
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Save this Record"
         Top             =   3270
         Width           =   675
      End
      Begin VB.CommandButton cmdDelJobDet 
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
         Height          =   765
         Left            =   1860
         MouseIcon       =   "frmSMIS_Files_jobRequest.frx":138E
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Files_jobRequest.frx":14E0
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Delete Entry"
         Top             =   3270
         Width           =   675
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label labprevamount 
         Caption         =   "0.00"
         Height          =   345
         Left            =   300
         TabIndex        =   58
         Top             =   3600
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LABDETID 
         Caption         =   "0"
         Height          =   465
         Left            =   270
         TabIndex        =   57
         Top             =   3150
         Visible         =   0   'False
         Width           =   1515
      End
      Begin XtremeShortcutBar.ShortcutCaption capAccessories 
         Height          =   330
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   4695
         _Version        =   655364
         _ExtentX        =   8281
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Job Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
         Height          =   195
         Left            =   90
         TabIndex        =   45
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By:(Sc Name):"
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
         Left            =   120
         TabIndex        =   52
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label lblamount 
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   55
         Top             =   2805
         Width           =   825
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   50
         Top             =   900
         Width           =   1185
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   7275
      Left            =   -30
      ScaleHeight     =   7275
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   -30
      Width           =   9285
      Begin VB.Frame Frame2 
         Caption         =   "List Of Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   2820
         TabIndex        =   32
         Top             =   3750
         Width           =   5715
         Begin MSComctlLib.ListView listjob 
            Height          =   2385
            Left            =   30
            TabIndex        =   33
            Top             =   180
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4207
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
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VI"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ITEM"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPTION"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "AMOUNT"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3975
         Left            =   2790
         ScaleHeight     =   3975
         ScaleWidth      =   5865
         TabIndex        =   1
         Top             =   0
         Width           =   5865
         Begin VB.TextBox txtCustomer 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   270
            Width           =   5625
         End
         Begin VB.TextBox txtUnitModel 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   900
            Width           =   5625
         End
         Begin VB.Frame Frame4 
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
            ForeColor       =   &H00000000&
            Height          =   2535
            Left            =   2610
            TabIndex        =   17
            Top             =   1230
            Width           =   3135
            Begin VB.TextBox txtJvin 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   23
               Text            =   "Text1"
               Top             =   1485
               Width           =   2955
            End
            Begin VB.TextBox txtJmodel 
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
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "Text1"
               Top             =   915
               Width           =   2955
            End
            Begin VB.TextBox txtCS 
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
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   360
               Width           =   2955
            End
            Begin VB.TextBox txtJcolor 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   90
               Locked          =   -1  'True
               TabIndex        =   25
               Text            =   "Text1"
               Top             =   2070
               Width           =   2955
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Model"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   20
               Top             =   705
               Width           =   420
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CS #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   18
               Top             =   150
               Width           =   345
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vin"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   22
               Top             =   1275
               Width           =   240
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   24
               Top             =   1860
               Width           =   375
            End
         End
         Begin VB.Frame Frame1 
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
            ForeColor       =   &H00000000&
            Height          =   2535
            Left            =   60
            TabIndex        =   6
            Top             =   1230
            Width           =   2505
            Begin VB.TextBox txtVINO 
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
               Height          =   375
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "Text1"
               Top             =   360
               Width           =   1125
            End
            Begin VB.TextBox txtSONO 
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
               Height          =   375
               Left            =   1230
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "Text1"
               Top             =   360
               Width           =   1155
            End
            Begin VB.TextBox txtDateRel 
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
               Height          =   375
               Left            =   60
               TabIndex        =   12
               Text            =   "Text1"
               Top             =   930
               Width           =   2325
            End
            Begin VB.TextBox txtTimeRel 
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
               Height          =   375
               Left            =   60
               TabIndex        =   14
               Text            =   "Text1"
               Top             =   1500
               Width           =   2325
            End
            Begin VB.TextBox txtTransaction 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "Text1"
               Top             =   2070
               Width           =   2325
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice #"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   60
               TabIndex        =   7
               Top             =   150
               Width           =   645
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Order#"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   1230
               TabIndex        =   8
               Top             =   150
               Width           =   960
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date of Release"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   60
               TabIndex        =   11
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time of Release"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   60
               TabIndex        =   13
               Top             =   1290
               Width           =   1155
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transaction Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   60
               TabIndex        =   15
               Top             =   1875
               Width           =   1260
            End
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
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
            Height          =   225
            Left            =   90
            TabIndex        =   2
            Top             =   30
            Width           =   1380
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Detail"
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
            Height          =   225
            Left            =   90
            TabIndex        =   4
            Top             =   660
            Width           =   855
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   7620
         Left            =   0
         ScaleHeight     =   7620
         ScaleWidth      =   2865
         TabIndex        =   26
         Top             =   0
         Width           =   2865
         Begin VB.OptionButton Option2 
            Caption         =   "&Client"
            Height          =   225
            Left            =   1020
            TabIndex        =   29
            Top             =   390
            Width           =   945
         End
         Begin VB.OptionButton Option1 
            Caption         =   " &VI#"
            Height          =   225
            Left            =   210
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   60
            TabIndex        =   30
            Top             =   660
            Width           =   2685
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   6105
            Left            =   60
            TabIndex        =   31
            Top             =   1050
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   10769
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
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":180B
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VINO"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMERNAME"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label11 
            Caption         =   "Search By"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   90
            Width           =   1665
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   2940
         ScaleHeight     =   900
         ScaleWidth      =   5655
         TabIndex        =   34
         Top             =   6390
         Width           =   5655
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
            Left            =   4890
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":196D
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":1ABF
            Style           =   1  'Graphical
            TabIndex        =   42
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
            Left            =   4890
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":1E25
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":1F77
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Print this Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
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
            Left            =   4200
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":22DD
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":242F
            Style           =   1  'Graphical
            TabIndex        =   38
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
            Left            =   3510
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":275A
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":28AC
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   705
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
            Left            =   2820
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":2C08
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":2D5A
            Style           =   1  'Graphical
            TabIndex        =   39
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
            Left            =   2130
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":306D
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":31BF
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Find a Record"
            Top             =   30
            Width           =   705
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
            Left            =   1440
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":34B9
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":360B
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "P&rev"
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
            Left            =   750
            MouseIcon       =   "frmSMIS_Files_jobRequest.frx":3963
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Files_jobRequest.frx":3AB5
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   705
         End
      End
   End
   Begin VB.Label LABVINO 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8670
      TabIndex        =   62
      Top             =   1380
      Width           =   1695
   End
End
Attribute VB_Name = "frmSMIS_Files_jobRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ISFree                                                            As Boolean
Dim theFREE                                                           As String
Public VI_NO                                                          As String
Dim rsJOBS                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim rsSO                                                              As ADODB.Recordset
Dim SALESAGENT                                                        As String

Function GETJOBDESC(XXX As String)
    Dim RS                                                            As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DESCRIPTION FROM SMIS_JOBS WHERE JOBCODE='" & Repleys(XXX) & "'")
    If Not (RS.EOF Or RS.BOF) Then
        GETJOBDESC = Null2String(RS!Description)
        RS.MoveNext
    End If
    Set RS = Nothing
End Function

Sub DisplayJobs()
    Listview_Loadval Listjob.ListItems, gconDMIS.Execute("SELECT VINO,ITEM,DESCRIPTION,PRICE,ID FROM SMIS_JOBREQUEST WHERE  VINO='" & VI_NO & "'")
    If Listjob.ListItems.Count = 0 Then
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
    Else
        cmdDelete.Enabled = True
        cmdEdit.Enabled = True
    End If
End Sub

Sub FillGrid()
    Dim XXX                                                           As String
    ListView1.Sorted = False
    If Option1.Value = True Then
        XXX = Format(Text1, "000000")
        Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT VI_NO , CUSTNAME FROM SMIS_SALESORDER  WHERE VI_NO LIKE '" & Repleys(XXX) & "%' AND STATUS='P' order by vi_no desc")
    Else
        XXX = Text1
        Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT VI_NO , CUSTNAME FROM SMIS_SALESORDER  WHERE CUSTNAME LIKE '" & Repleys(XXX) & "%' AND STATUS='P' order by custname asc")
    End If
End Sub

Sub InitCbo()
    Dim RS                                                            As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT JOBCODE FROM SMIS_JOBS order by JOBCODE ASC")
    cboJobsCode.Clear
    Do While Not RS.EOF
        cboJobsCode.AddItem RS!jobcode
        RS.MoveNext
    Loop

    Set RS = Nothing
End Sub

Sub iNITMEMVARDETAIL()
    cmdDelJobDet.Enabled = False
    LABDETID = ""
    txtSCName.Text = ""
    cboJobsCode = ""
    txtJobDescription.Text = ""
    txtAmount.Text = ""

End Sub

Sub InitMemVars()
    Listjob.ListItems.Clear
    SALESAGENT = ""
    txtCS = ""
    txtCustomer = ""
    txtDateRel = ""
    txtTimeRel = ""
    txtVINO = ""
    txtSONO = ""
    txtTransaction = ""
    txtJcolor = ""
    txtJmodel = ""
    txtUnitModel = ""

End Sub

Sub rsRefresh()
    Set rsJOBS = New ADODB.Recordset
    Set rsSO = New ADODB.Recordset
    Call rsSO.Open("SELECT VI_NO , CUSTNAME FROM SMIS_SALESORDER  WHERE VI_NO LIKE '" & Text1 & "%' AND STATUS='P'", gconDMIS, adOpenKeyset, adLockReadOnly)

    Call rsJOBS.Open("SELECT DISTINCT VINO  FROM SMIS_JobRequest ORDER BY vino DESC", gconDMIS, adOpenKeyset, adLockReadOnly)

    If VI_NO = "" Then
        If Not (rsSO.EOF Or rsSO.BOF) Then
            rsSO.MoveFirst
            VI_NO = Null2String(rsSO!VI_NO)
        End If
    End If
End Sub

Sub StoreMemVars()
    InitMemVars

    Dim rsInvoice                                                     As ADODB.Recordset
    Set rsInvoice = gconDMIS.Execute("SELECT CUSTNAME, MODELDESCRIPTION , VI_NO ,SALESAE ,SO_NO ,DATERELEASED,TERM,IGNKEY_NO,COLOR,MODEL,VINO,ID FROM SMIS_SALESORDER WHERE VI_NO='" & VI_NO & "'")
    If Not (rsInvoice.EOF Or rsInvoice.BOF) Then
        txtCustomer = Null2String(rsInvoice!CustName)
        txtUnitModel = Null2String(rsInvoice!modeldescription)
        txtVINO = Null2String(rsInvoice!VI_NO)
        txtSONO = Null2String(rsInvoice!SO_NO)
        SALESAGENT = UCase(Null2String(rsInvoice!salesae))
        If IsDate(rsInvoice!DATERELEASED) = True Then
            txtDateRel = DateValue(rsInvoice!DATERELEASED)
            txtTimeRel = TimeValue(rsInvoice!DATERELEASED)
        End If
        txtTransaction = Null2String(rsInvoice!TERM)
        txtCS = Null2String(rsInvoice!ignkey_no)
        txtJcolor = Null2String(rsInvoice!Color)
        txtJmodel = Null2String(rsInvoice!Model)
        txtJvin = Null2String(rsInvoice!Vino)
        DisplayJobs
    End If

End Sub

Private Sub cboJobsCode_Change()
    txtJobDescription = GETJOBDESC(cboJobsCode)
End Sub

Private Sub cboJobsCode_Click()
    txtJobDescription = GETJOBDESC(cboJobsCode)
End Sub

Private Sub cboJobsCode_GotFocus()
    If AddorEdit = "ADD" Then
        VBComBoBoxDroppedDown cboJobsCode
    End If
End Sub

Private Sub cboJobsCode_Validate(Cancel As Boolean)
    cboJobsCode.ListIndex = SelectCombo(cboJobsCode, cboJobsCode)
End Sub

Private Sub chkIsFree_Click()
    If chkIsFree.Value = 1 Then
        txtAmount.Enabled = False
        txtAmount.Text = "0.00"
    Else
        txtAmount.Text = labprevamount
        txtAmount.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "SALES JOB REQUEST") = False Then Exit Sub
    AddorEdit = "ADD"
    iNITMEMVARDETAIL
    ShowHidePictureBox2 picJobDetails, True, picMain
    On Error Resume Next
    cboJobsCode.SetFocus
    If txtSCName = "" Then
        txtSCName = SALESAGENT
    End If
End Sub

Private Sub cmdAddjob_Click()
    frmSMIS_FILE_jobMasterFile.Show 1
    InitCbo
End Sub

Private Sub cmdClose1_Click()
    AddorEdit = ""
    ShowHidePictureBox2 picJobDetails, False, picMain
    On Error Resume Next
    If Listjob.ListItems.Count > 0 Then
        Listjob.ListItems(1).Selected = True
        Listjob.ListItems(1).EnsureVisible
        Listjob.SetFocus
    End If
End Sub

Private Sub cmdClose2_Click()
    cmdClose1.Value = True

End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SALES JOB REQUEST") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = ("DELETE FROM SMIS_JOBREQUEST WHERE vino='" & txtVINO & "'")
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "X", "SALES JOB REQUEST", SQL_STATEMENT, N2Str2Null(LABDETID), "", "SO NO:" & txtSONO, "", ""

        LogAudit "X", "JOB REQUEST ", "FOR VIN" & txtVINO
        rsRefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdDelJobDet_Click()
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = ("DELETE FROM SMIS_JOBREQUEST WHERE ID=" & LABDETID)
        gconDMIS.Execute (SQL_STATEMENT)
        '**********NEW LOG AUDIT**********
        NEW_LogAudit "X", "SALES JOB REQUEST", SQL_STATEMENT, N2Str2Null(LABDETID), "", "SO NO:" & txtSONO, "", ""
        '**********NEW LOG AUDIT**********

        DisplayJobs
        cmdClose1.Value = True
    End If
End Sub

Private Sub cmdEdit_Click()

    listjob_DblClick
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsSO.MoveNext
    If rsSO.EOF Then
        rsSO.MoveLast
        ShowLastRecordMsg
    End If
    If Not rsSO.EOF Then
        VI_NO = (rsSO!VI_NO)
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSO.MovePrevious
    If rsSO.BOF Then
        rsSO.MoveFirst
        ShowFirstRecordMsg
    End If
    If Not rsSO.BOF Then
        VI_NO = (rsSO!VI_NO)
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_print", "SALES JOB REQUEST") = False Then Exit Sub

End Sub

Private Sub cmdSave_Click()
    Dim SQL                                                           As String
    Dim vVI_NO, vSO_NO, vcboJobsCode, vtxtJobDescription, vtxtDateRel, vtxtTimeRel, vtxtTransaction, vtxtSCName, theFREE


    If txtJobDescription.Text = "" Then
        ShowIsRequiredMsg "Job Description"
        txtJobDescription.SetFocus
        Exit Sub
    End If


    If txtSCName.Text = "" Then
        ShowIsRequiredMsg "Requested By "
        txtSCName.SetFocus
        Exit Sub
    End If


    vVI_NO = N2Str2Null(VI_NO)
    vSO_NO = N2Str2Null(txtSONO)
    vcboJobsCode = N2Str2Null(cboJobsCode)
    vtxtJobDescription = N2Str2Null(txtJobDescription)

    If IsDate(txtDateRel) = True Then
        vtxtDateRel = N2Str2Null(txtDateRel)
    Else
        vtxtDateRel = "''"
    End If

    If IsDate(txtTimeRel) = True Then
        vtxtTimeRel = N2Str2Null(txtTimeRel)
    Else
        vtxtTimeRel = "''"
    End If


    vtxtTransaction = N2Str2Null(txtTransaction)
    vtxtSCName = N2Str2Null(txtSCName)


    If chkIsFree.Value = 1 Then
        theFREE = "'FREE'"
    Else
        theFREE = NumericVal(txtAmount)
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO SMIS_jobrequest (vino,so_no,SCNAME,item,description,Price,timeofrelease,dateofrelease,Mtransaction) VALUES (" & vVI_NO & _
                        "," & vSO_NO & "," & vtxtSCName & "," & vcboJobsCode & "," & vtxtJobDescription & "," & theFREE & "," & vtxtTimeRel _
                      & "," & vtxtDateRel & "," & vtxtTransaction & ") "
        gconDMIS.Execute (SQL_STATEMENT)

        '**********NEW LOG AUDIT**********
        NEW_LogAudit "A", "SALES JOB REQUEST", SQL_STATEMENT, FindTransactionID(N2Str2Null(VI_NO), "VINO", "SMIS_JOBREQUEST"), "", "SO NO:" & txtSONO, "", ""
        '**********NEW LOG AUDIT**********

        iNITMEMVARDETAIL
        DisplayJobs
        AddorEdit = "ADD"
        MessagePop RecSaveOk, "RECORD UPDATED", "Job Request Added"
        LogAudit "A", "JOB REQUEST ", "FOR VIN" & txtVINO
    Else
        SQL_STATEMENT = "UPDATE SMIS_jobrequest SET vino=" & vVI_NO & "," _
                      & " so_no=" & vSO_NO & "," _
                      & " SCNAME=" & vtxtSCName & "," _
                      & " item=" & vcboJobsCode & "," _
                      & " price=" & theFREE & "," _
                      & " description=" & vtxtJobDescription & "," _
                      & " dateofrelease=" & vtxtDateRel & "," _
                      & " timeofrelease=" & vtxtTimeRel & "," _
                      & " Mtransaction=" & vtxtTransaction & "" _
                      & " where id=" & LABDETID

        gconDMIS.Execute (SQL_STATEMENT)

        '**********NEW LOG AUDIT**********
        NEW_LogAudit "E", "SALES JOB REQUEST", SQL_STATEMENT, N2Str2Null(LABDETID), "", "SO NO:" & txtSONO, "", ""
        '**********NEW LOG AUDIT**********
        LogAudit "E", "JOB REQUEST ", "FOR VIN" & txtVINO

        MessagePop RecSaveOk, "RECORD UPDATED", "Job Request Updated"
        DisplayJobs
        cmdClose1.Value = True
    End If

End Sub

Private Sub Command1_Click()
    Dim vVI_NO, vSO_NO, vtxtTimeRel, vtxtDateRel, vtxtTransaction, SQL
    vVI_NO = N2Str2Null(VI_NO)
    vSO_NO = N2Str2Null(VI_NO)
    vtxtTimeRel = N2Str2Null(txtTimeRel)
    vtxtDateRel = N2Str2Null(txtDateRel)
    vtxtTransaction = N2Str2Null(txtTransaction)


    SQL = "INSERT INTO SMIS_jobrequest (vino,so_no, timeofrelease,dateofrelease,Mtransaction) VALUES (" & vVI_NO & _
          "," & vSO_NO & "," & vtxtTimeRel & "," & vtxtDateRel & "," & vtxtTransaction & ") "
    gconDMIS.Execute (SQL)
    rsRefresh
    DisplayJobs
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And picJobDetails.Visible = True Then
        cmdClose1.Value = True
    Else
        MoveKeyPress (KeyCode)
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES JOB REQUEST)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(LABDETID), "SALES JOB REQUEST")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitMemVars
    iNITMEMVARDETAIL
    InitCbo
    rsRefresh
    StoreMemVars
    FillGrid
    If rsSO.EOF Or rsSO.BOF Then
        MsgBox "There Are No Invoices to Add Job Request. Job Request Now Will Unload", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VI_NO = ""
End Sub

Private Sub listjob_DblClick()
    If Function_Access(LOGID, "Acess_EDIT", "SALES JOB REQUEST") = False Then Exit Sub
    If Listjob.SelectedItem Is Nothing Then: Exit Sub
    Dim rsdet                                                         As ADODB.Recordset
    Set rsdet = gconDMIS.Execute("Select * from SMIS_JobRequest where id=" & Listjob.SelectedItem.ListSubItems(4).Text)
    If Not rsdet.EOF Or Not rsdet.BOF Then
        LABDETID = rsdet!ID
        cboJobsCode = Null2String(rsdet!Item)
        txtJobDescription = Null2String(rsdet!Description)
        If Null2String(rsdet!Price) = "FREE" Then
            chkIsFree.Value = 1
        Else
            chkIsFree.Value = 0
        End If
        txtSCName = Null2String(rsdet!SCNAME)
        txtAmount = FormatNumber(N2Str2Zero(rsdet!Price))
        labprevamount = txtAmount
        AddorEdit = "EDIT"
        cmdDelJobDet.Enabled = True
        If txtSCName = "" Then
            txtSCName = SALESAGENT
        End If
        ShowHidePictureBox2 picJobDetails, True, picMain
        On Error Resume Next
        cboJobsCode.SetFocus
    End If
End Sub

Private Sub listjob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then listjob_DblClick
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
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

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSO.MoveFirst
    rsSO.Find ("vi_no='" & Item.Text & "'")
    VI_NO = Item.Text
    StoreMemVars
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    Text1.SetFocus
    Text1 = ""
End Sub

Private Sub Option2_Click()
    On Error Resume Next
    Text1.SetFocus
    Text1 = ""
End Sub

Private Sub Text1_Change()
    FillGrid
End Sub

Private Sub txtAmount_GotFocus()
    If txtAmount = "0.00" Then: txtAmount = ""
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtAmount_LostFocus()
    If IsNumeric(txtAmount) = True Then
        txtAmount = FormatNumber(txtAmount)
    Else
        txtAmount = "0.00"
    End If
End Sub

Private Sub txtJobDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSCName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

