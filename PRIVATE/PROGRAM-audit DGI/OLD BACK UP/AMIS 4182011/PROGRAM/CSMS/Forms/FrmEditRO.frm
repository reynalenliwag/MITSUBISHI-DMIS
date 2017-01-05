VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSEditRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Repair Order"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmEditRO.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAppointment 
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
      Height          =   7305
      Left            =   30
      ScaleHeight     =   7305
      ScaleWidth      =   8025
      TabIndex        =   19
      Top             =   30
      Width           =   8025
      Begin VB.ComboBox Cbo_Rotype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "FrmEditRO.frx":09AA
         Left            =   4710
         List            =   "FrmEditRO.frx":09AC
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change Service Adviser"
         Height          =   285
         Left            =   6030
         TabIndex        =   63
         Top             =   2340
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtRECC 
         Height          =   1065
         Left            =   30
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "FrmEditRO.frx":09AE
         Top             =   4800
         Width           =   7905
      End
      Begin VB.TextBox txtInst 
         Height          =   1335
         Left            =   4020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "FrmEditRO.frx":09B4
         Top             =   3240
         Width           =   3915
      End
      Begin VB.TextBox txtModel 
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
         Height          =   345
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   6240
         Width           =   2745
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6840
         Width           =   7875
      End
      Begin VB.TextBox txtnote 
         Height          =   1335
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "FrmEditRO.frx":09BA
         Top             =   3240
         Width           =   3915
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   810
         Width           =   7875
      End
      Begin VB.TextBox txtSektion 
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
         Height          =   345
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   10
         Top             =   2640
         Width           =   2235
      End
      Begin VB.TextBox txtKm_rdg 
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
         Height          =   345
         Left            =   60
         MaxLength       =   9
         TabIndex        =   9
         Top             =   2640
         Width           =   2235
      End
      Begin VB.TextBox txtSvc_No 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9090
         MaxLength       =   1
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   7110
         Width           =   2055
      End
      Begin VB.TextBox txtROType 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9090
         MaxLength       =   1
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   6750
         Width           =   2055
      End
      Begin VB.TextBox txtAcct_No 
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
         Height          =   360
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtPlate_No 
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
         Height          =   345
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   5
         Top             =   6240
         Width           =   2055
      End
      Begin VB.TextBox txtNiym 
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
         Height          =   360
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4845
      End
      Begin VB.TextBox txtRep_Or 
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
         Height          =   360
         Left            =   90
         MaxLength       =   15
         TabIndex        =   0
         Top             =   240
         Width           =   1875
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4650
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "cboRecd_by"
         Top             =   2010
         Width           =   3345
      End
      Begin VB.TextBox txtDte_comp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   13230
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   6960
         Width           =   2235
      End
      Begin VB.TextBox txtDte_Rel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   12930
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   7350
         Width           =   2235
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   14010
         TabIndex        =   25
         Top             =   6780
         Width           =   1455
         Begin VB.TextBox txtInvoiceNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   150
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.TextBox txtMake 
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
         Height          =   345
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   6240
         Width           =   2955
      End
      Begin VB.TextBox txtCertific8 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   7080
         Width           =   2835
      End
      Begin VB.TextBox txtParticipat 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9720
         MaxLength       =   6
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   5520
         Width           =   1125
      End
      Begin VB.CheckBox chkParticipat 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
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
         Left            =   8250
         TabIndex        =   22
         Top             =   5460
         Width           =   225
      End
      Begin VB.TextBox txtVIN 
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
         ForeColor       =   &H00A00000&
         Height          =   360
         Left            =   12840
         MaxLength       =   35
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   7500
         Width           =   2295
      End
      Begin VB.ComboBox txtTerm 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmEditRO.frx":09C0
         Left            =   9090
         List            =   "FrmEditRO.frx":09CA
         TabIndex        =   20
         Top             =   6600
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtPromised 
         Height          =   345
         Left            =   2340
         TabIndex        =   13
         Top             =   2010
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
         Format          =   20316163
         CurrentDate     =   38936
      End
      Begin MSComCtl2.DTPicker txtDte_recd 
         Height          =   345
         Left            =   60
         TabIndex        =   12
         Top             =   2010
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM. dd, yyyy"
         Format          =   20316163
         CurrentDate     =   38936
      End
      Begin VB.Label lbl_rodescription 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   5850
         TabIndex        =   66
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label lbl_rotype 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RO Type"
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
         Height          =   210
         Left            =   4740
         TabIndex        =   65
         Top             =   2430
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suggestions/Recommendation"
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
         Height          =   210
         Index           =   2
         Left            =   45
         TabIndex        =   58
         Top             =   4590
         Width           =   2550
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service advisor Instruction"
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
         Height          =   210
         Index           =   1
         Left            =   4035
         TabIndex        =   57
         Top             =   3030
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   210
         Left            =   60
         TabIndex        =   55
         Top             =   6630
         Width           =   945
      End
      Begin VB.Label lblCN 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Telephone Number/Mobile Number/Home Phone"
         Top             =   1380
         Width           =   5745
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
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
         Height          =   210
         Left            =   60
         TabIndex        =   54
         Top             =   1170
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Request"
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
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   53
         Top             =   3030
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11910
         TabIndex        =   52
         Top             =   6450
         Width           =   1395
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11730
         TabIndex        =   51
         Top             =   7050
         Width           =   1815
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Recorded"
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
         Height          =   210
         Left            =   60
         TabIndex        =   50
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Advisor"
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
         Height          =   210
         Left            =   4680
         TabIndex        =   49
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Section No."
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
         Height          =   210
         Left            =   2340
         TabIndex        =   48
         Top             =   2430
         Width           =   915
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "KM Reading"
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
         Height          =   210
         Left            =   90
         TabIndex        =   47
         Top             =   2430
         Width           =   960
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Term"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8160
         TabIndex        =   46
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Service"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8370
         TabIndex        =   45
         Top             =   6840
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ROType"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   8340
         TabIndex        =   44
         Top             =   6780
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Code"
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
         Left            =   10080
         TabIndex        =   43
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
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
         Height          =   210
         Left            =   60
         TabIndex        =   42
         Top             =   6030
         Width           =   705
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   210
         Left            =   90
         TabIndex        =   41
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RO Number"
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
         Height          =   210
         Left            =   120
         TabIndex        =   40
         Top             =   30
         Width           =   930
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer "
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
         Height          =   210
         Left            =   2040
         TabIndex        =   39
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Label51 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE NO."
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
         Height          =   315
         Left            =   12720
         TabIndex        =   38
         Top             =   6840
         Width           =   1425
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Height          =   210
         Left            =   4980
         TabIndex        =   37
         Top             =   6000
         Width           =   450
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Height          =   210
         Left            =   2220
         TabIndex        =   36
         Top             =   6000
         Width           =   510
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "    Warranty Certificate Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8430
         TabIndex        =   35
         Top             =   6510
         Width           =   2865
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VIN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12270
         TabIndex        =   34
         Top             =   7530
         Width           =   585
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Participation"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8520
         TabIndex        =   33
         Top             =   5550
         Width           =   1065
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F12"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   8310
         TabIndex        =   32
         Top             =   6660
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Promised Date/Time"
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
         Height          =   210
         Left            =   2370
         TabIndex        =   31
         Top             =   1800
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   7320
      MouseIcon       =   "FrmEditRO.frx":09D8
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditRO.frx":0B2A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancel"
      Top             =   7350
      Width           =   735
   End
   Begin VB.TextBox txtEstimateno 
      BackColor       =   &H00FF8080&
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
      Height          =   345
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   990
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   6600
      MouseIcon       =   "FrmEditRO.frx":0E68
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditRO.frx":0FBA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Save Entry"
      Top             =   7350
      Width           =   735
   End
   Begin VB.CommandButton cmdEditVehicle 
      Caption         =   "Edit Customer Vehicle"
      Height          =   795
      Left            =   4830
      Picture         =   "FrmEditRO.frx":130A
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   7350
      Width           =   1785
   End
   Begin VB.CommandButton cmdSelectCustomer 
      Caption         =   "F2 - Select Customer"
      Height          =   795
      Left            =   3060
      Picture         =   "FrmEditRO.frx":238C
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7350
      Width           =   1785
   End
   Begin VB.Label labID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
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
      Left            =   150
      TabIndex        =   62
      Top             =   7740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblOLDRO 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   60
      Top             =   7350
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmCSMSEditRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTL                                                 As Control
Dim rsS_Model                                           As ADODB.Recordset
Dim rsEmpNo                                             As ADODB.Recordset
Dim OLD_RO_Number                                       As String
Dim AUDIT_SQL                                           As String
Dim xBEFORE_SAVE                                        As String
Dim xLOCAL_RO                                           As String
Dim WithEvents frm                                      As frmCSMSROCusveh
Attribute frm.VB_VarHelpID = -1
Dim WithEvents FRMx                                     As frmCSMS_MasterSearchCustomer
Attribute FRMx.VB_VarHelpID = -1
Public Event SaveEditRO()

Public Sub PassRepairOrderNo(XRONO As String, xID As Long)
    xLOCAL_RO = XRONO
    txtRep_Or.Text = XRONO
    OLD_RO_Number = XRONO
    labid.Caption = xID
End Sub

Function SetMake(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select model,make from CSMS_S_Model where ltrim(rtrim(model)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetMake = Null2String(rsS_Model!Make) Else SetMake = ""
    Set rsS_Model = Nothing
End Function

Function SetSA(emp As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where code = '" & emp & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSA = Null2String(rsEmpNo!NAYM)
    Set rsEmpNo = Nothing
End Function

Function SetCodeSA(nam As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!Code)
    Set rsEmpNo = Nothing
End Function

Sub SetVehicleInfo(XXX As String)
    Dim rsCusVeh                                       As New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CUSVEH where PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        txtPlate_No.Text = Null2String(rsCusVeh!PLATE_NO)
        txtModel.Text = Null2String(rsCusVeh!Model)
        TXTMAKE.Text = Null2String(rsCusVeh!Make)
        txtDescription.Text = Null2String(rsCusVeh!Description)
    End If
End Sub

Sub InitCbo()
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo order by naym asc")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        cboRecd_by.Clear
        cboRecd_by.Text = Null2String(rsEmpNo!NAYM)
        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!NAYM)
            rsEmpNo.MoveNext
        Loop
    End If
    Call fillCboType
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtAcct_No.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtNiym.Text = Null2String(rsCustomer!AcctName)
    End If
End Sub

'Updated: IEVB
'Description: calls the function of Displaying of Ro Type Description
Private Sub Cbo_Rotype_Change()
    ShowROTYPEdescription
End Sub
'Updated: IEVB
'Description: calls the function of Displaying of Ro Type Description

'Updated: IEVB
'Description: calls the function of Displaying of Ro Type Description
Private Sub Cbo_Rotype_Click()
    ShowROTYPEdescription
End Sub
'Updated: IEVB
'Description: calls the function of Displaying of Ro Type Description

Private Sub Command1_Click()
    If Module_Access(LOGID, "CHANGE SERVICE ADVISER", "SYSTEM") = False Then Exit Sub
    cboRecd_by.Enabled = True
End Sub

Private Sub frm_SaveChanges(xPLATE_NO As String, xWARR_CER As String, xMake As String, xMODEL As String, xSERIAL As String, xDESCRIPTION As Variant, FromFrom As String)
    If FromFrom = "EDIT RO" Then
        txtPlate_No.Text = xPLATE_NO
        TXTMAKE.Text = xMake
        txtModel.Text = xMODEL
        txtDescription.Text = xDESCRIPTION
        
        Unload frm
    End If
End Sub

Private Sub FRMx_SelectionMade(ByVal Xcode As String, xName As String, FromForm As String)
    If FromForm = "EDIT RO" Then
        txtAcct_No.Text = Xcode
        txtNiym.Text = xName
        
        Unload FRMx
    End If
End Sub

Private Sub txtKm_rdg_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtmodel_Change()
    TXTMAKE.Text = SetMake(txtModel.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEditVehicle_Click()
    If txtRep_Or.Text = "R-00000000" Then Exit Sub
    
    'EDIT_RO = txtPlate_No.Text
    'frmCSMSEditROVehicle.labCustCode = txtAcct_No.Text
    'frmCSMSEditROVehicle.labCustomer = txtNiym.Text
    'frmCSMSEditROVehicle.Show 1
    
    Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "EDIT RO", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
    frm.Show 1
End Sub

Function GetPlateId(XPLATENO As String) As Long
    Dim rstmp                                       As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(XPLATENO) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetPlateId = rstmp!ID
    End If
    Set rstmp = Nothing
End Function

Private Sub cmdSave_Click()
    If txtRep_Or.Text = "R-00000000" Then Exit Sub
    If Function_Access(LOGID, "Acess_Edit", "REPAIR ORDER") = False Then Exit Sub

    Dim VTXTREP_OR                                      As String
    Dim VTXTestimateno                                  As String
    Dim VTXTROType                                      As String
    Dim VTXTSvc_No                                      As String
    Dim VTXTAcct_No                                     As String
    Dim VTXTNiym                                        As String
    Dim VTXTPlate_No                                    As String
    Dim VTXTModel                                       As String
    Dim VTXTMake                                        As String
    Dim VTXTTerm                                        As String
    Dim VTXTSektion                                     As String
    Dim VTXTKm_rdg                                      As String
    Dim VTXTDte_recd                                    As String
    Dim VTXTCertific8                                   As String
    Dim VTXTDte_comp                                    As String
    Dim VTXTDte_Rel                                     As String
    Dim VtxtAddress                                     As String
    Dim VtxtVIN                                         As String
    Dim VTXTParticipat                                  As String
    Dim VcboRecd_by                                     As String
    Dim XNOTE                                           As String
    Dim Vusercode                                       As String
    Dim VLastUpdate                                     As String
    Dim VLastUpdateTime                                 As String
    Dim XINST                                           As String
    Dim XRECC                                           As String

    If Left(txtRep_Or.Text, 2) = "R-" Then
        txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text) - 2)), "00000000")
    Else
        If VALID_COMPANY_CODE_FORHAI = True Then
        Else
            txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text))), "00000000")
        End If
    End If

    VTXTREP_OR = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateNo.Text)
    VTXTROType = N2Str2Null(txtROType.Text)
    VTXTSvc_No = N2Str2Null(txtSvc_No.Text)
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    Dim kAdd                                           As Integer
    For kAdd = 1 To Len(txtAddress.Text)
        If Mid(txtAddress.Text, kAdd, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" Then Exit For
        VtxtAddress = VtxtAddress & Mid(txtAddress.Text, kAdd, 1)
    Next
    VtxtAddress = N2Str2Null(VtxtAddress)
    VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
    VTXTModel = N2Str2Null(txtModel.Text)
    VTXTMake = N2Str2Null(TXTMAKE.Text)
    VTXTTerm = N2Str2Null(txtTerm.Text)
    VTXTSektion = N2Str2Null(txtSektion.Text)
    VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
    VTXTDte_recd = N2Date2Null(txtDte_recd)
    VTXTCertific8 = N2Str2Null(txtCertific8.Text)
    VTXTDte_comp = N2Date2Null(txtDte_comp.Text)
    VTXTDte_Rel = N2Date2Null(txtDte_Rel.Text)
    VtxtVIN = N2Str2Null(txtVIN.Text)
    VTXTParticipat = N2Str2Null(txtParticipat.Text)
    VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    XNOTE = N2Str2Null(txtNote)
    XINST = N2Str2Null(txtInst.Text)
    XRECC = N2Str2Null(txtRECC.Text)
    'Updated by: IEVB 06282010 1118AM
    'Description:
    If COMPANY_CODE = "HCI" Then
        VTXTROType = N2Str2Null(Cbo_Rotype.Text)
    Else
        VTXTROType = N2Str2Null("")
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT REP_OR, ID FROM CSMS_REPOR WHERE REP_OR = " & VTXTREP_OR & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If Not labid.Caption = rstmp!ID Then
            MsgBox "Repair Order no. Already Exist", vbExclamation, "Info."
            txtRep_Or.SetFocus
            Exit Sub
        End If
    End If

    'UPDATE BY   : MJP 10062008 0327 PM
    'DESCRIPTION : TO ENSURE WHEN THE USER EDIT THE REPAIR ORDER THAT THE PLATE NO HE WILL BE ENCODE IS EXISTING IN THE VEHICLE MASTER FILE
    Dim RSPLATE                                        As New ADODB.Recordset
    Set RSPLATE = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_CUSVEH WHERE PLATE_NO = '" & txtPlate_No & "'")
    If (RSPLATE.BOF And RSPLATE.EOF) Then
        MsgBox "Plate no. not existing in Vehicle Master File", vbExclamation, "CSMS"
        txtPlate_No.SetFocus
        Exit Sub
    End If
    Set RSPLATE = Nothing
    'UPDATE BY   : MJP 10062008 0327 PM

    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " REP_OR = " & VTXTREP_OR & ", estimateno = " & VTXTestimateno & "," & _
        " NOTE = " & XNOTE & ", INSTRUCTION = " & XINST & ", RECOMMENDATION = " & XRECC & "," & _
        " rotype = " & VTXTROType & "," & _
        " svc_no = " & VTXTSvc_No & "," & _
        " acct_no = " & VTXTAcct_No & "," & _
        " niym = " & VTXTNiym & "," & _
        " plate_no = " & VTXTPlate_No & "," & _
        " model = " & VTXTModel & "," & _
        " term = " & VTXTTerm & "," & _
        " sektion = " & VTXTSektion & "," & _
        " recd_by = " & VcboRecd_by & "," & _
        " km_rdg = " & VTXTKm_rdg & "," & _
        " dte_recd = " & VTXTDte_recd & "," & _
        " certific8 = " & VTXTCertific8 & "," & _
        " dte_comp = " & VTXTDte_comp & "," & _
        " dte_rel = " & VTXTDte_Rel & "," & _
        " dte_pro = " & N2Str2Null(dtPromised) & "," & _
        " VIN = " & VtxtVIN & "," & _
        " participat = " & VTXTParticipat & "," & _
        " status = 'N'" & "," & _
        " USERCDE = " & Vusercode & "," & _
        " SAVEDATE = " & VLastUpdate & "," & _
        " SAVETIME = " & VLastUpdateTime & _
        " where REP_OR = '" & OLD_RO_Number & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(OLD_RO_Number), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & OLD_RO_Number, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
        " RO_NO = " & VTXTREP_OR & "," & _
        " acct_no = " & VTXTAcct_No & "," & _
        " plate_no = " & VTXTPlate_No & "," & _
        " model = " & VTXTModel & "," & _
        " AppointmentDate = " & VTXTDte_recd & "," & _
        " DateFinish = " & VTXTDte_comp & "," & _
        " PromiseDate = '" & dtPromised & "'," & _
        " Writer = '" & cboRecd_by.Text & "'" & _
        " Where RO_NO = '" & OLD_RO_Number & "'"

    SQL_STATEMENT = "Update CSMS_RO_Det Set REP_OR = " & VTXTREP_OR & " where REP_OR = '" & OLD_RO_Number & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(OLD_RO_Number), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & OLD_RO_Number, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    gconDMIS.Execute "UPDATE CSMS_JOBCLOCK SET RO_NO = " & VTXTREP_OR & " WHERE RO_NO = '" & OLD_RO_Number & "'"
    SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET ASSIGNEDRO = " & VTXTREP_OR & " WHERE ASSIGNEDRO = '" & OLD_RO_Number & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(OLD_RO_Number), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & OLD_RO_Number, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    cmdCancel.Value = True
    RaiseEvent SaveEditRO
    

    'Unload Me

    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelectCustomer_Click()
    'If txtRep_Or.Text = "R-00000000" Then Exit Sub
    'frmCSMS_EditROCustomerSearch.Show 1
        
    Call FRMx.PassVariable("EDIT RO")
    FRMx.Show 1
End Sub

Private Sub Form_Load()
    Dim rstmp                                          As New ADODB.Recordset
    
    Set frm = New frmCSMSROCusveh
    Set FRMx = New frmCSMS_MasterSearchCustomer
    
    With frmCSMSEditRO
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
    End With

    Call InitCbo
    Dim RSUPLOAD                                       As New ADODB.Recordset
    'Set RSUPLOAD = gconDMIS.Execute("Select * from CSMS_Repor where " & _
    '    " REP_OR = '" & frmCSMSServiceCounter.grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text & "'")
    
    Set RSUPLOAD = gconDMIS.Execute("Select * from CSMS_Repor where REP_OR = '" & xLOCAL_RO & "'")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        txtRECC.Text = Null2String(RSUPLOAD!RECOMMENDATION)
        txtNote = Null2String(RSUPLOAD!NOTE)
        txtInst.Text = Null2String(RSUPLOAD!INSTRUCTION)
        OLD_RO_Number = Null2String(RSUPLOAD!REP_OR)
        txtRep_Or.Text = Null2String(RSUPLOAD!REP_OR)
        txtInvoiceNo.Text = Null2String(RSUPLOAD!invoice)
        txtDte_Rel.Text = Null2String(RSUPLOAD!dte_rel)
        txtEstimateNo.Text = Null2String(RSUPLOAD!EstimateNo)
        txtROType.Text = Null2String(RSUPLOAD!ROTYPE)
        txtSvc_No.Text = Null2String(RSUPLOAD!svc_no)
        txtAcct_No.Text = Null2String(RSUPLOAD!ACCT_NO)
        txtNiym.Text = Null2String(RSUPLOAD!NIYM)
        txtPlate_No.Text = Null2String(RSUPLOAD!PLATE_NO)
        SetVehicleInfo txtPlate_No.Text
        txtTerm.Text = Null2String(RSUPLOAD!TERM)
        txtSektion.Text = Null2String(RSUPLOAD!sektion)
        cboRecd_by.Text = SetSA(Null2String(RSUPLOAD!RECD_BY))
        txtKm_rdg.Text = Null2String(RSUPLOAD!km_rdg)
        txtDte_recd.Value = Null2String(RSUPLOAD!DTE_RECD)
        txtCertific8.Text = Null2String(RSUPLOAD!certific8)
        txtDte_comp.Text = Null2String(RSUPLOAD!dte_comp)
        txtVIN.Text = Null2String(RSUPLOAD!Vin)
        txtParticipat.Text = Null2String(RSUPLOAD!participat)
'Updated by: IEBV 06282010 1102AM
'Description:   To enable and display the rotype of the HCI
        If COMPANY_CODE = "HCI" Then
            Cbo_Rotype.Visible = True
            lbl_rotype.Visible = True
            lbl_rodescription(1).Visible = True
            If Null2String(RSUPLOAD!ROTYPE) = "" Then
                Cbo_Rotype.ListIndex = 0
            Else
                Cbo_Rotype.Text = Null2String(RSUPLOAD!ROTYPE)
            End If
        End If
'Updated by: IEBV 06282010 1102AM
'Description:   To enable and display the rotype of the HCI

    End If

    Set rstmp = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & txtAcct_No.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        lblCN.Caption = Null2String(rstmp!HomePhone) & "\" & Null2String(rstmp!TelephoneNo) & "\" & Null2String(rstmp!Mobile)
    End If
    Set rstmp = Nothing
    
    Set RSUPLOAD = New ADODB.Recordset
    'Set RSUPLOAD = gconDMIS.Execute("Select PromiseDate from [CSMS_RepairOrder] where RO_No = '" & frmCSMSServiceCounter.grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text & "'")
    Set RSUPLOAD = gconDMIS.Execute("Select PromiseDate from [CSMS_RepairOrder] where RO_No = '" & xLOCAL_RO & "'")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        dtPromised.Value = Null2String(RSUPLOAD!PromiseDate)
    End If
End Sub

Private Sub txtRep_Or_LostFocus()
If VALID_COMPANY_CODE_FORHAI = True Then
    'DO NOTHING
Else
    If Left(txtRep_Or.Text, 2) = "R-" Then
        txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text) - 2)), "00000000")
    Else
        txtRep_Or.Text = "R-" & Format(NumericVal(Right(txtRep_Or.Text, Len(txtRep_Or.Text))), "00000000")
    End If
End If
End Sub
Sub fillCboType()
    Cbo_Rotype.Clear
    Cbo_Rotype.AddItem ""
    Cbo_Rotype.AddItem "WTY"
    Cbo_Rotype.AddItem "BRP"
    Cbo_Rotype.AddItem "JET"
    Cbo_Rotype.AddItem "GJ"
    Cbo_Rotype.AddItem "O/H"
    Cbo_Rotype.AddItem "PP"
    Cbo_Rotype.AddItem "RF"
    Cbo_Rotype.AddItem "QS"
    Cbo_Rotype.AddItem "AC"
    Cbo_Rotype.AddItem "PDI"
    Cbo_Rotype.AddItem "QC"
    Cbo_Rotype.AddItem "FI"
    Cbo_Rotype.AddItem "DET"
End Sub
'Updated: IEVB
'Description: Displaying of Ro Type Description
Private Sub ShowROTYPEdescription()
    If (Cbo_Rotype) = "WTY" Then
       lbl_rodescription(1).Caption = "Warranty"
    ElseIf (Cbo_Rotype) = "BRP" Then
       lbl_rodescription(1).Caption = "Body Repair and Painting"
    ElseIf (Cbo_Rotype) = "JET" Then
       lbl_rodescription(1).Caption = "Jet Service"
    ElseIf (Cbo_Rotype) = "GJ" Then
       lbl_rodescription(1).Caption = "General Job"
    ElseIf (Cbo_Rotype) = "O/H" Then
       lbl_rodescription(1).Caption = "Overhauling"
    ElseIf (Cbo_Rotype) = "PP" Then
       lbl_rodescription(1).Caption = "Painting Protection"
    ElseIf (Cbo_Rotype) = "RF" Then
       lbl_rodescription(1).Caption = "Rustproofing"
    ElseIf (Cbo_Rotype) = "QS" Then
       lbl_rodescription(1).Caption = "Quick Service"
    ElseIf (Cbo_Rotype) = "AC" Then
       lbl_rodescription(1).Caption = "Aircon/Electrical"
    ElseIf (Cbo_Rotype) = "PDI" Then
       lbl_rodescription(1).Caption = "Pre-delivery Inspection"
    ElseIf (Cbo_Rotype) = "QC" Then
       lbl_rodescription(1).Caption = "Quality Control"
    ElseIf (Cbo_Rotype) = "FI" Then
       lbl_rodescription(1).Caption = "Final Inspection"
    ElseIf (Cbo_Rotype) = "DET" Then
       lbl_rodescription(1).Caption = "Detailing"
    Else
       lbl_rodescription(1).Caption = ""
    End If
End Sub
'Updated: IEVB
'Description: Displaying of Ro Type Description

