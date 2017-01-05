VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSEditAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Customer Appointment"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
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
   Icon            =   "FrmEditAppointMent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   7110
      MouseIcon       =   "FrmEditAppointMent.frx":09AA
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditAppointMent.frx":0AFC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancel"
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   6390
      MouseIcon       =   "FrmEditAppointMent.frx":0E3A
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditAppointMent.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save Entry"
      Top             =   6600
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
      Left            =   8670
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   6420
      Width           =   2055
   End
   Begin VB.CommandButton cmdEditVehicle 
      Caption         =   "Edit Customer Vehicle"
      Height          =   795
      Left            =   5100
      TabIndex        =   56
      Top             =   6600
      Width           =   1305
   End
   Begin VB.CommandButton cmdSelectCustomer 
      Caption         =   "F2 - Select Customer"
      Height          =   795
      Left            =   3810
      TabIndex        =   57
      Top             =   6600
      Width           =   1305
   End
   Begin VB.PictureBox picAppointment 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   60
      ScaleHeight     =   6525
      ScaleWidth      =   7725
      TabIndex        =   17
      Top             =   30
      Width           =   7755
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
         Left            =   8700
         TabIndex        =   23
         Top             =   8370
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
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.ComboBox dtPromised 
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
         Left            =   5340
         TabIndex        =   13
         Text            =   "dtPromised"
         Top             =   3300
         Width           =   2265
      End
      Begin VB.TextBox txtModel 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtDescription 
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
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6090
         Width           =   6315
      End
      Begin VB.TextBox txtnote 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "FrmEditAppointMent.frx":12DC
         Top             =   3870
         Width           =   7545
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
         Top             =   1500
         Width           =   7575
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
         Left            =   2550
         TabIndex        =   10
         Top             =   2730
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
         Left            =   90
         MaxLength       =   9
         TabIndex        =   9
         Top             =   2730
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
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   9420
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
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   8700
         Width           =   2055
      End
      Begin VB.TextBox txtPlate_No 
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
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   5
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox txtRep_Or 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         Top             =   300
         Width           =   2325
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
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
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "cboRecd_by"
         Top             =   3300
         Width           =   2925
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
         Left            =   5940
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   8070
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
         Left            =   5940
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   8850
         Width           =   2235
      End
      Begin VB.TextBox txtMake 
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
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5520
         Width           =   2055
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
         Left            =   690
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   9390
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
         Left            =   8490
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   8820
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
         Left            =   9840
         TabIndex        =   20
         Top             =   7890
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
         Left            =   5880
         MaxLength       =   35
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   8610
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
         ItemData        =   "FrmEditAppointMent.frx":12E2
         Left            =   1470
         List            =   "FrmEditAppointMent.frx":12EC
         TabIndex        =   18
         Top             =   8550
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtDte_recd 
         Height          =   345
         Left            =   3030
         TabIndex        =   12
         Top             =   3300
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM. dd, yyyy"
         Format          =   53542915
         CurrentDate     =   38936
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
         Height          =   345
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   6195
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
         Height          =   345
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   4
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   52
         Top             =   5910
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
         Top             =   2070
         Width           =   4725
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   51
         Top             =   1890
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
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
         TabIndex        =   50
         Top             =   3660
         Width           =   375
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
         Left            =   5520
         TabIndex        =   49
         Top             =   9240
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
         Left            =   3810
         TabIndex        =   48
         Top             =   8430
         Width           =   1815
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Appointment"
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
         Left            =   3060
         TabIndex        =   47
         Top             =   3120
         Width           =   1695
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   46
         Top             =   3120
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2580
         TabIndex        =   45
         Top             =   2550
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   44
         Top             =   2520
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
         Left            =   1170
         TabIndex        =   43
         Top             =   8070
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
         Left            =   750
         TabIndex        =   42
         Top             =   8790
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
         Left            =   720
         TabIndex        =   41
         Top             =   8730
         Width           =   1035
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   40
         Top             =   5310
         Width           =   705
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   39
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment No."
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
         Left            =   75
         TabIndex        =   38
         Top             =   90
         Width           =   1380
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   37
         Top             =   690
         Width           =   885
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
         Left            =   6330
         TabIndex        =   36
         Top             =   9630
         Width           =   1425
      End
      Begin VB.Label Label25 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4350
         TabIndex        =   35
         Top             =   5280
         Width           =   450
      End
      Begin VB.Label Label24 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2220
         TabIndex        =   34
         Top             =   5280
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
         Left            =   810
         TabIndex        =   33
         Top             =   8460
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
         Left            =   5310
         TabIndex        =   32
         Top             =   8640
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
         Left            =   9000
         TabIndex        =   31
         Top             =   8490
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
         Left            =   690
         TabIndex        =   30
         Top             =   8610
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time of Appointment"
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
         Left            =   5400
         TabIndex        =   29
         Top             =   3120
         Width           =   1755
      End
   End
   Begin VB.Label lblOLDAPPTNO 
      BackColor       =   &H000000FF&
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
      Left            =   150
      TabIndex        =   55
      Top             =   7770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labID 
      BackColor       =   &H00E0E0E0&
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
      Left            =   180
      TabIndex        =   54
      Top             =   6270
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCSMSEditAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTL                                                 As Control
Dim rsS_Model                                           As ADODB.Recordset
Dim rsEmpNo                                             As ADODB.Recordset
Dim OldApptTime                                         As String
Dim OldApptNo                                           As String
Dim OldTranDate                                         As String
Dim WithEvents frm                                      As frmCSMSROCusveh
Attribute frm.VB_VarHelpID = -1
Dim WithEvents FRMx                                     As frmCSMS_MasterSearchCustomer
Attribute FRMx.VB_VarHelpID = -1

Function CheckIfScheduleAlreadyScheduleToOther() As Boolean
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_APPOINTMENT WHERE TRANDATE  = '" & txtDte_recd.Value & "' AND APPTTIME = '" & dtPromised.Text & "'")
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        If labID.Caption = RSTMP!ID Then
            CheckIfScheduleAlreadyScheduleToOther = False
        Else
            If Null2String(RSTMP!CUSCDE) = "" Then
                CheckIfScheduleAlreadyScheduleToOther = False
            Else
                CheckIfScheduleAlreadyScheduleToOther = True
                MsgBox "The Date and Time You want to re-schedule to this customer is already schedule to " & vbCrLf & RSTMP!CUSNAM & vbCrLf & "Plate no. " & RSTMP!PLATE_NO & "", vbInformation, "CSMS"
            End If
        End If
    Else
        CheckIfScheduleAlreadyScheduleToOther = False
    End If

    Set RSTMP = Nothing
End Function

Function FindTheNewApptNo()
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT APPTNO FROM CSMS_APPOINTMENT ORDER BY APPTNO DESC")
        'WHERE TRANDATE = '" & txtDte_recd.Value &         "' AND APPTTIME = '" & dtPromised.Text & "'")
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        FindTheNewApptNo = Format(NumericVal(RSTMP!APPTNO) + 1, "000000000")
    Else
        FindTheNewApptNo = Format(1, "000000000")
    End If
    
    Set RSTMP = Nothing
End Function

Function SetMake(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select model,make from CSMS_S_Model where ltrim(rtrim(model)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        SetMake = Null2String(rsS_Model!Make)
    Else
        SetMake = ""
    End If
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
        txtModel.Text = Null2String(rsCusVeh!MODEL)
        txtMake.Text = Null2String(rsCusVeh!Make)
        txtDescription.Text = Null2String(rsCusVeh!Description)
    End If
End Sub

Sub StoreAppInfo(XXX As String)
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSUPLOAD = gconDMIS.Execute("Select * from CSMS_Repor where REP_OR = '" & XXX & "'")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Set RSTMP = gconDMIS.Execute("SELECT ID FROM CSMS_APPOINTMENT WHERE APPTNO = '" & XXX & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            labID = Null2String(RSTMP!ID)
        End If
        Set RSTMP = Nothing
        txtnote = Null2String(RSUPLOAD!NOTE)
        txtRep_Or.Text = Null2String(RSUPLOAD!REP_OR)
        
        'COMMENT BY  : MJP 06152009 0944AM
        'DESCRIPTION : INSTEAD OF .CAPTION TO BE PUT, THE PROGRAMMER TYPE INCORRECT (.TABINDEX )
            'lblOLDAPPTNO.TabIndex = Null2String(RSUPLOAD!rep_OR)
        'COMMENT BY  : MJP 06152009 0944AM
        
        'UPDATE BY   : MJP 06152009 0944AM
        'DESCRIPTION : INSTEAD OF .CAPTION TO BE PUT, THE PROGRAMMER TYPE INCORRECT (.TABINDEX )
            lblOLDAPPTNO.Caption = Null2String(RSUPLOAD!REP_OR)
        'UPDATE BY   : MJP 06152009 0944AM
        
        OldApptNo = Null2String(RSUPLOAD!REP_OR)
        txtInvoiceNo.Text = Null2String(RSUPLOAD!INVOICE)
        txtDte_Rel.Text = Null2String(RSUPLOAD!dte_rel)
        txtEstimateno.Text = Null2String(RSUPLOAD!EstimateNo)
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
        txtVIN.Text = Null2String(RSUPLOAD!VIN)
        txtParticipat.Text = Null2String(RSUPLOAD!participat)
    End If

    Dim rsAppointment                                  As New ADODB.Recordset
    Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptNo = '" & XXX & "'")
    If Not rsAppointment.EOF And Not rsAppointment.BOF Then
        txtDte_recd.Value = Null2String(rsAppointment!TRANDATE)
        OldTranDate = Null2String(rsAppointment!TRANDATE)
        OldApptTime = Null2String(rsAppointment!APPTTIME)
        dtPromised = Null2String(rsAppointment!APPTTIME)
    End If

    Set RSTMP = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & txtAcct_No.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblCN.Caption = Null2String(RSTMP!HomePhone) & "\" & Null2String(RSTMP!TelephoneNo) & "\" & Null2String(RSTMP!Mobile)
    End If
    Set RSTMP = Nothing
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
End Sub

Sub StoreApptTme()
    Dim rsCSMS_ApptSchedule                            As New ADODB.Recordset
    Set rsCSMS_ApptSchedule = gconDMIS.Execute("Select * from CSMS_ApptSchedule Order by ID asc")
    If Not rsCSMS_ApptSchedule.EOF And Not rsCSMS_ApptSchedule.BOF Then
        rsCSMS_ApptSchedule.MoveFirst: dtPromised.Clear
        Do While Not rsCSMS_ApptSchedule.EOF
            dtPromised.AddItem Null2String(rsCSMS_ApptSchedule!TimeInterval)
            rsCSMS_ApptSchedule.MoveNext
        Loop
    End If
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtAcct_No.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtNiym.Text = Null2String(rsCustomer!ACCTNAME)
    End If
End Sub

Private Sub frm_SaveChanges(xPLATE_NO As String, xWARR_CER As String, xMake As String, xMODEL As String, xSERIAL As String, xDESCRIPTION As Variant, FromFrom As String)
    If FromFrom = "APPOINTMENT" Then
        txtPlate_No.Text = xPLATE_NO
        txtModel.Text = xMODEL
        txtMake.Text = xMake
        txtDescription.Text = xDESCRIPTION
        
        Unload frm
    End If
End Sub

Private Sub FRMx_SelectionMade(ByVal xCode As String, xName As String, FromForm As String)
    If FromForm = "EDIT APPOINTMENT" Then
        txtAcct_No.Text = xCode
        txtNiym.Text = xName
        
        Unload FRMx
    End If
End Sub

Private Sub txtmodel_Change()
    txtMake.Text = SetMake(txtModel.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEditVehicle_Click()
    'COMMENT BY  : MJP07222009 0107PM
    'DESCRIPTION : TO UNIFY THE EDITING OF VEHICLE INFO
        'EDIT_RO = txtPlate_No.Text
        'frmCSMSEditAppVehicle.labCustCode = txtAcct_No.Text
        'frmCSMSEditAppVehicle.labCustomer = txtNiym.Text
        'frmCSMSEditAppVehicle.Show 1
    'COMMENT BY  : MJP07222009 0107PM
    
    'UPDATE BY   : MJP07222009 0107PM
    'DESCRIPTION : TO UNIFY THE EDITING OF VEHICLE INFO
        Call frm.SelectSQl("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(txtPlate_No) & "", "EDIT APPOINTMENT", GetPlateId(txtPlate_No), txtAcct_No, txtNiym, txtPlate_No)
        frm.Show 1
    'UPDATE BY   : MJP07222009 0107PM
End Sub

Function GetPlateId(XPLATENO As String) As Long
    Dim RSTMP                                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT ID FROM CSMS_CUSVEH WHERE PLATE_NO = " & N2Str2Null(XPLATENO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetPlateId = RSTMP!ID
    End If
    Set RSTMP = Nothing
End Function

Private Sub cmdSave_Click()
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
    Dim vNEWAPPTNO                                      As String
    Dim temp                                            As String
    
    If CheckIfScheduleAlreadyScheduleToOther = True Then
        Exit Sub
    End If
    
    temp = txtRep_Or
    vNEWAPPTNO = FindTheNewApptNo
    If vNEWAPPTNO = "" Then
        vNEWAPPTNO = temp
    End If
    VTXTREP_OR = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)
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
    VTXTMake = N2Str2Null(txtMake.Text)
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
    XNOTE = N2Str2Null(txtnote)

    SQL_STATEMENT = "update CSMS_RepOr set REP_OR = '" & vNEWAPPTNO & _
        "', estimateno = " & VTXTestimateno & ",APPTNO = '" & vNEWAPPTNO & _
        "', NOTE = " & XNOTE & "," & _
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
        " where REP_OR = '" & lblOLDAPPTNO.Caption & "'"
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblOLDAPPTNO), "APPTNO", "CSMS_REPOR"), "", "APP NO: " & lblOLDAPPTNO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
        " RO_NO = '" & vNEWAPPTNO & _
        "',APPTNO = '" & vNEWAPPTNO & _
        "', acct_no = " & VTXTAcct_No & "," & _
        " plate_no = " & VTXTPlate_No & "," & _
        " model = " & VTXTModel & "," & _
        " AppointmentDate = " & VTXTDte_recd & "," & _
        " Writer = " & N2Str2Null(cboRecd_by.Text) & "," & _
        " DateFinish = " & VTXTDte_comp & "," & _
        " PromiseDate = '" & dtPromised & "'" & _
        " where RO_NO = '" & lblOLDAPPTNO.Caption & "'"

    gconDMIS.Execute ("UPDATE CSMS_RO_DET SET " & _
        " REP_OR = '" & vNEWAPPTNO & _
        "', APPTNO = '" & vNEWAPPTNO & _
        "' WHERE REP_OR = '" & lblOLDAPPTNO.Caption & "'")

    '**************************************************************************************
    'UPDATE BY   : MJP 08042008 11:44 PM
    'DESCRIPTION : TO UPDATE ALSO THE PMS JOB DETAILS
        gconDMIS.Execute ("UPDATE CSMS_PMS_JOB_DET SET " & _
            " APPTNO = '" & vNEWAPPTNO & _
            "' WHERE APPTNO = '" & lblOLDAPPTNO.Caption & "'")
    'UPDATE BY   : MJP 08042008 11:44 PM
    '**************************************************************************************

'    If lblOLDAPPTNO.Caption = vNEWAPPTNO Then
'        gconDMIS.Execute "update CSMS_Appointment set " & _
'            " CusCde = " & VTXTAcct_No & _
'            ", Plate_no = " & VTXTPlate_No & _
'            ", CUSNAM = " & VTXTNiym & _
'            ", MODEL = " & VTXTModel & _
'            ", NOTE = " & N2Str2Null(txtnote) & _
'            ", MAKE = " & VTXTMake & _
'            ", km_rdg = " & VTXTKm_rdg & _
'            " where APPTNO = '" & vNEWAPPTNO & "'"
'    Else
'        gconDMIS.Execute ("UPDATE CSMS_APPOINTMENT SET " & _
'            " CUSCDE = NULL " & _
'            ", PLATE_NO = NULL " & _
'            ", CUSNAM = NULL " & _
'            ", MODEL = NULL " & _
'            ", NOTE = NULL " & _
'            ", MAKE = NULL " & _
'            ", KM_RDG = NULL " & _
'            " WHERE APPTNO = '" & lblOLDAPPTNO.Caption & "'")
        
'        gconDMIS.Execute "update CSMS_Appointment set " & _
'            " CusCde = " & VTXTAcct_No & _
'            ", Plate_no = " & VTXTPlate_No & _
'            ", CUSNAM = " & VTXTNiym & _
'            ", MODEL = " & VTXTModel & _
'            ", NOTE = " & N2Str2Null(txtnote) & _
'            ", MAKE = " & VTXTMake & _
'            ", km_rdg = " & VTXTKm_rdg & _
'            " where ApptTime = '" & dtPromised & _
'            "' AND TranDate = " & VTXTDte_recd
'    End If
    gconDMIS.Execute ("DELETE FROM CSMS_APPOINTMENT WHERE APPTNO = " & N2Str2Null(lblOLDAPPTNO) & "")
    
    gconDMIS.Execute ("INSERT INTO CSMS_Appointment (APPTNO, CusCde, Plate_no, CUSNAM, MODEL, NOTE, MAKE, km_rdg, ApptTime, TranDate) " & _
        " VALUES(" & N2Str2Null(vNEWAPPTNO) & _
        ", " & VTXTAcct_No & _
        ", " & VTXTPlate_No & _
        ", " & VTXTNiym & _
        ", " & VTXTModel & _
        ", " & N2Str2Null(txtnote) & _
        ", " & VTXTMake & _
        ", " & VTXTKm_rdg & _
        ", '" & dtPromised & _
        "', " & VTXTDte_recd & ")")
            
    MessagePop InfoFriend, "Appointmet Information Updated", "Appointment Information Sucessfully Updated!", 1000
    cmdCancel.Value = True

    Exit Sub

Errorcode:
    Call ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelectCustomer_Click()
    'frmCSMS_EditAppCustomerSearch.Show 1
    
    Call FRMx.PassVariable("EDIT APPOINTMENT")
    FRMx.Show 1
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Dim RSTMP                                          As New ADODB.Recordset
    
    Set frm = New frmCSMSROCusveh
    Set FRMx = New frmCSMS_MasterSearchCustomer
        
    With frmCSMSEditAppointment
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
    End With
    Call InitCbo
End Sub
