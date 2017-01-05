VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSEditAppointment 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Customer Appointment"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   ForeColor       =   &H00FFC0C0&
   Icon            =   "FrmEditAppointMent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAppointment 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   60
      ScaleHeight     =   6405
      ScaleWidth      =   7905
      TabIndex        =   17
      Top             =   -480
      Width           =   7905
      Begin VB.ComboBox dtPromised 
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
         Left            =   5550
         TabIndex        =   56
         Text            =   "dtPromised"
         Top             =   3900
         Width           =   2175
      End
      Begin VB.TextBox txtModel 
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
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3450
         Width           =   2055
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4350
         Width           =   6375
      End
      Begin VB.CommandButton cmdEditVehicle 
         BackColor       =   &H0080C0FF&
         Caption         =   "Edit Customer Vehicle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2610
         Width           =   2055
      End
      Begin VB.CommandButton cmdSelectCustomer 
         BackColor       =   &H0080C0FF&
         Caption         =   "F2 - Select Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5850
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
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
         Height          =   1515
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "FrmEditAppointMent.frx":09AA
         Top             =   4800
         Width           =   6765
      End
      Begin VB.TextBox txtAddress 
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
         Height          =   435
         Left            =   1410
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1560
         Width           =   6375
      End
      Begin VB.TextBox txtSektion 
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
         Left            =   5550
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2640
         Width           =   2235
      End
      Begin VB.TextBox txtKm_rdg 
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
         Left            =   5550
         MaxLength       =   9
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2220
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
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   28
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
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   6750
         Width           =   2055
      End
      Begin VB.TextBox txtAcct_No 
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
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2220
         Width           =   2055
      End
      Begin VB.TextBox txtPlate_No 
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
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3030
         Width           =   2055
      End
      Begin VB.TextBox txtNiym 
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
         Height          =   405
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4365
      End
      Begin VB.TextBox txtRep_Or 
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   630
         Width           =   2325
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   5550
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "cboRecd_by"
         Top             =   3060
         Width           =   2235
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
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   26
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
         Left            =   5340
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   25
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   6120
         TabIndex        =   23
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
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.TextBox txtMake 
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3900
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
         Left            =   420
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
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
         Left            =   6660
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   4920
         Width           =   1125
      End
      Begin VB.CheckBox chkParticipat 
         BackColor       =   &H00D8E9EC&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6420
         TabIndex        =   20
         Top             =   4920
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
         Left            =   5280
         MaxLength       =   35
         TabIndex        =   19
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
         ItemData        =   "FrmEditAppointMent.frx":09B0
         Left            =   1200
         List            =   "FrmEditAppointMent.frx":09BA
         TabIndex        =   18
         Top             =   6600
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtDte_recd 
         Height          =   345
         Left            =   5550
         TabIndex        =   13
         Top             =   3450
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMM. dd, yyyy"
         Format          =   51904515
         CurrentDate     =   38936
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   270
         TabIndex        =   54
         Top             =   4410
         Width           =   1755
      End
      Begin VB.Label lblCN 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4830
         TabIndex        =   1
         ToolTipText     =   "Telephone Number/Mobile Number/Home Phone"
         Top             =   630
         Width           =   2955
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
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
         Height          =   255
         Left            =   3810
         TabIndex        =   52
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE :"
         Height          =   345
         Left            =   360
         TabIndex        =   51
         Top             =   4830
         Width           =   795
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
         Left            =   4020
         TabIndex        =   50
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
         Left            =   3840
         TabIndex        =   49
         Top             =   7050
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Appointment"
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
         Left            =   3690
         TabIndex        =   48
         Top             =   3540
         Width           =   1965
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Advisor"
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
         Left            =   4290
         TabIndex        =   47
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Section No."
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
         Height          =   405
         Left            =   4500
         TabIndex        =   46
         Top             =   2730
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "KM Reading"
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
         Height          =   405
         Left            =   4470
         TabIndex        =   45
         Top             =   2280
         Width           =   1185
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
         Left            =   270
         TabIndex        =   44
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
         Left            =   480
         TabIndex        =   43
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
         Left            =   450
         TabIndex        =   42
         Top             =   6780
         Width           =   1035
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Code"
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
         Height          =   405
         Left            =   300
         TabIndex        =   41
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
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
         Left            =   480
         TabIndex        =   40
         Top             =   3090
         Width           =   1035
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   570
         TabIndex        =   39
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment No."
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
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer "
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
         Height          =   435
         Left            =   420
         TabIndex        =   37
         Top             =   1170
         Width           =   1635
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
         Left            =   4830
         TabIndex        =   36
         Top             =   6840
         Width           =   1425
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Left            =   780
         TabIndex        =   35
         Top             =   3960
         Width           =   1035
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   720
         TabIndex        =   34
         Top             =   3510
         Width           =   1035
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
         Left            =   540
         TabIndex        =   33
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
         Left            =   4710
         TabIndex        =   32
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
         Left            =   5280
         TabIndex        =   31
         Top             =   4950
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
         Left            =   420
         TabIndex        =   30
         Top             =   6660
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time of Appointment"
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
         Left            =   3690
         TabIndex        =   29
         Top             =   3990
         Width           =   1845
      End
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
      Left            =   7110
      MouseIcon       =   "FrmEditAppointMent.frx":09C8
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditAppointMent.frx":0B1A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancel"
      Top             =   6000
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
      Left            =   6390
      MouseIcon       =   "FrmEditAppointMent.frx":0E58
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditAppointMent.frx":0FAA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save Entry"
      Top             =   6000
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
      Left            =   3390
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   4830
      Width           =   2055
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
      TabIndex        =   57
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
Dim ctl                                 As Control
Dim rsS_Model                           As ADODB.Recordset
Dim rsEmpNo                             As ADODB.Recordset
Dim OldApptTime                         As String
Dim OldApptNo                           As String
Dim OldTranDate                         As String

Private Sub txtmodel_Change()
    txtMake.Text = SetMake(txtModel.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEditVehicle_Click()
    EDIT_RO = txtPlate_No.Text
    frmCSMSEditAppVehicle.labCustCode = txtAcct_No.Text
    frmCSMSEditAppVehicle.labCustomer = txtNiym.Text
    frmCSMSEditAppVehicle.Show 1
End Sub

Sub SetVehicleInfo(XXX As String)
    Dim rsCusVeh                        As ADODB.Recordset
    Set rsCusVeh = New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CUSVEH where PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        txtPlate_No.Text = Null2String(rsCusVeh!Plate_no)
        txtModel.Text = Null2String(rsCusVeh!Model)
        txtMake.Text = Null2String(rsCusVeh!Make)
        txtDescription.Text = Null2String(rsCusVeh!Description)
    End If
End Sub

Private Sub cmdSave_Click()
    Dim VTXTrep_or, VTXTestimateno, VTXTROType As String
    Dim VTXTSvc_No, VTXTAcct_No, VTXTNiym As String
    Dim VTXTPlate_No, VTXTModel, VTXTMake As String
    Dim VTXTTerm, VTXTSektion, VTXTKm_rdg As String
    Dim VTXTDte_recd, VTXTCertific8, VTXTDte_comp As String
    Dim VTXTDte_Rel, VtxtAddress        As String
    Dim VtxtVIN                         As String
    Dim VTXTParticipat, VcboRecd_by     As String
    Dim XNOTE, Vusercode, VLastUpdate, VLastUpdateTime As String

    VTXTrep_or = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)
    VTXTROType = N2Str2Null(txtROType.Text)
    VTXTSvc_No = N2Str2Null(txtSvc_No.Text)
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    Dim kAdd                            As Integer
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

    gconDMIS.Execute "update CSMS_RepOr set" & _
                   " estimateno = " & VTXTestimateno & "," & _
                   " NOTE = " & XNOTE & "," & _
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
                   " where REP_OR =" & VTXTrep_or & ""

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
                   " acct_no = " & VTXTAcct_No & "," & _
                   " plate_no = " & VTXTPlate_No & "," & _
                   " model = " & VTXTModel & "," & _
                   " AppointmentDate = " & VTXTDte_recd & "," & _
                   " DateFinish = " & VTXTDte_comp & "," & _
                   " PromiseDate = '" & dtPromised & "'" & _
                   " where RO_NO = " & VTXTrep_or & ""
        LogAudit "E", "CUSTOMERS APPOINTMENT", txtRep_Or & " APPOINTMENT TIME " & txtDte_recd
    If OldApptTime <> dtPromised Or OldTranDate <> txtDte_recd Then
        Dim rsAppointment               As ADODB.Recordset
        Dim TimeApptNo                  As String
        Set rsAppointment = New ADODB.Recordset
        Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptTime = '" & dtPromised & "' and Trandate = " & VTXTDte_recd)
        If Not rsAppointment.EOF And Not rsAppointment.BOF Then
            TimeApptNo = Null2String(rsAppointment!ApptNo)
        End If
        gconDMIS.Execute "update CSMS_Appointment set ApptTime ='" & dtPromised & "', TranDate = " & VTXTDte_recd & ", ApptNo = " & VTXTrep_or & ", CusCde = " & VTXTAcct_No & ", Plate_no = " & VTXTPlate_No & ", CUSNAM = " & VTXTNiym & ", MODEL = " & VTXTModel & ", MAKE = " & VTXTMake & " ,NOTE=" & N2Str2Null(txtnote) & ", km_rdg = " & VTXTKm_rdg & " where apptno= '" & txtRep_Or & "'"
    Else
        gconDMIS.Execute "update CSMS_Appointment set CusCde = " & VTXTAcct_No & ", Plate_no = " & VTXTPlate_No & ", CUSNAM = " & VTXTNiym & ", MODEL = " & VTXTModel & ", MAKE = " & VTXTMake & " ,NOTE=" & N2Str2Null(txtnote) & ",  km_rdg = " & VTXTKm_rdg & " where ApptTime = '" & dtPromised & "' AND TranDate = " & VTXTDte_recd
    End If
    cmdCancel.Value = True
    MsgBox "Information been Update!", vbInformation, "Information"
    LogAudit "E", "EDIT CUSTOMER APPOINTMENT", "APPOINTMENT NO." & VTXTrep_or
    Exit Sub

Errorcode:

    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelectCustomer_Click()
    frmCSMS_EditAppCustomerSearch.Show 1
End Sub

Private Sub Form_Load()
    Dim rsTmp                           As New ADODB.Recordset
    With frmCSMSEditAppointment
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next ctl
    End With
    InitCbo
End Sub

Sub StoreAppInfo(XXX As String)
    Dim rsUpload                        As ADODB.Recordset
    Dim rsTmp                           As ADODB.Recordset
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select * from CSMS_Repor where REP_OR = '" & XXX & "'")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        labID = Null2String(rsUpload!ID)
        txtnote = Null2String(rsUpload!NOTE)
        txtRep_Or.Text = Null2String(rsUpload!REP_OR)
        OldApptNo = Null2String(rsUpload!REP_OR)
        txtInvoiceNo.Text = Null2String(rsUpload!invoice)
        txtDte_Rel.Text = Null2String(rsUpload!dte_rel)
        txtEstimateno.Text = Null2String(rsUpload!EstimateNo)
        txtROType.Text = Null2String(rsUpload!rotype)
        txtSvc_No.Text = Null2String(rsUpload!svc_no)
        txtAcct_No.Text = Null2String(rsUpload!ACCT_NO)
        txtNiym.Text = Null2String(rsUpload!niym)
        txtPlate_No.Text = Null2String(rsUpload!Plate_no)
        SetVehicleInfo txtPlate_No.Text
        txtTerm.Text = Null2String(rsUpload!TERM)
        txtSektion.Text = Null2String(rsUpload!sektion)
        cboRecd_by.Text = SetSA(Null2String(rsUpload!RECD_BY))
        txtKm_rdg.Text = Null2String(rsUpload!KM_RDG)
        txtDte_recd.Value = Null2String(rsUpload!dte_recd)

        txtCertific8.Text = Null2String(rsUpload!certific8)
        txtDte_comp.Text = Null2String(rsUpload!dte_comp)
        txtVIN.Text = Null2String(rsUpload!Vin)
        txtParticipat.Text = Null2String(rsUpload!participat)
    End If
    Dim rsAppointment                   As ADODB.Recordset
    Set rsAppointment = New ADODB.Recordset
    Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptNo = '" & XXX & "'")
    If Not rsAppointment.EOF And Not rsAppointment.BOF Then
        txtDte_recd.Value = Null2String(rsAppointment!trandate)
        OldTranDate = Null2String(rsAppointment!trandate)
        OldApptTime = Null2String(rsAppointment!ApptTime)
        dtPromised = Null2String(rsAppointment!ApptTime)
    End If
    Set rsTmp = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & txtAcct_No.Text & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        lblCN.Caption = Null2String(rsTmp!HomePhone) & "\" & Null2String(rsTmp!TelephoneNo) & "\" & Null2String(rsTmp!Mobile)
    End If
    Set rsTmp = Nothing
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
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSA = Null2String(rsEmpNo!naym)
    Set rsEmpNo = Nothing
End Function

Sub InitCbo()
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo order by naym asc")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        cboRecd_by.Clear
        cboRecd_by.Text = Null2String(rsEmpNo!naym)
        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!naym)
            rsEmpNo.MoveNext
        Loop
    End If
End Sub

Sub StoreApptTme()
    Dim rsCSMS_ApptSchedule             As ADODB.Recordset
    Set rsCSMS_ApptSchedule = New ADODB.Recordset
    Set rsCSMS_ApptSchedule = gconDMIS.Execute("Select * from CSMS_ApptSchedule Order by ID asc")
    If Not rsCSMS_ApptSchedule.EOF And Not rsCSMS_ApptSchedule.BOF Then
        rsCSMS_ApptSchedule.MoveFirst: dtPromised.Clear
        Do While Not rsCSMS_ApptSchedule.EOF
            dtPromised.AddItem Null2String(rsCSMS_ApptSchedule!timeInterval)
            rsCSMS_ApptSchedule.MoveNext
        Loop
    End If
End Sub

Function SetCodeSA(nam As String)
    Set rsEmpNo = New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!code)
    Set rsEmpNo = Nothing
End Function

Sub SetCustomer()
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtAcct_No.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtNiym.Text = Null2String(rsCustomer!AcctName)
    End If
End Sub

