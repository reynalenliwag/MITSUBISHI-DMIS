VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSNewAppointment 
   BackColor       =   &H00FCD5BC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Entry"
   ClientHeight    =   8580
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "FrmNewAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMaster 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   30
      ScaleHeight     =   8475
      ScaleWidth      =   10485
      TabIndex        =   7
      Top             =   30
      Width           =   10545
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   2370
         ScaleHeight     =   705
         ScaleWidth      =   7875
         TabIndex        =   12
         Top             =   90
         Width           =   7875
         Begin VB.TextBox txtCustName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   1860
            TabIndex        =   13
            Top             =   90
            Width           =   5895
         End
         Begin VB.TextBox txtID 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   1320
         Top             =   6900
      End
      Begin VB.TextBox txtTranNo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtEstimateEdit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2130
         TabIndex        =   10
         Text            =   "txtEstimateEdit"
         Top             =   7560
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtAppointmentEdit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2130
         TabIndex        =   9
         Text            =   "txtAppointmentEdit"
         Top             =   7890
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   855
         Left            =   9600
         MouseIcon       =   "FrmNewAppointment.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "FrmNewAppointment.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel Transaction"
         Top             =   7500
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Height          =   855
         Left            =   8820
         Picture         =   "FrmNewAppointment.frx":1512
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Next"
         Top             =   7500
         Width           =   795
      End
      Begin VB.CommandButton cmdBack 
         Height          =   855
         Left            =   8040
         Picture         =   "FrmNewAppointment.frx":1860
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Previous "
         Top             =   7500
         Width           =   795
      End
      Begin VB.PictureBox picAppointment 
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
         Height          =   6165
         Left            =   2250
         ScaleHeight     =   6135
         ScaleWidth      =   8025
         TabIndex        =   144
         Top             =   930
         Visible         =   0   'False
         Width           =   8055
         Begin VB.ComboBox Cbo_Rotype 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   6540
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   212
            Top             =   450
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Change Service Adviser"
            Height          =   315
            Left            =   1380
            TabIndex        =   207
            Top             =   5760
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtAddress 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   2040
            Width           =   6495
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
            Left            =   6750
            MaxLength       =   3
            TabIndex        =   164
            Top             =   5370
            Width           =   1035
         End
         Begin VB.TextBox txtAcct_No 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   1650
            Width           =   1065
         End
         Begin VB.TextBox txtEstimateno 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   162
            Top             =   900
            Width           =   1755
         End
         Begin VB.TextBox txtNiym 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   161
            Top             =   1650
            Width           =   5385
         End
         Begin VB.TextBox txtRep_Or 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   160
            Top             =   540
            Width           =   1755
         End
         Begin VB.ComboBox cboRecd_by 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1380
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   5370
            Width           =   3855
         End
         Begin VB.TextBox txtParticipat 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   158
            Text            =   "Text1"
            Top             =   6330
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.Frame Frame1 
            Caption         =   "Vehicle information"
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
            Height          =   2475
            Left            =   90
            TabIndex        =   147
            Top             =   2460
            Width           =   7815
            Begin VB.TextBox cboModel 
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   152
               Top             =   1050
               Width           =   6225
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   151
               Top             =   660
               Width           =   2295
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
               Left            =   1290
               Locked          =   -1  'True
               MaxLength       =   7
               TabIndex        =   150
               Top             =   1440
               Width           =   1755
            End
            Begin VB.TextBox txtYear 
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   149
               Top             =   270
               Width           =   915
            End
            Begin VB.TextBox txtVIN 
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
               Left            =   1290
               Locked          =   -1  'True
               TabIndex        =   148
               Top             =   1830
               Width           =   2985
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
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   735
               TabIndex        =   157
               Top             =   1110
               Width           =   510
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
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   795
               TabIndex        =   156
               Top             =   750
               Width           =   450
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
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   540
               TabIndex        =   155
               Top             =   1530
               Width           =   705
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Year"
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
               Left            =   870
               TabIndex        =   154
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "VIN No."
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
               Left            =   675
               TabIndex        =   153
               Top             =   1920
               Width           =   570
            End
         End
         Begin VB.TextBox txtApointmentNo 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   146
            Top             =   1260
            Width           =   1755
         End
         Begin VB.TextBox txtKm_rdg 
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
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   6750
            MaxLength       =   10
            TabIndex        =   145
            Top             =   4980
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker txtDte_recd 
            Height          =   345
            Left            =   1410
            TabIndex        =   166
            Top             =   4980
            Width           =   1995
            _ExtentX        =   3519
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
            CalendarBackColor=   16777215
            CustomFormat    =   "MMM. dd, yyyy"
            Format          =   56426499
            CurrentDate     =   38936
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
            Left            =   5820
            TabIndex        =   214
            Top             =   570
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lbl_rodescription 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   5700
            TabIndex        =   213
            Top             =   825
            Visible         =   0   'False
            Width           =   2175
         End
         Begin XtremeShortcutBar.ShortcutCaption SCCap 
            Height          =   375
            Left            =   0
            TabIndex        =   208
            Top             =   -30
            Width           =   8025
            _Version        =   655364
            _ExtentX        =   14155
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   " Repair Order Information"
            ForeColor       =   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   4194304
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Repair Order Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   4500
            TabIndex        =   177
            Top             =   30
            Width           =   2820
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor"
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
            TabIndex        =   176
            Top             =   5430
            Width           =   1305
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
            Left            =   5790
            TabIndex        =   175
            Top             =   5430
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Estimate No."
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
            Left            =   360
            TabIndex        =   174
            Top             =   990
            Width           =   1020
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   660
            TabIndex        =   173
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label Label7 
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
            Left            =   450
            TabIndex        =   172
            Top             =   630
            Width           =   930
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Left            =   495
            TabIndex        =   171
            Top             =   1770
            Width           =   840
         End
         Begin VB.Label Label21 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Insurance             Participation"
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
            Height          =   435
            Left            =   210
            TabIndex        =   170
            Top             =   6390
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label11 
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
            Left            =   165
            TabIndex        =   169
            Top             =   5070
            Width           =   1200
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Appointment"
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
            Left            =   300
            TabIndex        =   168
            Top             =   1350
            Width           =   1080
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
            Left            =   5730
            TabIndex        =   167
            Top             =   5070
            Width           =   960
         End
      End
      Begin VB.PictureBox picVehicle 
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
         Height          =   6165
         Left            =   2250
         ScaleHeight     =   6135
         ScaleWidth      =   8025
         TabIndex        =   136
         Top             =   930
         Visible         =   0   'False
         Width           =   8055
         Begin VB.CommandButton cmdAddVeh 
            Caption         =   "&Add/Edit/Delete Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   138
            Top             =   5640
            Width           =   2355
         End
         Begin VB.TextBox txtVehName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   600
            Width           =   6945
         End
         Begin MSComctlLib.ListView lstVehicle 
            Height          =   4485
            Left            =   120
            TabIndex        =   139
            Top             =   1080
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   7911
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmNewAppointment.frx":1BAE
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Model"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Plate No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Serial No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Engine"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Prod'n. No."
               Object.Width           =   2540
            EndProperty
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Left            =   0
            TabIndex        =   209
            Top             =   0
            Width           =   8025
            _Version        =   655364
            _ExtentX        =   14155
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   " Select Vehicle"
            ForeColor       =   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   4194304
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   141
            Top             =   90
            Width           =   1605
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   6
            Left            =   150
            TabIndex        =   140
            Top             =   660
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
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
         Left            =   2250
         TabIndex        =   93
         Top             =   930
         Width           =   8055
         Begin VB.TextBox textSearch 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2010
            TabIndex        =   132
            Top             =   180
            Width           =   5865
         End
         Begin VB.OptionButton optFN 
            Caption         =   "First Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3180
            TabIndex        =   131
            Top             =   660
            Width           =   1185
         End
         Begin VB.OptionButton optLN 
            Caption         =   "Last Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1950
            TabIndex        =   130
            Top             =   660
            Width           =   1185
         End
         Begin VB.OptionButton optFullName 
            Caption         =   "Full Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            TabIndex        =   129
            Top             =   660
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddeditCustomer 
            Caption         =   "Add/Edit/Delete Customer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   128
            ToolTipText     =   "Add/Edit/Delete Customer"
            Top             =   5670
            Width           =   2475
         End
         Begin VB.OptionButton optEndUser 
            Caption         =   "End User"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5700
            TabIndex        =   127
            Top             =   660
            Width           =   1035
         End
         Begin VB.OptionButton optPlate 
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
            Height          =   285
            Left            =   6870
            TabIndex        =   95
            Top             =   660
            Width           =   1035
         End
         Begin VB.CheckBox chkSettingAll 
            Caption         =   "Show All Customer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5430
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Click to Search all Customer"
            Top             =   5670
            Value           =   1  'Checked
            Width           =   2475
         End
         Begin MSComctlLib.ListView lstCustomer 
            Height          =   4515
            Left            =   180
            TabIndex        =   133
            Top             =   1020
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   7964
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmNewAppointment.frx":1D10
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Last Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "First Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Address"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Province"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Phone No."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "CusName"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.PictureBox picPlate 
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
            Height          =   4515
            Left            =   180
            ScaleHeight     =   4485
            ScaleWidth      =   7725
            TabIndex        =   96
            Top             =   1020
            Visible         =   0   'False
            Width           =   7755
            Begin VB.Frame Frame6 
               Caption         =   "Vehicle Information"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   2145
               Left            =   90
               TabIndex        =   106
               Top             =   2220
               Width           =   7545
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   8
                  Left            =   5310
                  TabIndex        =   122
                  Top             =   1290
                  Width           =   1875
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   7
                  Left            =   1860
                  TabIndex        =   121
                  Top             =   1680
                  Width           =   5325
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   6
                  Left            =   1860
                  TabIndex        =   120
                  Top             =   1320
                  Width           =   2085
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   5
                  Left            =   1860
                  TabIndex        =   119
                  Top             =   960
                  Width           =   2085
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   4
                  Left            =   1860
                  TabIndex        =   118
                  Top             =   630
                  Width           =   2085
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   3
                  Left            =   1860
                  TabIndex        =   117
                  Top             =   270
                  Width           =   5325
               End
               Begin VB.Label LBLcap 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Selling Dealer :"
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
                  Index           =   8
                  Left            =   570
                  TabIndex        =   116
                  Top             =   1710
                  Width           =   1215
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Engine no. : "
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
                  Index           =   7
                  Left            =   795
                  TabIndex        =   115
                  Top             =   990
                  Width           =   990
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Km rdg. :"
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
                  Index           =   6
                  Left            =   4485
                  TabIndex        =   114
                  Top             =   1350
                  Width           =   735
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Color :"
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
                  Index           =   5
                  Left            =   1245
                  TabIndex        =   113
                  Top             =   1350
                  Width           =   540
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vin no. :"
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
                  Index           =   4
                  Left            =   1125
                  TabIndex        =   112
                  Top             =   660
                  Width           =   660
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Model :"
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
                  Index           =   3
                  Left            =   1200
                  TabIndex        =   111
                  Top             =   330
                  Width           =   600
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Plate no. :"
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
                  Index           =   9
                  Left            =   4410
                  TabIndex        =   110
                  Top             =   660
                  Width           =   795
               End
               Begin VB.Label LBLcap 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "CS no. : "
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
                  Index           =   10
                  Left            =   4620
                  TabIndex        =   109
                  Top             =   990
                  Width           =   660
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   10
                  Left            =   5310
                  TabIndex        =   108
                  Top             =   630
                  Width           =   1875
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   11
                  Left            =   5310
                  TabIndex        =   107
                  Top             =   960
                  Width           =   1875
               End
            End
            Begin VB.CommandButton cmdSPlate 
               Caption         =   "Search"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   6450
               TabIndex        =   124
               Top             =   180
               Width           =   1155
            End
            Begin VB.TextBox txtSPlate 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3420
               TabIndex        =   123
               Top             =   210
               Width           =   2955
            End
            Begin VB.Frame Frame5 
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
               ForeColor       =   &H00800000&
               Height          =   1335
               Left            =   90
               TabIndex        =   98
               Top             =   720
               Width           =   7545
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   9
                  Left            =   2040
                  TabIndex        =   105
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   104
                  Top             =   900
                  Width           =   4725
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   1
                  Left            =   2040
                  TabIndex        =   103
                  Top             =   570
                  Width           =   5325
               End
               Begin VB.Label lblINFO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   0
                  Left            =   3090
                  TabIndex        =   102
                  Top             =   240
                  Width           =   4275
               End
               Begin VB.Label LBLcap 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Contact no. :"
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
                  Left            =   750
                  TabIndex        =   101
                  Top             =   930
                  Width           =   1020
               End
               Begin VB.Label LBLcap 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Address :"
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
                  Left            =   120
                  TabIndex        =   100
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.Label LBLcap 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Name :"
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
                  Left            =   300
                  TabIndex        =   99
                  Top             =   300
                  Width           =   1440
               End
            End
            Begin VB.ComboBox cboSearchBy 
               Height          =   330
               Left            =   1350
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   210
               Width           =   1965
            End
            Begin VB.Label lblPlateFind 
               BackColor       =   &H000000FF&
               Caption         =   "Plate Find"
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
               Left            =   5070
               TabIndex        =   126
               Top             =   -360
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Search By :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   150
               TabIndex        =   125
               Top             =   270
               Width           =   1095
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Customer"
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
            Index           =   0
            Left            =   240
            TabIndex        =   135
            Top             =   300
            Width           =   1560
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F3 - SEARCH"
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
            Left            =   270
            TabIndex        =   134
            Top             =   690
            Width           =   975
         End
      End
      Begin VB.PictureBox picEstimate 
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
         Height          =   6165
         Left            =   2250
         ScaleHeight     =   6135
         ScaleWidth      =   8025
         TabIndex        =   39
         Top             =   930
         Visible         =   0   'False
         Width           =   8055
         Begin VB.Frame Frame4 
            Caption         =   "Estimated Cost "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2505
            Left            =   210
            TabIndex        =   43
            Top             =   2880
            Width           =   7665
            Begin VB.TextBox txtVatTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6750
               Locked          =   -1  'True
               TabIndex        =   77
               Top             =   2070
               Width           =   825
            End
            Begin VB.TextBox txtVatAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6750
               Locked          =   -1  'True
               TabIndex        =   76
               Top             =   1680
               Width           =   825
            End
            Begin VB.TextBox txtVatParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6750
               Locked          =   -1  'True
               TabIndex        =   75
               Top             =   900
               Width           =   825
            End
            Begin VB.TextBox txtVatLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6750
               Locked          =   -1  'True
               TabIndex        =   74
               Top             =   510
               Width           =   825
            End
            Begin VB.TextBox txtDiscTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5850
               TabIndex        =   73
               Top             =   2070
               Width           =   855
            End
            Begin VB.TextBox txtDiscAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5850
               TabIndex        =   72
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtDiscParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5850
               TabIndex        =   71
               Top             =   900
               Width           =   855
            End
            Begin VB.TextBox txtDiscLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5850
               TabIndex        =   70
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox txtWarLaborTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   2070
               Width           =   975
            End
            Begin VB.TextBox txtWarLaborAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   68
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtWarParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   67
               Top             =   900
               Width           =   975
            End
            Begin VB.TextBox txtWarLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   66
               Top             =   510
               Width           =   975
            End
            Begin VB.TextBox txtSalesTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   65
               Top             =   2070
               Width           =   825
            End
            Begin VB.TextBox txtSalesAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   1680
               Width           =   795
            End
            Begin VB.TextBox txtSalesParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   900
               Width           =   795
            End
            Begin VB.TextBox txtSalesLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   510
               Width           =   795
            End
            Begin VB.TextBox txtCompTotal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   2070
               Width           =   975
            End
            Begin VB.TextBox txtCompAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtCompPart 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   59
               Top             =   900
               Width           =   975
            End
            Begin VB.TextBox txtCompLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   510
               Width           =   975
            End
            Begin VB.TextBox txtTotalAmt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   1230
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   2070
               Width           =   975
            End
            Begin VB.TextBox txtEstAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   1230
               TabIndex        =   56
               Top             =   1680
               Width           =   975
            End
            Begin VB.TextBox txtEstParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   1230
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   900
               Width           =   975
            End
            Begin VB.TextBox txtEstLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   1230
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   510
               Width           =   975
            End
            Begin VB.TextBox txtRateLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5130
               TabIndex        =   53
               Top             =   510
               Width           =   675
            End
            Begin VB.TextBox txtRateparts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5130
               TabIndex        =   52
               Top             =   900
               Width           =   675
            End
            Begin VB.TextBox txtRateAces 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5130
               TabIndex        =   51
               Top             =   1680
               Width           =   675
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5130
               TabIndex        =   50
               Top             =   1290
               Width           =   675
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   1230
               TabIndex        =   49
               Top             =   1290
               Width           =   975
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   1290
               Width           =   975
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1290
               Width           =   795
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   4110
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1290
               Width           =   975
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   5850
               TabIndex        =   45
               Top             =   1290
               Width           =   855
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   6750
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   1290
               Width           =   825
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VAT"
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
               Left            =   7110
               TabIndex        =   89
               Top             =   270
               Width           =   345
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
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
               Left            =   6000
               TabIndex        =   88
               Top             =   270
               Width           =   720
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Warranty"
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
               Left            =   4200
               TabIndex        =   87
               Top             =   270
               Width           =   735
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales"
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
               Left            =   3480
               TabIndex        =   86
               Top             =   270
               Width           =   450
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company"
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
               Left            =   2400
               TabIndex        =   85
               Top             =   270
               Width           =   780
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
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
               Left            =   1440
               TabIndex        =   84
               Top             =   270
               Width           =   660
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL"
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
               Height          =   285
               Left            =   570
               TabIndex        =   83
               Top             =   2130
               Width           =   615
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Accessories"
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
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   82
               Top             =   1770
               Width           =   1035
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Parts"
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
               Height          =   195
               Left            =   750
               TabIndex        =   81
               Top             =   990
               Width           =   435
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Labor"
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
               Height          =   195
               Left            =   705
               TabIndex        =   80
               Top             =   600
               Width           =   480
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Disc.Rate"
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
               Left            =   5130
               TabIndex        =   79
               Top             =   270
               Width           =   750
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materials"
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
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   78
               Top             =   1380
               Width           =   1005
            End
         End
         Begin VB.CommandButton cmdAddAcc 
            Caption         =   "Add Accessories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5340
            TabIndex        =   40
            Top             =   5550
            Width           =   2565
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2085
            Left            =   180
            TabIndex        =   90
            Top             =   540
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   3678
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
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
            MouseIcon       =   "FrmNewAppointment.frx":1E72
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Part No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6879
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "SRP"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.CommandButton cmdAddMat 
            Caption         =   "Add Materials"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2790
            TabIndex        =   41
            Top             =   5550
            Width           =   2565
         End
         Begin VB.CommandButton cmdAddParts 
            Caption         =   "Add Parts"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   42
            Top             =   5550
            Width           =   2565
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   0
            TabIndex        =   210
            Top             =   0
            Width           =   8025
            _Version        =   655364
            _ExtentX        =   14155
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Parts and Accessories Estimate"
            ForeColor       =   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   4194304
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOUBLE CLICK TO REMOVE DETAIL(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   1
            Left            =   4815
            TabIndex        =   91
            Top             =   2640
            Width           =   3000
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Parts and Accessories Estimate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   3270
            TabIndex        =   92
            Top             =   60
            Width           =   3615
         End
      End
      Begin VB.PictureBox picReason 
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
         Height          =   6165
         Left            =   2250
         ScaleHeight     =   6135
         ScaleWidth      =   8025
         TabIndex        =   16
         Top             =   930
         Visible         =   0   'False
         Width           =   8055
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Job"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   25
            Top             =   5640
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   150
            TabIndex        =   19
            Top             =   4980
            Width           =   7755
            Begin VB.TextBox txtRecorded 
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
               Height          =   345
               Left            =   1560
               TabIndex        =   20
               Top             =   150
               Width           =   1425
            End
            Begin MSComCtl2.DTPicker dtPromised 
               Height          =   345
               Left            =   5220
               TabIndex        =   21
               Top             =   150
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
               Format          =   56426499
               CurrentDate     =   38936
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
               Left            =   210
               TabIndex        =   23
               Top             =   240
               Width           =   1200
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
               Left            =   3450
               TabIndex        =   22
               Top             =   240
               Width           =   1680
            End
         End
         Begin VB.TextBox txtRecomendation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   150
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   4140
            Width           =   3885
         End
         Begin VB.TextBox txtInst 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   4230
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   4140
            Width           =   3675
         End
         Begin MSComctlLib.ListView lblJob4Service 
            Height          =   1515
            Left            =   120
            TabIndex        =   24
            Top             =   450
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   2672
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmNewAppointment.frx":1FD4
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Type"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Jobs Description"
               Object.Width           =   8467
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Flat Rate"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Std.Time"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Discount"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Charged To"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Note"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "is_war"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "BP_TYPE"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "QS"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "PMS_READING"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstPMSDet 
            Height          =   1305
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   2302
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmNewAppointment.frx":2136
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Type"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   11289
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Model"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add &Canned Labor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4950
            TabIndex        =   27
            Top             =   5640
            Width           =   1665
         End
         Begin VB.CommandButton cmdPMS 
            Caption         =   "Add &PMS Jobs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3300
            TabIndex        =   28
            Top             =   5640
            Width           =   1665
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "Add &Other Jobs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1740
            TabIndex        =   29
            Top             =   5640
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddJobs 
            Caption         =   "Add &General Job"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   30
            Top             =   5640
            Width           =   1605
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
            Height          =   375
            Left            =   0
            TabIndex        =   211
            Top             =   0
            Width           =   8025
            _Version        =   655364
            _ExtentX        =   14155
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Jobs for Service"
            ForeColor       =   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   4194304
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Jobs for Service"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   60
            Width           =   1860
         End
         Begin VB.Label lblStdHrs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7020
            TabIndex        =   37
            Top             =   2010
            Width           =   765
         End
         Begin VB.Label lbltlFaltRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   4590
            TabIndex        =   36
            Top             =   2010
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&PMS/Canned Job Details :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   8
            Left            =   150
            TabIndex        =   35
            Top             =   2100
            Width           =   2460
         End
         Begin VB.Label labNotes 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Request"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   34
            Top             =   3870
            Width           =   1725
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Flat Rate :"
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
            Left            =   3330
            TabIndex        =   33
            Top             =   2100
            Width           =   1215
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Total STD Time :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            TabIndex        =   32
            Top             =   2100
            Width           =   1305
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Advisor Instruction"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4200
            TabIndex        =   31
            Top             =   3870
            Width           =   2535
         End
      End
      Begin VB.Label lblDate 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   4050
         TabIndex        =   206
         Top             =   7920
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblTime 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   4080
         TabIndex        =   205
         Top             =   7470
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label labType 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   188
         Top             =   7530
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   2040
         X2              =   2040
         Y1              =   60
         Y2              =   8250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Search"
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
         Index           =   0
         Left            =   150
         TabIndex        =   187
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle"
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
         Left            =   150
         TabIndex        =   186
         Top             =   3630
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order"
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
         Left            =   150
         TabIndex        =   185
         Top             =   4140
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs"
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
         Left            =   150
         TabIndex        =   184
         Top             =   5160
         Width           =   405
      End
      Begin VB.Shape shpCustomer 
         BackColor       =   &H00D8E9EC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   3030
         Width           =   1875
      End
      Begin VB.Image Image1 
         Height          =   1560
         Left            =   150
         Picture         =   "FrmNewAppointment.frx":2298
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1725
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   240
         X2              =   10410
         Y1              =   7350
         Y2              =   7350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate"
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
         Left            =   150
         TabIndex        =   183
         Top             =   5670
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment"
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
         Left            =   150
         TabIndex        =   182
         Top             =   4650
         Width           =   1080
      End
      Begin VB.Label labType 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Index           =   0
         Left            =   135
         TabIndex        =   181
         Top             =   7560
         Width           =   1845
      End
      Begin VB.Label labTranType 
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Order No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   180
         Top             =   1860
         Width           =   2295
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   765
         Left            =   2250
         TabIndex        =   179
         Top             =   60
         Width           =   8055
         Size            =   "14208;1349"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label labEdit 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   180
         TabIndex        =   178
         Top             =   6150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape shpVehicle 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   3540
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Shape shpRO 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   4050
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Shape ShpAppointment 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Shape shpJobs 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   5070
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Shape ShpEstimate 
         BackColor       =   &H00C8CBFD&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   405
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   5580
         Visible         =   0   'False
         Width           =   1875
      End
   End
   Begin VB.PictureBox PicDet 
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
      Height          =   5775
      Left            =   1080
      ScaleHeight     =   5745
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   1403
      Visible         =   0   'False
      Width           =   8445
      Begin VB.ComboBox cboSModel 
         Height          =   330
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   720
         Width           =   2865
      End
      Begin VB.CheckBox chkModel 
         Caption         =   "By Model"
         Height          =   255
         Left            =   4080
         TabIndex        =   203
         Top             =   750
         Width           =   1335
      End
      Begin VB.PictureBox picADD 
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
         Height          =   2655
         Left            =   1500
         ScaleHeight     =   2625
         ScaleWidth      =   5415
         TabIndex        =   191
         Top             =   1680
         Visible         =   0   'False
         Width           =   5445
         Begin VB.CommandButton Command4 
            Caption         =   "Close"
            Height          =   735
            Left            =   4560
            MouseIcon       =   "FrmNewAppointment.frx":10E0A
            MousePointer    =   99  'Custom
            Picture         =   "FrmNewAppointment.frx":10F5C
            Style           =   1  'Graphical
            TabIndex        =   194
            ToolTipText     =   "Cancel"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txtEQTY 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1260
            TabIndex        =   192
            Top             =   1560
            Width           =   1605
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "Add"
            Height          =   735
            Left            =   3840
            MouseIcon       =   "FrmNewAppointment.frx":1129A
            MousePointer    =   99  'Custom
            Picture         =   "FrmNewAppointment.frx":113EC
            Style           =   1  'Graphical
            TabIndex        =   193
            ToolTipText     =   "Save Entry"
            Top             =   1800
            Width           =   735
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   202
            Top             =   0
            Width           =   8085
            _Version        =   655364
            _ExtentX        =   14261
            _ExtentY        =   450
            _StockProps     =   14
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   0
            GradientColorLight=   4210752
            GradientColorDark=   12632256
            ForeColor       =   16777215
         End
         Begin VB.Label lblRES 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Index           =   2
            Left            =   1260
            TabIndex        =   201
            Top             =   1140
            Width           =   1605
         End
         Begin VB.Label lblRES 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Index           =   1
            Left            =   1260
            TabIndex        =   200
            Top             =   750
            Width           =   4005
         End
         Begin VB.Label lblRES 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Index           =   0
            Left            =   1260
            TabIndex        =   199
            Top             =   330
            Width           =   2385
         End
         Begin VB.Label LBLcap 
            AutoSize        =   -1  'True
            Caption         =   "QTY"
            Height          =   210
            Index           =   14
            Left            =   840
            TabIndex        =   198
            Top             =   1650
            Width           =   330
         End
         Begin VB.Label LBLcap 
            AutoSize        =   -1  'True
            Caption         =   "SRP"
            Height          =   210
            Index           =   13
            Left            =   870
            TabIndex        =   197
            Top             =   1230
            Width           =   300
         End
         Begin VB.Label LBLcap 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock Description"
            Height          =   390
            Index           =   12
            Left            =   180
            TabIndex        =   196
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label LBLcap 
            AutoSize        =   -1  'True
            Caption         =   "Stock No"
            Height          =   210
            Index           =   11
            Left            =   540
            TabIndex        =   195
            Top             =   420
            Width           =   645
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   735
         Left            =   7620
         MouseIcon       =   "FrmNewAppointment.frx":11687
         MousePointer    =   99  'Custom
         Picture         =   "FrmNewAppointment.frx":117D9
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancel"
         Top             =   4860
         Width           =   735
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Description"
         Height          =   195
         Left            =   1110
         TabIndex        =   5
         Top             =   780
         Width           =   1665
      End
      Begin VB.OptionButton optCode 
         Caption         =   "Code"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   780
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   330
         Width           =   8265
      End
      Begin MSComctlLib.ListView lsvDET 
         Height          =   3705
         Left            =   60
         TabIndex        =   3
         Top             =   1050
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6535
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "SRP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Model Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Genuine"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Onhand"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "DOUBLE CLICK/ OR PRESS ENTER TO ADD ITEM"
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
         Left            =   90
         TabIndex        =   190
         Top             =   4800
         Width           =   3660
      End
      Begin VB.Label lblTYPE 
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
         Height          =   375
         Left            =   120
         TabIndex        =   189
         Top             =   5190
         Visible         =   0   'False
         Width           =   1425
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   0
         Left            =   -30
         TabIndex        =   1
         Top             =   0
         Width           =   8475
         _Version        =   655364
         _ExtentX        =   14949
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "  ADD DETAILS"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   0
         GradientColorLight=   4210752
         GradientColorDark=   12632256
         ForeColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmCSMSNewAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AUDIT_SQL                                          As String
Dim rsAddRepor                                         As ADODB.Recordset
Dim rsAddRepor2                                        As ADODB.Recordset
Dim rsFind                                             As ADODB.Recordset
Dim CTL                                                As Control
Dim bevvy                                              As Long
Dim tlHrs                                              As Double
Dim tlFR                                               As Double
Dim xPartsAmt                                          As Double
Dim xAcesAmt                                           As Double
Dim xMatAmt                                            As Double
Dim xTransType                                         As String
Dim xApptNo                                            As String
Dim xESTIMATENO                                        As String

Dim JobTotal                                           As Double
Dim JobComTotal                                        As Double
Dim JobSalesTotal                                      As Double
Dim JobWarTotal                                        As Double
Dim JobDiscTotal                                       As Double
Dim JobVatTotal                                        As Double

Dim PartsTotal                                         As Double
Dim PartsComTotal                                      As Double
Dim PartsSalesTotal                                    As Double
Dim PartsWarTotal                                      As Double
Dim PartsDiscTotal                                     As Double
Dim PartsVatTotal                                      As Double

Dim MatTotal                                           As Double
Dim MatComTotal                                        As Double
Dim MatSalesTotal                                      As Double
Dim MatWarTotal                                        As Double
Dim MatDiscTotal                                       As Double
Dim MatVatTotal                                        As Double

Dim ACCTotal                                           As Double
Dim ACCComTotal                                        As Double
Dim ACCSalesTotal                                      As Double
Dim ACCWarTotal                                        As Double
Dim ACCDiscTotal                                       As Double
Dim ACCVatTotal                                        As Double

Dim COMTotal                                           As Double
Dim SALESTotal                                         As Double
Dim WARTotal                                           As Double
Dim VATTotal                                           As Double
Dim ROTotal                                            As Double

Dim EndUserCode                                        As String
Dim str_MSG                                            As String

Function ComputeResultOfRatenTimeWhenJobDelete()
    Dim X                                              As Integer
    Dim FR                                             As Double
    Dim ST                                             As Double

    For X = 1 To lblJob4Service.ListItems.Count
        FR = FR + lblJob4Service.ListItems(X).SubItems(3)
        ST = ST + lblJob4Service.ListItems(X).SubItems(4)
    Next

    lbltlFaltRate.Caption = FR
    lblStdHrs.Caption = ST
End Function

Function IsBodyOrSublet(XXX As String) As Boolean
    Dim rsJOBS                                         As New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("Select * from CSMS_Jobs Where JCode = '" & XXX & "'")
    If Not rsJOBS.EOF And Not rsJOBS.BOF Then
        If Trim(Null2String(rsJOBS!MAIN_CAT)) = "60" Or Trim(Null2String(rsJOBS!MAIN_CAT)) = "99" Or Left(Trim(Null2String(rsJOBS!JCode)), 2) = "SR" Then
            IsBodyOrSublet = True
        Else
            IsBodyOrSublet = False
        End If
    End If
End Function

Function GetJobCat(XXX As Variant)
    Dim rsGetJC                                        As New ADODB.Recordset
    Set rsGetJC = gconDMIS.Execute("Select JobCategory from CSMS_vw_Jobs where [jcode] = '" & XXX & "'")
    If Not rsGetJC.EOF And Not rsGetJC.BOF Then
        GetJobCat = Null2String(rsGetJC!CSMS_JobCategory)
    End If
    Set rsGetJC = Nothing
End Function

Function GetColor(CCC As String)
    Dim rsColor                                        As New ADODB.Recordset
    rsColor.Open "select COLOR_DESC from ALL_Color where COLOR_CODE = '" & CCC & "'", gconDMIS
    If Not rsColor.EOF And Not rsColor.BOF Then
        GetColor = Null2String(rsColor!color_desc)
    Else
        GetColor = ""
    End If
End Function

Function GetSellingDealer(XXX)
    Dim temprs                                         As New ADODB.Recordset
    Set temprs = gconDMIS.Execute("select Dealername from CSMS_SellingDealer where dealerCode='" & XXX & "'")
    If Not temprs.BOF Or Not temprs.EOF Then
        GetSellingDealer = Null2String(temprs!dealername)
    End If
    Set temprs = Nothing
End Function

Function SetCodeSA(nam As String) As String
    Dim rsEmpNo                                        As New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym,empno from CSMS_vw_EmpNo where naym = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetCodeSA = Null2String(rsEmpNo!Code)
    Set rsEmpNo = Nothing
End Function

Function SetSAname(nam As String)
    Dim rsEmpNo                                        As New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("Select code,naym,empno from CSMS_vw_EmpNo where empno = '" & nam & "'")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then SetSAname = Null2String(rsEmpNo!NAYM)
    Set rsEmpNo = Nothing
End Function

Function SetEndUserName(XXX As String) As String
    Dim rsEndUser                                      As New ADODB.Recordset
    Set rsEndUser = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE = '" & XXX & "'")
    If Not rsEndUser.EOF And Not rsEndUser.BOF Then
        SetEndUserName = Null2String(rsEndUser!lastname) & ", " & Null2String(rsEndUser!Firstname)
    End If
    Set rsEndUser = Nothing
End Function

Function SetEndACCTName(XXX As String) As String
    Dim rsCustomer                                     As New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE CUSCDE = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetEndACCTName = Null2String(rsCustomer!lastname) & ", " & Null2String(rsCustomer!Firstname)
    End If
    Set rsCustomer = Nothing
End Function

Sub EditTransaction()
    optFullName.Value = True
    Set rsFind = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        labTranType.Caption = "Repair Order  No."
        xTransType = "R"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where RO_NO = '" & txtTranNo & "'")
    ElseIf labType(0).Caption = "Estimate" Then
        labTranType.Caption = "Estimate No."
        xTransType = "E"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where ESTIMATENO = '" & txtTranNo & "'")
    ElseIf labType(0).Caption = "Appointment" Then
        labTranType.Caption = "Appointment No."
        xTransType = "A"
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where ApptNo = '" & txtTranNo & "'")
    End If
    If Not rsFind.EOF And Not rsFind.BOF Then
        If labType(0).Caption = "Repair Order" Then
            txtRep_Or = txtTranNo
        ElseIf labType(0).Caption = "Estimate" Then
            txtEstimateno = txtTranNo
        ElseIf labType(0).Caption = "Appointment" Then
            txtApointmentNo = txtTranNo
        End If
        txtEstimateEdit = Null2String(rsFind![EstimateNo])
        txtAppointmentEdit = Null2String(rsFind![APPTNO])
        textSearch = Null2String(rsFind![Customer])
        txtCustName = Null2String(UCase(rsFind![Customer]))
        txtNiym = Null2String(rsFind![Customer])
        txtID = Null2String(rsFind![ACCT_NO])
        txtAddress = Null2String(rsFind![CUSTOMERADD])
        txtAcct_No = Null2String(rsFind![ACCT_NO])
        txtDte_recd.Value = Null2String(rsFind![AppointmentDate])
        dtPromised.Value = Null2String(rsFind![PromiseDate])

        Dim rsVehicleKo                                As New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where Cuscde = '" & txtID & "' and plate_no = '" & rsFind![PLATE_NO] & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![PLATE_NO])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtYear = Null2String(rsVehicleKo![YER])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If

        Set rsVehicleKo = New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where Cuscde = '" & txtID & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            Listview_Loadval Me.lstVehicle.ListItems, rsVehicleKo
            lstCustomer.Refresh
        End If
        'JOBS
        Set rsFind = New ADODB.Recordset
        If labType(0).Caption = "Repair Order" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditRO where rep_OR = '" & txtTranNo & "' order by jobtype,detdsc asc")
        ElseIf labType(0).Caption = "Estimate" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditEstimate where ESTIMATENO = '" & txtTranNo & "' order by jobtype,detdsc asc")
        ElseIf labType(0).Caption = "Appointment" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EditAppt where ApptNo = '" & txtTranNo & "' order by jobtype,detdsc asc")
        End If
        If Not rsFind.EOF And Not rsFind.BOF Then
            txtKm_rdg.Text = Null2String(rsFind![km_rdg])
            cboRecd_by.Text = SetSAname(Null2String(rsFind![RECD_BY]))
            txtSektion.Text = Null2String(rsFind![sektion])
            txtRecorded = Null2String(rsFind![DTE_RECD])
            txtDte_recd.Value = Null2String(rsFind![DTE_RECD])
            txtParticipat.Text = Null2String(rsFind![participat])
            txtRecomendation = Null2String(rsFind![NOTE])
            Do Until rsFind.EOF
                With lblJob4Service
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![DETCDE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![JOBTYPE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![DETDSC])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , NumericVal(rsFind![FLATRATE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , NumericVal(rsFind![DET_HRS])
                    .ListItems(.ListItems.Count).ListSubItems.Add 5, , NumericVal(rsFind![discrate])
                    .ListItems(.ListItems.Count).ListSubItems.Add 6, , Null2String(rsFind![wCode])
                    .ListItems(.ListItems.Count).ListSubItems.Add 7, , Null2String(rsFind![Detail])
                End With
                rsFind.MoveNext
            Loop
        End If
        
        'PMS DETAILS
        Set rsFind = New ADODB.Recordset
        If labType(0).Caption = "Repair Order" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where rep_OR = '" & txtTranNo & "' order by pms_model,detcde asc")
        ElseIf labType(0).Caption = "Estimate" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where ESTIMATENO = '" & txtTranNo & "' order by pms_model,detcde asc")
        ElseIf labType(0).Caption = "Appointment" Then
            Set rsFind = gconDMIS.Execute("select * from CSMS_PMS_Job_Det where ApptNo = '" & txtTranNo & "' order by pms_model,detcde asc")
        End If
        If Not rsFind.EOF And Not rsFind.BOF Then
            Do Until rsFind.EOF
                With lstPMSDet
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![DETCDE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![JOBTYPE])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![DETDSC])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , Null2String(rsFind![PMS_Model])
                End With
                rsFind.MoveNext
            Loop
        End If
        
        'ESTIMATE
        Set rsFind = New ADODB.Recordset
        Set rsFind = gconDMIS.Execute("select * from CSMS_vw_EstimateDetails where ESTIMATENO = '" & txtTranNo & "'")
        If Not rsFind.EOF And Not rsFind.BOF Then
            Do Until rsFind.EOF
                With ListView1
                    .Sorted = False
                    .ListItems.Add , , Null2String(rsFind![Type])
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind![partno])
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsFind![PartDesc])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , NumericVal(rsFind![QTY])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , NumericVal(rsFind![SRP])
                End With
                rsFind.MoveNext
            Loop
        End If
    End If

    Exit Sub
ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub COMPUTETOTALITEM()
    Dim X                                              As Integer

    txtEstParts.Text = "0.00"
    Text2.Text = "0.00"
    txtEstAces.Text = "0.00"
    txtTotalAmt.Text = "0.00"

    'TYPE, PART NO, PART DESCRIPTION, QTY, SRP
    For X = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(X).Text = "P" Then
            txtEstParts = txtEstParts + (CCur(ListView1.ListItems(X).SubItems(3)) * CCur(ListView1.ListItems(X).SubItems(4)))
        ElseIf ListView1.ListItems(X).Text = "M" Then
            Text2 = Text2 + (CCur(ListView1.ListItems(X).SubItems(3)) * CCur(ListView1.ListItems(X).SubItems(4)))
        Else
            txtEstAces = txtEstAces + (CCur(ListView1.ListItems(X).SubItems(3)) * CCur(ListView1.ListItems(X).SubItems(4)))
        End If
    Next

    txtEstParts.Text = Format(txtEstParts.Text, MAXIMUM_DIGIT)
    Text2.Text = Format(Text2.Text, MAXIMUM_DIGIT)
    txtEstAces.Text = Format(txtEstAces.Text, MAXIMUM_DIGIT)
    txtTotalAmt.Text = Format(CCur(txtEstParts.Text) + CCur(Text2.Text) + CCur(txtEstAces.Text) + CCur(txtEstLabor.Text), MAXIMUM_DIGIT)
End Sub

Sub RemovePMSJobDetails(Code As String)
    Dim X                                              As Integer

    For X = lstPMSDet.ListItems.Count To 1 Step -1
        If Code = lstPMSDet.ListItems(X).SubItems(3) Then
            lstPMSDet.ListItems.Remove (X)
        End If
    Next
End Sub

Sub SaveAllInfo()
    If cboModel = "" And xTransType <> "A" Then
        MsgBox "Please select vehicle", vbInformation, "CSMS"
        Exit Sub
    End If
    If txtNiym = "" Then
        MsgBox "Please select Customer", vbInformation, "CSMS"

        Exit Sub
    End If
    If txtRep_Or = "" And txtEstimateno = "" And txtApointmentNo = "" Then
        MsgBox "Please Check RO No./Estimate No. or Appointment No.", vbInformation, "CSMS"
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgBox "Please Select Sales Advisor", vbInformation, "CSMS"
        Exit Sub
    End If
    'Updated by: IEBV 62820010 1018AM
    'Description:   For the HCI, to disable rotype for an appoinment transaction type
    '++++++++++++++++++++++++++++++++++++++++++
    If COMPANY_CODE = "HCI" Then
        If labType(0).Caption = "Appointment" Then
            Cbo_Rotype.ListIndex = 0
        Else
            If Cbo_Rotype.Text = "" Then
                Cbo_Rotype.ListIndex = 0
            End If
        End If
    End If
    'Updated by: IEBV 62820010 1018AM
    '++++++++++++++++++++++++++++++++++++++++++
    If labType(0).Caption = "Appointment" Then
        If MsgBox("Save This Appointment, Are You Sure", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
    ElseIf labType(0).Caption = "Estimate" Then
        If MsgBox("Save This Estimate, Are You Sure", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
    Else
        If MsgBox("Save This Repair Order, Are You Sure", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbNo Then Exit Sub
    End If


    Dim rstmp                                          As New ADODB.Recordset
    If labType(0).Caption = "Appointment" Then
        Set rstmp = gconDMIS.Execute("SELECT APPTNO FROM CSMS_REPOR WHERE APPTNO = '" & txtTranNo.Text & "' and TRANSTYPE = 'A'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            MsgBox "Appointment No. already Exist", vbExclamation, "Info."
            txtTranNo.SetFocus
            Exit Sub
        End If
    ElseIf labType(0).Caption = "Repair Order" Then
        Set rstmp = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE REP_OR = '" & txtTranNo.Text & "' AND TRANSTYPE = 'R'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            MsgBox "Repair Order No. already Exist", vbExclamation, "Info."
            txtTranNo.SetFocus
            Exit Sub
        End If
    Else
        Set rstmp = gconDMIS.Execute("SELECT ESTIMATENO FROM CSMS_ESTHD WHERE ESTIMATENO = '" & txtTranNo.Text & "' AND TRANSTYPE = 'E'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            MsgBox "Estimate No. already Exist", vbExclamation, "Info."
            txtTranNo.SetFocus
            Exit Sub
        End If
    End If


    'JUN 02/10/2008
    If labType(0).Caption = "Appointment" Then
        LogAudit "A", "DATA ENTRY - ADDED NEW APPOINTMENT", "Appointment No." & txtTranNo
    ElseIf labType(0).Caption = "Repair Order" Then
        LogAudit "A", "DATA ENTRY - ADDED NEW REPAIR ORDER", "RO No." & txtTranNo
    Else
        LogAudit "A", "DATA ENTRY - ADDED NEW ESTIMATE ", "Estimate No." & txtTranNo
    End If

    If labEdit.Caption = "Edit" Then
        If labType(0).Caption = "Repair Order" Then   'R/O
            gconDMIS.Execute "delete from CSMS_Repor where REP_OR = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where REP_OR = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where RO_No = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where REP_OR = '" & txtTranNo & "'"
        ElseIf labType(0).Caption = "Estimate" Then   'ESTIMATE
            gconDMIS.Execute "delete from CSMS_Repor where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_EstHD where ESTIMATENO = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_EstDETAILS where ESTIMATENO = '" & txtTranNo & "'"

        ElseIf labType(0).Caption = "Appointment" Then    'APPOINTMENT
            gconDMIS.Execute "delete from CSMS_Repor where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_Ro_Det where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_RepairOrder where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from CSMS_PMS_Job_Det where ApptNo = '" & txtTranNo & "'"
            gconDMIS.Execute "delete from csms_appointment where ApptNo = '" & txtTranNo & "'"
        End If
    End If

    Call SaveRepor
End Sub

Sub SaveEstimate()
    'On Error GoTo ERROR_CODE
    Screen.MousePointer = 11
    Dim xESTIMATENO                             As String
    Dim xACCT_NO                                As String
    Dim XTYPE                                   As String
    Dim xPARTNO                                 As String
    Dim xPARTDESC                               As String
    Dim xQTY                                    As Double
    Dim xSRP                                    As Double
    Dim X                                       As Long
    Dim vLIVIL                                  As String
    
    xESTIMATENO = N2Str2Null(txtEstimateno)
    xACCT_NO = N2Str2Null(txtAcct_No)

    If ListView1.ListItems.Count() <= 0 Then Exit Sub
    For X = 1 To ListView1.ListItems.Count
        XTYPE = N2Str2Null(ListView1.ListItems(X))
        xPARTNO = N2Str2Null(ListView1.ListItems(X).SubItems(1))
        xPARTDESC = N2Str2Null(ListView1.ListItems(X).SubItems(2))
        xQTY = NumericVal(ListView1.ListItems(X).SubItems(3))
        xSRP = NumericVal(ListView1.ListItems(X).SubItems(4))
        
        If ListView1.ListItems(X).Text = "P" Then vLIVIL = N2Str2Null("2")
        If ListView1.ListItems(X).Text = "M" Then vLIVIL = N2Str2Null("3")
        If ListView1.ListItems(X).Text = "A" Then vLIVIL = N2Str2Null("4")

        gconDMIS.Execute "insert into CSMS_EstDETAILS " & _
            "(TRANSTYPE, LIVIL, LINE_NO, DETCDE, DETDSC, DETVOL, DETPRC, DETAMT, DET_AMT, EstimateNo, REP_OR, TAXRATE, TAXVAL)" & _
            " values ('E' " & _
            ", " & vLIVIL & _
            ", " & N2Str2Null(Format(X, "00")) & _
            ", " & xPARTNO & _
            ", " & xPARTDESC & _
            ", " & xQTY & _
            ", " & xSRP & _
            ", " & (xQTY * xSRP) & _
            ", " & (xQTY * xSRP) & _
            ", " & xESTIMATENO & _
            ", " & xESTIMATENO & _
            ", " & VAT_RATE & _
            ", " & xSRP * 0.12 & ")"
    Next X
ERROR_CODE:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub SaveROMonitoring()
    Dim xMODEL                                          As String
    Dim xAppointmentDate                                As String
    Dim xRO_NO                                          As String
    Dim xACCT_NO                                        As String
    Dim xPLATE_NO                                       As String
    Dim xjCode                                          As String
    Dim xDescreption                                    As String
    Dim xRecommendation                                 As String
    Dim xStatus                                         As String
    Dim xWriter                                         As String
    Dim xPromiseDate                                    As String
    Dim xPromiseTime                                    As String
    Dim Vusercode                                       As String
    Dim VLastUpdate                                     As String
    Dim VLastUpdateTime                                 As String
    Dim xHours                                          As Double

    'On Error GoTo ErrorCode

    xAppointmentDate = N2Str2Null(Format(txtDte_recd, "MM/dd/yyyy"))
    xRO_NO = N2Str2Null(txtRep_Or)
    xACCT_NO = N2Str2Null(txtAcct_No)
    xPLATE_NO = N2Str2Null(txtPlate_No)
    xMODEL = N2Str2Null(cboModel)
    xRecommendation = "''"
    xHours = NumericVal(lblStdHrs)
    xStatus = "'Park'"
    xWriter = N2Str2Null(cboRecd_by)
    xRecommendation = N2Str2Null(txtRecomendation)
    xPromiseDate = N2Str2Null(dtPromised)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtApointmentNo)
    xESTIMATENO = N2Str2Null(txtEstimateno)
    If txtEstimateEdit <> "" Then
        xESTIMATENO = N2Str2Null(txtEstimateEdit)
    End If
    If txtAppointmentEdit <> "" Then
        xApptNo = N2Str2Null(txtAppointmentEdit)
    End If

    gconDMIS.Execute "insert into CSMS_RepairOrder " & _
        "(ESTIMATENO, ApptNo, TransType, model, AppointmentDate, RO_No, ACCT_NO, PLATE_NO, Recommendation, Hours, Status, Writer, PromiseDate, USERCDE, SAVEDATE, savetime)" & _
        " values (" & xESTIMATENO & _
        ", " & xApptNo & _
        ", '" & xTransType & _
        "', " & xMODEL & _
        ", " & xAppointmentDate & _
        ", " & xRO_NO & _
        ", " & xACCT_NO & _
        ", " & xPLATE_NO & _
        ", " & xRecommendation & _
        ", " & xHours & _
        ", " & xStatus & _
        ", " & xWriter & _
        ", " & xPromiseDate & _
        ", " & Vusercode & _
        ", " & VLastUpdate & _
        ", " & VLastUpdateTime & ")"
    
    Exit Sub

'ErrorCode:
'    str_MSG = Replace(str_MSG, "@UTX83912839123", "Update Master file")
'    MsgBox str_MSG, vbCritical, "Creating RO Error"
'    gconDMIS.RollbackTrans
'    Screen.MousePointer = 0
'    ShowVBError
'    Exit Sub
End Sub

Sub SaveRepor()
Dim str_MSG                                        As String
    If txtNiym.Text = "" Then
        MsgSpeechBox "Customer must have a name"
        On Error Resume Next
        txtNiym.SetFocus
        Exit Sub
    End If
    If cboRecd_by.Text = "" Then
        MsgSpeechBox "Service Advisor must not be Empty!"
        On Error Resume Next
        cboRecd_by.SetFocus
        Exit Sub
    Else
        Dim rsEmpNo                                    As New ADODB.Recordset
        Set rsEmpNo = gconDMIS.Execute("select naym from CSMS_vw_EmpNo where naym = '" & cboRecd_by.Text & "'")
        If rsEmpNo.EOF And rsEmpNo.BOF Then
            MsgSpeechBox "Invalid Service Advisor"
            On Error Resume Next
            Exit Sub
        End If
        Set rsEmpNo = Nothing
    End If

    Dim rsDupRepor                                     As New ADODB.Recordset
    Set rsDupRepor = gconDMIS.Execute("select rep_or from CSMS_RepOr where rep_or = " & N2Str2Null(txtTranNo.Text))
    If Not rsDupRepor.EOF And Not rsDupRepor.BOF Then
        If UCase(labType(0).Caption) = "ESTIMATE" Then
            MsgSpeechBox "Estimate Number Already Exist!"
        ElseIf UCase(labType(0).Caption) = "APPOINTMENT" Then
            MsgSpeechBox "Appointment Number Already Exist!"
        Else
            MsgSpeechBox "Repair Order Number Already Exist!"
        End If

        On Error Resume Next
        txtRep_Or.SetFocus
        Exit Sub
    End If
    Set rsDupRepor = Nothing
    
    str_MSG = "Error Appear In During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    
    gconDMIS.BeginTrans
    If finalsave = False Then
        str_MSG = Replace(str_MSG, "@UTX83912839123", "Creating" & UCase(labType(0).Caption))
        MsgBox str_MSG, vbCritical, "Saving Error"
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    gconDMIS.CommitTrans
    Unload Me
    If labType(0).Caption = "Estimate" Then frmCSMSEstimateEntry.Show
    If xTransType <> "A" Then TrigerTheRefresh        'BTT - 05212007
    Unload Me
    Exit Sub
    
ErrorCode:

    Screen.MousePointer = 0
    'Screen.MousePointer = 0:
    ShowVBError
Exit Sub

End Sub

Function finalsave() As Boolean

On Error GoTo IVANEXEQUIELVALENCIA

    Dim VTXTREP_OR                                      As String
    Dim VTXTestimateno                                  As String
    Dim VTXTROType                                      As String
    Dim VTXTSvc_No                                      As String
    Dim VTXTAcct_No                                     As String
    Dim VTXTNiym                                        As String
    Dim VTXTPlate_No                                    As String
    Dim VcboModel                                       As String
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
    Dim VLastUpdateTime                                 As String
    Dim Vusercode                                       As String
    Dim VLastUpdate                                     As String
    Dim VTXTParticipat                                  As String
    Dim VcboRecd_by                                     As String
    Dim XNOTE                                           As String
    Dim XDTE_PRO                                        As String
    Dim kAdd                                            As Integer
    Dim XINST                                           As String
    Dim xrotype                                         As String
    Dim rsRO_DET                                        As New ADODB.Recordset

    
    VTXTREP_OR = N2Str2Null(txtRep_Or.Text)
    VTXTestimateno = N2Str2Null(txtEstimateno.Text)
    'Updated by IEBV 06282010 1022AM
    '++++++++++++++++++++++++++++++++++++++++++++
    'VTXTROType = "''"
    VTXTROType = N2Str2Null(Cbo_Rotype.Text)
    '++++++++++++++++++++++++++++++++++++++++++++
    'Updated by IEBV 06282010 1022AM
    VTXTSvc_No = "''"
    VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
    VTXTNiym = N2Str2Null(txtNiym.Text)
    
    
    For kAdd = 1 To Len(txtAddress.Text)
        If Mid(txtAddress.Text, kAdd, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" And Mid(txtAddress.Text, kAdd + 1, 1) = "-" Then Exit For
        VtxtAddress = VtxtAddress & Mid(txtAddress.Text, kAdd, 1)
    Next
    
    VtxtAddress = N2Str2Null(VtxtAddress)
    VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
    VcboModel = N2Str2Null(cboModel.Text)
    VTXTMake = N2Str2Null(txtMake.Text)
    VTXTTerm = "''"
    VTXTSektion = N2Str2Null(txtSektion.Text)
    VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
    VTXTDte_recd = N2Date2Null(txtDte_recd)
    VTXTCertific8 = "''"
    VTXTDte_comp = "''"
    VTXTDte_Rel = "''"
    VtxtVIN = N2Str2Null(txtVIN.Text)
    VTXTParticipat = N2Str2Null(txtParticipat.Text)
    VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtApointmentNo)
    xESTIMATENO = N2Str2Null(txtEstimateno)
    XNOTE = N2Str2Null(txtRecomendation)
    XDTE_PRO = N2Str2Null(dtPromised.Value)
    XINST = N2Str2Null(txtInst.Text)
    xrotype = N2Str2Null(Cbo_Rotype.Text)
    
    If xTransType = "E" Then
        xESTIMATENO = N2Str2Null(txtEstimateno)
        xApptNo = N2Str2Null("")
    ElseIf xTransType = "A" Then
        xApptNo = N2Str2Null(txtApointmentNo)
        xESTIMATENO = N2Str2Null("")
    Else
        xApptNo = N2Str2Null("")
        xESTIMATENO = N2Str2Null("")
    End If
    SQL_STATEMENT = "insert into CSMS_RepOr " & _
        "(ESTIMATENO, ApptNo, TransType, [note], rep_or, rotype, svc_no, acct_no, niym, plate_no, model, term, sektion, Recd_by, km_rdg, dte_recd, DTE_PRO, certific8, VIN, participat, status, USERCDE, SAVEDATE, SAVETIME, INSTRUCTION)" & _
        " values (" & xESTIMATENO & _
        ", " & xApptNo & _
        ", '" & xTransType & _
        "', " & XNOTE & _
        ", " & VTXTREP_OR & _
        ", " & VTXTROType & _
        ", " & VTXTSvc_No & _
        ", " & VTXTAcct_No & _
        ", " & VTXTNiym & _
        ", " & VTXTPlate_No & _
        ", " & VcboModel & _
        ", " & VTXTTerm & _
        ", " & VTXTSektion & _
        ", " & VcboRecd_by & _
        ", " & VTXTKm_rdg & _
        ", " & VTXTDte_recd & _
        ", " & XDTE_PRO & _
        ", " & VTXTCertific8 & _
        ", " & VtxtVIN & _
        ", " & VTXTParticipat & _
        ", 'N', " & Vusercode & _
        ", " & VLastUpdate & _
        ", " & VLastUpdateTime & "," & XINST & ")"
    gconDMIS.Execute (SQL_STATEMENT)


    'NEW LOG AUDIT+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dim VTRANID                                        As String
    
        If labType(0).Caption = "Repair Order" Then
            VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "REP_OR", "CSMS_REPOR")
            Call NEW_LogAudit("A", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, "R", "RO NO: " & txtTranNo, "", "")
        ElseIf labType(0).Caption = "Estimate" Then
            VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "ESTIMATENO", "CSMS_REPOR")
            Call NEW_LogAudit("A", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, "E", "EST NO: " & txtTranNo, "", "")
        ElseIf labType(0).Caption = "Appointment" Then
            VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "APPTNO", "CSMS_REPOR")
            Call NEW_LogAudit("A", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, "A", "APPT NO: " & txtTranNo, "", "")
        End If
    'NEW LOG AUDIT+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    Call SaveROMonitoring
    Call SaveJobs
    Call SavePMSJObs

    If labType(0).Caption = "Estimate" Then
        SQL_STATEMENT = "insert into CSMS_ESTHD " & _
            "(UPLOAD_STATUS, SERVICE_ADVISER, ESTIMATENO, ApptNo, TransType, [note], rep_or, rotype, svc_no, acct_no, niym, plate_no, model, term, sektion, Recd_by, km_rdg, dte_recd, DTE_PRO, certific8, VIN, participat, status, USERCDE, SAVEDATE, SAVETIME, INSTRUCTION)" & _
            " values ('N', " & N2Str2Null(cboRecd_by) & ", " & xESTIMATENO & _
            ", " & xApptNo & _
            ", '" & xTransType & _
            "', " & XNOTE & _
            ", " & VTXTREP_OR & _
            ", " & VTXTROType & _
            ", " & VTXTSvc_No & _
            ", " & VTXTAcct_No & _
            ", " & VTXTNiym & _
            ", " & VTXTPlate_No & _
            ", " & VcboModel & _
            ", " & VTXTTerm & _
            ", " & VTXTSektion & _
            ", " & VcboRecd_by & _
            ", " & VTXTKm_rdg & _
            ", " & VTXTDte_recd & _
            ", " & XDTE_PRO & _
            ", " & VTXTCertific8 & _
            ", " & VtxtVIN & _
            ", " & VTXTParticipat & _
            ", 'N', " & Vusercode & _
            ", " & VLastUpdate & _
            ", " & VLastUpdateTime & "," & XINST & ")"
        gconDMIS.Execute (SQL_STATEMENT)
    
        Call SaveEstimate
    ElseIf labType(0).Caption = "Appointment" Then
        Call UpdateAppointmentSkid
    End If

    
    TOTJOBAMT = 0
    TOTJOBDISC = 0
    TOTJOBDISCVAL = 0
    TOTJOBTAX = 0
    JobComTotal = 0
    JobSalesTotal = 0
    JobWarTotal = 0

    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '1' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '1' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '1' order by LINE_NO asc")
    End If
    
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    
    TOTJOBAMT = Round(TOTJOBAMT, 2)
    TOTJOBDISC = Round(TOTJOBDISC, 2)
    TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2)
    TOTJOBTAX = Round(TOTJOBTAX, 2)
    TOTPARTSAMT = 0
    TOTPARTSDISC = 0
    TOTPARTSDISCVAL = 0
    TOTPARTSTAX = 0
    PartsComTotal = 0
    PartsSalesTotal = 0
    PartsWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '2' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '2' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '2' order by LINE_NO asc")
    End If
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2)
    TOTPARTSDISC = Round(TOTPARTSDISC, 2)
    TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2)
    TOTPARTSTAX = Round(TOTPARTSTAX, 2)

    TOTMATAMT = 0
    TOTMATDISC = 0
    TOTMATDISCVAL = 0
    TOTMATTAX = 0
    MatComTotal = 0
    MatSalesTotal = 0
    MatWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '3' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Estimate" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where EstimateNo = " & VTXTestimateno & " and livil = '3' order by LINE_NO asc")
    ElseIf labType(0).Caption = "Appointment" Then
        Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where APPTNO = " & xApptNo & " and livil = '3' order by LINE_NO asc")
    End If
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    
    TOTMATAMT = Round(TOTMATAMT, 2)
    TOTMATDISC = Round(TOTMATDISC, 2)
    TOTMATDISCVAL = Round(TOTMATDISCVAL, 2)
    TOTMATTAX = Round(TOTMATTAX, 2)
    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT

    Dim FilterType                                     As String
    Set rsRO_DET = New ADODB.Recordset

    If labType(0).Caption = "Repair Order" Then
        FilterType = "rep_or = " & VTXTREP_OR
    ElseIf labType(0).Caption = "Estimate" Then
        FilterType = "EstimateNo = " & VTXTestimateno
    ElseIf labType(0).Caption = "Appointment" Then
        FilterType = "APPTNO = " & xApptNo
    End If

    gconDMIS.Execute "update CSMS_RepOr set " & _
        " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & _
        ", l_amtvalue = " & Round(TOTJOBAMT, 2) & _
        ", l_disc = " & Round(TOTJOBDISCVAL, 2) & _
        ", l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & _
        ", l_taxval = " & Round(TOTJOBTAX, 2) & _
        ", l_discount = " & Round(TOTJOBDISC, 2) & _
        ", amount = " & Round(ROTotal - (TOTJOBDISC + TOTPARTSDISC + TOTMATDISC + TOTACCDISC), 2) & _
        ", rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & _
        ", wl_amt = " & 0 & _
        ", ro_amount = " & Round(ROTotal - (TOTJOBDISC + TOTPARTSDISC + TOTMATDISC + TOTACCDISC), 2) & _
        " where " & FilterType
   
    
    finalsave = True
    Exit Function
IVANEXEQUIELVALENCIA:
    finalsave = False
End Function
Sub UpdateAppointmentSkid()
    Dim xTranDate                                       As String
    Dim xApptTime                                       As String
    Dim xCuscde                                         As String
    Dim xCUSNAM                                         As String
    Dim xPLATE_NO                                       As String
    Dim xMODEL                                          As String
    Dim xMake                                           As String
    Dim XNOTE                                           As String
    Dim xKM_RDG                                         As Double
    
    xApptNo = N2Str2Null(txtTranNo)
    xCuscde = N2Str2Null(txtAcct_No)
    xCUSNAM = N2Str2Null(txtNiym)
    xPLATE_NO = N2Str2Null(txtPlate_No)
    xMODEL = N2Str2Null(cboModel)
    xMake = N2Str2Null(txtMake)
    xKM_RDG = NumericVal(txtKm_rdg)
    XNOTE = N2Str2Null(txtRecomendation)
    xApptTime = N2Str2Null(lblTime)
    xTranDate = N2Str2Null(lblDate)
    
    gconDMIS.Execute ("INSERT INTO CSMS_Appointment (APPTNO, CUSCDE, CUSNAM, PLATE_NO, MODEL, MAKE, KM_RDG, NOTE, TRANDATE, APPTTIME)" & _
        " VALUES(" & xApptNo & _
        ", " & xCuscde & _
        ", " & xCUSNAM & _
        ", " & xPLATE_NO & _
        ", " & xMODEL & _
        ", " & xMake & _
        ", " & xKM_RDG & _
        ", " & XNOTE & _
        ", " & xTranDate & _
        ", " & xApptTime & ")")
End Sub

Sub SaveJobs()
    'On Error GoTo error_code
    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim VLastUpdateTime                                 As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim VLastUpdate                                     As String
    Dim Vusercode                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim xFLATRATE                                       As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As String
    Dim xJobType                                        As String
    Dim X                                               As Long
    Dim BP_TYPE                                         As String
    Dim xrotype                                         As String
    Dim QUICK_SERVICE                                   As String
    Dim PMS_READING                                     As Long

    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

    xApptNo = N2Str2Null(txtTranNo)
    JOBLINE_NO = "0"
    Dim XITEMNO                                         As Integer
    X = 0

    gconDMIS.Execute "delete from CSMS_RO_Det where ApptNo = " & xApptNo & ""
    For X = 1 To lblJob4Service.ListItems.Count
        Dim IS_WARRANTY                                As String

        JOBREP_OR = N2Str2Null(txtRep_Or)
        JOBLEVEL = "'1'"
        XITEMNO = XITEMNO + 1
        JOBLINE_NO = N2Str2Null(Format(XITEMNO, "00"))
        JOBDETCDE = N2Str2Null(lblJob4Service.ListItems(X))
        xJobType = N2Str2Null(lblJob4Service.ListItems(X).SubItems(1))
        JOBDETDSC = Replace(N2Str2Null(Mid(lblJob4Service.ListItems(X).SubItems(2), 1, 500)), vbCrLf, " ")
        xFLATRATE = NumericVal(lblJob4Service.ListItems(X).SubItems(3))
        JOBDET_HRS = NumericVal(lblJob4Service.ListItems(X).SubItems(4))
        JOBDISCRATE = NumericVal(lblJob4Service.ListItems(X).SubItems(5)) / 100
        JOBWCODE = N2Str2Null(lblJob4Service.ListItems(X).SubItems(6))
        IS_WARRANTY = N2Str2Null(lblJob4Service.ListItems(X).SubItems(8))
        BP_TYPE = N2Str2Null(lblJob4Service.ListItems(X).SubItems(9))
        QUICK_SERVICE = N2Str2Null(lblJob4Service.ListItems(X).SubItems(10))
        PMS_READING = NumericVal(lblJob4Service.ListItems(X).SubItems(11))
        
        JOBDETUNT = "NULL"
        JOBDETVOL = NumericVal(0)
        JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
        JOBCODE = "NULL"
        JOBTAXRATE = (VAT_RATE / 100)
        JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
        If Left(lblJob4Service.ListItems(X).SubItems(1), 6) = "Starex" Then
            JOBPOCODE = "'PM'"
        Else
            JOBPOCODE = "NULL"
        End If
        JOBRep_Or2 = "NULL"
        JOBDETAIL = Replace(N2Str2Null(CheckChar(lblJob4Service.ListItems(X).SubItems(7))), vbCrLf, " ")
        JOBDET_AMT = JOBDETPRC
        JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
        JOBREMARKS = N2Str2Null(CheckChar(txtRecomendation.Text))
        JOBTECHNICIAN = "NULL"
        
        'COMMENT BY  : MJP
        'DESCRIPTION : DOUBLE VAT
            'JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
        'COMMENT BY  : MJP
        'UPDATE BY   : MJP
        'DESCRIPTION : DOUBLE VAT
            JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
        'UPDATE BY   : MJP
        
        Vusercode = "" & N2Str2Null(LOGCODE) & ""
        VLastUpdate = "'" & LOGDATE & "'"
        VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
        xApptNo = N2Str2Null(txtApointmentNo)
        xESTIMATENO = N2Str2Null(txtEstimateno)
        xrotype = N2Str2Null(Cbo_Rotype)
        If txtEstimateEdit <> "" Then
            xESTIMATENO = N2Str2Null(txtEstimateEdit)
        End If
        If txtAppointmentEdit <> "" Then
            xApptNo = N2Str2Null(txtAppointmentEdit)
        End If

        SQL_STATEMENT = "Insert Into CSMS_RO_Det " & _
            " (QUICK_SERVICE, TRANSTATUS, ESTIMATENO, JobType, TransType, ApptNo, FLATRATE, Rep_or, Livil, LINE_NO, Detcde, Detdsc, Technician, Det_hrs, Detunt, Detvol, Detprc, Detamt, Code, Wcode, Taxrate, Discrate, Taxval, Disval, Pocode, Rep_or2, Detail, Det_amt, Dis_val, Discount_2, USERCDE, SAVEDATE, SAVETIME, STATUS1,rotype, PMS_READING) " & _
            " values (" & QUICK_SERVICE & ", " & BP_TYPE & "," & xESTIMATENO & _
            "," & xJobType & ",'" & xTransType & _
            "'," & xApptNo & "," & xFLATRATE & _
            "," & JOBREP_OR & ", " & JOBLEVEL & _
            ", " & JOBLINE_NO & "," & JOBDETCDE & _
            "," & JOBDETDSC & "," & JOBTECHNICIAN & _
            "," & JOBDET_HRS & "," & JOBDETUNT & _
            ", " & JOBDETVOL & "," & JOBDETPRC & _
            ", " & JOBDETAMT & ", " & JOBCODE & _
            ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & _
            ", " & (JOBDISCRATE * 100) & "," & JOBTAXVAL & _
            ", " & JOBDISVAL & ", " & JOBPOCODE & _
            ", " & JOBRep_Or2 & ", " & JOBDETAIL & _
            ", " & JOBDET_AMT & ", " & JOBDIS_VAL & _
            ", " & JOBDISCOUNT_2 & ", " & Vusercode & _
            ", " & VLastUpdate & ", " & VLastUpdateTime & "," & IS_WARRANTY & "," & xrotype & ", " & PMS_READING & ")"
        gconDMIS.Execute (SQL_STATEMENT)
                
        'UPDATE BY   : MJP 09212000 1011AM
        'DESCRIPTION : TO SAVE ALSO IN THE CSMS_ESTDETAILS THE JOBS DETAILS
            If labType(0).Caption = "Estimate" Then
                SQL_STATEMENT = "INSERT INTO CSMS_ESTDETAILS " & _
                    " (QUICK_SERVICE, TRANSTATUS, ESTIMATENO, JobType, TransType, ApptNo, FLATRATE, Rep_or, Livil, LINE_NO, Detcde, Detdsc, Technician, Det_hrs, Detunt, Detvol, Detprc, Detamt, Code, Wcode, Taxrate, Discrate, Taxval, Disval, Pocode, Rep_or2, Detail, Det_amt, Dis_val, Discount_2, USERCDE, SAVEDATE, SAVETIME, STATUS1) " & _
                    " values (" & QUICK_SERVICE & ", " & BP_TYPE & "," & xESTIMATENO & _
                    "," & xJobType & ",'" & xTransType & _
                    "'," & xApptNo & "," & xFLATRATE & _
                    "," & JOBREP_OR & ", " & JOBLEVEL & _
                    ", " & JOBLINE_NO & "," & JOBDETCDE & _
                    "," & JOBDETDSC & "," & JOBTECHNICIAN & _
                    "," & JOBDET_HRS & "," & JOBDETUNT & _
                    ", " & JOBDETVOL & "," & JOBDETPRC & _
                    ", " & JOBDETAMT & ", " & JOBCODE & _
                    ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & _
                    ", " & (JOBDISCRATE * 100) & "," & JOBTAXVAL & _
                    ", " & JOBDISVAL & ", " & JOBPOCODE & _
                    ", " & JOBRep_Or2 & ", " & JOBDETAIL & _
                    ", " & JOBDET_AMT & ", " & JOBDIS_VAL & _
                    ", " & JOBDISCOUNT_2 & ", " & Vusercode & _
                    ", " & VLastUpdate & ", " & VLastUpdateTime & "," & IS_WARRANTY & ")"
                gconDMIS.Execute (SQL_STATEMENT)
            End If
        'UPDATE BY   : MJP 09212000 1011AM
        
        'NEW LOG AUDIT +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Dim MODENAME                                   As String
            Dim VTRANID                                    As String
            Dim VDETID                                     As String
            Dim VAPPTYPE                                   As String
            
            If labType(0).Caption = "Repair Order" Then
                MODENAME = "BILLING SYSTEM"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "REP_OR", "CSMS_REPOR")
                VAPPTYPE = "R"
            ElseIf labType(0).Caption = "Estimate" Then
                MODENAME = "JOB ESTIMATE"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "ESTIMATENO", "CSMS_REPOR")
                VAPPTYPE = "E"
            ElseIf labType(0).Caption = "Appointment" Then
                MODENAME = "APPOINTMENT"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "APPTNO", "CSMS_REPOR")
                VAPPTYPE = "A"
            End If
    
    
            If lblJob4Service.ListItems(X).SubItems(1) = "CND" Then
                Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, VAPPTYPE, "JOB CODE: " & Null2String(JOBDETCDE), LTrim(RTrim(lblJob4Service.ListItems(X).SubItems(1))), VDETID)
            ElseIf lblJob4Service.ListItems(X).SubItems(1) = "PMS" Then
                Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, VAPPTYPE, "JOB CODE: " & Null2String(JOBDETCDE), LTrim(RTrim(lblJob4Service.ListItems(X).SubItems(1))), VDETID)
            Else
                Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, VAPPTYPE, "JOB CODE: " & Null2String(JOBDETCDE), LTrim(RTrim(lblJob4Service.ListItems(X).SubItems(1))), VDETID)
            End If
        'NEW LOG AUDIT +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Next X
'error_code:
'    str_MSG = Replace(str_MSG, "@UTX83912839123", "Update Master file")
'    MsgBox str_MSG, vbCritical, "Creating RO Error"
'    gconDMIS.RollbackTrans
'    Screen.MousePointer = 0
'    ShowVBError
End Sub

Sub ComputeMe()
    tlHrs = 0: tlFR = 0
    For bevvy = 1 To Me.lblJob4Service.ListItems.Count
        tlHrs = tlHrs + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(4))
        tlFR = tlFR + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(3))
    Next bevvy
    
    lblStdHrs.Caption = tlHrs
    lbltlFaltRate.Caption = tlFR
    txtEstLabor = tlFR
    xPartsAmt = 0: xAcesAmt = 0: xMatAmt = 0

    For bevvy = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(bevvy).Text = "P" Then
            xPartsAmt = xPartsAmt + (NumericVal(ListView1.ListItems(bevvy).SubItems(3)) * NumericVal(ListView1.ListItems(bevvy).SubItems(4)))
        ElseIf ListView1.ListItems(bevvy).Text = "M" Then
            xMatAmt = xMatAmt + (NumericVal(ListView1.ListItems(bevvy).SubItems(3)) * NumericVal(ListView1.ListItems(bevvy).SubItems(4)))
        Else
            xAcesAmt = xAcesAmt + (NumericVal(ListView1.ListItems(bevvy).SubItems(3)) * NumericVal(ListView1.ListItems(bevvy).SubItems(4)))
        End If
    Next bevvy

    txtEstParts.Text = xPartsAmt
    Text2.Text = xMatAmt
    txtEstAces.Text = xAcesAmt
    txtTotalAmt.Text = Val(txtEstLabor) + Val(txtEstParts) + Val(Text2) + Val(txtEstAces)
End Sub

Sub ViewJobs()
    Dim RSUPLOAD                                       As New ADODB.Recordset
    
    lblJob4Service.Sorted = False: lblJob4Service.ListItems.Clear
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,det_hrs from CSMS_Ro_Det where REP_OR = '" & txtRep_Or & "' Order by det_hrs  desc")    '[LINE_NO]
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lblJob4Service.ListItems, RSUPLOAD
    End If

    tlHrs = 0: tlFR = 0
    For bevvy = 1 To Me.lblJob4Service.ListItems.Count
        tlHrs = tlHrs + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(3))
        tlFR = tlFR + NumericVal(lblJob4Service.ListItems(bevvy).SubItems(2))
    Next bevvy
    lblStdHrs.Caption = tlHrs
    lbltlFaltRate.Caption = tlFR
End Sub

Sub GetDefaultTransactionType()
    Set rsAddRepor = New ADODB.Recordset
    If labType(0).Caption = "Repair Order" Then
        xTransType = "R"
        rsAddRepor.Open "select id,rep_or from CSMS_RepOr where TransType = 'R' order by rep_or desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsAddRepor.EOF And Not rsAddRepor.BOF Then
            rsAddRepor.MoveFirst
            txtRep_Or.Text = Format(NumericVal(Mid$(rsAddRepor!REP_OR, 3, 8)) + 1, "R-00000000")
        Else
            txtRep_Or.Text = "R-00000001"
        End If
        labTranType.Caption = "Repair Order  No."
        txtTranNo.Text = txtRep_Or
        txtTranNo.Locked = False
    ElseIf labType(0).Caption = "Estimate" Then       '
        xTransType = "E"
        rsAddRepor.Open "select id,ESTIMATENO from CSMS_ESTHD WHERE ESTIMATENO IS NOT NULL order by ESTIMATENO desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not (rsAddRepor.EOF And rsAddRepor.BOF) Then
            rsAddRepor.MoveFirst
            txtEstimateno.Text = Format(NumericVal(Mid$(rsAddRepor!EstimateNo, 3, 8)) + 1, "E-00000000")
        Else
            txtEstimateno.Text = "E-00000001"
        End If
        labTranType.Caption = "Estimate  No."
        txtTranNo.Text = txtEstimateno
        txtTranNo.Locked = False
    ElseIf labType(0).Caption = "Appointment" Then
        xTransType = "A"
        labTranType.Caption = "Appointment  No."
        txtApointmentNo.Text = txtTranNo
        txtTranNo.Locked = False
    End If
End Sub

Sub GetVehicleforCustomer()
    txtPlate_No = "": cboModel = "": txtMake = "": txtYear = "": txtVIN = "": txtVehName = ""
    Dim rsVehicle                                      As New ADODB.Recordset

    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear
    Set rsVehicle = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where (Cuscde = '" & txtID & "' OR EndUser = '" & txtID & "')")
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsVehicle
        lstCustomer.Refresh

        Dim rsVehicleKo                                As New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where (Cuscde = '" & txtID & "' OR ENDUSER = '" & txtID & "') and plate_no = '" & lstVehicle.SelectedItem.SubItems(1) & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![PLATE_NO])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtYear = Null2String(rsVehicleKo![YER])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If
    End If
End Sub

Sub GetVehicleforEndUser(XXX As String)
    txtPlate_No = "": cboModel = "": txtMake = "": txtYear = "": txtVIN = "": txtVehName = ""
    Dim rsVehicle                                      As New ADODB.Recordset

    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear
    Set rsVehicle = gconDMIS.Execute("select model,Plate_no,serial,engine,prodno from CSMS_CusVeh where PLATE_NO = '" & XXX & "'")
    If Not (rsVehicle.EOF And rsVehicle.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsVehicle
        lstCustomer.Refresh

        Dim rsVehicleKo                                As New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where plate_no = '" & XXX & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            txtPlate_No = Null2String(rsVehicleKo![PLATE_NO])
            cboModel = Null2String(rsVehicleKo![Model])
            txtMake = Null2String(rsVehicleKo![Make])
            txtYear = Null2String(rsVehicleKo![YER])
            txtVIN = Null2String(rsVehicleKo![Vin])
            txtVehName = Trim(cboModel) & "   " & txtPlate_No
        End If

    End If
End Sub

'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------
'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
Sub CleanPLateInfoLabel()
    Dim X                                              As Integer

    For X = 0 To 9
        lblinfo(X).Caption = ""
    Next
End Sub
'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------

Sub FillGrid()
    Dim rsCustomer                                     As New ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    
    If chkSettingAll.Value = 1 Then
        Set rsCustomer = gconDMIS.Execute("select TOP 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer order by lastname asc")
    Else
        Set rsCustomer = gconDMIS.Execute("select top 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer order by lastname asc")
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
    If optEndUser.Value = True Then
        lstCustomer.ColumnHeaders(1).Text = "EndUser Name"
        lstCustomer.ColumnHeaders(2).Text = "Account Name"
        lstCustomer.ColumnHeaders(3).Text = "Plate No"
        lstCustomer.ColumnHeaders(4).Text = "Model"
        lstCustomer.ColumnHeaders(5).Text = "Description"
    Else
        lstCustomer.ColumnHeaders(1).Text = "Last Name"
        lstCustomer.ColumnHeaders(2).Text = "First Name"
        lstCustomer.ColumnHeaders(3).Text = "Address"
        lstCustomer.ColumnHeaders(4).Text = "Province"
        lstCustomer.ColumnHeaders(5).Text = "Phone No."
    End If

End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer                                     As New ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    XXX = Repleys(LTrim(RTrim(XXX)))
    
    If chkSettingAll.Value = 1 Then
        If optLN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where lastname like'" & XXX & "%' order by lastname asc")
        ElseIf optFN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where firstname like'" & XXX & "%' order by firstname asc")
        ElseIf optFullName.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select TOP 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where AcctName like'" & XXX & "%' order by AcctName asc")
        ElseIf optEndUser.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select CSMS_CUSVEH.ENDUSER,CSMS_CUSVEH.CUSCDE,CSMS_CusVeh.Plate_No,CSMS_CusVeh.Model,CSMS_CusVeh.Description,CSMS_CusVeh.CusCde,All_Customer.CustomerAdd from ALL_Customer inner join CSMS_CusVeh on ALL_Customer.CusCde = CSMS_CusVeh.EndUser where (ALL_Customer.LastName like '" & XXX & "%' OR ALL_Customer.FirstName like '" & XXX & "%') order by ALL_Customer.AcctName asc")
            lstCustomer.ColumnHeaders(1).Text = "EndUser Name"
            lstCustomer.ColumnHeaders(2).Text = "Account Name"
            lstCustomer.ColumnHeaders(3).Text = "Plate No"
            lstCustomer.ColumnHeaders(4).Text = "Model"
            lstCustomer.ColumnHeaders(5).Text = "Description"
        End If
    Else
        If optLN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select top 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where lastname like'" & XXX & "%' order by lastname asc")
        ElseIf optFN.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select top 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where firstname like'" & XXX & "%' order by firstname asc")
        ElseIf optFullName.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select top 100 lastname,firstname,CustomerAdd,ProvincialAdd,TelephoneNo,AcctName,CusCde from ALL_Customer where AcctName like'" & XXX & "%' order by AcctName asc")
        ElseIf optEndUser.Value = True Then
            Set rsCustomer = gconDMIS.Execute("select top 100 CSMS_CUSVEH.ENDUSER,CSMS_CUSVEH.CUSCDE,CSMS_CusVeh.Plate_No,CSMS_CusVeh.Model,CSMS_CusVeh.Description,CSMS_CusVeh.CusCde,All_Customer.CustomerAdd from ALL_Customer inner join CSMS_CusVeh on ALL_Customer.CusCde = CSMS_CusVeh.EndUser where (ALL_Customer.LastName like '" & XXX & "%' OR ALL_Customer.FirstName like '" & XXX & "%') order by ALL_Customer.AcctName asc")
            lstCustomer.ColumnHeaders(1).Text = "EndUser Name"
            lstCustomer.ColumnHeaders(2).Text = "Account Name"
            lstCustomer.ColumnHeaders(3).Text = "Plate No"
            lstCustomer.ColumnHeaders(4).Text = "Model"
            lstCustomer.ColumnHeaders(5).Text = "Description"
        End If
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        If optEndUser.Value = True Then
            rsCustomer.MoveFirst
            Do While Not rsCustomer.EOF
                With lstCustomer
                    .Sorted = False
                    .ListItems.Add , , SetEndUserName(Null2String(rsCustomer![ENDUSER]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , SetEndACCTName(Null2String(rsCustomer![CUSCDE]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , Null2String(rsCustomer![PLATE_NO])
                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , Null2String(rsCustomer![Model])
                    .ListItems(.ListItems.Count).ListSubItems.Add 4, , Null2String(rsCustomer![Description])
                    .ListItems(.ListItems.Count).ListSubItems.Add 5, , SetEndACCTName(Null2String(rsCustomer![CUSCDE]))
                    .ListItems(.ListItems.Count).ListSubItems.Add 6, , Null2String(rsCustomer![CUSCDE])
                End With
                rsCustomer.MoveNext
            Loop
        Else
            Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
            lstCustomer.Refresh
        End If
    End If
End Sub

Sub SavePMSJObs()
    'On Error GoTo error_code
    Dim X                                               As Long
    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim VLastUpdateTime                                 As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim VLastUpdate                                     As String
    Dim Vusercode                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim xFLATRATE                                       As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As String
    Dim xJobType                                        As String
    Dim xPMD_Model                                      As String
    
    JOBDISVAL = 0
    JOBTAXVAL = 0
    JOBDETAMT = 0
    JOBDIS_VAL = 0
    JOBDISCOUNT_2 = 0
    JOBDISCRATE = 0
    xApptNo = "NULL"
    JOBLINE_NO = "0"
    
    For X = 1 To lstPMSDet.ListItems.Count
        JOBREP_OR = N2Str2Null(txtRep_Or)
        JOBLEVEL = "'1'"
        JOBLINE_NO = N2Str2Null(Format(Val(JOBLINE_NO) + 1, "00"))
        JOBDETCDE = N2Str2Null(lstPMSDet.ListItems(X))
        xJobType = N2Str2Null(lstPMSDet.ListItems(X).SubItems(1))
        JOBDETDSC = N2Str2Null(Mid(lstPMSDet.ListItems(X).SubItems(2), 1, 500))
        JOBDETUNT = "NULL"
        JOBDETVOL = NumericVal(0)
        JOBDET_HRS = NumericVal(lbltlFaltRate)
        xFLATRATE = NumericVal(lblStdHrs)
        JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
        JOBCODE = "NULL"
        JOBWCODE = "NULL"
        JOBTAXRATE = (VAT_RATE / 100)
        JOBDISCRATE = NumericVal(0)
        JOBDETAMT = Round(JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE), 2)
        JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
        JOBPOCODE = "NULL"
        JOBRep_Or2 = "NULL"
        JOBDETAIL = "NULL"
        JOBDET_AMT = JOBDETPRC
        JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
        JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
        JOBREMARKS = "NULL"
        JOBTECHNICIAN = "NULL"
        JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
        Vusercode = "" & N2Str2Null(LOGCODE) & ""
        VLastUpdate = "'" & LOGDATE & "'"
        VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
        xApptNo = N2Str2Null(txtApointmentNo)
        xESTIMATENO = N2Str2Null(txtEstimateno)
        xPMD_Model = N2Str2Null(N2Str2Null(lstPMSDet.ListItems(X).SubItems(3)))

        If txtEstimateEdit <> "" Then
            xESTIMATENO = N2Str2Null(txtEstimateEdit)
        End If
        If txtAppointmentEdit <> "" Then
            xApptNo = N2Str2Null(txtAppointmentEdit)
        End If

        SQL_STATEMENT = "insert into CSMS_PMS_Job_Det " & _
            "(PMS_Model, ESTIMATENO, ApptNo, JobType, TransType, rep_or, LINE_NO, detcde, detdsc)" & _
            " values (" & xPMD_Model & _
            ", " & xESTIMATENO & _
            ", " & xApptNo & _
            ", " & xJobType & _
            ", '" & xTransType & _
            "', " & JOBREP_OR & _
            ", " & JOBLINE_NO & _
            ", " & JOBDETCDE & _
            ", " & JOBDETDSC & ")"
        gconDMIS.Execute (SQL_STATEMENT)

        'NEW LOG AUDIT+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Dim VTRANID                                    As String
            Dim VDETID                                     As String
            Dim VAPPTYPE                                   As String
    
            If labType(0).Caption = "Repair Order" Then
                VAPPTYPE = "R"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "REP_OR", "CSMS_REPOR")
            ElseIf labType(0).Caption = "Estimate" Then
                VAPPTYPE = "E"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "ESTIMATENO", "CSMS_REPOR")
            ElseIf labType(0).Caption = "Appointment" Then
                VAPPTYPE = "A"
                VTRANID = FindTransactionID(N2Str2Null(txtTranNo), "APPTNO", "CSMS_REPOR")
            End If
            Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, VAPPTYPE, "JOB CODE: " & LTrim(RTrim(lstPMSDet.ListItems(X).Text)), "PMS", "")
        'NEW LOG AUDIT+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Next X
'error_code:
'    str_MSG = Replace(str_MSG, "@UTX83912839123", "Update Master file")
'    MsgBox str_MSG, vbCritical, "Creating RO Error"
'    gconDMIS.RollbackTrans
'    Screen.MousePointer = 0
'    ShowVBError
'    Exit Sub
End Sub

Sub TrigerTheRefresh()
   frmCSMS_ServiceCounter.cmdRefresh.Value = True

End Sub

Private Sub Cbo_Rotype_Change()
    ShowROTYPEdescription
End Sub

Private Sub Cbo_Rotype_Click()
    ShowROTYPEdescription
End Sub

Private Sub cboRecd_by_LostFocus()
    Dim i
    For i = 0 To cboRecd_by.ListCount
        If UCase(cboRecd_by) = UCase(cboRecd_by.List(i)) Then
            Exit Sub
        End If
    Next
    MsgBox "Please Select Proper Service Advisor From the List", vbInformation
    cboRecd_by = ""
    cboRecd_by.SetFocus
End Sub

Private Sub cboSearchBy_Click()
    lblPlateFind.Caption = "NO"
    txtSPlate.Text = ""
    lblinfo(9).Caption = ""
    lblinfo(0).Caption = ""
    txtCustName.Text = ""
    txtNiym.Text = ""
    lblinfo(1).Caption = ""
    txtAddress.Text = ""
    lblinfo(2).Caption = ""
    txtID.Text = ""
    txtAcct_No.Text = ""
    txtYear.Text = ""
    txtMake.Text = ""
    cboModel.Text = ""
    txtPlate_No.Text = ""
    txtVIN.Text = ""
    lblinfo(3).Caption = ""
    lblinfo(4).Caption = ""
    lblinfo(5).Caption = ""
    lblinfo(6).Caption = ""
    lblinfo(7).Caption = ""
    lblinfo(8).Caption = ""
    lblinfo(10).Caption = ""
    lblinfo(11).Caption = ""
End Sub

Private Sub cboSModel_Change()
    Call FillModelStock
End Sub

Sub FillModelStock()
    If chkModel.Value = 1 Then
        Dim rstmp As New ADODB.Recordset
        Set rstmp = gconDMIS.Execute("SELECT TOP 100 STOCKNO, STOCKDESC, ISNULL(SRP,0) AS SRP, " & _
                " ISNULL(MODELCODE,'') AS MODELCODE ,ISNULL(GENUINE,'N') GENUINE, " & _
                " CASE WHEN ISNULL(ONHAND,0) <= 0 THEN 'N' " & _
                " Else 'Y' END AS ONHAND FROM PMIS_STOCKMAS WHERE TYPE = " & N2Str2Null(lblTYPE) & " AND MODELCODE LIKE '" & Left(cboSModel, 2) & "%'")
        Call Listview_Loadval(lsvDET.ListItems, rstmp)
    End If
End Sub

Private Sub cboSModel_Click()
    Call FillModelStock
End Sub

Private Sub cboSModel_LostFocus()
    Call FillModelStock
End Sub

Private Sub chkSettingAll_Click()
    Call SaveSetting("DMIS 2.0", "CSMS", "SHOW ALL CUSTOMER SEARCH", chkSettingAll.Value)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdAddAcc_Click()
    picDET.Visible = True
    picDET.ZOrder 0
    picMaster.Enabled = False
    Text8.Text = "a":    Text8.Text = ""
    lblTYPE.Caption = "A"
    On Error Resume Next
    Text8.SetFocus
End Sub

Private Sub cmdAddMat_Click()
    picDET.Visible = True
    picDET.ZOrder 0
    picMaster.Enabled = False
    Text8.Text = "a":    Text8.Text = ""
    lblTYPE.Caption = "M"
    On Error Resume Next
    Text8.SetFocus
End Sub

Private Sub cmdAddParts_Click()
    picDET.Visible = True
    picDET.ZOrder 0
    picMaster.Enabled = False
    Text8.Text = "a":    Text8.Text = ""
    lblTYPE.Caption = "P"
    On Error Resume Next
    Text8.SetFocus
End Sub

Private Sub cmdOK_Click()
    If NumericVal(txtEQTY) = 0 Then
        MsgBox "Invalid Qty", vbInformation, "DMIS"
        txtEQTY.SetFocus
        Exit Sub
    End If
    
    Dim Item As ListItem
    
    Set Item = ListView1.ListItems.Add(, , lblTYPE)
    Item.SubItems(1) = lblRES(0)
    Item.SubItems(2) = lblRES(1)
    Item.SubItems(3) = NumericVal(txtEQTY)
    Item.SubItems(4) = lblRES(2)
        
    Call Command4_Click
    Call COMPUTETOTALITEM
End Sub

Private Sub cmdOther_Click()
    frmMain.MousePointer = 11

    With frmCSMSOtherJobs
        .txtCustomer = txtNiym
        .txtActNo = txtAcct_No
        .txtROno = txtRep_Or
        .txtAppt = "NewAppt"
        .txtCheckMe = ""
        .txtCheckMe = "ro"
        .txtVehicle = cboModel.Text
    End With
    frmCSMSOtherJobs.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "CHANGE SERVICE ADVISER", "SYSTEM") = False Then Exit Sub
    cboRecd_by.Enabled = True
End Sub

Private Sub Command3_Click()
    picMaster.Enabled = True
    picDET.Visible = False
    picDET.ZOrder 1
End Sub

Private Sub Command4_Click()
    Text8.Enabled = True
    lsvDET.Enabled = True
    Command3.Enabled = True
    
    picADD.ZOrder 1
    picADD.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If shpCustomer.Visible = True Then
        If KeyCode = vbKeyF3 Then textSearch.SetFocus
    End If
End Sub

Private Sub Image1_Click()
'txtTranNo.Enabled = True
'txtTranNo.Locked = False
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub

    On Error Resume Next
    If MsgBox("Delete this parts " & ListView1.SelectedItem.SubItems(2) & "", vbYesNo + vbQuestion + vbDefaultButton1, "Are You Sure") = vbNo Then Exit Sub

    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
    Call COMPUTETOTALITEM
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If ListView1.ListItems.Count = 0 Then Exit Sub
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optEndUser.Value = True Then
        txtCustName = UCase(lstCustomer.SelectedItem.SubItems(5))
        txtNiym = lstCustomer.SelectedItem.SubItems(5)
        txtID = lstCustomer.SelectedItem.SubItems(6)
        txtAddress = lstCustomer.SelectedItem.SubItems(2)
        txtAcct_No = lstCustomer.SelectedItem.SubItems(6)
        GetVehicleforEndUser lstCustomer.SelectedItem.SubItems(2)
    Else
        txtCustName = UCase(lstCustomer.SelectedItem.SubItems(5))
        txtNiym = lstCustomer.SelectedItem.SubItems(5)
        txtID = lstCustomer.SelectedItem.SubItems(6)
        txtAddress = lstCustomer.SelectedItem.SubItems(2)
        txtAcct_No = lstCustomer.SelectedItem.SubItems(6)
        GetVehicleforCustomer
    End If
End Sub

Private Sub cmdAddeditCustomer_Click()
    frmAllCustomer.ZOrder 0
    Me.Hide
    frmAllCustomer.Show
End Sub

Private Sub cmdAddJobs_Click()
    frmMain.MousePointer = 11

    With frmCSMSReqJobs
        .txtCustomer.Text = txtNiym.Text
        .txtActNo.Text = txtAcct_No.Enabled
        .txtROno.Text = txtRep_Or.Text
        .txtAppt.Text = "NewAppt"
        .txtCheckMe.Text = ""
        .txtCheckMe.Text = "ro"
    End With
    frmCSMSReqJobs.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub cmdAddVeh_Click()
    If optEndUser.Value = True Then
        With frmCSMSAddVehicleEndUser
            .CustomerCode = txtPlate_No.Text
            .labCustCode.Caption = txtAcct_No.Text
            .labCustomer.Caption = txtCustName
        End With
        frmCSMSAddVehicleEndUser.Show 1
    Else
        With frmCSMSAddVehicle
            .CustomerCode = txtID
            .labCustCode.Caption = txtAcct_No.Text
            .labCustomer.Caption = txtCustName
        End With
        frmCSMSAddVehicle.Show 1
        GetVehicleforCustomer
    End If
End Sub

Private Sub cmdBack_Click()
    If shpVehicle.Visible = True Then
        shpVehicle.Visible = False
        shpCustomer.Visible = True
        Frame3.Visible = True
        If lstCustomer.Enabled = True And lstCustomer.ListItems.Count > 0 Then
            lstCustomer.SetFocus
        End If
        cmdAddeditCustomer.Visible = True
        picVehicle.Visible = False
    ElseIf shpRO.Visible = True Then
        'UPDATE BY : MJP 09-12-2007 02:28 AM ------------------------------------------------------
        'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SERCH BY PLATE.
            If optPlate.Value = True Then                 'SEARCH BY PLATE
                ShpAppointment.Visible = False
                picAppointment.Visible = False
                shpRO.Visible = False
                shpCustomer.Visible = True
                Frame3.Visible = True
                Exit Sub
            End If
        'UPDATE BY : MJP 09-12-2007 02:28 AM ------------------------------------------------------

        picVehicle.Visible = True
        picAppointment.Visible = False
        shpRO.Visible = False
        shpVehicle.Visible = True
        If lstVehicle.ListItems.Count > 0 And lstVehicle.Enabled = True Then
            lstVehicle.SetFocus
        End If
    ElseIf ShpAppointment.Visible = True Then         '
        ShpAppointment.Visible = False
        picVehicle.Visible = True
        picAppointment.Visible = False
        shpRO.Visible = False
        shpVehicle.Visible = True
        If lstVehicle.Enabled = True And lstVehicle.ListItems.Count > 0 Then
            lstVehicle.SetFocus
        End If
    ElseIf shpJobs.Visible = True Then
        If labType(0).Caption = "Appointment" Then
            ShpAppointment.Visible = True
            picReason.Visible = False
            shpJobs.Visible = False
            picAppointment.Visible = True
            cmdNext.Caption = "&Next  >>"
            Label6(4).Caption = "Appointment Information"
            SCCap.Caption = "Appointment Information"
        Else
            picReason.Visible = False
            shpJobs.Visible = False
            picAppointment.Visible = True
            shpRO.Visible = True
            cmdNext.Caption = "&Next  >>"
            Label6(4).Caption = "Repair Order Information"
            SCCap.Caption = "Repair Order Information"
        End If
    ElseIf ShpEstimate.Visible = True Then
        shpJobs.Visible = True
        ShpEstimate.Visible = False
        picEstimate.Visible = False
        cmdNext.Caption = "&Next  >>"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    If MsgBox("Delete this job " & lblJob4Service.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirm") = vbNo Then
        Exit Sub
    End If

    Screen.MousePointer = 11
    RemovePMSJobDetails lblJob4Service.SelectedItem.Text

    Me.lblJob4Service.ListItems.Remove Me.lblJob4Service.SelectedItem.Index

    Call ComputeResultOfRatenTimeWhenJobDelete

    Screen.MousePointer = 0
End Sub

Private Sub cmdEdit_Click()
    With frmCSMSJobSelected
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
        .cboJobChargeTo.Clear
        .cboJobChargeTo.AddItem "W"
        .cboJobChargeTo.AddItem "S"
        .cboJobChargeTo.AddItem "C"
        
        Dim RSUPLOAD                                   As New ADODB.Recordset
        Set RSUPLOAD = gconDMIS.Execute("Select * from CSMS_Ro_Det where REP_OR = '" & txtRep_Or & "' and detcde = '" & lblJob4Service.SelectedItem & "'")
        If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
            .txtCustomer = txtNiym
            .txtROno = txtRep_Or
            .txtJobCat = GetJobCat(RSUPLOAD![DETCDE])
            .txtJobDesc = Null2String(RSUPLOAD![DETDSC])
            .txtjCode = Null2String(RSUPLOAD![DETCDE])
            .txtflatrate = NumericVal(RSUPLOAD![DetPrc])
            .txtstdrate = NumericVal(RSUPLOAD![DET_HRS])
            .txtNote = Null2String(RSUPLOAD![Detail])
            .cboJobChargeTo = Null2String(RSUPLOAD![wCode])
            .txtJobDiscount = Null2String(RSUPLOAD![discrate])
            .txtSaveorEdit = "Edit"
            If IsBodyOrSublet(Trim(RSUPLOAD![DETCDE])) = True Then
                .txtDetCost.Visible = True
                .labDetCost.Visible = True
            Else
                .txtDetCost.Visible = False
                .labDetCost.Visible = False
            End If
        End If
    End With
    frmCSMSJobSelected.Show 1
End Sub

'UPDATE BY : MJP 09-12-2007 02:17 AM----------------------------------------------------------------
'DESCRIPTION : REQUEST OF SIR ARIEL OCAMPO SEARCH BY PLATE NO.
Private Sub cmdSPlate_Click()
    Dim rstmp                                          As New ADODB.Recordset
    Dim rsCUS                                          As New ADODB.Recordset
    Dim SPLATENO                                       As String

    Dim ToSearch                                       As String
    If cboSearchBy.Text = "Plate No." Then ToSearch = "Plate_No"
    If cboSearchBy.Text = "CS No." Then ToSearch = "VCOND_NO"
    If cboSearchBy.Text = "VIN" Then ToSearch = "VIN"
    If Not txtSPlate.Text = "" Then
        SPLATENO = N2Str2Null(txtSPlate.Text)
        Set rstmp = gconDMIS.Execute("Select * From CSMS_CusVeh Where " & ToSearch & " = " & SPLATENO & "")
        If Not (rstmp.BOF And rstmp.EOF) Then
            lblPlateFind.Caption = "YES"
            lblinfo(9).Caption = Null2String(rstmp!CUSCDE)
            Set rsCUS = gconDMIS.Execute("Select LastName,FirstName,CustomerAdd,ISNULL(ProvincialAdd ,'')   + ISNULL(CITY,'')  ,TelephoneNo,HomePhone,CusCde From All_Customer_Table Where CUSCDE = '" & lblinfo(9).Caption & "'")
            If Not (rsCUS.BOF And rsCUS.EOF) Then
                lblinfo(0).Caption = Null2String(rsCUS!lastname) & "," & Null2String(rsCUS!Firstname)
                txtCustName.Text = lblinfo(0).Caption
                txtNiym.Text = lblinfo(0).Caption

                lblinfo(1).Caption = Null2String(rsCUS!CUSTOMERADD)
                txtAddress.Text = lblinfo(1).Caption

                lblinfo(2).Caption = Null2String(rsCUS!HomePhone) & "/" & Null2String(rsCUS!TelephoneNo)

                txtID.Text = lblinfo(9).Caption
                txtAcct_No.Text = lblinfo(9).Caption
            End If
            txtYear.Text = Null2String(rstmp!YER)
            txtMake.Text = Null2String(rstmp!Make)
            cboModel.Text = Null2String(rstmp!Model)
            txtPlate_No.Text = Null2String(rstmp!PLATE_NO)
            txtVIN.Text = Null2String(rstmp!Vin)

            lblinfo(3).Caption = Null2String(rstmp!YER) & " " & Null2String(rstmp!Make) & " " & Null2String(rstmp!Model)
            lblinfo(4).Caption = Null2String(rstmp!Vin)
            lblinfo(5).Caption = Null2String(rstmp!Engine)
            lblinfo(6).Caption = GetColor(Null2String(rstmp!ClrCde))
            lblinfo(7).Caption = GetSellingDealer(Null2String(rstmp!Selling_Dealer))
            lblinfo(8).Caption = Null2String(rstmp!KMReading)
            lblinfo(10).Caption = Null2String(rstmp!PLATE_NO)
            lblinfo(11).Caption = Null2String(rstmp!VCOND_NO)
        Else
            lblPlateFind.Caption = "NO"

            lblinfo(9).Caption = ""
            lblinfo(0).Caption = ""
            txtCustName.Text = ""
            txtNiym.Text = ""

            lblinfo(1).Caption = ""
            txtAddress.Text = ""

            lblinfo(2).Caption = ""

            txtID.Text = ""
            txtAcct_No.Text = ""
            txtYear.Text = ""
            txtMake.Text = ""
            cboModel.Text = ""
            txtPlate_No.Text = ""
            txtVIN.Text = ""

            lblinfo(3).Caption = ""
            lblinfo(4).Caption = ""
            lblinfo(5).Caption = ""
            lblinfo(6).Caption = ""
            lblinfo(7).Caption = ""
            lblinfo(8).Caption = ""
            lblinfo(10).Caption = ""
            lblinfo(11).Caption = ""

            MsgBox cboSearchBy.Text & " not found", vbInformation, "Search " & cboSearchBy.Text
            txtSPlate.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Enter a " & cboSearchBy.Text & " to Search", vbInformation, "PLate No. Search"
        txtSPlate.SetFocus
    End If
    Set rstmp = Nothing
End Sub
'UPDATE BY : MJP 09-12-2007 02:17 AM----------------------------------------------------------------

Private Sub cmdNext_Click()
    Dim Flag_mode                                      As Boolean
    Flag_mode = False

    If shpCustomer.Visible = True Then
        If txtCustName.Text = "" Then
            MsgBox "Choose a customer first", vbInformation, "Info."
            Exit Sub
        End If
        'UPDATE BY: MJP 09-12-2007 01:42 AM--------------------------------------------------------
        'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
            If optPlate.Value = True Then
                If lblPlateFind.Caption = "YES" Then
                    Frame3.Visible = False
                    picAppointment.Visible = True
                    shpCustomer.Visible = False
    
                    If labType(0).Caption = "Appointment" Then
                        ShpAppointment.Visible = True
                    Else
                        shpRO.Visible = True
                    End If
    
                    picAppointment.Visible = True
                    txtKm_rdg.SetFocus
                    Exit Sub
                Else
                    MsgBox "Search First A Plate no. Before Proceed", vbInformation, "Search PLate No."
                    txtSPlate.SetFocus
                    Exit Sub
                End If
            End If
        'UPDATE BY: MJP 09-12-2007 01:42 AM--------------------------------------------------------

        shpCustomer.Visible = False
        shpVehicle.Visible = True
        picVehicle.Visible = True
        Frame3.Visible = False
        cmdAddeditCustomer.Visible = False
        If lstVehicle.ListItems.Count > 0 And lstVehicle.Enabled = True Then
            lstVehicle.SetFocus
        End If
    ElseIf shpVehicle.Visible = True Then
        If txtVehName = "" And xTransType <> "A" Then
            MsgBox "This vehiclce has no PLATE NO." & vbCrLf & "Please input the PLATE NO. or CONDUCTION STICKER NO.", vbInformation
            Exit Sub
        End If

        If xTransType = "E" Then
            'do nothing
        Else

            Dim rsVehicle                              As New ADODB.Recordset
            Set rsVehicle = gconDMIS.Execute("SELECT Status,ro_no from CSMS_RepairOrder where TransType = '" & xTransType & "' AND PLATE_No = '" & Trim(txtPlate_No.Text) & "' AND UPPER(STATUS) <> 'RELEASED'")
            If IsNull(rsVehicle!RO_NO) Then
                Flag_mode = True
            End If

            If xTransType = "R" Then
                If Flag_mode = False Then
                    If Not rsVehicle.BOF And Not rsVehicle.EOF Then
                        If MsgBox("Repair Order is already open. Continue anyway?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                            frmPlateNumberVerification.Show 1
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        If labType(0).Caption = "Appointment" Then
 'Updated by:   IEBV 06282010 1005AM
 'Description:  To disable selecting ROtype in appointment for the HCI
            If COMPANY_CODE = "HCI" Then
                Cbo_Rotype.Visible = False
                lbl_rotype.Visible = False
                lbl_rodescription(1).Visible = False
            End If
            If COMPANY_CODE = "HCI" Then
                If Cbo_Rotype.Text = "" Then
                    Cbo_Rotype.ListIndex = 0
                End If
            End If
 'Updated by:   IEBV 06282010 1005AM
 'Description:  To disable selecting ROtype in appointment for the HCI
            shpVehicle.Visible = False
            picVehicle.Visible = False
            picAppointment.Visible = True
            ShpAppointment.Visible = True
            Label6(4).Caption = "Appointment Information"
            SCCap.Caption = "Appointment Information"
        Else
            shpVehicle.Visible = False
            picVehicle.Visible = False
            picAppointment.Visible = True
            shpRO.Visible = True
            Label6(4).Caption = "Repair Order Information"
            SCCap.Caption = "Repair Order Information"
        End If

        On Error Resume Next
        txtKm_rdg.SetFocus
    ElseIf ShpAppointment.Visible = True Then
        If cboRecd_by = "" Then
            MsgBox "Please select Service Advisor assigned...", vbInformation, "Select"
            Exit Sub
        End If

        If labType(0).Caption = "Repair Order" Then
            If txtKm_rdg.Text = "" Then
                ShowIsRequiredMsg "KM Reading Cannot be Blank"
                txtKm_rdg.SetFocus
                Exit Sub
            End If
        End If

        ShpAppointment.Visible = False
        picAppointment.Visible = False
        shpRO.Visible = False
        picReason.Visible = True
        shpJobs.Visible = True
        cmdNext.Caption = "Finish"
    ElseIf shpRO.Visible = True Then                  '
        If cboRecd_by = "" Then
            MsgBox "Please select Service Adviser assigned...", vbInformation, "Select"
            cboRecd_by.SetFocus
            Exit Sub
        End If

        If cboRecd_by.ListIndex = -1 Then
            MsgBox "Please select Service Adviser From The List Provided...", vbInformation, "Select"
            cboRecd_by.SetFocus
            Exit Sub
        End If

        picAppointment.Visible = False
        shpRO.Visible = False
        picReason.Visible = True
        shpJobs.Visible = True
        If labType(0).Caption <> "Estimate" Then
            cmdNext.Caption = "Finish"
        End If
    ElseIf shpJobs.Visible = True Then
        If labType(0).Caption = "Estimate" Then
            shpJobs.Visible = False
            ShpEstimate.Visible = True
            picEstimate.Visible = True
            picEstimate.ZOrder 0
            cmdNext.Caption = "Finish"
        Else
            Call SaveAllInfo
        End If
    ElseIf ShpEstimate.Visible = True Then
        Call SaveAllInfo
    End If
End Sub

Private Sub cmdPMS_Click()
    Screen.MousePointer = 11

    frmCSMSPMS.txtCHECK.Text = "AddJobs"
    frmCSMSPMS.Show 1

    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    frmMain.MousePointer = 11

    frmCSMSGetCannedLabor.txtCheckMe = "ro"
    frmCSMSGetCannedLabor.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Call ComputeMe
End Sub

Private Sub Form_Load()
    picVehicle.Top = 990
    picVehicle.Left = 2310
    picAppointment.Top = 990
    picAppointment.Left = 2310
    picReason.Top = 990
    picReason.Left = 2310
    picEstimate.Top = 990
    picEstimate.Left = 2310
    optLN.Value = True

    dtPromised.Value = DateValue(Now) & " " & TimeValue(Now)
    txtDte_recd.Value = Format(Now(), "MM/dd/yyyy")
    txtRecorded.Text = Format(Now(), "MM/dd/yyyy")
    Dim CTL                                            As Control
    With frmCSMSNewAppointment
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
    End With
    lblJob4Service.Sorted = False: lblJob4Service.ListItems.Clear
    
    Dim rsEmpNo                                        As New ADODB.Recordset
    Set rsEmpNo = gconDMIS.Execute("SELECT (LASTNAME + ', ' + FIRSTNAME + ' ' + LEFT(MIDDLENAME,1) + '.') AS NAYM " & _
        " FROM HRMS_EMPINFO WHERE IS_SERVICE_ADVISER = '1' AND ACTIVEINACTIVE = 'A' AND RESIGNED IS NULL")
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        cboRecd_by.Clear
        Do While Not rsEmpNo.EOF
            cboRecd_by.AddItem Null2String(rsEmpNo!NAYM)
            rsEmpNo.MoveNext
        Loop
    End If
 'Updated by:   IEBV 06282010 1005AM
 'Description:  To enable selecting ROtype for HCI
    If COMPANY_CODE = "HCI" Then
        Cbo_Rotype.Visible = True
        lbl_rotype.Visible = True
        lbl_rodescription(1).Visible = True
        Call ADDROTYPE
    End If
 'Updated by:   IEBV 06282010 1005AM
 'Description:  To enable selecting ROtype for HCI
    
    'UPDATE BY   : MJP 11092009 0134PM
    'DESCRIPTION : CRF 108
'        If COMPANY_CODE = "HGC" Then
'            Dim RSDSA                                       As New ADODB.Recordset
'            Dim RSHRMS                                      As New ADODB.Recordset
'            Set RSDSA = gconDMIS.Execute("SELECT EMPNO FROM ALL_RAMS_USERS WHERE USERCODE = " & N2Str2Null(LTrim(RTrim(LOGCODE))) & "")
'            If Not (RSDSA.BOF And RSDSA.EOF) Then
'                Set RSHRMS = gconDMIS.Execute("SELECT (LASTNAME + ', ' + FIRSTNAME + ' ' + LEFT(MIDDLENAME,1) + '.') AS NAYM FROM HRMS_EMPINFO WHERE EMPNO = '" & Null2String(RSDSA!EMPNO) & "'")
'                If Not (RSHRMS.EOF And RSHRMS.BOF) Then
'                    cboRecd_by.Text = Null2String(RSHRMS!NAYM)
'                End If
'            End If
'            Set RSDSA = Nothing
'            cboRecd_by.Enabled = False
'        Else
'            cboRecd_by.Enabled = True
'        End If
    'UPDATE BY   : MJP 11092009 0134PM
    
    '******************************************************************
        Dim RSMODEL As New ADODB.Recordset
        Set RSMODEL = gconDMIS.Execute("SELECT LTRIM(RTRIM(CODE)) + ' - '+ DESCRIPTION    FROM ALL_MODELCODE")
        Call Combo_Loadval(cboSModel, RSMODEL)
    '******************************************************************
    
    txtEstimateEdit.Text = "": txtAppointmentEdit.Text = ""
    chkSettingAll.Value = GetSetting("DMIS 2.0", "CSMS", "SHOW ALL CUSTOMER SEARCH", 1)
    Call FillGrid
    SendKeys "{end}"
End Sub

Private Sub lblJob4Service_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        If MsgBox("DELETE! this job...  " & lblJob4Service.SelectedItem.SubItems(2) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
            Exit Sub
        End If
        Me.lblJob4Service.ListItems.Remove Me.lblJob4Service.SelectedItem.Index
        Call ComputeMe
    End If
End Sub

Private Sub lstCustomer_DblClick()
    If Not lstCustomer.ListItems.Count = 0 Then
        cmdNext.Value = True
    End If
End Sub

Private Sub lstCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstCustomer_DblClick
End Sub

Private Sub lstPMSDet_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        If MsgBox("DELETE! this job...  " & lstPMSDet.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
            Exit Sub
        End If
        Me.lstPMSDet.ListItems.Remove Me.lstPMSDet.SelectedItem.Index
    End If
End Sub

Private Sub lstVehicle_DblClick()
    cmdNext.Value = True
End Sub

Private Sub lstVehicle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsVehicleKo                                    As New ADODB.Recordset
    Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where (Cuscde = '" & txtID & "' OR ENDUSER = '" & txtID & "') and plate_no = '" & lstVehicle.SelectedItem.SubItems(1) & "'")
    If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
        txtPlate_No = Null2String(rsVehicleKo![PLATE_NO])
        cboModel = Null2String(rsVehicleKo![Model])
        txtMake = Null2String(rsVehicleKo![Make])
        txtYear = Null2String(rsVehicleKo![YER])
        txtVIN = Null2String(rsVehicleKo![Vin])
        txtVehName = Trim(cboModel) & "   " & txtPlate_No
    Else
        txtPlate_No = ""
        cboModel = ""
        txtMake = ""
        txtYear = ""
        txtVIN = ""
        txtVehName = ""
    End If
End Sub

Private Sub lsvDET_DblClick()
    If lsvDET.ListItems.Count = 0 Then Exit Sub
    
    Dim Index As Integer
    
    Text8.Enabled = False
    lsvDET.Enabled = False
    Command3.Enabled = False
    Index = lsvDET.SelectedItem.Index
    lblRES(0).Caption = lsvDET.ListItems(Index).Text
    lblRES(1).Caption = lsvDET.ListItems(Index).ListSubItems(1)
    lblRES(2).Caption = lsvDET.ListItems(Index).ListSubItems(2)
    txtEQTY.Text = ""
    picADD.Visible = True
    picADD.ZOrder 0
    On Error Resume Next
    txtEQTY.SetFocus
End Sub

Private Sub optCode_Click()
    On Error Resume Next
    Text8.SetFocus
End Sub

Private Sub optDesc_Click()
    On Error Resume Next
    Text8.SetFocus
End Sub

Private Sub optEndUser_Click()
    picPlate.Visible = False
    lstCustomer.Visible = True

    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------
    'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
        Label6(0).Visible = True
        textSearch.Visible = True
        cmdAddeditCustomer.Visible = True
        txtSPlate.Text = ""
    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------

    Call textSearch_Change
End Sub

Private Sub optFN_Click()
    picPlate.Visible = False
    lstCustomer.Visible = True

    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------
    'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
        Label6(0).Visible = True
        textSearch.Visible = True
        cmdAddeditCustomer.Visible = True
        txtSPlate.Text = ""
    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------

    Call textSearch_Change
End Sub

Private Sub optFullName_Click()
    picPlate.Visible = False
    lstCustomer.Visible = True

    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------
    'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
        Label6(0).Visible = True
        textSearch.Visible = True
        cmdAddeditCustomer.Visible = True
        txtSPlate.Text = ""
    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------

    Call textSearch_Change
End Sub

Private Sub optLN_Click()
    picPlate.Visible = False
    lstCustomer.Visible = True

    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------
    'DESCRIPTION : REQUEST BY SIR ARIEL OCAMPO SEARCH BY PLATE NO.
        Label6(0).Visible = True
        textSearch.Visible = True
        cmdAddeditCustomer.Visible = True
        txtSPlate.Text = ""
    'UPDATE BY : MJP 09-12-2007 02:42 AM -----------------------------------------------------------

    Call textSearch_Change
End Sub

Private Sub optPlate_Click()
    If optPlate.Value = True Then
        Label6(0).Visible = False
        textSearch.Visible = False

        lstCustomer.Visible = False
        cmdAddeditCustomer.Visible = False
        picPlate.Visible = True
        txtSPlate.Text = ""
        txtSPlate.SetFocus

        cboSearchBy.Clear
        cboSearchBy.AddItem "Plate No."
        cboSearchBy.AddItem "CS No."
        cboSearchBy.AddItem "VIN"
        cboSearchBy.Text = "Plate No."
        CleanPLateInfoLabel
    End If
End Sub

Private Sub Text8_Change()
    Call FillLsvDetails(Text8, lblTYPE)
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstCustomer.SetFocus
End Sub

Private Sub Timer1_Timer()
    If labType(0).Visible = True Then labType(0).Visible = False Else labType(0).Visible = True
End Sub

Private Sub txtDte_recd_Click()
    txtRecorded = txtDte_recd
End Sub

Private Sub txtEQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtKm_rdg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtRateLabor_Change()
    txtDiscLabor = NumericVal(txtEstLabor) * (NumericVal(txtRateLabor) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtRateparts_Change()
    txtDiscParts = NumericVal(txtEstParts) * (NumericVal(txtRateparts) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtRateAces_Change()
    txtDiscAces = NumericVal(txtEstAces) * (NumericVal(txtRateAces) / 100)
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtDiscAces_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtDiscLabor_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtDiscParts_Change()
    txtDiscTotal = NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscAces)
End Sub

Private Sub txtTranNo_Change()
    txtRep_Or.Text = txtTranNo.Text
    txtEstimateno.Text = txtTranNo.Text
End Sub

Private Sub txtTranNo_LostFocus()
    If xTransType <> "A" Then
        If Left(txtTranNo.Text, 2) = xTransType & "-" Then
            txtTranNo.Text = xTransType & "-" & Format(NumericVal(Right(txtTranNo.Text, Len(txtTranNo.Text) - 2)), "00000000")
        Else
            If VALID_COMPANY_CODE_FORHAI = True Then
            
            Else
                txtTranNo.Text = xTransType & "-" & Format(NumericVal(Right(txtTranNo.Text, Len(txtTranNo.Text))), "00000000")
            End If
        End If
    Else
    End If
    txtRep_Or.Text = txtTranNo
End Sub

Sub FillLsvDetails(XXX As String, XTYPE As String)
    Dim rstmp                       As New ADODB.Recordset
    Dim xCOND                       As String
    Dim xMODEL                      As String
    
    If XXX = "" Then
        Set rstmp = gconDMIS.Execute("SELECT TOP 100 STOCKNO, STOCKDESC, ISNULL(SRP,0) AS SRP, " & _
            " ISNULL(MODELCODE,'') AS MODELCODE ,ISNULL(GENUINE,'N') GENUINE, " & _
            " CASE WHEN ISNULL(ONHAND,0) <= 0 THEN 'N' " & _
            " Else 'Y' END AS ONHAND FROM PMIS_STOCKMAS " & _
            " WHERE TYPE = " & N2Str2Null(XTYPE) & "")
    Else
        If optCode.Value = True Then xCOND = " AND STOCKNO LIKE '" & XXX & "%'"
        If optDesc.Value = True Then xCOND = " AND STOCKDESC LIKE '" & XXX & "%'"
        If chkModel.Value = 1 Then xMODEL = " AND MODELCODE LIKE '%" & Left(cboSModel, 2) & "%'"
        If chkModel.Value = 0 Then xMODEL = ""
        
        Set rstmp = gconDMIS.Execute("SELECT TOP 100 STOCKNO, STOCKDESC, ISNULL(SRP,0) AS SRP, " & _
            " ISNULL(MODELCODE,'') AS MODELCODE ,ISNULL(GENUINE,'N') GENUINE, " & _
            " CASE WHEN ISNULL(ONHAND,0) <= 0 THEN 'N' " & _
            " Else 'Y' END AS ONHAND FROM PMIS_STOCKMAS " & _
            " WHERE TYPE = " & N2Str2Null(XTYPE) & "" & xCOND & xMODEL & "")
        
    End If
    Call Listview_Loadval(lsvDET.ListItems, rstmp)
End Sub

Function ADDROTYPE()
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
End Function
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

