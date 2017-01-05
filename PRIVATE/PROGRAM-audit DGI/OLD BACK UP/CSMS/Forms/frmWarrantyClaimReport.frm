VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmWarrantyClaimReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dealer Warranty Claim Report"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13245
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWarrantyClaimReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   13245
   Begin VB.PictureBox aclpic 
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
      Height          =   1185
      Left            =   5340
      ScaleHeight     =   1155
      ScaleWidth      =   3135
      TabIndex        =   42
      Top             =   2550
      Width           =   3165
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
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
         Height          =   195
         Left            =   -60
         ScaleHeight     =   165
         ScaleWidth      =   8145
         TabIndex        =   46
         Top             =   -60
         Width           =   8175
      End
      Begin VB.TextBox theACL 
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
         Left            =   900
         TabIndex        =   45
         Top             =   210
         Width           =   2085
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1740
         TabIndex        =   44
         Top             =   570
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2370
         TabIndex        =   43
         Top             =   570
         Width           =   645
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "ACL NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   60
      ScaleHeight     =   4755
      ScaleWidth      =   13095
      TabIndex        =   5
      Top             =   60
      Width           =   13125
      Begin VB.PictureBox thePicClaim 
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
         Height          =   4305
         Left            =   3990
         ScaleHeight     =   4275
         ScaleWidth      =   9015
         TabIndex        =   7
         Top             =   360
         Width           =   9045
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4245
            Left            =   3570
            TabIndex        =   8
            Top             =   0
            Width           =   5385
            Begin VB.Label lblEng 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1290
               TabIndex        =   27
               Top             =   2250
               Width           =   3735
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   315
               TabIndex        =   26
               Top             =   360
               Width           =   900
            End
            Begin VB.Label lblclaimNo 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1290
               TabIndex        =   25
               Top             =   690
               Width           =   3705
            End
            Begin VB.Label lblvin 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1290
               TabIndex        =   24
               Top             =   1830
               Width           =   3735
            End
            Begin VB.Label lblclaimType 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1290
               TabIndex        =   23
               Top             =   1440
               Width           =   3735
            End
            Begin VB.Label lblRO 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1290
               TabIndex        =   22
               Top             =   1050
               Width           =   1875
            End
            Begin VB.Label lblCustomer 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1290
               TabIndex        =   21
               Top             =   300
               Width           =   3705
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Eng No:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   495
               TabIndex        =   20
               Top             =   2340
               Width           =   720
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VIN:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   840
               TabIndex        =   19
               Top             =   1890
               Width           =   375
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Claim Type:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   135
               TabIndex        =   18
               Top             =   1500
               Width           =   1080
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ro No:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   585
               TabIndex        =   17
               Top             =   1140
               Width           =   630
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Claim No:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   345
               TabIndex        =   16
               Top             =   750
               Width           =   870
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Parts:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   705
               TabIndex        =   15
               Top             =   2820
               Width           =   510
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Labor:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   645
               TabIndex        =   14
               Top             =   3240
               Width           =   570
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   600
               TabIndex        =   13
               Top             =   3720
               Width           =   615
            End
            Begin VB.Label lblLabor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   1290
               TabIndex        =   12
               Top             =   3150
               Width           =   2115
            End
            Begin VB.Label Lblsublet 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   1290
               TabIndex        =   11
               Top             =   3630
               Width           =   2115
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   5415
               _Version        =   655364
               _ExtentX        =   9551
               _ExtentY        =   450
               _StockProps     =   14
               Caption         =   "INFORMATION"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GradientColorLight=   8421504
               GradientColorDark=   8421504
            End
            Begin VB.Label lblParts 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   375
               Left            =   1290
               TabIndex        =   9
               Top             =   2700
               Width           =   2115
            End
         End
         Begin VB.TextBox txtAcl 
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
            Left            =   1200
            TabIndex        =   33
            Top             =   330
            Width           =   2235
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
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
            Height          =   255
            Left            =   -60
            ScaleHeight     =   225
            ScaleWidth      =   9045
            TabIndex        =   32
            Top             =   -30
            Width           =   9075
         End
         Begin VB.TextBox txtsublet 
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
            Left            =   1200
            TabIndex        =   31
            Top             =   1050
            Width           =   2235
         End
         Begin VB.TextBox txtpwatype 
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
            Left            =   1200
            TabIndex        =   30
            Top             =   1410
            Width           =   2235
         End
         Begin VB.TextBox txtprevRO 
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
            Left            =   1200
            TabIndex        =   29
            Top             =   1770
            Width           =   2235
         End
         Begin VB.TextBox txtPrevAcl 
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
            Left            =   1200
            TabIndex        =   28
            Top             =   690
            Width           =   2235
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACL NO."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   510
            TabIndex        =   40
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet type."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   39
            Top             =   1110
            Width           =   900
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PWA Type."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   330
            TabIndex        =   38
            Top             =   1500
            Width           =   795
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prev.RO NO"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   37
            Top             =   1860
            Width           =   885
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prev ACL NO"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   36
            Top             =   750
            Width           =   945
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F5 - Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   2790
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Esc - Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   34
            Top             =   2790
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ListClaim 
         Height          =   4305
         Left            =   60
         TabIndex        =   41
         Top             =   360
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   7594
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
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmWarrantyClaimReport.frx":0E42
         NumItems        =   23
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Claiim No"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ro No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vin"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Eng No"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Part No"
            Object.Width           =   3526
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "del date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Repdate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "compdate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "KMG"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "partNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Ncode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ccode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "totallts"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "condition"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "cause"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "remedy"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "parts"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "TOtalPArts"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "GrandTotal"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "pwdtype"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "sublettype"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "prevno"
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   -30
         TabIndex        =   6
         Top             =   -30
         Width           =   13155
         _Version        =   655364
         _ExtentX        =   23204
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "LIST OF ALL CLAIM"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox cboRO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6240
      Width           =   2535
   End
   Begin Crystal.CrystalReport rptwarrantyClaim 
      Left            =   270
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
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
      Left            =   12450
      MouseIcon       =   "frmWarrantyClaimReport.frx":0FA4
      MousePointer    =   99  'Custom
      Picture         =   "frmWarrantyClaimReport.frx":10F6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   4920
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
      Left            =   11730
      MouseIcon       =   "frmWarrantyClaimReport.frx":1541
      MousePointer    =   99  'Custom
      Picture         =   "frmWarrantyClaimReport.frx":1693
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Information: Double Click the list to add to Warranty Claim"
      Height          =   270
      Left            =   90
      TabIndex        =   4
      Top             =   4950
      Width           =   4860
   End
   Begin VB.Label ro 
      Alignment       =   1  'Right Justify
      Caption         =   "Repair Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   6285
      Width           =   1170
   End
End
Attribute VB_Name = "frmWarrantyClaimReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theClaim, theRo, theClaimType, TheVin, theEng, thePrevAcl, theDelDate, theRepDate, theCompDate As String
Attribute theRo.VB_VarUserMemId = 1073938432
Attribute theClaimType.VB_VarUserMemId = 1073938432
Attribute TheVin.VB_VarUserMemId = 1073938432
Attribute theEng.VB_VarUserMemId = 1073938432
Attribute thePrevAcl.VB_VarUserMemId = 1073938432
Attribute theDelDate.VB_VarUserMemId = 1073938432
Attribute theRepDate.VB_VarUserMemId = 1073938432
Attribute theCompDate.VB_VarUserMemId = 1073938432
Dim theKmgReading, thePartNo, theNcode, theCCode, theReplacePart, theReplacePartName, theopcode, theQty, theLts As String
Attribute theKmgReading.VB_VarUserMemId = 1073938441
Attribute thePartNo.VB_VarUserMemId = 1073938441
Attribute theNcode.VB_VarUserMemId = 1073938441
Attribute theCCode.VB_VarUserMemId = 1073938441
Attribute theReplacePart.VB_VarUserMemId = 1073938441
Attribute theReplacePartName.VB_VarUserMemId = 1073938441
Attribute theopcode.VB_VarUserMemId = 1073938441
Attribute theQty.VB_VarUserMemId = 1073938441
Attribute theLts.VB_VarUserMemId = 1073938441
Dim thecondition, theCause, theRemedy                  As String
Attribute thecondition.VB_VarUserMemId = 1073938450
Attribute theCause.VB_VarUserMemId = 1073938450
Attribute theRemedy.VB_VarUserMemId = 1073938450
Dim theParts, theLabor, thesublet, thepwdtype, thesubletType, theprevRo As String
Attribute theParts.VB_VarUserMemId = 1073938453
Attribute theLabor.VB_VarUserMemId = 1073938453
Attribute thesublet.VB_VarUserMemId = 1073938453
Attribute thepwdtype.VB_VarUserMemId = 1073938453
Attribute thesubletType.VB_VarUserMemId = 1073938453
Attribute theprevRo.VB_VarUserMemId = 1073938453

Sub FillRO()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT Rep_or From CSMS_repor WHERE Transtype='R' order by Rep_or asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cboRO.Clear

    Do While Not RS.EOF

        cboRO.AddItem Null2String(RS!rep_OR)

        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub loadClaim()
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String
    Dim ITEM                                           As ListItem

    SQL = "SELECT DLR_CQIR_Referenceno,Ro_no,vinno,engineno,causalpartno,Customer,pwa_Date,inspectiondate,repairdate,mileage,causalpartno,naturecode,causecode,totallts,description,analysis,correctiveAction,totallaborcost,totalsubletrepair,grandtotal,totalcost from CSMS_CQIR"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListClaim.ListItems.Clear
    Do While Not RS.EOF
        Set ITEM = ListClaim.ListItems.Add(, , RS!DLR_CQIR_REFERENCENO)
        ITEM.SubItems(1) = Null2String(RS!Customer)
        ITEM.SubItems(2) = Null2String(RS!RO_NO)
        ITEM.SubItems(3) = Null2String(RS!VINNO)
        ITEM.SubItems(4) = Null2String(RS!EngineNo)
        ITEM.SubItems(5) = Null2String(RS!CAUSALPARTNO)
        ITEM.SubItems(6) = Null2String(RS!pwa_date)
        ITEM.SubItems(7) = Null2String(RS!InspectionDate)
        ITEM.SubItems(8) = Null2String(RS!RepairDate)
        ITEM.SubItems(9) = Null2String(RS!MILEAGE)
        ITEM.SubItems(10) = Null2String(RS!CAUSALPARTNO)
        ITEM.SubItems(11) = Null2String(RS!NATURECODE)
        ITEM.SubItems(12) = Null2String(RS!CAUSECODE)
        ITEM.SubItems(13) = Null2String(RS!totalLTS)
        ITEM.SubItems(14) = Null2String(RS!Description)
        ITEM.SubItems(15) = Null2String(RS!ANALYSIS)
        ITEM.SubItems(16) = Null2String(RS!correctiveAction)
        ITEM.SubItems(17) = Null2String(RS!TotalLaborCost)
        ITEM.SubItems(18) = Null2String(RS!totalcost)
        ITEM.SubItems(20) = Null2String(RS!TotalSUBLETREPAIR)
        ITEM.SubItems(19) = Null2String(RS!grandtotal)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub saveClaim()
    Dim SQL                                            As String

    theClaimType = "W"
    If txtAcl.Text = "" Then
        MsgBox "Pls Input Acl No!", vbExclamation, "WARNING"
        txtAcl.SetFocus
        Exit Sub
    End If
    thePrevAcl = Null
    theReplacePart = "not"
    theReplacePartName = "NOT"
    theopcode = "Not"
    theQty = 0
    thepwdtype = txtpwatype.Text
    thesubletType = txtsublet.Text
    theprevRo = txtPREVRO.Text

    SQL = "INSERT INTO CSMS_CQIRClaims VALUES('" & txtAcl.Text & "','" & theClaim & "','" & theRo & "','" & theClaimType & "','" & TheVin & _
          "','" & theEng & "','" & thePrevAcl & "','" & theDelDate & "','" & theRepDate & "','" & theCompDate & _
          "','" & theKmgReading & "','" & thePartNo & "','" & theNcode & "','" & theCCode & _
          "','" & theReplacePart & "','" & theReplacePartName & "','" & theopcode & _
          "','" & theQty & "','" & theLts & "','" & thecondition & _
          "','" & theCause & "','" & theRemedy & "','" & theParts & "','" & theLabor & _
          "','" & thesublet & "','" & thepwdtype & "','" & thesubletType & "','" & theprevRo & "')"

    gconDMIS.Execute (SQL)
    initMemvars
    MsgBox "All information has been save..", vbInformation, "Confirm"

End Sub

Sub initMemvars()
    txtAcl.Text = ""
    txtPREVACL.Text = ""
    txtsublet.Text = ""
    txtpwatype.Text = ""
    txtPREVRO.Text = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    thePicClaim.Visible = False
    ListClaim.Enabled = True
End Sub

Private Sub cmdOK_Click()
    If Function_Access(LOGID, "Acess_PRINT", "WARRANTY CLAIM REPORT") = False Then Exit Sub

    If theACL.Text = "" Then
        MsgBox "Please input a Accl No!", vbExclamation, "WARNING"
        theACL.SetFocus
        Exit Sub
    End If

    Dim SQL                                            As String
    Dim test                                           As String
    Dim theRo                                          As String
    Dim RS                                             As New ADODB.Recordset




    SQL = "SELECT * FROM CSMS_CQIRCLAIMS where acl_no='" & theACL.Text & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    Screen.MousePointer = 11
    aclpic.Visible = False
    'JUN 02/05/2005
    rptwarrantyClaim.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptwarrantyClaim.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptwarrantyClaim, CSMS_REPORT_PATH & "WarrantyClaim.rpt", "{CSMS_CQIRclaims.ACL_no}='" & theACL.Text & "'", CSMS_REPORT_CONNECTION, 1

    LogAudit "V", "WARRANTY CLAIM REPORTS - REPORTS ", theACL
    Screen.MousePointer = 0


End Sub

Private Sub cmdPrint_Click()
    aclpic.Visible = True


End Sub

Private Sub Command2_Click()
End Sub

Private Sub cmsSave_Click()
    Call saveClaim
End Sub

Private Sub Command1_Click()
    aclpic.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ans                                            As String
    If KeyCode = vbKeyF5 Then
        If thePicClaim.Visible = True Then
            ans = MsgBox("Are you Sure Do you want to save this claim?", vbQuestion + vbYesNo)
            If ans = vbYes Then
                saveClaim
                thePicClaim.Visible = False
            End If

        End If
    End If

    If KeyCode = vbKeyEscape Then
        If thePicClaim.Visible = True Then
            thePicClaim.Visible = False
            ListClaim.Enabled = True
        End If
    End If

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    FillRO
    loadClaim
    thePicClaim.Visible = False
    aclpic.Visible = False
End Sub

Private Sub ListClaim_DblClick()
    On Error Resume Next
    Dim thecust                                        As String
    theClaim = ListClaim.ListItems(ListClaim.SelectedItem.Index).Text
    thecust = ListClaim.SelectedItem.SubItems(1)
    theRo = ListClaim.SelectedItem.SubItems(2)
    TheVin = ListClaim.SelectedItem.SubItems(3)
    theEng = ListClaim.SelectedItem.SubItems(4)
    thePartNo = ListClaim.SelectedItem.SubItems(5)
    theDelDate = ListClaim.SelectedItem.SubItems(6)
    theRepDate = ListClaim.SelectedItem.SubItems(7)
    theCompDate = ListClaim.SelectedItem.SubItems(8)
    theKmgReading = ListClaim.SelectedItem.SubItems(9)
    theNcode = ListClaim.SelectedItem.SubItems(11)
    theCCode = ListClaim.SelectedItem.SubItems(12)
    theLts = ListClaim.SelectedItem.SubItems(13)
    thecondition = ListClaim.SelectedItem.SubItems(14)
    theCause = ListClaim.SelectedItem.SubItems(15)
    theRemedy = ListClaim.SelectedItem.SubItems(16)
    theLabor = ListClaim.SelectedItem.SubItems(17)
    theParts = ListClaim.SelectedItem.SubItems(18)
    thesublet = ListClaim.SelectedItem.SubItems(20)

    thePicClaim.Visible = True
    ListClaim.Enabled = False

    lblCustomer.Caption = thecust
    lblRO.Caption = theRo
    lblvin.Caption = TheVin
    lblEng.Caption = theEng
    lblclaimNo.Caption = theClaim
    lblParts = Format(theParts, "#,###,##0.00")
    lblLabor = Format(theLabor, "#,###,##0.00")
    Lblsublet = Format(thesublet, "#,###,#0.00")
End Sub

