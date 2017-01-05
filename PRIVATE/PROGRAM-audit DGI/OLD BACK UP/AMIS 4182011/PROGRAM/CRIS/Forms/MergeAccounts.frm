VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCRIS_MergeAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Merge"
   ClientHeight    =   8535
   ClientLeft      =   525
   ClientTop       =   795
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "MergeAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   12270
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9030
      Left            =   0
      ScaleHeight     =   9030
      ScaleWidth      =   3135
      TabIndex        =   129
      Top             =   0
      Width           =   3135
      Begin VB.ComboBox cboMerge 
         Height          =   345
         Left            =   0
         TabIndex        =   133
         Top             =   240
         Width           =   3075
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3825
         Left            =   0
         TabIndex        =   130
         Top             =   870
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   6747
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merging account is irreversible process. Please view transaction history to verify Transaction Details of Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   915
         Left            =   0
         TabIndex        =   141
         Top             =   4770
         Width           =   3075
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merging account Consolidates All the Transaction That customer has done and registered in company database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   885
         Left            =   0
         TabIndex        =   140
         Top             =   5715
         Width           =   3075
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"MergeAccounts.frx":08CA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1155
         Left            =   0
         TabIndex        =   139
         Top             =   6645
         Width           =   3075
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Please Fill Information of Merger Account."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   0
         TabIndex        =   138
         Top             =   7830
         Width           =   3075
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Merger Account to "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   0
         TabIndex        =   132
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Merge Account Of "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   0
         TabIndex        =   131
         Top             =   600
         Width           =   2205
      End
   End
   Begin VB.TextBox labOLDCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   14220
      TabIndex        =   102
      Top             =   2640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   14220
      TabIndex        =   74
      Top             =   2130
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   9315
      Left            =   3060
      ScaleHeight     =   9315
      ScaleWidth      =   9405
      TabIndex        =   0
      Top             =   -60
      Width           =   9405
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
         Left            =   8400
         MouseIcon       =   "MergeAccounts.frx":095B
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":0AAD
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Cancel"
         Top             =   7740
         Width           =   705
      End
      Begin VB.PictureBox picToolFrame 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   9855
         TabIndex        =   1
         Top             =   -150
         Width           =   9855
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   8550
            Top             =   900
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmdCUSTINFO_Contact 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3600
            MouseIcon       =   "MergeAccounts.frx":0DEB
            MousePointer    =   99  'Custom
            Picture         =   "MergeAccounts.frx":0F3D
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Contact Information"
            Top             =   240
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Child 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2100
            MouseIcon       =   "MergeAccounts.frx":162F
            MousePointer    =   99  'Custom
            Picture         =   "MergeAccounts.frx":1781
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "View Customers Number of Children"
            Top             =   240
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Credit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            MouseIcon       =   "MergeAccounts.frx":1DA2
            MousePointer    =   99  'Custom
            Picture         =   "MergeAccounts.frx":1EF4
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Update Credit and Terms of Customers"
            Top             =   240
            Width           =   585
         End
         Begin VB.Label labCustCode 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A00001"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   7020
            TabIndex        =   116
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label labCustInfo_Contact 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   4260
            MouseIcon       =   "MergeAccounts.frx":2557
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   420
            Width           =   1995
         End
         Begin VB.Label labCustInfo_Child 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Children"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2700
            MouseIcon       =   "MergeAccounts.frx":2861
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Label labCustInfo_Credit 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit && Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   780
            MouseIcon       =   "MergeAccounts.frx":2B6B
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   405
            Width           =   1335
         End
         Begin VB.Label labCustCode2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A00001"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   6990
            TabIndex        =   117
            Top             =   390
            Width           =   1425
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7125
         Left            =   60
         TabIndex        =   8
         Top             =   570
         Width           =   9165
         Begin VB.Frame Frame4 
            Caption         =   "Delivery Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1335
            Left            =   4290
            TabIndex        =   59
            Top             =   2580
            Width           =   4995
            Begin VB.TextBox txtDeliveryAddress 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   825
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   420
               Width           =   4755
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Same As above"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3480
               TabIndex        =   60
               Top             =   150
               Width           =   1305
            End
         End
         Begin VB.TextBox txtAcctName 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   5190
            MaxLength       =   100
            TabIndex        =   12
            Top             =   195
            Width           =   3885
         End
         Begin VB.ComboBox cboCustType 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   345
            ItemData        =   "MergeAccounts.frx":2E75
            Left            =   1320
            List            =   "MergeAccounts.frx":2E77
            TabIndex        =   9
            Text            =   "cboCustType"
            Top             =   180
            Width           =   2835
         End
         Begin VB.Frame Frame3 
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3165
            Left            =   4290
            TabIndex        =   62
            Top             =   3900
            Width           =   4815
            Begin VB.TextBox txtNotes 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   2865
               Left            =   60
               MaxLength       =   300
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   63
               Top             =   240
               Width           =   4725
            End
         End
         Begin VB.Frame fraEntity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   2130
            Left            =   30
            TabIndex        =   13
            Top             =   450
            Width           =   9075
            Begin VB.ComboBox cboPersonalCity 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   4200
               TabIndex        =   27
               Text            =   "cboApod"
               Top             =   1020
               Width           =   1995
            End
            Begin VB.TextBox txtPersonalStreet 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   120
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   1020
               Width           =   4035
            End
            Begin VB.TextBox txtPersonalState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   6240
               MaxLength       =   30
               TabIndex        =   28
               Top             =   1020
               Width           =   1695
            End
            Begin VB.TextBox txtPersonalZIP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   7980
               MaxLength       =   6
               TabIndex        =   29
               Top             =   1020
               Width           =   1005
            End
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1020
               TabIndex        =   19
               ToolTipText     =   "LAST NAME OR COMPANY NAME"
               Top             =   420
               Width           =   2715
            End
            Begin VB.ComboBox cboApod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   18
               Text            =   "cboApod"
               Top             =   420
               Width           =   855
            End
            Begin VB.TextBox txtMiddleName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   6435
               MaxLength       =   50
               TabIndex        =   21
               Top             =   420
               Width           =   2550
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3780
               TabIndex        =   20
               Top             =   420
               Width           =   2625
            End
            Begin VB.ComboBox cboSex 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2040
               TabIndex        =   34
               Text            =   "cboSex"
               Top             =   1650
               Width           =   855
            End
            Begin VB.TextBox txtBirthDate 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   120
               MaxLength       =   10
               TabIndex        =   33
               Top             =   1650
               Width           =   1875
            End
            Begin VB.TextBox txtSpouse 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   2940
               MaxLength       =   100
               TabIndex        =   35
               Top             =   1650
               Width           =   6045
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "City"
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
               Left            =   4200
               TabIndex        =   24
               Top             =   810
               Width           =   315
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Street"
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
               Left            =   120
               TabIndex        =   23
               Top             =   810
               Width           =   525
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "State/Province"
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
               Left            =   6240
               TabIndex        =   25
               Top             =   810
               Width           =   1245
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Zip Code"
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
               Left            =   7980
               TabIndex        =   22
               Top             =   780
               Width           =   735
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Salutation"
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
               Left            =   120
               TabIndex        =   14
               Top             =   150
               Width           =   855
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
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
               Left            =   6420
               TabIndex        =   17
               Top             =   210
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
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
               Left            =   3780
               TabIndex        =   16
               Top             =   210
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
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
               Left            =   1050
               TabIndex        =   15
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
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
               Left            =   2010
               TabIndex        =   32
               Top             =   1440
               Width           =   330
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
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
               Left            =   120
               TabIndex        =   30
               Top             =   1410
               Width           =   840
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Name"
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
               Left            =   2910
               TabIndex        =   31
               Top             =   1410
               Width           =   1185
            End
         End
         Begin VB.Frame fraMiscellenous 
            Height          =   4455
            Left            =   60
            TabIndex        =   36
            Top             =   2610
            Width           =   4185
            Begin VB.TextBox txtTin 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               MaxLength       =   15
               TabIndex        =   38
               Top             =   210
               Width           =   2775
            End
            Begin VB.TextBox txtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   54
               Top             =   3300
               Width           =   2775
            End
            Begin VB.TextBox txtHomePhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   52
               Top             =   2925
               Width           =   2775
            End
            Begin VB.TextBox txtMobile 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   50
               Top             =   2550
               Width           =   2775
            End
            Begin VB.TextBox txtCusphon1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   48
               Top             =   2175
               Width           =   2775
            End
            Begin VB.TextBox txtAsstPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   58
               Top             =   3975
               Width           =   2775
            End
            Begin VB.TextBox txtAssistant 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   56
               Top             =   3675
               Width           =   2775
            End
            Begin VB.TextBox txtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   46
               Top             =   1785
               Width           =   2775
            End
            Begin VB.TextBox txtDepartment 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               TabIndex        =   44
               Top             =   1380
               Width           =   2775
            End
            Begin VB.TextBox txtTitle 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   42
               Top             =   1005
               Width           =   2775
            End
            Begin VB.ComboBox cboLeadSource 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   40
               Text            =   "cboLeadSource"
               Top             =   615
               Width           =   2775
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Tin Number"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   37
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   870
               TabIndex        =   53
               Top             =   3255
               Width           =   285
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   51
               Top             =   2925
               Width           =   1095
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   615
               TabIndex        =   49
               Top             =   2205
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   47
               Top             =   2565
               Width           =   1050
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Asst. Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   165
               TabIndex        =   57
               Top             =   3885
               Width           =   990
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Assistant"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   390
               TabIndex        =   55
               Top             =   3525
               Width           =   765
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   675
               TabIndex        =   45
               Top             =   1815
               Width           =   480
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   43
               Top             =   1365
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   480
               TabIndex        =   41
               Top             =   1005
               Width           =   675
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Lead Source"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   39
               Top             =   645
               Width           =   1050
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Acct. Name"
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
            TabIndex        =   11
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Merge"
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
         Left            =   7710
         MouseIcon       =   "MergeAccounts.frx":2E79
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":2FCB
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Save this Record"
         Top             =   7740
         Width           =   705
      End
      Begin VB.CommandButton cmdAdjustments 
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7020
         MouseIcon       =   "MergeAccounts.frx":3440
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":3592
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "View Customer Transaction History"
         Top             =   7740
         Width           =   705
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview Customer Transaction History>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3090
         TabIndex        =   135
         Top             =   8040
         Width           =   3855
      End
   End
   Begin VB.PictureBox picContactAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFCCCF&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4185
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4305
      ScaleWidth      =   4350
      TabIndex        =   81
      Top             =   1815
      Visible         =   0   'False
      Width           =   4380
      Begin VB.TextBox txtContactName 
         Height          =   345
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   85
         Top             =   390
         Width           =   3045
      End
      Begin VB.CommandButton cmdCloseContactsAE 
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
         Height          =   675
         Index           =   1
         Left            =   3600
         MouseIcon       =   "MergeAccounts.frx":3C5C
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":3DAE
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Cancel Entry"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdSaveContact 
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
         Height          =   675
         Left            =   2970
         MouseIcon       =   "MergeAccounts.frx":40EC
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":423E
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Save Details"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdDeleteContact 
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
         Height          =   675
         Left            =   2340
         MouseIcon       =   "MergeAccounts.frx":458E
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":46E0
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Delect Details"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseContactsAE 
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
         Height          =   315
         Index           =   0
         Left            =   3990
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cboContactRelation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "MergeAccounts.frx":4A0B
         Left            =   1140
         List            =   "MergeAccounts.frx":4A0D
         TabIndex        =   87
         Top             =   790
         Width           =   3045
      End
      Begin VB.TextBox txtContactPosition 
         Height          =   345
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   89
         Top             =   1190
         Width           =   3045
      End
      Begin VB.TextBox txtContactDepartment 
         Height          =   345
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   90
         Top             =   1590
         Width           =   3045
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   345
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   92
         Top             =   1990
         Width           =   3045
      End
      Begin VB.TextBox txtContactMobile 
         Height          =   345
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   94
         Top             =   2390
         Width           =   3045
      End
      Begin VB.TextBox txtContactAddress 
         Height          =   645
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   2790
         Width           =   3045
      End
      Begin VB.Label labIDContacts 
         Height          =   555
         Left            =   1350
         TabIndex        =   98
         Top             =   3570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   375
         TabIndex        =   86
         Top             =   870
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   330
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "CONTACTS INFORMATION"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   570
         TabIndex        =   84
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   375
         TabIndex        =   88
         Top             =   1290
         Width           =   735
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   60
         TabIndex        =   91
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   525
         TabIndex        =   93
         Top             =   2130
         Width           =   585
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   510
         TabIndex        =   95
         Top             =   2550
         Width           =   600
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   345
         TabIndex        =   96
         Top             =   2970
         Width           =   765
      End
   End
   Begin VB.PictureBox picChildList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   3450
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   67
      Top             =   1560
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdCancelChildList 
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
         Height          =   645
         Left            =   5040
         MouseIcon       =   "MergeAccounts.frx":4A0F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":4B61
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton cmdSelectChild 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4350
         MouseIcon       =   "MergeAccounts.frx":4E9F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":4FF1
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Select"
         Top             =   4080
         Width           =   705
      End
      Begin MSComctlLib.ListView lvChildList 
         Height          =   3735
         Left            =   60
         TabIndex        =   69
         Top             =   330
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6588
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MergeAccounts.frx":532D
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddChildInfo 
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
         Height          =   645
         Left            =   3660
         MouseIcon       =   "MergeAccounts.frx":548F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":55E1
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Add Children/Dependent"
         Top             =   4080
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   ":: LIST OF CHILDRENS/DEPENDENTS::"
         ForeColor       =   -2147483630
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
   End
   Begin VB.PictureBox picChildAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFFDFD&
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   4185
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2475
      ScaleWidth      =   4350
      TabIndex        =   103
      Top             =   2730
      Visible         =   0   'False
      Width           =   4380
      Begin VB.CommandButton cmdCloseChildAE 
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
         Height          =   675
         Index           =   3
         Left            =   3480
         MouseIcon       =   "MergeAccounts.frx":58F4
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":5A46
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Cancel Entry"
         Top             =   1650
         Width           =   645
      End
      Begin VB.CommandButton cmdSaveChild 
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
         Height          =   675
         Left            =   2850
         MouseIcon       =   "MergeAccounts.frx":5D84
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":5ED6
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Save Children Information"
         Top             =   1650
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseChildAE 
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
         Height          =   315
         Index           =   2
         Left            =   3990
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdDeleteChild 
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
         Height          =   675
         Left            =   2220
         MouseIcon       =   "MergeAccounts.frx":6226
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":6378
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Add Children Information"
         Top             =   1650
         Width           =   645
      End
      Begin VB.TextBox txtChildName 
         Height          =   345
         Left            =   1200
         TabIndex        =   107
         Top             =   390
         Width           =   3015
      End
      Begin VB.ComboBox cboChildSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "MergeAccounts.frx":66A3
         Left            =   1200
         List            =   "MergeAccounts.frx":66B0
         TabIndex        =   111
         Top             =   1170
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtChildDate 
         Height          =   345
         Left            =   1200
         TabIndex        =   109
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "mm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   570
         TabIndex        =   106
         Top             =   390
         Width           =   540
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   104
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::CHILDREN INFORMATION::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Brith:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   45
         TabIndex        =   108
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label labIdCHILD 
         Height          =   555
         Left            =   1290
         TabIndex        =   115
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SEX:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   720
         TabIndex        =   110
         Top             =   1200
         Width           =   390
      End
   End
   Begin VB.PictureBox picCredit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00A9B8C2&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   4665
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2625
      ScaleWidth      =   3390
      TabIndex        =   118
      Top             =   2655
      Visible         =   0   'False
      Width           =   3420
      Begin VB.CommandButton cmdCloseTerm 
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
         Height          =   315
         Index           =   0
         Left            =   3090
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton Command12 
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
         Height          =   675
         Left            =   1830
         MouseIcon       =   "MergeAccounts.frx":66BD
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":680F
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Save Entry"
         Top             =   1770
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseTerm 
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
         Height          =   675
         Index           =   1
         Left            =   2460
         MouseIcon       =   "MergeAccounts.frx":6B5F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":6CB1
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Cancel Entry"
         Top             =   1770
         Width           =   645
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   420
         Width           =   1875
      End
      Begin VB.TextBox txtCreditDays 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   840
         Width           =   1875
      End
      Begin VB.CheckBox chkZeroRated 
         Appearance      =   0  'Flat
         BackColor       =   &H00A9B8C2&
         Caption         =   "Zero Rate  Customer"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1200
         TabIndex        =   119
         Top             =   1290
         Width           =   2205
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   330
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Width           =   3405
         _Version        =   655364
         _ExtentX        =   6006
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::CREDITS AND TERMS::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   690
         TabIndex        =   127
         Top             =   480
         Width           =   465
      End
      Begin VB.Label labTermID 
         Height          =   555
         Left            =   360
         TabIndex        =   126
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   180
         TabIndex        =   125
         Top             =   930
         Width           =   1020
      End
   End
   Begin VB.PictureBox picContactList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   3450
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   75
      Top             =   1560
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdCancelContactList 
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
         Height          =   645
         Left            =   5010
         MouseIcon       =   "MergeAccounts.frx":6FEF
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":7141
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Cancel"
         Top             =   4110
         Width           =   705
      End
      Begin VB.CommandButton cmdEditContact 
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
         Height          =   645
         Left            =   4320
         MouseIcon       =   "MergeAccounts.frx":747F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":75D1
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Edit Contact"
         Top             =   4110
         Width           =   705
      End
      Begin MSComctlLib.ListView lvContactList 
         Height          =   3735
         Left            =   60
         TabIndex        =   77
         Top             =   330
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6588
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MergeAccounts.frx":792D
         NumItems        =   0
      End
      Begin VB.CommandButton Command4 
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
         Height          =   645
         Left            =   3630
         MouseIcon       =   "MergeAccounts.frx":7A8F
         MousePointer    =   99  'Custom
         Picture         =   "MergeAccounts.frx":7BE1
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Add Contact"
         Top             =   4110
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   1
         Left            =   -30
         TabIndex        =   76
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   ":: LIST OF CONTACTS::"
         ForeColor       =   -2147483630
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
   End
   Begin VB.Label labid 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   14220
      TabIndex        =   64
      Top             =   420
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   14220
      TabIndex        =   66
      Top             =   1260
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQMAX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   14220
      TabIndex        =   73
      Top             =   1710
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labPrev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   14220
      TabIndex        =   65
      Top             =   870
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmCRIS_MergeAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS                                                 As ADODB.Recordset
Dim rsCusCtl                                           As ADODB.Recordset
Dim AddorEdit                                          As String
Dim AccountCode                                        As String
Dim RSMERGER                                           As ADODB.Recordset
Dim CUSCDE                                             As String
Dim CustType                                           As String
Private Sub cboApod_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub cboCustType_Click()
    Select Case cboCustType.Text
    Case "Personal"
        Label1.Caption = "Last Name"
        Label2.Visible = True: Label3.Visible = True
        txtLastName.Width = 2625: txtFirstName.Visible = True: txtMiddleName.Visible = True
        Label7.Caption = "Birth Date"
        Label24.Caption = "Spouse Name"
        CustType = "P"

        cmdCustInfo_Child.Enabled = True
        labCustInfo_Child.Enabled = True

    Case "Company/Agency"
        Label7.Caption = "Est Date"
        Label2.Visible = False: Label3.Visible = False
        txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
        Label1.Caption = "Company Name"
        Label24.Caption = "Contact Person"
        CustType = "C"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False
    Case "Fleet Account"
        CustType = "F"
        Label7.Caption = "Est Date"
        Label2.Visible = False: Label3.Visible = False
        txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
        Label1.Caption = "Company Name"
        Label24.Caption = "Contact Person"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False

    Case "Government"
        Label7.Caption = "Est Date"
        Label2.Visible = False: Label3.Visible = False
        txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
        Label1.Caption = "Establisment Name"
        Label24.Caption = "Contact Person"
        CustType = "G"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False
    End Select
End Sub



Private Sub cboMerge_Click()
    RSMERGER.MoveFirst
    RSMERGER.filter = ("ID<>" & cboMerge.ItemData(cboMerge.ListIndex))
    RefreshMegeres
    CUSCDE = cboMerge.Text
    StoreMemVars
End Sub

Private Sub cboMerge_LostFocus()
    cboMerge.ListIndex = SelectCombo(cboMerge, cboMerge)
End Sub

Function SelectCombo(C As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                                              As Long
    Dim ItemDataX                                      As Long
    If ByItemData = False Then
        For i = 0 To C.ListCount - 1
            If UCase(C.List(i)) = UCase(Trim(STR)) Then
                SelectCombo = i
                Exit Function
            End If
        Next
    Else
        If STR = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If

        ItemDataX = CLng(STR)

        For i = 0 To C.ListCount - 1
            If C.ItemData(i) = STR Then
                SelectCombo = i
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function

Private Sub cmdAddChildInfo_Click()
    cmdDeleteChild.Enabled = False
    txtChildDate = ""
    txtChildName = ""
    cboChildSex = ""
    labIdCHILD = ""
    ShowPictureBox picChildAE, True, picMain
    On Error Resume Next
    txtChildName.SetFocus
End Sub

Private Sub cmdAdjustments_Click()
    Dim frmTraHist                                     As frmCRIS_Inquiry_CustomerTransHistory
    Set frmTraHist = New frmCRIS_Inquiry_CustomerTransHistory
    frmTraHist.SHOWTRANSACTION txtCuscde
    frmTraHist.Show
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelChildList_Click()
    ShowPictureBox picChildList, False, picMain
End Sub

Private Sub cmdCancelContactList_Click()
    ShowPictureBox picContactList, False, picMain
End Sub

Private Sub cmdCloseChildAE_Click(Index As Integer)
    ShowPictureBox picChildAE, False, picMain
End Sub

Private Sub cmdCloseContactsAE_Click(Index As Integer)
    ShowPictureBox picContactAE, False, picMain
End Sub

Private Sub cmdCloseTerm_Click(Index As Integer)
    ShowPictureBox picCredit, False, picMain
End Sub

Private Sub cmdDeleteChild_Click()

    On Error GoTo ErrorCode:
    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub

    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CHILD WHERE id=" & labIdCHILD
    ShowPictureBox picChildAE, False, picMain
    LogAudit "X", "CUSTOMER CHILD", labIdCHILD
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdDeleteContact_Click()

    On Error GoTo ErrorCode:

    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub
    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CONTACTS WHERE id=" & labIDContacts
    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If

    ShowPictureBox picContactAE, False, picMain
    LogAudit "X", "CUSTOMER CONTACT", labIDContacts
    Exit Sub
ErrorCode:
    ShowVBError

End Sub


Private Sub cmdEditContact_Click()
    lvContactList_KeyPress 13
End Sub









Private Sub cmdSave_Click()

    If MsgBox("Do you want to Merge this Account ", vbInformation + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrorCode:
    If txtAcctName = "" Then
        ShowIsRequiredMsg "Account Name "
        On Error Resume Next
        txtAcctName.SetFocus
        Exit Sub
    End If
    If CustType = "P" And txtLastName = "" Then
        ShowIsRequiredMsg "Last Name"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If

    If CustType = "C" And txtLastName = "" Then
        ShowIsRequiredMsg "Company Name"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If
    If cboMerge.Text = "" Then
        ShowIsRequiredMsg "Merger Account Name"
        cboMerge.SetFocus
        Exit Sub
    End If

    Dim i
    Dim Exists                                         As Boolean
    i = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            Exists = True
        End If
    Next

    If Exists = False Then
        ShowIsRequiredMsg "Mergee Name"
        ListView1.SetFocus
        Exit Sub
    End If

    i = 0
    Dim Xcode
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            Xcode = ListView1.ListItems(i).Text
            If FormExist("frmAllCustomer") Then
                frmAllCustomer.cmdExit.Value = True
            End If

            On Error GoTo ErrorCode:
            Load frmSplash
            Screen.MousePointer = 11
            frmSplash.labCon.Caption = "Updating / Checking Customer Record(s)... Please wait..."

            frmSplash.Show
            DoEvents


            'GOT RECEIPTS
            gconDMIS.Execute ("UPDATE CMIS_OFF_HD SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-RECIEPTS", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE CMIS_OFF_DT SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-RECIEPTS", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE CSMS_CUSVEH SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-CUSTOMER VEHICLE", " FROM " & Xcode & "  TO " & txtCuscde
            'GOT APPOINTMENT
            gconDMIS.Execute ("UPDATE CSMS_APPOINTMENT SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-SERVICE CUSTOMER APPOINTMENT", " FROM " & Xcode & "  TO " & txtCuscde
            'GOT PARTS TRANS
            gconDMIS.Execute ("UPDATE pmis_ord_hist SET CUSTCODE=" & N2Str2Null(txtCuscde) & "  WHERE CUSTCODE=" & N2Str2Null(Xcode))

            gconDMIS.Execute ("UPDATE pmis_ord_hd SET CUSTCODE=" & N2Str2Null(txtCuscde) & "  WHERE CUSTCODE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-PARTS TRANSACTION", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE amis_openinvoice SET CUSTCODE=" & N2Str2Null(txtCuscde) & "  WHERE CUSTCODE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-ACCOUNTING OPEN INVOICE", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE smis_salesorder SET CODE=" & N2Str2Null(txtCuscde) & "  WHERE CODE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-CUSTOMER SALES INVOICES/SALES ORDER", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE CSMS_repairorder SET ACCT_NO=" & N2Str2Null(txtCuscde) & "  WHERE ACCT_NO=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CSMS_REPOR SET ACCT_NO=" & N2Str2Null(txtCuscde) & "  WHERE ACCT_NO=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-REPAIR ORDER", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE CSMS_ESTDETAILS SET ACCT_NO=" & N2Str2Null(txtCuscde) & "  WHERE ACCT_NO=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CSMS_ESTHD SET ACCT_NO=" & N2Str2Null(txtCuscde) & "  WHERE ACCT_NO=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE ALL_CUSTOMER_CHILD SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            '            gconDMIS.Execute ("UPDATE CMIS_NONVAT SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_LGM SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            '            gconDMIS.Execute ("UPDATE CMIS_OFF_HD_DEPOSITED SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE SMIS_PO SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE SMIS_MRRINV SET CUSTOMERCODE=" & N2Str2Null(txtCuscde) & "  WHERE CUSTOMERCODE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET CUSTOMERCODE=" & N2Str2Null(txtCuscde) & "  WHERE CUSTOMERCODE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECT_CALLS SET CSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECT_EMAIL SET CSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECT_LETTER SET CSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECT_VISITS SET CSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_REMINDERS SET CSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CSCDE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(txtCuscde) & "  WHERE CUSCDE=" & N2Str2Null(Xcode))
            LogAudit "E", "MERGE ACCOUNT-PROSPECT", " FROM " & Xcode & "  TO " & txtCuscde
            gconDMIS.Execute ("UPDATE SMIS_LOANINDIV SET APLCODE=" & N2Str2Null(txtCuscde) & "  WHERE APLCODE=" & N2Str2Null(Xcode))

            gconDMIS.Execute ("UPDATE SMIS_LOANCORP SET APLCODE=" & N2Str2Null(txtCuscde) & "  WHERE APLCODE=" & N2Str2Null(Xcode))
            gconDMIS.Execute ("DELETE FROM ALL_CUSTOMER_TABLE WHERE ID=" & ListView1.ListItems(i).ListSubItems(2).Text)
        End If

    Next


    Dim vtxtCusCde                                     As String
    Dim VTXTLASTNAME                                   As String
    Dim VTXTFIRSTNAME                                  As String
    Dim vtxtMiddleInitial                              As String
    Dim vtxtCUSCOMP                                    As String

    Dim vcboSex                                        As String
    Dim vtxtCusadd1                                    As String
    Dim vtxtCusadd2                                    As String
    Dim vtxtCuszipc                                    As String
    Dim vtxtCusphon1                                   As String
    Dim vtxtAcctName                                   As String
    Dim vcboApod                                       As String
    Dim vcboLeadSource                                 As String
    Dim vtxtTitle                                      As String
    Dim vtxtDepartment                                 As String
    Dim vtxtEmail                                      As String
    Dim vtxtMobile                                     As String
    Dim vtxtHomePhone                                  As String
    Dim VtxtFax                                        As String
    Dim vtxtAssistant                                  As String
    Dim vtxtAsstPhone                                  As String
    Dim vtxtCity                                       As String
    Dim vTxtBirthDate                                  As String
    Dim vTxtSpouse                                     As String
    Dim vtxtDescription                                As String
    Dim vtxtCustType                                   As String
    Dim vtxtCompanyAdd                                 As String
    Dim TEMPSQL                                        As String
    Dim vtxtDeliveryAddress                            As String
    Dim VtxtTIN                                        As String
    VtxtTIN = N2Str2Null(txtTin)
    vtxtCompanyAdd = N2Str2Null(UCase(txtPersonalStreet))
    vtxtCustType = N2Str2Null(CustType)
    vcboApod = N2Str2Null(UCase(cboApod))
    vtxtCusCde = N2Str2Null(txtCuscde)
    VTXTLASTNAME = N2Str2Null((txtLastName))
    VTXTFIRSTNAME = N2Str2Null((txtFirstName))
    vtxtMiddleInitial = N2Str2Null(txtMiddleName)
    vtxtAcctName = N2Str2Null(txtAcctName)
    vtxtCUSCOMP = N2Str2Null((txtLastName))
    vcboSex = N2Str2Null(cboSex)
    vtxtCusadd1 = N2Str2Null(Trim(UCase(txtPersonalStreet)))
    vtxtCusadd2 = N2Str2Null(UCase(txtPersonalState))
    vtxtCuszipc = N2Str2Null(txtPersonalZIP)
    vtxtCusphon1 = N2Str2Null(txtCusphon1)

    vcboLeadSource = N2Str2Null(cboLeadSource)
    vtxtTitle = N2Str2Null(UCase(txtTitle))
    vtxtDepartment = N2Str2Null(UCase(txtDepartment))
    vtxtEmail = N2Str2Null(txtEmail)
    vtxtMobile = N2Str2Null(txtMobile)
    vtxtHomePhone = N2Str2Null(txtHomePhone)
    VtxtFax = N2Str2Null(txtFax)
    vtxtAssistant = N2Str2Null(txtAssistant)
    vtxtAsstPhone = N2Str2Null(txtAsstPhone)

    vtxtCity = N2Str2Null(UCase(cboPersonalCity))
    vTxtBirthDate = N2Str2Null(txtBirthDate)
    vTxtSpouse = N2Str2Null(txtSpouse)
    vtxtDescription = N2Str2Null(txtNotes)

    vtxtDeliveryAddress = N2Str2Null(txtDeliveryAddress)





    TEMPSQL = "UPDATE ALL_CUSTOMER SET" & vbCrLf
    TEMPSQL = TEMPSQL & " CUSCOMP = " & vtxtCUSCOMP & "," & vbCrLf
    TEMPSQL = TEMPSQL & " TIN = " & VtxtTIN & "," & vbCrLf
    TEMPSQL = TEMPSQL & " COMPANYADD = " & vtxtCompanyAdd & "," & vbCrLf
    TEMPSQL = TEMPSQL & " APOD = " & vcboApod & "," & vbCrLf
    TEMPSQL = TEMPSQL & " LASTNAME = " & VTXTLASTNAME & "," & vbCrLf
    TEMPSQL = TEMPSQL & " FIRSTNAME = " & VTXTFIRSTNAME & "," & vbCrLf
    TEMPSQL = TEMPSQL & " MIDDLEINITIAL = " & vtxtMiddleInitial & "," & vbCrLf
    TEMPSQL = TEMPSQL & " ACCTNAME = " & vtxtAcctName & "," & vbCrLf
    TEMPSQL = TEMPSQL & " SEX = " & vcboSex & "," & vbCrLf
    TEMPSQL = TEMPSQL & " CUSTOMERADD = " & vtxtCusadd1 & "," & vbCrLf
    TEMPSQL = TEMPSQL & " PROVINCIALADD = " & vtxtCusadd2 & "," & vbCrLf
    TEMPSQL = TEMPSQL & " ZIPCODE = " & vtxtCuszipc & "," & vbCrLf
    TEMPSQL = TEMPSQL & " LEADSOURCE = " & vcboLeadSource & "," & vbCrLf
    TEMPSQL = TEMPSQL & " TITLE = " & vtxtTitle & "," & vbCrLf
    TEMPSQL = TEMPSQL & " DEPARTMENT = " & vtxtDepartment & "," & vbCrLf
    TEMPSQL = TEMPSQL & " EMAIL = " & vtxtEmail & "," & vbCrLf
    TEMPSQL = TEMPSQL & " MOBILE = " & vtxtMobile & "," & vbCrLf
    TEMPSQL = TEMPSQL & " TELEPHONENO  = " & vtxtCusphon1 & "," & vbCrLf
    TEMPSQL = TEMPSQL & " HOMEPHONE = " & vtxtHomePhone & "," & vbCrLf
    TEMPSQL = TEMPSQL & " FAX = " & VtxtFax & "," & vbCrLf
    TEMPSQL = TEMPSQL & " ASSISTANT = " & vtxtAssistant & "," & vbCrLf
    TEMPSQL = TEMPSQL & " ASSTPHONE = " & vtxtAsstPhone & "," & vbCrLf
    TEMPSQL = TEMPSQL & " CITY = " & vtxtCity & "," & vbCrLf
    TEMPSQL = TEMPSQL & " BIRTHDATE = " & vTxtBirthDate & "," & vbCrLf
    TEMPSQL = TEMPSQL & " SPOUSE = " & vTxtSpouse & "," & vbCrLf
    TEMPSQL = TEMPSQL & " CUSTYPE = " & vtxtCustType & "," & vbCrLf
    TEMPSQL = TEMPSQL & " DESCRIPTION = " & vtxtDescription & "," & vbCrLf

    TEMPSQL = TEMPSQL & " DeliveryAddress = " & vtxtDeliveryAddress & vbCrLf
    TEMPSQL = TEMPSQL & " WHERE CUSCDE = '" & txtCuscde & "'" & vbCrLf
    gconDMIS.Execute TEMPSQL








    LogAudit "E", "CUSTOMER MASTER MERGED", labCustCode & " ACCOUNT NAME" & txtAcctName



    Dim k                                              As Integer
    Dim NewCtlCde                                      As String
    Dim rsCustomer
    For k = 65 To 90
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!Code, 2, 5)) + 1, "00000")
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
    Next
    MsgBox "Account(s) Successfully Merged", vbInformation
    Unload frmSplash

    Screen.MousePointer = 0
    If FormExist("frmCRIS_Inquiry_PossibleDuplicates") Then
        frmCRIS_Inquiry_PossibleDuplicates.cmdFind.Value = True

    End If
    cmdSave.Enabled = False

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveContact_Click()
    On Error GoTo ErrorCode:
    If RTrim(LTrim(txtContactName)) = "" Then
        ShowIsRequiredMsg "CONTACT NAME"
        On Error Resume Next
        txtContactName.SetFocus
        Exit Sub
    End If
    Dim vtxtCusCde, ContactName, Relation, ContactPosition, Department, Phone, Mobile, Address, SQL

    vtxtCusCde = N2Str2Null(txtCuscde)
    ContactName = N2Str2Null(txtContactName)
    Relation = N2Str2Null(cboContactRelation)
    ContactPosition = N2Str2Null(txtContactPosition)
    Department = N2Str2Null(txtContactDepartment)
    Phone = N2Str2Null(txtContactPhone)
    Mobile = N2Str2Null(txtContactMobile)
    Address = N2Str2Null(txtContactAddress)

    If NumericVal(labIDContacts) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CONTACTS "
        SQL = SQL & "(ContactName , CUSCDE, Relation, ContactPosition, Department, Phone, Mobile, Address) VALUES ("
        SQL = SQL & ContactName & " ,"
        SQL = SQL & vtxtCusCde & " ,"
        SQL = SQL & Relation & " ,"
        SQL = SQL & ContactPosition & " ,"
        SQL = SQL & Department & " ,"
        SQL = SQL & Phone & " ,"
        SQL = SQL & Mobile & " ,"
        SQL = SQL & Address & " )"

        LogAudit "A", "CONTACTS INFORMAION"
    Else
        SQL = "UPDATE ALL_CUSTOMER_CONTACTS SET "
        SQL = SQL & " ContactName =" & ContactName & ", "
        SQL = SQL & " Relation =" & Relation & ", "
        SQL = SQL & " ContactPosition =" & ContactPosition & ", "
        SQL = SQL & " Department =" & Department & ", "
        SQL = SQL & " Phone =" & Phone & ", "
        SQL = SQL & " Address =" & Address & ", "
        SQL = SQL & " Mobile =" & Mobile
        SQL = SQL & "  where id=" & labIDContacts
        LogAudit "E", "CONTACTS INFORMAION"
    End If

    gconDMIS.Execute SQL
    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If

    Unload Me
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelectChild_Click()
    lvChildList_KeyPress 13
End Sub

Private Sub cmdCUSTINFO_Child_Click()
    Dim temprs                                         As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT CHILDNAME,SEX,DOB, DATEDIFF(YEAR, DOB, GETDATE()) ,ID FROM ALL_CUSTOMER_CHILD WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    Listview_Loadval lvChildList.ListItems, temprs

    ShowPictureBox picChildList, True, picMain
    If lvChildList.ListItems.Count = 0 Then
        cmdSelectChild.Enabled = False
    Else
        cmdSelectChild.Enabled = True
    End If
End Sub


Private Sub Command3_Click()
    txtDeliveryAddress = txtPersonalStreet & "," & txtPersonalState & "," & cboPersonalCity & "," & txtPersonalZIP
End Sub

Private Sub Command4_Click()
    labCUSTINFO_Contact_Click
    ShowPictureBox picContactAE, True
End Sub

Private Sub cmdSaveChild_Click()
    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtChildName)) = "" Then
        ShowIsRequiredMsg "Children Name "
        On Error Resume Next
        txtChildName.SetFocus
        Exit Sub
    End If

    Dim vtxtCHILDNAME, vtxtSEX, vtxtDOB, SQL, vtxtCusCde
    vtxtCHILDNAME = N2Str2Null(txtChildName)
    vtxtCusCde = N2Str2Null(txtCuscde)
    If cboChildSex = "M" Then
        vtxtSEX = "'M'"
    ElseIf cboChildSex = "F" Then

        vtxtSEX = "'F'"
    Else
        vtxtSEX = "'U'"
    End If

    vtxtDOB = N2Date2Null(txtChildDate)

    If NumericVal(labIdCHILD) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CHILD (CUSCDE,CHILDNAME,SEX,DOB)VALUES("
        SQL = SQL & vtxtCusCde
        SQL = SQL & "," & vtxtCHILDNAME & " ,"
        SQL = SQL & vtxtSEX & " ,"
        SQL = SQL & vtxtDOB & " )"
        LogAudit "A", "CUSTOMER CHILD"
    Else
        SQL = "UPDATE ALL_CUSTOMER_CHILD SET "
        SQL = SQL & " CHILDNAME =" & vtxtCHILDNAME & " , "
        SQL = SQL & " SEX=" & vtxtSEX & " , "
        SQL = SQL & " DOB=" & vtxtDOB
        SQL = SQL & " where id=" & labIdCHILD
        LogAudit "E", "CUSTOMER CHILD"
    End If

    gconDMIS.Execute SQL

    If picChildList.Visible = True Then
        cmdCUSTINFO_Child_Click
    End If
    ShowPictureBox picChildAE, False, picMain

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCUSTINFO_Contact_Click()
    Dim temprs                                         As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT  CONTACTNAME, RELATION,PHONE, MOBILE, CONTACTPOSITION, DEPARTMENT, ADDRESS, ID FROM ALL_CUSTOMER_CONTACTS WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    Listview_Loadval lvContactList.ListItems, temprs

    ShowPictureBox picContactList, True, picMain
    If lvContactList.ListItems.Count = 0 Then
        cmdEditContact.Enabled = False
    Else
        cmdEditContact.Enabled = True
    End If
End Sub

Private Sub cmdCUSTINFO_CREDIT_Click()
    If Module_Access(LOGID, "CUSTOMER CREDIT LIMIT", "DATA ENTRY") = False Then Exit Sub
    ShowPictureBox picCredit, True, picMain
    On Error Resume Next
    txtCreditLimit.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picChildAE.Visible = True And KeyCode = 27 Then
        cmdCloseChildAE_Click 0
    ElseIf picChildList.Visible = True And KeyCode = 27 Then
        cmdCancelChildList_Click

    ElseIf picContactAE.Visible = True And KeyCode = 27 Then
        cmdCloseContactsAE_Click 0
    ElseIf picContactList.Visible = True And KeyCode = 27 Then
        cmdCancelContactList_Click

    ElseIf picCredit.Visible = True And KeyCode = 27 Then
        cmdCloseTerm_Click 0
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1

    picMain.Enabled = True
    initMemvars
    InitData
    With ListView1.ColumnHeaders
        .Add 1, , "Code", 0.3 * ListView1.Width
        .Add 2, , "Name", 0.7 * ListView1.Width
    End With

    With lvChildList.ColumnHeaders
        .Add 1, , "ChildName", 0.5 * lvChildList.Width
        .Add 2, , "SEX", 0.15 * lvChildList.Width
        .Add 3, , "DATEOFBIRTH", 0.15 * lvChildList.Width
        .Add 4, , "AGE", 0.15 * lvChildList.Width
    End With
    With lvContactList.ColumnHeaders
        .Add 1, , "CONTACTNAME", 0.4 * lvChildList.Width
        .Add 2, , "RELATION", 0.2 * lvChildList.Width
        .Add 3, , "PHONE", 0.17 * lvChildList.Width
        .Add 4, , "MOBILE", 0.17 * lvChildList.Width
    End With
    FIllMerger
    StoreMemVars
    Screen.MousePointer = 0
End Sub
Sub FIllMerger()
    Dim LST
    RSMERGER.MoveFirst
    ListView1.ListItems.Clear
    While Not RSMERGER.EOF
        Set LST = ListView1.ListItems.Add(, , RSMERGER.Fields(0).Value)
        cboMerge.AddItem RSMERGER.Fields(0).Value
        cboMerge.ItemData(cboMerge.NewIndex) = RSMERGER.Fields(2).Value
        LST.ListSubItems.Add , , RSMERGER.Fields(1).Value
        LST.ListSubItems.Add , , RSMERGER.Fields(2).Value
        LST.Checked = True
        RSMERGER.MoveNext
    Wend
    cboMerge.ListIndex = 0
    RSMERGER.MoveFirst
End Sub
Sub RefreshMegeres()
    Dim LST
    RSMERGER.MoveFirst
    ListView1.ListItems.Clear
    While Not RSMERGER.EOF
        Set LST = ListView1.ListItems.Add(, , RSMERGER.Fields(0).Value)
        LST.ListSubItems.Add , , RSMERGER.Fields(1).Value
        LST.ListSubItems.Add , , RSMERGER.Fields(2).Value
        LST.Checked = True
        RSMERGER.MoveNext
    Wend
    RSMERGER.MoveFirst

End Sub

Sub MergeAccount(RS As ADODB.Recordset)
    Set RSMERGER = RS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsCusCtl = Nothing
    AddorEdit = vbNullString
    AccountCode = vbNullString
    CustType = vbNullString
    Set RSMERGER = Nothing

End Sub

Sub InitData()
    Combo_Loadval cboPersonalCity, gconDMIS.Execute("Select Distinct CITY FROM ALL_CUSTOMER WHERE CITY IS NOT NULL")
    With cboCustType
        .AddItem ("Personal")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet Account")
        .ListIndex = 0
    End With
End Sub

Sub initMemvars()
    Dim temprs                                         As ADODB.Recordset
    labSEQ.Caption = gconDMIS.Execute("SELECT isnull(MAX(ID),0) FROM ALL_CUSTOMER").Collect(0)
    labCustCode = ""
    labCustCode2 = ""
    txtCuscde.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAcctName.Text = ""
    cboLeadSource.Text = ""
    cboSex.Text = ""
    txtTitle.Text = ""
    txtDepartment.Text = ""
    txtEmail.Text = ""
    txtCusphon1.Text = ""
    txtMobile.Text = ""
    txtHomePhone.Text = ""
    txtFax.Text = ""
    txtAssistant.Text = ""
    txtAsstPhone.Text = ""
    txtPersonalStreet.Text = ""
    cboPersonalCity.Text = ""
    txtPersonalState.Text = ""
    txtPersonalZIP.Text = ""
    txtBirthDate.Text = ""
    txtSpouse.Text = ""
    txtNotes.Text = ""
    cboApod.Clear

    txtTin = ""
    txtDeliveryAddress = ""
    txtCreditDays = "0"
    txtCreditLimit = "0.00"

    Dim rsAPOD                                         As ADODB.Recordset
    Set rsAPOD = New ADODB.Recordset
    Set rsAPOD = gconDMIS.Execute("Select distinct apod from ALL_CustMaster_Smis Where APOD is Not Null")

    If Not rsAPOD.EOF And Not rsAPOD.BOF Then
        rsAPOD.MoveFirst
        Do While Not rsAPOD.EOF
            cboApod.AddItem Null2String(rsAPOD!APOD)
            rsAPOD.MoveNext
        Loop
    End If
    Set rsAPOD = Nothing
    Set temprs = gconDMIS.Execute("Select DataDesc from CRIS_vW_MasterPullDown where  Masterdesc='Lead Source'")
    cboLeadSource.Clear

    While Not temprs.EOF
        cboLeadSource.AddItem temprs.Collect(0)
        temprs.MoveNext
    Wend
    cboSex.Clear
    cboSex.AddItem "NA"
    cboSex.AddItem "M"
    cboSex.AddItem "F"
End Sub

Private Sub labCustInfo_Child_Click()
    cmdAddChildInfo_Click
End Sub

Private Sub labCustInfo_Child_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = &H400000
    labCustInfo_Child.FontBold = True
End Sub

Private Sub labCUSTINFO_CREDIT_Click()
    cmdCUSTINFO_CREDIT_Click
End Sub

Private Sub labCUSTINFO_CREDIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Credit.ForeColor = &H400000
    labCustInfo_Credit.FontBold = True
End Sub

Private Sub labCUSTINFO_Contact_Click()
    labIDContacts = 0
    txtContactName = ""
    cboContactRelation = ""
    txtContactPosition = ""
    txtContactDepartment = ""
    txtContactPhone = ""
    txtContactMobile = ""
    txtContactAddress = ""
    cmdDeleteContact.Enabled = False
End Sub

Private Sub labCUSTINFO_Contact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Contact.ForeColor = &H400000
    labCustInfo_Contact.FontBold = True
End Sub




Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i
    cmdSave.Enabled = False
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
            cmdSave.Enabled = True
        End If
    Next
End Sub

Private Sub lvChildList_DblClick()
    lvChildList_KeyPress 13
End Sub

Private Sub lvChildList_KeyPress(KeyAscii As Integer)
    If lvChildList.SelectedItem Is Nothing Then Exit Sub
    On Error GoTo ADDER:

    If KeyAscii <> 13 Then Exit Sub
    txtChildName = lvChildList.SelectedItem
    cboChildSex = lvChildList.SelectedItem.ListSubItems(1).Text
    txtChildDate = lvChildList.SelectedItem.ListSubItems(2).Text
    labIdCHILD = lvChildList.SelectedItem.ListSubItems(3).Text
    cmdDeleteChild.Enabled = True

    ShowPictureBox picChildList, False
    ShowPictureBox picChildAE, True, picMain
    On Error Resume Next
    txtChildName.SetFocus
    Exit Sub
ADDER:
    ShowVBError
End Sub

Private Sub lvContactList_DblClick()
    lvContactList_KeyPress 13
End Sub

Private Sub lvContactList_KeyPress(KeyAscii As Integer)
    If lvContactList.SelectedItem Is Nothing Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    ShowPictureBox picContactAE, True, picMain
    With lvContactList.SelectedItem
        txtContactName = .Text
        cboContactRelation = .ListSubItems(1).Text
        txtContactPhone = .ListSubItems(2).Text
        txtContactMobile = .ListSubItems(3).Text
        txtContactPosition = .ListSubItems(4).Text
        txtContactDepartment = .ListSubItems(5).Text
        txtContactAddress = .ListSubItems(6).Text
        labIDContacts = .ListSubItems(7).Text

    End With

    cmdDeleteContact.Enabled = True
    On Error Resume Next

    txtContactName.SetFocus
End Sub


Sub SetCustomerAccountName()
    If AddorEdit = "EDIT" Or AddorEdit = "" Then: Exit Sub
    txtAcctName = UCase(txtLastName & IIf(txtFirstName = "", "", ",") & txtFirstName & IIf(txtMiddleName = "", "", ".") & Left(txtMiddleName, 1))
End Sub

Sub StoreMemVars()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM ALL_CUSTOMER WHERE CUSCDE IN('" & CUSCDE & "')", gconDMIS, adOpenKeyset, adLockReadOnly
    If Not RS.EOF And Not RS.BOF Then
        labID.Caption = RS!ID
        txtCuscde.Text = Null2String(RS!CUSCDE)
        txtAcctName = Null2String(RS!AcctName)
        labCustCode = Null2String(RS!CUSCDE)
        labCustCode2 = Null2String(RS!CUSCDE)

        cboApod.Text = Null2String(RS!APOD)
        txtLastName.Text = Null2String(RS!lastname)
        txtFirstName.Text = Null2String(RS!Firstname)
        txtMiddleName.Text = Null2String(RS!MiddleInitial)
        txtTin = Null2String(RS!TIN)
        cboSex.Text = Null2String(RS!Sex)
        txtPersonalStreet.Text = Null2String(RS!CUSTOMERADD)
        txtPersonalState.Text = Null2String(RS!provincialadd)
        txtPersonalZIP.Text = Null2String(RS!ZIPCODE)
        txtCusphon1.Text = Null2String(RS!TelephoneNo)
        cboLeadSource.Text = Null2String(RS!LeadSource)
        txtTitle.Text = Null2String(RS!TITLE)
        txtDepartment.Text = Null2String(RS!Department)
        txtEmail.Text = Null2String(RS!EMAIL)
        txtMobile.Text = Null2String(RS!Mobile)
        txtHomePhone.Text = Null2String(RS!HomePhone)
        txtFax.Text = Null2String(RS!Fax)
        txtAssistant.Text = Null2String(RS!Assistant)
        txtAsstPhone.Text = Null2String(RS!AsstPhone)
        cboPersonalCity.Text = Null2String(RS!CITY)
        txtBirthDate.Text = Null2String(RS!BirthDate)
        txtSpouse.Text = Null2String(RS!Spouse)
        txtDeliveryAddress = Null2String(RS!DELIVERYADDRESS)
        txtNotes.Text = Null2String(RS!DESCRIPTION)
        txtCreditDays = NumericVal(RS!CREDITDAYS)
        txtCreditLimit = FormatNumber(NumericVal(RS!CreditLimit))

        If Null2String(RS!CUSTYPE) = "P" Then
            cboCustType.ListIndex = 0
        ElseIf Null2String(RS!CUSTYPE) = "C" Then
            cboCustType.ListIndex = 1
        ElseIf Null2String(RS!CUSTYPE) = "G" Then
            cboCustType.ListIndex = 2
        ElseIf Null2String(RS!CUSTYPE) = "F" Then
            cboCustType.ListIndex = 3
        Else
            cboCustType.ListIndex = 0
        End If
        If Null2String(RS!CUSCAT) = "Z" Then
            chkZeroRated.Value = 1
        Else
            chkZeroRated.Value = 0
        End If
    Else
        ShowNoRecord
    End If
End Sub






Private Sub picToolFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = vbBlack
    labCustInfo_Child.FontBold = False
    labCustInfo_Credit.ForeColor = vbBlack
    labCustInfo_Credit.FontBold = False
    labCustInfo_Contact.ForeColor = vbBlack
    labCustInfo_Contact.FontBold = False
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCreditDays_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditDays_GotFocus()
    If NumericVal(txtCreditDays.Text) <= 0 Then txtCreditDays = ""
End Sub

Private Sub txtCreditDays_LostFocus()
    If NumericVal(txtCreditDays) <= 0 Then txtCreditDays = "0"
    txtCreditDays = NumericVal(txtCreditDays)
End Sub

Private Sub txtCreditLimit_GotFocus()
    If NumericVal(txtCreditLimit.Text) <= 0 Then txtCreditLimit = ""
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditLimit_LostFocus()
    If NumericVal(txtCreditLimit) <= 0 Then txtCreditLimit = "0.00"
    txtCreditLimit = FormatNumber(NumericVal(txtCreditLimit))
End Sub



Private Sub txtFirstName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtLastName_Change()
    If AddorEdit = "ADD" And LTrim(RTrim(txtLastName)) <> "" Then
        txtCuscde = GetCustomerCode(txtLastName)
        SetCustomerAccountName
    End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtMiddleName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    UpperAscii (KeyAscii)
End Sub

Private Sub txtPersonalStreet_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtPersonalZIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub


Sub ShowPictureBox(cntl As Object, State As Boolean, Optional ByVal MasterObject As Object)
    cntl.Visible = State
    If Not (MasterObject Is Nothing) Then
        MasterObject.Enabled = Not State
    End If
    If State = True Then
        cntl.ZOrder 0
    Else
        cntl.ZOrder 1
    End If
End Sub

Function GetCustomerCode(lastname As String) As String
    Dim temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = Left(Trim(lastname), 1)
    Set temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (temprs.EOF Or temprs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerCode = Left(lastname, 1) & "00001"
    End If
End Function

Public Sub FillCustomerListView(RS As Recordset, grd As ListView, Optional WithSN As Boolean = False, Optional WITHCOLUMNHEADER As Boolean = False)
    Dim fld                                            As Field
    Dim j                                              As Long
    Dim ijx                                            As Integer
    Dim LST                                            As ListItem
    Dim i                                              As Integer

    grd.Enabled = False

    grd.ListItems.Clear

    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend

    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then

        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend

    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    Else
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , Null2String(RS.Fields(0).Value))
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    End If
    grd.Enabled = True
    Set LST = Nothing
End Sub
