VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRIS_Customer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   6480
   ClientLeft      =   525
   ClientTop       =   840
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "ALLCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   2925
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   57
      Top             =   5550
      Width           =   9225
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
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
         MouseIcon       =   "ALLCustomers.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   58
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
         Left            =   785
         MouseIcon       =   "ALLCustomers.frx":0D7B
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":0ECD
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   45
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
         Left            =   1540
         MouseIcon       =   "ALLCustomers.frx":1225
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":1377
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   45
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
         Left            =   3865
         MouseIcon       =   "ALLCustomers.frx":1671
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":17C3
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Left            =   4620
         MouseIcon       =   "ALLCustomers.frx":1AD6
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":1C28
         Style           =   1  'Graphical
         TabIndex        =   60
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
         Left            =   5375
         MouseIcon       =   "ALLCustomers.frx":1F84
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":20D6
         Style           =   1  'Graphical
         TabIndex        =   61
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
         Left            =   6130
         MouseIcon       =   "ALLCustomers.frx":2401
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":2553
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   30
         Width           =   705
      End
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
         Left            =   6885
         MouseIcon       =   "ALLCustomers.frx":28B9
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":2A0B
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         Left            =   3080
         MouseIcon       =   "ALLCustomers.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   45
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   2295
         MouseIcon       =   "ALLCustomers.frx":3213
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":3365
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9750
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   68
      Top             =   5625
      Width           =   2580
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
         MouseIcon       =   "ALLCustomers.frx":36C3
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":3815
         Style           =   1  'Graphical
         TabIndex        =   69
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
         Left            =   765
         MouseIcon       =   "ALLCustomers.frx":3B65
         MousePointer    =   99  'Custom
         Picture         =   "ALLCustomers.frx":3CB7
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6315
      Left            =   60
      TabIndex        =   53
      Top             =   -30
      Width           =   2325
      Begin VB.TextBox txtSearch 
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
         Height          =   375
         Left            =   60
         MaxLength       =   35
         TabIndex        =   54
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   5565
         Left            =   30
         TabIndex        =   55
         Top             =   660
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   9816
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ALLCustomers.frx":3FF5
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CUSTOMER NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5505
      Left            =   2460
      TabIndex        =   2
      Top             =   -60
      Width           =   8835
      Begin VB.ComboBox cboSex 
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
         ForeColor       =   &H00973640&
         Height          =   315
         Left            =   3750
         TabIndex        =   17
         Text            =   "cboSex"
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox txtDescription 
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
         Height          =   1125
         Left            =   4620
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Text            =   "ALLCustomers.frx":4157
         Top             =   4320
         Width           =   4125
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other Info"
         Height          =   1095
         Left            =   4620
         TabIndex        =   45
         Top             =   2910
         Width           =   4155
         Begin VB.TextBox txtSpouse 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   1140
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   630
            Width           =   2925
         End
         Begin VB.TextBox txtBirthDate 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   1140
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   270
            Width           =   2925
         End
         Begin VB.Label Label24 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   660
            Width           =   2085
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   300
            Width           =   2085
         End
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   4410
         Width           =   3165
      End
      Begin VB.TextBox txtHomePhone 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   4050
         Width           =   3165
      End
      Begin VB.TextBox txtMobile 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3690
         Width           =   3165
      End
      Begin VB.TextBox txtCusphon1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3330
         Width           =   3165
      End
      Begin VB.TextBox txtAsstPhone 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   52
         Top             =   5130
         Width           =   3165
      End
      Begin VB.TextBox txtAssistant 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4770
         Width           =   3165
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   2970
         Width           =   3165
      End
      Begin VB.Frame Frame2 
         Caption         =   "Address Information"
         Height          =   2415
         Left            =   4620
         TabIndex        =   36
         Top             =   480
         Width           =   4155
         Begin VB.TextBox txtCity 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   1140
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   1290
            Width           =   2925
         End
         Begin VB.TextBox txtCusadd1 
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
            Height          =   705
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Text            =   "ALLCustomers.frx":415D
            Top             =   540
            Width           =   3885
         End
         Begin VB.TextBox txtCusadd2 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   1140
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   1650
            Width           =   2925
         End
         Begin VB.TextBox txtCuszipc 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   1140
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   2010
            Width           =   855
         End
         Begin VB.Label Label22 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   2085
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Mailing Street"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   270
            Width           =   1875
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Province"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   1680
            Width           =   2085
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.TextBox txtDepartment 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2610
         Width           =   3165
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2250
         Width           =   3165
      End
      Begin VB.ComboBox cboLeadSource 
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
         ForeColor       =   &H00973640&
         Height          =   315
         Left            =   1380
         TabIndex        =   16
         Text            =   "cboLeadSource"
         Top             =   1890
         Width           =   1875
      End
      Begin VB.ComboBox cboApod 
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
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Text            =   "cboApod"
         Top             =   450
         Width           =   1035
      End
      Begin VB.TextBox txtAcctName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1530
         Width           =   3165
      End
      Begin VB.TextBox txtCusnam3 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   4050
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox txtCusnam2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1170
         Width           =   2625
      End
      Begin VB.TextBox txtCusnam1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   1380
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   810
         Width           =   2625
      End
      Begin VB.TextBox txtCuscde 
         Appearance      =   0  'Flat
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
         Left            =   2940
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4650
         TabIndex        =   30
         Top             =   4050
         Width           =   3615
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Phone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Office Phone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label lblCap 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Asst. Phone"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   50
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Assistant"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Lead Source"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   3495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   30
         Top             =   150
         Width           =   8745
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Acct. Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3330
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4110
         TabIndex        =   9
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   585
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   285
         Left            =   4620
         Top             =   4020
         Width           =   4125
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   225
      Left            =   750
      TabIndex        =   56
      Top             =   5820
      Width           =   1455
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
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
      Left            =   270
      TabIndex        =   1
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCRIS_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim rsCustomer                               As ADODB.Recordset
'Dim rsPurchAgree                             As ADODB.Recordset
'Dim rsCusCtl                                 As ADODB.Recordset
'Dim AddorEdit                                As String
'
'Private Sub cmdAdd_Click()
'    AddorEdit = "ADD"
'    Frame1.Enabled = True
'    Picture1.Visible = False
'    Picture2.Visible = True
'    lstCustomer.Enabled = False
'    InitMemvars
'    On Error Resume Next
'    txtCusnam1.SetFocus
'End Sub
'
'Private Sub cmdAddVehicle_Click()
''    With FrmCSMSAddVehicle
''        .labCustCode.Caption = txtCuscde
''        .labCustomer.Caption = txtAcctName
''        .cmdAdd.Value = True
''    End With
''    FrmCSMSAddVehicle.Show
'End Sub
'
'Private Sub cmdCancel_Click()
'    Frame1.Enabled = False
'    Picture1.Visible = True
'    Picture2.Visible = False
'    lstCustomer.Enabled = True
'    AddorEdit = ""
'    StoreMemVars
'End Sub
'
'Private Sub cmdDelete_Click()
''    On Error GoTo ErrorCode
''    If ShowConfirmDelete = True Then
''        oConSQL.Execute "delete from ALL_CustMaster_Smis where Code = '" & txtCuscde.Text & "'"
''        oConSQL.Execute "delete from ALL_CustMaster_Smis where Code = '" & txtCuscde.Text & "'"
''        oConSQL.Execute "delete from ALL_CustMaster_Smis where Code = '" & txtCuscde.Text & "'"
''        oConSQL.Execute "delete from ALL_CustMaster_Smis where Code = '" & txtCuscde.Text & "'"
''        Screen.MousePointer = 11
''        oConSQL.Execute "delete from ALL_CusCtl"
''        oConSQL.Execute "delete from ALL_CusCtl"
''        oConSQL.Execute "delete from ALL_CusCtl"
''        oConSQL.Execute "delete from ALL_CusCtl"
''        Dim k                                As Integer
''        Dim NewCtlCde                        As String
''        For k = 65 To 90
''            Set rsCustomer = New ADODB.Recordset
''            rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", oConSQL
''            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
''                NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!code, 2, 5)) + 1, "00000")
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
''            Else
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
''                oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
''            End If
''        Next
''        Screen.MousePointer = 0
''        ShowDeletedMsg
''    End If
''    rsRefresh
''    StoreMemVars
''    FillGrid
''    Exit Sub
''
''ErrorCode:
''    ShowVBError
''    Exit Sub
'End Sub
'
'Private Sub cmdEdit_Click()
'    AddorEdit = "EDIT"
'    Frame1.Enabled = True
'    Picture1.Visible = False
'    Picture2.Visible = True
'    lstCustomer.Enabled = False
'    On Error Resume Next
'    txtCusnam1.SetFocus
'End Sub
'
'Private Sub cmdExit_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdFind_CLick()
'    txtSearch.SetFocus
'End Sub
'
'Private Sub cmdFirst_Click()
'    rsCustomer.MoveFirst
'    StoreMemVars
'End Sub
'
'Private Sub cmdLast_Click()
'    rsCustomer.MoveLast
'    StoreMemVars
'End Sub
'
'Private Sub cmdNext_Click()
'        rsCustomer.MoveNext
'    If rsCustomer.EOF Then
'        rsCustomer.MoveLast
'        ShowLastRecordMsg
'    End If
'    StoreMemVars
'End Sub
'
'Private Sub cmdPrevious_Click()
'    rsCustomer.MovePrevious
'    If rsCustomer.BOF Then
'        rsCustomer.MoveFirst
'        ShowFirstRecordMsg
'    End If
'    StoreMemVars
'End Sub
'
'Private Sub cmdSave_Click()
'    On Error GoTo ErrorCode
'    If IsNull(txtCuscde.Text) = True Then
'        ShowIsRequiredMsg "Code"
'        On Error Resume Next
'        txtCuscde.SetFocus
'        Exit Sub
'    Else
'        If AddorEdit = "ADD" Then
'            Dim rsfindDup                    As ADODB.Recordset
'            Set rsfindDup = New ADODB.Recordset
'            rsfindDup.Open "select * from ALL_CustMaster_Smis where Code = '" & txtCuscde.Text & "'", oConSQL, adOpenForwardOnly, adLockReadOnly
'            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
'                MsgBox "Code already exist!"
'                On Error Resume Next
'                txtCuscde.SetFocus
'                Exit Sub
'            End If
'        End If
'    End If
'    If txtCusnam1.Text = "" Then
'        ShowIsRequiredMsg "Last Name"
'        On Error Resume Next
'        txtCusnam1.SetFocus
'        Exit Sub
'    End If
'    Dim CusCtlSql, NewCtlCde, CustomerNam    As String
'    NewCtlCde = Left(txtCuscde.Text, 1) & Format(NumericVal(MID(txtCuscde.Text, 2, 5)) + 1, "00000")
'    CustomerNam = N2Str2Null(Cap1st(txtCusnam1.Text))
'    Dim VTXTCuscde, VTXTLastName, VTXTFirstname, VTXTMiddleInitial As String
'    Dim vcboSex, vtxtCusadd1, vtxtCusadd2    As String
'    Dim VTXTCuszipc, VTXTCusphon1, vtxtAcctName As String
'
'    Dim vcboApod                             As String
'    Dim vcboLeadSource                       As String
'    Dim vtxtTitle                            As String
'    Dim vtxtDepartment                       As String
'    Dim vtxtEmail                            As String
'    Dim vtxtMobile                           As String
'    Dim vtxtHomePhone                        As String
'    Dim VtxtFax                              As String
'    Dim vtxtAssistant                        As String
'    Dim vtxtAsstPhone                        As String
'    Dim vtxtCity                             As String
'    Dim VtxtBirthDate                        As String
'    Dim VtxtSpouse                           As String
'    Dim VTXTDescription                      As String
'
'    vcboApod = N2Str2Null(cboApod.Text)
'    VTXTCuscde = N2Str2Null(txtCuscde.Text)
'    VTXTLastName = N2Str2Null(txtCusnam1.Text)
'    VTXTFirstname = N2Str2Null(txtCusnam2.Text)
'    VTXTMiddleInitial = N2Str2Null(txtCusnam3.Text)
'    vtxtAcctName = N2Str2Null(txtAcctName.Text)
'
'    vcboSex = N2Str2Null(cboSex.Text)
'    vtxtCusadd1 = N2Str2Null(txtCusadd1.Text)
'    vtxtCusadd2 = N2Str2Null(txtCusadd2.Text)
'    VTXTCuszipc = N2Str2Null(txtCuszipc.Text)
'    VTXTCusphon1 = N2Str2Null(txtCusphon1.Text)
'
'    vcboLeadSource = N2Str2Null(cboLeadSource.Text)
'    vtxtTitle = N2Str2Null(txtTitle.Text)
'    vtxtDepartment = N2Str2Null(txtDepartment.Text)
'    vtxtEmail = N2Str2Null(txtEmail.Text)
'    vtxtMobile = N2Str2Null(txtMobile.Text)
'    vtxtHomePhone = N2Str2Null(txtHomePhone.Text)
'    VtxtFax = N2Str2Null(txtFax.Text)
'    vtxtAssistant = N2Str2Null(txtAssistant.Text)
'    vtxtAsstPhone = N2Str2Null(txtAsstPhone.Text)
'    vtxtCity = N2Str2Null(txtCity.Text)
'    VtxtBirthDate = N2Str2Null(txtBirthDate.Text)
'    VtxtSpouse = N2Str2Null(txtSpouse.Text)
'    VTXTDescription = N2Str2Null(txtDescription.Text)
'
'    If AddorEdit = "ADD" Then
'        Dim rsCustomerDup                    As ADODB.Recordset
'        Set rsCustomerDup = New ADODB.Recordset
'        rsCustomerDup.Open "select * from ALL_CustMaster_Smis where Code <> '999999' order by id asc", oConSQL, adOpenForwardOnly, adLockReadOnly
'        If Not rsCustomerDup.EOF And Not rsCustomerDup.BOF Then
'            rsCustomerDup.MoveLast
'            labid.Caption = NumericVal(rsCustomerDup!ID) + 1
'        End If
'        oConSQL.Execute "Insert into ALL_CustMaster_Smis" & _
         '                      " (Apod,Code,lastname,firstname,middleinitial,ACCTNAME,sex,customeradd,provincialadd,zipcode,telephoneno,LeadSource,Title,Department,Email,Mobile,HomePhone,Fax,Assistant,AsstPhone,City,BirthDate,Spouse,Description)" & _
         '                      " values (" & vcboApod & "," & VTXTCuscde & ", " & VTXTLastName & ", " & VTXTFirstname & ", " & VTXTMiddleInitial & ", " & vtxtAcctName & "," & vcboSex & "," & _
         '                      " " & vtxtCusadd1 & ", " & vtxtCusadd2 & ", " & VTXTCuszipc & ", " & VTXTCusphon1 & "," & vcboLeadSource & "," & vtxtTitle & "," & vtxtDepartment & "," & vtxtEmail & "," & vtxtMobile & "," & vtxtHomePhone & "," & VtxtFax & "," & vtxtAssistant & "," & vtxtAsstPhone & "," & vtxtCity & "," & VtxtBirthDate & "," & VtxtSpouse & "," & VTXTDescription & ")"
'    Else
'        oConSQL.Execute "update ALL_CustMaster_Smis set" & _
         '                      " Apod = " & vcboApod & "," & _
         '                      " lastname = " & VTXTLastName & "," & _
         '                      " firstname = " & VTXTFirstname & "," & _
         '                      " middleinitial = " & VTXTMiddleInitial & "," & _
         '                      " AcctName = " & vtxtAcctName & "," & _
         '                      " sex = " & vcboSex & "," & _
         '                      " customeradd = " & vtxtCusadd1 & "," & _
         '                      " provincialadd = " & vtxtCusadd2 & "," & _
         '                      " zipcode = " & VTXTCuszipc & "," & _
         '                      " LeadSource = " & vcboLeadSource & "," & _
         '                      " Title = " & vtxtTitle & "," & _
         '                      " Department = " & vtxtDepartment & "," & _
         '                      " Email = " & vtxtEmail & "," & _
         '                      " Mobile = " & vtxtMobile & "," & _
         '                      " HomePhone = " & vtxtHomePhone & "," & _
         '                      " Fax = " & VtxtFax & "," & _
         '                      " Assistant = " & vtxtAssistant & "," & _
         '                      " AsstPhone = " & vtxtAsstPhone & "," & _
         '                      " City = " & vtxtCity & "," & _
         '                      " BirthDate = " & VtxtBirthDate & "," & _
         '                      " Spouse = " & VtxtSpouse & "," & _
         '                      " Description = " & VTXTDescription & _
         '                      " where Code = '" & txtCuscde.Text & "'"
'    End If
'    Screen.MousePointer = 11
'    oConSQL.Execute "delete from ALL_CusCtl"
'    Dim k                                    As Integer
'    For k = 65 To 90
'        Set rsCustomer = New ADODB.Recordset
'        rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", oConSQL
'        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
'            NewCtlCde = Chr(k) & Format(NumericVal(MID(rsCustomer!code, 2, 5)) + 1, "00000")
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
'        Else
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
'        End If
'    Next
'    Screen.MousePointer = 0
'    rsRefresh
'    On Error Resume Next
'    rsCustomer.Find "Code =" & VTXTCuscde
'    cmdCancel.Value = True
'    Exit Sub
'
'ErrorCode:
'    ShowVBError
'    Exit Sub
'End Sub
'
'Sub UpdateCusCtl()
'    Dim NewCtlCde                            As String
'    Dim rsCustomer                           As ADODB.Recordset
'    Dim k                                    As Integer
'    Screen.MousePointer = 11
'    oConSQL.Execute "delete from ALL_CusCtl"
'    oConSQL.Execute "delete from ALL_CusCtl"
'    oConSQL.Execute "delete from ALL_CusCtl"
'    oConSQL.Execute "delete from ALL_CusCtl"
'    For k = 65 To 90
'        Set rsCustomer = New ADODB.Recordset
'        rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", oConSQL
'        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
'            NewCtlCde = Chr(k) & Format(NumericVal(MID(rsCustomer!code, 2, 5)) + 1, "00000")
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
'        Else
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
'            oConSQL.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
'        End If
'    Next
'    Screen.MousePointer = 0
'End Sub
'
'Private Sub Command1_Click()
'    UpdateCusCtl
'End Sub
'
'
'
'Private Sub Form_Load()
'    Screen.MousePointer = 11
'    CenterMe frmMain, Me, 1
'    rsRefresh
'    Frame1.Enabled = False
'    InitMemvars
'    txtSearch.Text = ""
'    StoreMemVars
'    Screen.MousePointer = 0
'End Sub
'
'Sub InitMemvars()
'    txtCuscde.Text = ""
'    txtCusnam1.Text = ""
'    txtCusnam2.Text = ""
'    txtCusnam3.Text = ""
'    txtAcctName.Text = ""
'    cboLeadSource.Text = ""
'    cboSex.Text = ""
'    txtTitle.Text = ""
'    txtDepartment.Text = ""
'    txtEmail.Text = ""
'    txtCusphon1.Text = ""
'    txtMobile.Text = ""
'    txtHomePhone.Text = ""
'    txtFax.Text = ""
'    txtAssistant.Text = ""
'    txtAsstPhone.Text = ""
'    txtCusadd1.Text = ""
'    txtCity.Text = ""
'    txtCusadd2.Text = ""
'    txtCuszipc.Text = ""
'    txtBirthDate.Text = ""
'    txtSpouse.Text = ""
'    txtDescription.Text = ""
'    'cboApod.Clear
'    'cboApod.AddItem "Mr."
'    'cboApod.AddItem "Ms."
'    'cboApod.AddItem "Mrs."
'    'cboApod.AddItem "Dr."
'    'cboApod.AddItem "Dra."
'    'cboApod.AddItem "Prof."
'    'cboApod.AddItem "Engr."
'    'cboApod.AddItem "Atty."
'    Dim rsAPOD                               As ADODB.Recordset
'    Set rsAPOD = New ADODB.Recordset
'    Set rsAPOD = oConSQL.Execute("Select distinct apod from ALL_CustMaster_Smis")
'    If Not rsAPOD.EOF And Not rsAPOD.BOF Then
'        rsAPOD.MoveFirst: cboApod.Clear
'        Do While Not rsAPOD.EOF
'            cboApod.AddItem Null2String(rsAPOD!apod)
'            rsAPOD.MoveNext
'        Loop
'    End If
'    Set rsAPOD = Nothing
'    cboLeadSource.Clear
'    cboLeadSource.AddItem "Walk-In"
'    cboLeadSource.AddItem "Phone-In"
'    cboLeadSource.AddItem "Advertisement"
'    cboLeadSource.AddItem "Referral"
'    cboLeadSource.AddItem "Existing"
'    cboSex.Clear
'    cboSex.AddItem "NA"
'    cboSex.AddItem "Male"
'    cboSex.AddItem "Female"
'End Sub
'
'Sub StoreMemVars()
'    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
'        labid.Caption = rsCustomer!ID
'        cboApod.Text = Null2String(rsCustomer!apod)
'        txtCuscde.Text = Null2String(rsCustomer!code)
'        txtCusnam1.Text = Null2String(rsCustomer!LastName)
'        txtCusnam2.Text = Null2String(rsCustomer!FirstName)
'        txtCusnam3.Text = Null2String(rsCustomer!MiddleInitial)
'        cboSex.Text = Null2String(rsCustomer!Sex)
'        txtCusadd1.Text = Null2String(rsCustomer!CustomerAdd)
'        txtCusadd2.Text = Null2String(rsCustomer!provincialadd)
'        txtCuszipc.Text = Null2String(rsCustomer!ZIPCODE)
'        txtCusphon1.Text = Null2String(rsCustomer!telephoneno)
'
'        cboLeadSource.Text = Null2String(rsCustomer!LeadSource)
'        txtTitle.Text = Null2String(rsCustomer!Title)
'        txtDepartment.Text = Null2String(rsCustomer!Department)
'        txtEmail.Text = Null2String(rsCustomer!Email)
'        txtMobile.Text = Null2String(rsCustomer!Mobile)
'        txtHomePhone.Text = Null2String(rsCustomer!HomePhone)
'        txtFax.Text = Null2String(rsCustomer!fax)
'        txtAssistant.Text = Null2String(rsCustomer!Assistant)
'        txtAsstPhone.Text = Null2String(rsCustomer!AsstPhone)
'        txtCity.Text = Null2String(rsCustomer!City)
'        txtBirthDate.Text = Null2String(rsCustomer!BirthDate)
'        txtSpouse.Text = Null2String(rsCustomer!Spouse)
'        txtDescription.Text = Null2String(rsCustomer!Description)
'    Else
'        ShowNoRecord
'        cmdAdd.Value = True
'    End If
'End Sub
'
'Sub rsRefresh()
'    Set rsCustomer = New ADODB.Recordset
'    rsCustomer.Open "select * from ALL_CustMaster_Smis where Code <> '999999' order by firstname+' '+lastname asc", oConSQL, adOpenKeyset
'End Sub
'
'Private Sub txtCusnam1_Change()
'    If AddorEdit = "ADD" Then
'        If IsNumeric(Left(txtCusnam1.Text, 1)) = True Then
'            Set rsCusCtl = New ADODB.Recordset
'            rsCusCtl.Open "select * from ALL_CusCtl where left(ctlcde,1) = 'Z'", oConSQL, adOpenForwardOnly, adLockReadOnly
'            If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
'                txtCuscde.Text = Null2String(rsCusCtl!ctlcde)
'            End If
'        Else
'            Set rsCusCtl = New ADODB.Recordset
'            rsCusCtl.Open "select * from ALL_CusCtl where left(ctlcde,1) = '" & Left(txtCusnam1.Text, 1) & "'", oConSQL, adOpenForwardOnly, adLockReadOnly
'            If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
'                txtCuscde.Text = Null2String(rsCusCtl!ctlcde)
'            End If
'        End If
'    End If
'    If Trim(txtCusnam3.Text) <> "" Then
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text & " " & txtCusnam3.Text
'    Else
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text
'    End If
'End Sub
'
'Private Sub txtCusnam1_LostFocus()
'    If AddorEdit = "ADD" Then
'        Set rsCusCtl = New ADODB.Recordset
'        rsCusCtl.Open "select * from ALL_CusCtl where left(ctlcde,1) = '" & Left(txtCusnam1.Text, 1) & "'", oConSQL, adOpenForwardOnly, adLockReadOnly
'        If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
'            txtCuscde.Text = Null2String(rsCusCtl!ctlcde)
'        End If
'    End If
'End Sub
'
'Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "Code", lstCustomer.SelectedItem.SubItems(1)).Bookmark
'    StoreMemVars
'End Sub
'
'Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    With lstCustomer
'        .Sorted = True
'        If .SortKey = ColumnHeader.Index - 1 Then
'            If .SortOrder = lvwAscending Then
'                .SortOrder = lvwDescending
'            Else
'                .SortOrder = lvwAscending
'            End If
'        Else
'            .SortOrder = lvwAscending
'            .SortKey = ColumnHeader.Index - 1
'        End If
'    End With
'End Sub
'
'Private Sub lstCustomer_DblClick()
'    cmdEdit.Value = True
'End Sub
'
'Private Sub txtCusnam2_Change()
'    If Trim(txtCusnam3.Text) <> "" Then
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text & " " & txtCusnam3.Text
'    Else
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text
'    End If
'End Sub
'
'Private Sub txtCusnam3_Change()
'    If Trim(txtCusnam3.Text) <> "" Then
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text & " " & txtCusnam3.Text
'    Else
'        txtAcctName.Text = txtCusnam1.Text & ", " & txtCusnam2.Text
'    End If
'End Sub
'
'Private Sub txtSearch_Change()
'    If Trim(txtSearch.Text) = "" Then FillGrid Else FillSearchGrid (txtSearch.Text)
'End Sub
'
'Sub FillGrid2()
'    Dim rsCustomer2                          As ADODB.Recordset
'    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
'    Set rsCustomer2 = New ADODB.Recordset
'    Set rsCustomer2 = oConSQL.Execute("select firstname+' '+lastname as CustomerName,Code from ALL_CustMaster_Smis order by firstname+' '+lastname asc")
'    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
'        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
'        lstCustomer.Refresh
'    End If
'End Sub
'
'Sub FillSearchGrid2(XXX As String)
'    Dim rsCustomer2                          As ADODB.Recordset
'    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
'    Set rsCustomer2 = New ADODB.Recordset
'    Set rsCustomer2 = oConSQL.Execute("select firstname+' '+lastname as CustomerName,Code from ALL_CustMaster_Smis where firstname+' '+lastname like'" & XXX & "%' order by firstname+' '+lastname asc")
'    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
'        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
'        lstCustomer.Refresh
'    End If
'End Sub
'
'Private Sub FillGrid()
'    Dim rsCustomer2                          As ADODB.Recordset
'    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
'    Set rsCustomer2 = New ADODB.Recordset
'    Set rsCustomer2 = oConSQL.Execute("select lastname + ', ' + firstname as CustomerName,Code from ALL_CustMaster_Smis where Code <> '999999' order by lastname + ', ' + firstname asc")
'    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
'        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
'        lstCustomer.Refresh
'    End If
'End Sub
'
'Sub FillSearchGrid(XXX As String)
'    Dim rsCustomer2                          As ADODB.Recordset
'    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
'    Set rsCustomer2 = New ADODB.Recordset
'    XXX = Replace(XXX, "'", "")
'    Set rsCustomer2 = oConSQL.Execute("select firstname+' '+lastname as CustomerName,Code from ALL_CustMaster_Smis where Code <> '999999' and lastname + ', ' + firstname like'" & XXX & "%' order by firstname+' '+lastname asc")
'    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
'        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
'        lstCustomer.Refresh
'    End If
'End Sub
