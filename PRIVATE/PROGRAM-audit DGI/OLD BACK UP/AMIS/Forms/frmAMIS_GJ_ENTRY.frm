VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISJournalEntry_GJDetails 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Journal Entry"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   Icon            =   "frmAMIS_GJ_ENTRY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9720
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   7410
      ScaleHeight     =   795
      ScaleWidth      =   2340
      TabIndex        =   61
      Top             =   4590
      Width           =   2340
      Begin VB.CommandButton cmdGJCancel 
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
         Left            =   1500
         MouseIcon       =   "frmAMIS_GJ_ENTRY.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmAMIS_GJ_ENTRY.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmdGJSave 
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
         Left            =   750
         MouseIcon       =   "frmAMIS_GJ_ENTRY.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMIS_GJ_ENTRY.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmdGJDelete 
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
         Left            =   0
         MouseIcon       =   "frmAMIS_GJ_ENTRY.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "frmAMIS_GJ_ENTRY.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.TextBox txtJItemNo 
      Height          =   285
      Left            =   60
      TabIndex        =   60
      Top             =   4560
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picSearchInvoice 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   1290
      ScaleHeight     =   2175
      ScaleWidth      =   3345
      TabIndex        =   34
      Top             =   1860
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtSearchInvoice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         TabIndex        =   52
         Top             =   330
         Width           =   3285
      End
      Begin MSComctlLib.ListView lvwInvNo 
         Height          =   1425
         Left            =   30
         TabIndex        =   35
         Top             =   690
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   2514
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoice Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   30
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   2625
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   60
      ScaleHeight     =   2745
      ScaleWidth      =   9585
      TabIndex        =   13
      Top             =   60
      Width           =   9615
      Begin VB.CommandButton cmdCHANGE_DETAIL 
         Caption         =   "...."
         Height          =   375
         Left            =   5160
         TabIndex        =   59
         Top             =   1290
         Width           =   375
      End
      Begin VB.TextBox txtINVOICE_TYPE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1290
         Width           =   1905
      End
      Begin VB.TextBox txtINVOICE_DETAIL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1290
         Width           =   1935
      End
      Begin VB.TextBox txtOTH_NO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   900
         Width           =   1725
      End
      Begin VB.ComboBox cboJTYPE 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "frmAMIS_GJ_ENTRY.frx":1679
         Left            =   1230
         List            =   "frmAMIS_GJ_ENTRY.frx":1695
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   870
         Width           =   1875
      End
      Begin VB.CheckBox chkOther 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6720
         TabIndex        =   50
         Top             =   930
         Width           =   225
      End
      Begin RichTextLib.RichTextBox txtADJ_Remarks 
         Height          =   915
         Left            =   1200
         TabIndex        =   49
         Top             =   1770
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1614
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmAMIS_GJ_ENTRY.frx":16BF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboCDJNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4380
         TabIndex        =   45
         Top             =   870
         Width           =   1935
      End
      Begin VB.TextBox txtJtype 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8100
         TabIndex        =   39
         Top             =   30
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdVendor 
         Caption         =   "Select Vendor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   60
         Width           =   1845
      End
      Begin VB.CommandButton cmdCustomer 
         Caption         =   "Select Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   1725
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   8295
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   1425
      End
      Begin VB.TextBox txtInvoiceNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2490
         TabIndex        =   19
         Top             =   480
         Width           =   1905
      End
      Begin VB.TextBox txtInvoiceType 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5850
         TabIndex        =   20
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -450
         TabIndex        =   58
         Top             =   1350
         Width           =   1605
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "J. Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -450
         TabIndex        =   54
         Top             =   930
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6090
         TabIndex        =   51
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -150
         TabIndex        =   47
         Top             =   1770
         Width           =   1245
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Press Enter key"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1230
         TabIndex        =   46
         Top             =   600
         Width           =   2265
      End
      Begin VB.Label lblCDJ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Journal No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3285
         TabIndex        =   44
         Top             =   930
         Width           =   1050
      End
      Begin VB.Label labClass 
         BackColor       =   &H000000FF&
         Caption         =   "labClass"
         Height          =   345
         Left            =   6480
         TabIndex        =   36
         Top             =   60
         Visible         =   0   'False
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -60
         TabIndex        =   18
         Top             =   510
         Width           =   1245
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cus. Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -60
         TabIndex        =   14
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label lblInvoiceType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         TabIndex        =   17
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label lblInvoiceNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   16
         Top             =   510
         Width           =   1245
      End
   End
   Begin VB.PictureBox picGJEntry 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   60
      ScaleHeight     =   1665
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   2850
      Width           =   9615
      Begin VB.TextBox txtGJCredit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   33
         Top             =   330
         Width           =   1365
      End
      Begin VB.TextBox txtGJDebit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6780
         TabIndex        =   32
         Top             =   330
         Width           =   1335
      End
      Begin VB.TextBox txtGJAccountName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   330
         Width           =   4425
      End
      Begin VB.PictureBox fraATC 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2370
         ScaleHeight     =   825
         ScaleWidth      =   4335
         TabIndex        =   23
         Top             =   690
         Visible         =   0   'False
         Width           =   4365
         Begin VB.TextBox txtTaxBase2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2550
            MaxLength       =   15
            TabIndex        =   26
            Top             =   360
            Width           =   1725
         End
         Begin VB.TextBox txtRATE2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   25
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox cboATC2 
            BackColor       =   &H00F1F6F5&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00973640&
            Height          =   330
            Left            =   60
            TabIndex        =   24
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label49 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Taxbase Amt."
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
            Height          =   225
            Left            =   2550
            TabIndex        =   30
            Top             =   90
            Width           =   1725
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RATE"
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
            Height          =   225
            Left            =   1380
            TabIndex        =   29
            Top             =   90
            Width           =   855
         End
         Begin VB.Label Label47 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "ATC Code"
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
            Height          =   225
            Left            =   120
            TabIndex        =   28
            Top             =   90
            Width           =   1365
         End
         Begin VB.Label Label46 
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
            Height          =   225
            Left            =   2190
            TabIndex        =   27
            Top             =   390
            Width           =   855
         End
      End
      Begin VB.ComboBox cboJVSupCust 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   2340
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   2100
         Width           =   4305
      End
      Begin RichTextLib.RichTextBox txtGJAccountParticulars 
         Height          =   885
         Left            =   2280
         TabIndex        =   5
         Top             =   3780
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   1561
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   0   'False
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmAMIS_GJ_ENTRY.frx":173B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboGJAccountNo 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Text            =   "cboGJAccountNo"
         Top             =   330
         Width           =   2235
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Press Enter key"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   330
         TabIndex        =   48
         Top             =   660
         Width           =   2265
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   225
         Left            =   2400
         TabIndex        =   12
         Top             =   60
         Width           =   2205
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
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
         Height          =   255
         Left            =   390
         TabIndex        =   11
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
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
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
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
         Height          =   225
         Left            =   6810
         TabIndex        =   9
         Top             =   60
         Width           =   885
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Height          =   255
         Left            =   8160
         TabIndex        =   8
         Top             =   60
         Width           =   795
      End
      Begin VB.Label labGJID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   225
         Left            =   1230
         TabIndex        =   7
         Top             =   4020
         Width           =   2205
      End
      Begin VB.Label labATC 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier :"
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
         Height          =   225
         Left            =   1170
         TabIndex        =   6
         Top             =   2160
         Width           =   1305
      End
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   60
      ScaleHeight     =   3705
      ScaleWidth      =   9585
      TabIndex        =   40
      Top             =   60
      Visible         =   0   'False
      Width           =   9615
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   2955
         Left            =   30
         TabIndex        =   41
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5212
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmAMIS_GJ_ENTRY.frx":17D2
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   11819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TYPE"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.TextBox txtSearchAccount 
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
         MaxLength       =   50
         TabIndex        =   2
         Top             =   330
         Width           =   9465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   43
         Top             =   30
         Width           =   405
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Chart Accounts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   42
         Top             =   60
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmAMISJournalEntry_GJDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                     As String
Dim WHAT_CLASS                                    As String
Dim ENTITYCODE                                    As String
Dim xJOURNALTYPE                                  As String
Dim rsATC                                         As ADODB.Recordset

Sub xADDorEDIT(xADDorEDIT As String)
    AddorEdit = xADDorEDIT
End Sub

Private Sub cboATC2_Click()
'UPDATED: ACL 06252010
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC2.Text))
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE2.Text = N2Str2Zero(rsATC!Rate)
        If NumericVal(txtRATE2.Text) > 0 Then
            txtGJCredit.Text = Round(NumericVal(txtTaxBase2.Text) * (NumericVal(txtRATE2.Text) / 100), 2)
        End If
    End If
    Set rsATC = Nothing
End Sub

Private Sub cboCDJNo_Click()
    Call GET_VOUCHERNO_JTYPE(cboCDJNo.Text, cboJTYPE.Text)
End Sub

Private Sub cboCDJNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboJTYPE.Text <> "CDJ" And cboJTYPE.Text <> "GJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "CRJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "SJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "COB" And cboJTYPE.Text <> "OTH" And cboJTYPE.Text <> "VPJ" Then Exit Sub

        If cboCDJNo.Text = "" Then
            txtInvoiceNo.Enabled = True
            txtInvoiceType.Enabled = True
            chkOther.Enabled = True
        Else
            Call GET_VOUCHERNO_JTYPE(cboCDJNo.Text, cboJTYPE.Text)
            txtInvoiceNo.Enabled = False
            txtInvoiceType.Enabled = False
            chkOther.Enabled = False
        End If

        If cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
            Call FIND_SUB_DETAIL_REFERENCE
        End If

        'txtSearchInvoice.SetFocus

    End If
End Sub

Function GET_VOUCHERNO_JTYPE(xVOUCHERNO As String, xJType As String) As String
    Dim rsVOUCHERNO_JTYPE                         As ADODB.Recordset
    Dim rsENTITY_NAME                             As ADODB.Recordset
    Set rsVOUCHERNO_JTYPE = New ADODB.Recordset
    If xJType = "CDJ" Or xJType = "APJ" Or xJType = "VPJ" Then
        rsVOUCHERNO_JTYPE.Open "SELECT VENDORCODE AS CODE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO='" & xVOUCHERNO & "' AND JTYPE= '" & xJType & "' AND STATUS='P'", gconDMIS, adOpenForwardOnly
        ENTITYCODE = "V"
    ElseIf xJType = "CRJ" Or xJType = "SJ" Or xJType = "COB" Then
        rsVOUCHERNO_JTYPE.Open "SELECT CUSTOMERCODE AS CODE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO='" & xVOUCHERNO & "' AND JTYPE= '" & xJType & "' AND STATUS='P'", gconDMIS, adOpenForwardOnly
        ENTITYCODE = "C"
    Else
        rsVOUCHERNO_JTYPE.Open "SELECT CUSTOMERCODE AS CODE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO='" & xVOUCHERNO & "' AND JTYPE= '" & xJType & "' AND STATUS='P'", gconDMIS, adOpenForwardOnly
        ENTITYCODE = ""
    End If
    If Not rsVOUCHERNO_JTYPE.EOF And Not rsVOUCHERNO_JTYPE.BOF Then
        If ENTITYCODE <> "" Then
            txtCode.Text = N2String(rsVOUCHERNO_JTYPE!Code)
            labClass = ENTITYCODE & txtCode.Text
        End If
        Set rsENTITY_NAME = New ADODB.Recordset
        rsENTITY_NAME.Open "SELECT ACCOUNTNAME FROM ALL_ENTITY WHERE COMPLET_CODE='" & Null2String(labClass) & "' ", gconDMIS, adOpenForwardOnly
        If Not rsENTITY_NAME.EOF And Not rsENTITY_NAME.BOF Then
            txtName.Text = N2String(rsENTITY_NAME!AccountName)
        End If
    End If
End Function

Private Sub cboCDJNo_LostFocus()
    If cboJTYPE.Text <> "CDJ" And cboJTYPE.Text <> "GJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "CRJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "SJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "COB" And cboJTYPE.Text <> "OTH" And cboJTYPE.Text <> "VPJ" Then Exit Sub

    If cboCDJNo.Text = "" Then
        txtInvoiceNo.Enabled = True
        txtInvoiceType.Enabled = True
        chkOther.Enabled = True
    Else
        Call GET_VOUCHERNO_JTYPE(cboCDJNo.Text, cboJTYPE.Text)
        txtInvoiceNo.Enabled = False
        txtInvoiceType.Enabled = False
        chkOther.Enabled = False
    End If

    If cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
        Call FIND_SUB_DETAIL_REFERENCE
    End If

    'txtSearchInvoice.SetFocus
End Sub

Private Sub cboGJAccountNo_Change()
'VALIDATE IF THE ACCT_CODE IS A SCHEDULED ACCOUNT IF YES CONTROL NO. WILL BE NEEDED
    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String
    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")


    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    GettheTaxBaseAmnt
    If cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        On Error Resume Next
        cboATC2.SetFocus
    Else
        fraATC.Visible = False
    End If

    If cboGJAccountNo.Text = "" Then
        If AddorEdit = "ADD" Then
            txtOTH_NO.Text = ""
        End If
        txtGJAccountName.Text = ""
    Else
        If AddorEdit = "ADD" Then
            If CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True And chkOther.Value = 1 Then
                Call GET_OTH_MAX_NO
            ElseIf cboCDJNo.Text = "" And cboJTYPE.Text = "" And CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True Then
                chkOther.Value = 1
                Call GET_OTH_MAX_NO
            Else
                txtOTH_NO.Text = ""
            End If
        End If
    End If
End Sub

Private Sub InitChart()
'DESCRIPTION: INITIALIZE THE CBO WITH CHART OF ACCOUNT CODE
    Dim rsInitAccountCode                         As ADODB.Recordset
    Set rsInitAccountCode = New ADODB.Recordset
    rsInitAccountCode.Open "Select AcctCode from Amis_ChartAccount order by AcctCode asc", gconDMIS, adOpenKeyset
    'cboGJAccountNo.Clear
    If Not rsInitAccountCode.EOF And Not rsInitAccountCode.BOF Then
        Do While Not rsInitAccountCode.EOF
            cboGJAccountNo.AddItem Null2String(rsInitAccountCode!ACCTCODE)
            rsInitAccountCode.MoveNext
        Loop
    End If
    Set rsInitAccountCode = Nothing
End Sub

Private Sub cboGJAccountNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        picChart.Visible = True
        picChart.ZOrder 0
        lstAccounts.Enabled = True
        FILL_CHARTACCOUNT
    End If
End Sub

Private Sub cboGJAccountNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl.Name = "cboGJAccountNo" Then
        On Error Resume Next
        txtSearchAccount.SetFocus
    End If
End Sub

Private Sub cboGJAccountNo_LostFocus()
    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String
    DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    If cboGJAccountNo.Text = DEALER_ITW_COMPENSATION Or cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
    Else
        fraATC.Visible = False
    End If
    '    cboGJAccountNo_Click

    If cboGJAccountNo.Text = "" Then
        If AddorEdit = "ADD" Then
            txtOTH_NO.Text = ""
        End If
        txtGJAccountName.Text = ""
    Else
        If AddorEdit = "ADD" Then
            If CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True And chkOther.Value = 1 Then
                Call GET_OTH_MAX_NO
            ElseIf cboCDJNo.Text = "" And cboJTYPE.Text = "" And CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True Then
                chkOther.Value = 1
                Call GET_OTH_MAX_NO
            Else
                txtOTH_NO.Text = ""
            End If
        End If
    End If

End Sub

Private Sub cboJTYPE_Change()
'Call FIND_JOURNAL_NO(RTrim(LTrim(cboJTYPE.Text)), RTrim(LTrim(txtCode.Text)))
    If cboJTYPE.Text = "" Then
    Else
        If cboJTYPE.Text <> "CDJ" And cboJTYPE.Text <> "GJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "CRJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "SJ" And cboJTYPE.Text <> "APJ" And cboJTYPE.Text <> "COB" And cboJTYPE.Text <> "OTH" And cboJTYPE.Text <> "VPJ" Then Exit Sub
        cboJTYPE.Text = UCase(cboJTYPE.Text)
        Call FIND_JOURNAL_NO(RTrim(LTrim(cboJTYPE.Text)))
    End If
End Sub

Private Sub cboJTYPE_Click()
    cboCDJNo.Clear
    'Call FIND_JOURNAL_NO(RTrim(LTrim(cboJTYPE.Text)), RTrim(LTrim(txtCode.Text)))
    Call FIND_JOURNAL_NO(RTrim(LTrim(cboJTYPE.Text)))
    txtJtype.Text = RTrim(LTrim(cboJTYPE.Text))
End Sub

Private Sub cboJTYPE_LostFocus()
    If cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
        Call FIND_SUB_DETAIL_REFERENCE
    End If
End Sub

Private Sub chkOther_Click()
    If chkOther.Value = 0 Then
        txtInvoiceNo.Enabled = True
        txtInvoiceType.Enabled = True
        cboCDJNo.Enabled = True
        cboJTYPE.Enabled = True
        txtOTH_NO.Text = ""
        txtJtype.Text = ""
    Else
        txtInvoiceNo.Enabled = False
        txtInvoiceType.Enabled = False
        cboCDJNo.Enabled = False
        cboJTYPE.Enabled = False
        cboCDJNo.Text = ""
        cboJTYPE.ListIndex = -1
        txtJtype.Text = "OTH"

        If AddorEdit = "ADD" Then
            If CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True And chkOther.Value = 1 Then
                GET_OTH_MAX_NO
            End If
        Else
            Dim rsGET_CON_NUM                     As ADODB.Recordset
            Set rsGET_CON_NUM = New ADODB.Recordset
            rsGET_CON_NUM.Open "SELECT INVOICENO FROM AMIS_JOURNAL_DET WHERE ID = '" & frmAMISJournalEntry_GJ.labDET.Caption & "' and IS_OTHERS = 1", gconDMIS, adOpenKeyset
            If Not rsGET_CON_NUM.EOF And Not rsGET_CON_NUM.BOF Then
                If IsNull(rsGET_CON_NUM!INVOICENO) = False Then
                    txtOTH_NO.Text = Null2String(rsGET_CON_NUM!INVOICENO)
                Else
                    GET_OTH_MAX_NO
                End If
            Else
                GET_OTH_MAX_NO
            End If
            Set rsGET_CON_NUM = Nothing
        End If
    End If
End Sub

Private Sub cmdCHANGE_DETAIL_Click()
'If txtINVOICE_DETAIL.Text <> "" And txtINVOICE_TYPE.Text <> "" Then
    FIND_SUB_DETAIL_REFERENCE
    picSearchInvoice.Visible = True
    picSearchInvoice.ZOrder 0
    'End If
End Sub

Private Sub cmdCustomer_Click()
    xJOURNALTYPE = "GJ"
    SelectEntity = "Customer"
    frmEntity.Caption = "SEARCH CUSTOMER"
    ENTITYCODE = "C"
    Call frmEntity.LoadJournal("GJ")
    frmEntity.Show
    frmEntity.txtSearch.SetFocus
End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdGJCancel_Click()
    frmAMISJournalEntry_GJ.StoreSearch (frmAMISJournalEntry_GJ.txtVoucherNo.Text)
    Unload Me
End Sub

Private Sub cmdGJDelete_Click()
    If MsgBox("Are you sure you want to Delete this Journal entry?", vbQuestion + vbYesNo, "Warning") = vbYes Then
        gconDMIS.Execute "Delete From Amis_Journal_det where ID = '" & frmAMISJournalEntry_GJ.labDET.Caption & "'"

        Dim cnt                                   As Integer
        Dim rsJournalDup                          As ADODB.Recordset
        Set rsJournalDup = New ADODB.Recordset
        rsJournalDup.Open "select id,JItemno,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(frmAMISJournalEntry_GJ.txtVoucherNo.Text) & " order by ID asc", gconDMIS
        If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
            rsJournalDup.MoveFirst
            cnt = 0
            Do While Not rsJournalDup.EOF
                cnt = cnt + 1
                SQL_STATEMENT = "update AMIS_Journal_Det set JItemno = '" & Format(cnt, "0000") & "' where id = " & rsJournalDup!ID
                gconDMIS.Execute SQL_STATEMENT
                rsJournalDup.MoveNext
                NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, frmAMISJournalEntry_GJ.labDET.Caption, "", frmAMISJournalEntry_GJ.txtVoucherNo, xJOURNALTYPE, frmAMISJournalEntry_GJ.txtJNo
            Loop
        End If
        MessagePop Delete, "INFORMATION", "Record Succesfully Deleted"
        frmAMISJournalEntry_GJ.StoreSearch (frmAMISJournalEntry_GJ.txtVoucherNo.Text)
        Unload Me
    End If
End Sub

Function DetailPosting() As Boolean
    On Error GoTo ErrorCode

    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String

    DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

    If cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        If cboATC2.Text = "" Then
            MsgBox "ATC Code must have a value", vbInformation, "System Message!"
            DetailPosting = True
            Exit Function
        End If
    End If

    If CHECK_IF_AR_SCHED(RTrim(LTrim(cboGJAccountNo.Text))) = True Then
        If txtCode.Text = "" Then
            If SelectEntity = "Customer" Then
                MessagePop InfoFriend, "INFORMATION", "Customer Code must have a value"
                DetailPosting = True
                Exit Function
            Else
                MessagePop InfoFriend, "INFORMATION", "Vendor Code must have a value"
                DetailPosting = True
                Exit Function
            End If
        End If

        If txtName.Text = "" Then
            If SelectEntity = "Customer" Then
                MessagePop InfoFriend, "INFORMATION", "Customer Name must have a value"
                DetailPosting = True
                Exit Function
            Else
                MessagePop InfoFriend, "INFORMATION", "Vendor Name must have a value"
                DetailPosting = True
                Exit Function
            End If
        End If
        If chkOther.Value = 1 Then
            If txtOTH_NO.Text = "" Then
                MessagePop InfoFriend, "INFORMATION", "Control # must have a value."
                DetailPosting = True
                Exit Function
            End If
        Else
            If cboCDJNo.Text = "" And cboJTYPE.Text = "" Then
                MessagePop InfoFriend, "INFORMATION", "Journal no. and Journal type must have a value."
                DetailPosting = True
                Exit Function
            End If
        End If
    Else
        'ALLOW TO SAVE THIS IS NOT AN AR SCHEDULE
    End If

    If cboGJAccountNo.Text = "" Then
        MessagePop InfoFriend, "INFORMATION", "Account Code must have a value."
        DetailPosting = True
        Exit Function
    End If

    If txtGJAccountName.Text = "" Then
        MessagePop InfoFriend, "INFORMATION", "Account Name must have a value."
        DetailPosting = True
        Exit Function
    End If

    If NumericVal(txtGJCredit.Text) = 0 And NumericVal(txtGJDebit.Text) = 0 Then
        MessagePop InfoFriend, "INFORMATION", "Debit or Credit must have a value."
        DetailPosting = True
        Exit Function
    End If

    If NumericVal(txtGJDebit.Text) > 0 And NumericVal(txtGJCredit.Text) > 0 Then
        MessagePop InfoFriend, "INFORMATION", "Invalid Journal Entry! Debit and Credit Amount can not both have an amount!"
        DetailPosting = True
        Exit Function
    End If

    'VALIDATE THE TYPE OF ADJUSTMENT THAT THERE SHOULD INVOICE OR CDJ NO OR OTHERS

    If cboJTYPE.Text = "" And cboCDJNo.Text = "" And chkOther.Value = 0 Then
        MessagePop InfoFriend, "INFORMATION", "Journal No. and Jtype or Others must have a value!"
        DetailPosting = True
        Exit Function
    End If


    'REMARKS IS REQUIRED IF IT IS ACCOUNT SCHEDULE

    If COMPANY_CODE = "HMH" Then
    Else
        If txtADJ_Remarks.Text = "" Then
            If CHECK_IF_AR_SCHED(RTrim(LTrim(cboGJAccountNo.Text))) = True Then
                MessagePop InfoFriend, "INFORMATION", "Remarks is required because you are adding a schedule account"
                DetailPosting = True
                Exit Function
            Else
                'ALLOW TO SAVE THE ENTRY
            End If
        End If
    End If

    If AddorEdit = "ADD" Then
        If ALREADY_ADJUSTED(txtInvoiceNo.Text, txtInvoiceType.Text, txtJtype) = True Then
            If MsgBox("This is entry was already adjusted." & vbCrLf & "Are you sure you want to proceed?", vbQuestion + vbYesNo, "INFORMATION") = vbYes Then
                'PROCEED
            Else
                'MessagePop InfoFriend, "INFORMATION", "This is entry was already adjusted"
                DetailPosting = True
                Exit Function
            End If

        End If
    End If

    'VALIDATE ACCT_CODE IF IT VALID

    If CHECK_ACCOUNTCODE(cboGJAccountNo.Text) = False Then
        MessagePop InfoFriend, "INFORMATION", "Please check account code invalid."
        DetailPosting = True
        Exit Function
    End If


    'VALIDATE THE JOURNAL NO IF EXISTING
    If cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
        Dim rsJOURNAL                             As ADODB.Recordset
        Set rsJOURNAL = New ADODB.Recordset
        If cboJTYPE.Text = "OTH" Then
            rsJOURNAL.Open "SELECT * FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = '" & cboCDJNo.Text & "' AND ADJ_JTYPE = '" & cboJTYPE.Text & "'", gconDMIS, adOpenKeyset
        Else
            rsJOURNAL.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = '" & cboJTYPE.Text & "'", gconDMIS, adOpenKeyset
        End If
        If Not rsJOURNAL.EOF And Not rsJOURNAL.BOF Then
        Else
            MessagePop InfoFriend, "INFORMATION", "Please check your Journal No. and Type not found."
            DetailPosting = True
            Exit Function
        End If
        Set rsJOURNAL = Nothing

        'TEMPORARY COMMENTED BY: JUN 12-10-2009-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'VALIDATE THE ACCT CODE OF THE IT IS MATCH TO THE ACCOUNT CODE OF THE VOUCHER THEY ARE ADJUSTING
        '                Dim rsVALIDATE_CODE As ADODB.Recordset
        '                    Set rsVALIDATE_CODE = New ADODB.Recordset
        '                        If cboJTYPE.Text = "OTH" Then
        '                            rsVALIDATE_CODE.Open "SELECT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = '" & cboGJAccountNo.Text & "' AND ADJ_VOUCHERNO = '" & cboCDJNo.Text & "' AND ADJ_JTYPE = '" & cboJTYPE.Text & "'", gconDMIS, adOpenKeyset
        '                        Else
        '                            rsVALIDATE_CODE.Open "Select ACCT_CODE FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = '" & cboGJAccountNo.Text & "' AND VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = '" & cboJTYPE.Text & "'", gconDMIS, adOpenKeyset
        '                        End If
        '                        If rsVALIDATE_CODE.EOF And rsVALIDATE_CODE.BOF Then
        '                           If CHECK_IF_AR_SCHED(RTrim(LTrim(cboGJAccountNo.Text))) = True And NumericVal(txtGJDebit.Text) <> 0 Then
        '                           Else
        '                                MessagePop InfoFriend, "INFORMATION", "Account Code did not macth to the acct code reference no."
        '                                exit function
        '                           End If
        '                        End If
        '                Set rsVALIDATE_CODE = Nothing
        'TEMPORARY COMMENTED BY: JUN 12-10-2009-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End If

    'VALIDATE IF VOUCHERNO NO. HAS ENTRY AND JTYPE HAS NO ENTRY
    If cboCDJNo.Text <> "" And cboJTYPE.Text = "" Then
        MessagePop InfoFriend, "INFORMATION", "Please check your Journal No. and Type not found."
        DetailPosting = True
        Exit Function
    End If

    If cboCDJNo.Text = "" And cboJTYPE.Text <> "" Then
        MessagePop InfoFriend, "INFORMATION", "Please check your Journal No. and Type not found."
        DetailPosting = True
        Exit Function
    End If

    'UPDATED BY: JUN -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DATE UPDATED: 11/11/2009
    'DESCRIPTION: VALIDATE THE SUM AR AMOUNT IN GJ AND CHECK IF IT MATCH TO THE AMOUNT OF THE REFERENCE NO
    If CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True Then
        If cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
            Dim rsAMOUNT                          As ADODB.Recordset
            Dim XXX_SUM                           As Double
            XXX_SUM = 0
            Set rsAMOUNT = New ADODB.Recordset
            rsAMOUNT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_DEBIT, ROUND(SUM(CREDIT),2) AS SUM_CREDIT " & _
                          "FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(cboCDJNo.Text) & " AND JTYPE = " & N2Str2Null(cboJTYPE.Text) & " AND ACCT_CODE = " & N2Str2Null(cboGJAccountNo.Text) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
            If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
                If NumericVal(rsAMOUNT!SUM_DEBIT) <> 0 Then
                    XXX_SUM = NumericVal(rsAMOUNT!SUM_DEBIT)
                ElseIf NumericVal(rsAMOUNT!SUM_CREDIT) <> 0 Then
                    XXX_SUM = NumericVal(rsAMOUNT!SUM_CREDIT)
                End If
            End If
            Set rsAMOUNT = Nothing

            Dim rsCOMPUTE_AMOUNT                  As ADODB.Recordset
            Dim GJ_XXX_SUM                        As Double
            GJ_XXX_SUM = 0
            Set rsCOMPUTE_AMOUNT = New ADODB.Recordset
            If NumericVal(txtGJDebit.Text) <> 0 Then
                rsCOMPUTE_AMOUNT.Open "SELECT ROUND(SUM(DEBIT),2) AS GJ_SUM " & _
                                      "FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & N2Str2Null(cboCDJNo.Text) & " AND ADJ_JTYPE = " & N2Str2Null(cboJTYPE.Text) & " AND ACCT_CODE = " & N2Str2Null(cboGJAccountNo.Text) & " AND ID <> " & NumericVal(frmAMISJournalEntry_GJ.labDET.Caption) & "", gconDMIS, adOpenKeyset
            ElseIf NumericVal(txtGJCredit.Text) <> 0 Then
                rsCOMPUTE_AMOUNT.Open "SELECT ROUND(SUM(CREDIT),2) AS GJ_SUM " & _
                                      "FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & N2Str2Null(cboCDJNo.Text) & " AND ADJ_JTYPE = " & N2Str2Null(cboJTYPE.Text) & " AND ACCT_CODE = " & N2Str2Null(cboGJAccountNo.Text) & " AND ID <> " & NumericVal(frmAMISJournalEntry_GJ.labDET.Caption) & "", gconDMIS, adOpenKeyset
            End If

            If Not rsCOMPUTE_AMOUNT.EOF And Not rsCOMPUTE_AMOUNT.BOF Then
                If NumericVal(txtGJDebit.Text) <> 0 Then
                    GJ_XXX_SUM = Round((NumericVal(rsCOMPUTE_AMOUNT!GJ_SUM) + NumericVal(txtGJDebit.Text)), 2)
                ElseIf NumericVal(txtGJCredit.Text) <> 0 Then
                    GJ_XXX_SUM = Round((NumericVal(rsCOMPUTE_AMOUNT!GJ_SUM) + NumericVal(txtGJCredit.Text)), 2)
                End If
            End If
            Set rsCOMPUTE_AMOUNT = Nothing

            If XXX_SUM <> 0 Then
                If NumericVal(GJ_XXX_SUM) > NumericVal(XXX_SUM) Then
                    MessagePop InfoFriend, "INFORMATION", "Sum of AR amount is greater than to the AR amount of the reference no."
                    DetailPosting = True
                    Exit Function
                End If
            End If
        End If
    End If
    'UPDATED BY: JUN -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim MULTI_DETAIL                              As Boolean
    MULTI_DETAIL = False
    'THIS IS TO VALIDATE IF THE VOUCHER NO HAS A MULTI LINE DETAIL
    Dim rsCHECK_MULTI_DETAIL                      As ADODB.Recordset
    Set rsCHECK_MULTI_DETAIL = New ADODB.Recordset
    If cboJTYPE.Text = "CRJ" Then
        rsCHECK_MULTI_DETAIL.Open "SELECT COUNT(VOUCHERNO) AS COUNT_VOUCHERNO FROM AMIS_CRJ_DETAIL ACD INNER JOIN AMIS_CHARTACCOUNT AC ON ACD.J_CLASS=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO = '" & cboCDJNo.Text & "' AND CR_TYPE = '" & cboJTYPE.Text & "'", gconDMIS, adOpenKeyset
    ElseIf cboJTYPE.Text = "CDJ" Then
        rsCHECK_MULTI_DETAIL.Open "SELECT COUNT(VOUCHERNO) AS COUNT_VOUCHERNO FROM AMIS_CV_DETAIL ACD INNER JOIN AMIS_CHARTACCOUNT AC ON ACD.J_CLASS=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = 'APJ'", gconDMIS, adOpenKeyset
    End If

    If cboJTYPE.Text = "CRJ" Or cboJTYPE.Text = "CDJ" Then
        If Not rsCHECK_MULTI_DETAIL.EOF And Not rsCHECK_MULTI_DETAIL.BOF Then
            If txtINVOICE_DETAIL.Text = "" And txtINVOICE_TYPE.Text = "" Then
                If NumericVal(rsCHECK_MULTI_DETAIL!COUNT_VOUCHERNO) > 1 Then
                    cmdCHANGE_DETAIL_Click
                    DetailPosting = True
                    MULTI_DETAIL = True
                    Exit Function
                End If
            End If
        End If
    End If
    Set rsCHECK_MULTI_DETAIL = Nothing

    Dim J_JDATE                                   As String
    Dim J_VOUCHERNO                               As String
    Dim J_JTYPE                                   As String
    Dim J_JNO                                     As String
    Dim J_ACCT_CODE                               As String
    Dim J_ACCT_NAME                               As String
    Dim J_STATUS                                  As String
    Dim J_JITEMNO                                 As String
    Dim xCUSCODE                                  As String
    Dim xCUSNAME                                  As String
    Dim xENTITY_CLASS                             As String

    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String

    Dim xAdj_type                                 As String
    Dim xADJ_VOUCHERNO                            As String

    Dim xIS_OTHERS                                As Integer
    Dim xADJ_REMARKS                              As String

    Dim J_DEBIT                                   As Double
    Dim J_CREDIT                                  As Double
    Dim J_TAX                                     As Double

    J_JDATE = N2Date2Null(frmAMISJournalEntry_GJ.txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(frmAMISJournalEntry_GJ.txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(GetJNo(frmAMISJournalEntry_GJ.txtVoucherNo.Text))
    'J_JITEMNO = N2Str2Null(GetItemNO(frmAMIS_GJ_JOURNAL_ENTRY.txtVoucherNo.Text))
    J_ACCT_CODE = N2Str2Null(cboGJAccountNo.Text)
    J_ACCT_NAME = N2Str2Null(txtGJAccountName.Text)
    J_DEBIT = Round(NumericVal(txtGJDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtGJCredit.Text), 2)
    'J_TAX = Round(NumericVal(txtTax.Text), 2)
    J_TAX = Round(NumericVal(txtTaxBase2.Text), 2)
    J_STATUS = "'N'"

    'xADJ_REMARKS = N2Str2Null(RTrim(LTrim(Replace(txtADJ_Remarks.Text, vbCrLf, ""))))
    xADJ_REMARKS = N2Str2Null(RTrim(LTrim(txtADJ_Remarks.Text)))

    If chkOther.Value = 1 Then
        xIS_OTHERS = 1
    Else
        xIS_OTHERS = 0
    End If

    If chkOther.Value = 1 Then
        'THIS IS FOR ADJUSTMENT WHICH HAS NO DETAIL

        If CHECK_IF_AR_SCHED(cboGJAccountNo.Text) = True Then
            'xADJ_VOUCHERNO = N2Str2Null(Format(NumericVal(txtOTH_NO.Text), "000000"))
            xADJ_VOUCHERNO = Format(N2Str2Null(txtOTH_NO.Text), "000000")
        Else
            xADJ_VOUCHERNO = N2Str2Null("")
        End If
    Else
        xADJ_VOUCHERNO = N2Str2Null(cboCDJNo.Text)
    End If
    '        If cboCDJNo.Text = "" Then
    '            xINVOICENO = N2Str2Null(txtInvoiceNo.Text)
    '        Else
    '            xINVOICENO = N2Str2Null(cboCDJNo.Text)
    '        End If


    xINVOICENO = N2Str2Null(txtINVOICE_DETAIL.Text)
    xInvoiceType = N2Str2Null(txtINVOICE_TYPE.Text)

    If MULTI_DETAIL = False Then
        'UPDATE BY: ACL 9202010
        'Description: Check the account code of the Reference Voucher if equal with GJ Entry
        Dim rsCheckReference                      As ADODB.Recordset
        Dim rsCheckEntity                         As ADODB.Recordset
        Set rsCheckReference = New ADODB.Recordset
        If NumericVal(txtGJDebit.Text) > 0 Then
            'rsCheckReference.Open "SELECT DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.VOUCHERNO='" & cboCDJNo.Text & "' AND HD.JTYPE='" & cboJTYPE.Text & "' AND ISNULL(DET.CREDIT,0) > 0 AND DET.ACCT_CODE='" & cboGJAccountNo.Text & "' AND HD.STATUS='P'", gconDMIS, adOpenForwardOnly
        'Updated by al for non-shecduled in GJ to be closed
            '    rsCheckReference.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.JTYPE,DET.ACCT_CODE,HD.STATUS,IS_SCHEDULE_ACCNT,CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY ELSE DET.CREDIT END AS CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE) T WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO='" & cboCDJNo.Text & "' AND JTYPE='" & cboJTYPE.Text & "' AND ISNULL(CREDIT,0) > 0 AND STATUS='P'", gconDMIS, adOpenForwardOnly
            rsCheckReference.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.JTYPE,DET.ACCT_CODE,HD.STATUS,IS_SCHEDULE_ACCNT,CASE WHEN HD.JTYPE='VPJ' THEN HD.AMOUNTTOPAY ELSE DET.CREDIT END AS CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE) T WHERE VOUCHERNO='" & cboCDJNo.Text & "' AND JTYPE='" & cboJTYPE.Text & "' AND ISNULL(CREDIT,0) > 0 AND STATUS='P' AND ACCT_CODE='" & cboGJAccountNo.Text & "'", gconDMIS, adOpenForwardOnly
        Else
            'rsCheckReference.Open "SELECT DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.VOUCHERNO='" & cboCDJNo.Text & "' AND HD.JTYPE='" & cboJTYPE.Text & "' AND ISNULL(DET.DEBIT,0) > 0 AND DET.ACCT_CODE='" & cboGJAccountNo.Text & "' AND HD.STATUS='P'", gconDMIS, adOpenForwardOnly
            'rsCheckReference.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.JTYPE,DET.ACCT_CODE,HD.STATUS,IS_SCHEDULE_ACCNT,CASE WHEN HD.JTYPE='COB' THEN HD.INVOICEAMT ELSE DET.DEBIT END AS DEBIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE) T WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO='" & cboCDJNo.Text & "' AND JTYPE='" & cboJTYPE.Text & "' AND ISNULL(DEBIT,0) > 0 AND STATUS='P'", gconDMIS, adOpenForwardOnly
        'Updated by al for non-shecduled in GJ to be closed
            rsCheckReference.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.JTYPE,DET.ACCT_CODE,HD.STATUS,IS_SCHEDULE_ACCNT,CASE WHEN HD.JTYPE='COB' THEN HD.INVOICEAMT ELSE DET.DEBIT END AS DEBIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE) T WHERE VOUCHERNO='" & cboCDJNo.Text & "' AND JTYPE='" & cboJTYPE.Text & "' AND ISNULL(DEBIT,0) > 0 AND STATUS='P' AND ACCT_CODE='" & cboGJAccountNo.Text & "'", gconDMIS, adOpenForwardOnly
        
        End If
        If Not rsCheckReference.EOF And Not rsCheckReference.BOF Then
            If cboGJAccountNo.Text <> rsCheckReference!Acct_code Then
                '            Else
                MessagePop InfoFriend, "INFORMATION", "Please check account code, not equal to reference Voucher No." & cboJTYPE.Text + "-" + cboCDJNo.Text
                DetailPosting = True
                Exit Function
            End If
        End If
        Set rsCheckReference = Nothing

        Set rsCheckEntity = New ADODB.Recordset
        rsCheckEntity.Open "SELECT ENTITY=(CASE WHEN JTYPE IN ('SJ','CRJ','COB') THEN CUSTOMERCODE WHEN JTYPE IN ('APJ','CDJ','VPJ') THEN VENDORCODE END) FROM AMIS_JOURNAL_HD WHERE VOUCHERNO='" & cboCDJNo.Text & "' AND JTYPE='" & cboJTYPE.Text & "'", gconDMIS, adOpenForwardOnly
        If Not rsCheckEntity.EOF And Not rsCheckEntity.BOF Then
            If rsCheckEntity!ENTITY <> txtCode.Text Then
                MessagePop InfoFriend, "INFORMATION", "Please check customer/vendor code, not equal to reference Voucher No." & cboJTYPE.Text + "-" + cboCDJNo.Text + "  with Entity Code: " + rsCheckEntity!ENTITY
                DetailPosting = True
                Exit Function
            End If
        End If
        Set rsCheckEntity = Nothing
    End If

    If AddorEdit = "ADD" Then
        If chkOther.Value = 0 Then
            xCUSCODE = N2Str2Null(Left(labClass, 1) & RTrim(LTrim(txtCode.Text)))
        Else
            If txtCode.Text = "" Then
                xCUSCODE = N2Str2Null("")
            Else
                xCUSCODE = N2Str2Null(Left(labClass, 1) & RTrim(LTrim(txtCode.Text)))
            End If
        End If
        J_JITEMNO = N2Str2Null(GetItemNO(frmAMISJournalEntry_GJ.txtVoucherNo.Text))
    Else
        If chkOther.Value = 0 Then
            xCUSCODE = N2Str2Null(Left(labClass, 1) & RTrim(LTrim(txtCode.Text)))
        Else
            If txtCode.Text = "" Then
                xCUSCODE = N2Str2Null("")
            Else
                xCUSCODE = N2Str2Null(Left(labClass, 1) & RTrim(LTrim(txtCode.Text)))
            End If
        End If
        J_JITEMNO = N2Str2Null(Format(txtJItemNo.Text, "0000"))
    End If

    If txtOTH_NO.Text = "" And chkOther.Value = 1 Then
        xAdj_type = N2Str2Null("")
    ElseIf txtOTH_NO.Text <> "" And chkOther.Value = 1 Then
        xAdj_type = N2Str2Null(txtJtype.Text)
    ElseIf cboCDJNo.Text <> "" And cboJTYPE.Text <> "" Then
        xAdj_type = N2Str2Null(txtJtype.Text)
    End If

    Dim J_SUPCODE, J_ATC                          As String
    Dim J_RATE, J_TAXBASE                         As Double

    If fraATC.Visible = True Then
        If cboATC2.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Please select ATC code"
            cboATC2.SetFocus
            DetailPosting = True
            Exit Function
        End If
        If txtTaxBase2.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Tax base amount must have a value"
            txtTaxBase2.SetFocus
            DetailPosting = True
            Exit Function
        End If
        If txtRATE2.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Tax rate must have a value"
            txtRATE2.SetFocus
            DetailPosting = True
            Exit Function
        End If
    End If

    If cboGJAccountNo.Text = DEALER_ITW_COMPENSATION Or cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        'J_SUPCODE = N2Str2Null(SetVendorCode(cboJVSupCust.Text))
        'xCUSCODE = N2Str2Null(WHAT_CLASS & txtCode.Text)
        J_ATC = N2Str2Null(cboATC2.Text)
        J_RATE = NumericVal(txtRATE2.Text)
        J_TAXBASE = NumericVal(txtTaxBase2.Text)
    Else
        'J_SUPCODE = "'999999'"
        'xCUSCODE = N2Str2Null(WHAT_CLASS & txtCode.Text)
        J_ATC = "NULL"
        J_RATE = 0
        J_TAXBASE = 0
    End If

    If AddorEdit = "ADD" Then
        '        If txtGJAccountParticulars.Text <> "" And txtGJAccountParticulars.Text <> "Pls Type Your Remarks Here!" Then
        '            gconDMIS.Execute "insert into AMIS_JV_Detail " & _
                     '                             "(JNo,VoucherNo,itemno,Particulars,status)" & _
                     '                           " values (" & J_JNO & ", " & J_VOUCHERNO & ", " & J_JITEMNO & _
                     '                             ", " & N2Str2Null(txtGJAccountParticulars.Text) & _
                     '                             ", " & J_STATUS & ")"
        '        End If
        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE,INVOICENO,INVOICETYPE,ENTITY,ADJ_VOUCHERNO,ADJ_JTYPE,Adj_Remarks,IS_OTHERS)" & _
                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                         ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & "," & xINVOICENO & "," & xInvoiceType & "," & xCUSCODE & "," & xADJ_VOUCHERNO & "," & xAdj_type & ", " & xADJ_REMARKS & "," & xIS_OTHERS & ")"

        MessagePop RecSave, "INFORMATION", "Record succedfully added"
    Else
        gconDMIS.Execute "update AMIS_Journal_Det set" & _
                         " jdate = " & J_JDATE & "," & _
                         " voucherno = " & J_VOUCHERNO & "," & _
                         " jtype = " & J_JTYPE & "," & _
                         " jno = " & J_JNO & "," & _
                         " jitemno = " & J_JITEMNO & "," & _
                         " acct_code = " & J_ACCT_CODE & "," & _
                         " acct_name = " & J_ACCT_NAME & "," & _
                         " debit = " & J_DEBIT & "," & _
                         " credit = " & J_CREDIT & "," & _
                         " tax = " & J_TAX & "," & _
                         " ATC = " & J_ATC & "," & _
                         " RATE = " & J_RATE & "," & _
                         " TAXBASE = " & J_TAXBASE & "," & _
                         " status = " & J_STATUS & "," & _
                         " invoiceno = " & xINVOICENO & "," & _
                         " invoicetype= " & xInvoiceType & "," & _
                         " entity = " & xCUSCODE & "," & _
                         " adj_voucherno = " & xADJ_VOUCHERNO & "," & _
                         " adj_jtype = " & xAdj_type & "," & _
                         " Adj_Remarks = " & xADJ_REMARKS & "," & _
                         " IS_OTHERS = " & xIS_OTHERS & _
                         " where id = " & frmAMISJournalEntry_GJ.labDET.Caption

        MessagePop RecSave, "INFORMATION", "Record Succesfully updated"
        frmAMISJournalEntry_GJ.labDET.Caption = ""
        '        gconDMIS.Execute "update AMIS_JV_Detail set" & _
                 '                       " Particulars = " & N2Str2Null(txtGJAccountParticulars.Text) & _
                 '                       " where JNo = " & J_JNO & " and ItemNo = " & J_JITEMNO
    End If
    Unload Me
    Call frmAMISJournalEntry_GJ.StoreSearch(J_VOUCHERNO)
    frmAMISJournalEntry_GJ.Picture1.Enabled = True

    DetailPosting = True
    Exit Function

ErrorCode:
    DetailPosting = False
End Function

Function CHECK_ACCOUNTCODE(xACCOUNTCODE As String) As Boolean
    Dim rsCHECK_ACCOUNTCODE                       As ADODB.Recordset
    Set rsCHECK_ACCOUNTCODE = New ADODB.Recordset
    rsCHECK_ACCOUNTCODE.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = '" & xACCOUNTCODE & "'", gconDMIS, adOpenKeyset
    If Not rsCHECK_ACCOUNTCODE.EOF And Not rsCHECK_ACCOUNTCODE.BOF Then
        CHECK_ACCOUNTCODE = True
    Else
        CHECK_ACCOUNTCODE = False
    End If
    Set rsCHECK_ACCOUNTCODE = Nothing
End Function

Function GetJNo(xVOUCHERNO As String) As String
'DESCRIPTION: GET THE THE HIGHEST JNO
    Dim rsgetJNO                                  As ADODB.Recordset
    Set rsgetJNO = gconDMIS.Execute("Select JNO From Amis_Journal_hd where Voucherno = '" & xVOUCHERNO & "' and Jtype = 'GJ'")
    If Not rsgetJNO.EOF And Not rsgetJNO.BOF Then
        GetJNo = Null2String(rsgetJNO!JNo)
    Else
        GetJNo = "000001"
    End If
    Set rsgetJNO = Nothing
End Function

Function ALREADY_ADJUSTED(xINVOICENO As String, xInvoiceType As String, xAdj_type As String) As Boolean
'DESCRIPTION: CHECK THE INVOICENO IS ALREADY ADJUSTED SAME WITH THE MRR NO
    Dim rsALREADY_ADJUSTED                        As ADODB.Recordset
    Set rsALREADY_ADJUSTED = New ADODB.Recordset
    rsALREADY_ADJUSTED.Open "Select InvoiceNo,InvoiceType,Adj_jtype from Amis_journal_det where InvoiceNo = '" & xINVOICENO & "' and InvoiceType = '" & xInvoiceType & "' and Adj_Jtype = '" & xAdj_type & "'", gconDMIS, adOpenKeyset
    If Not rsALREADY_ADJUSTED.EOF And Not rsALREADY_ADJUSTED.BOF Then
        ALREADY_ADJUSTED = True
    Else
        ALREADY_ADJUSTED = False
    End If
    Set rsALREADY_ADJUSTED = Nothing
End Function

Function GetItemNO(xVOUCHERNO As String) As String
'DESCRIPTION: GET THE HIGHEST ITEMNO
    Dim rsGetItemNO                               As ADODB.Recordset
    Set rsGetItemNO = gconDMIS.Execute("Select JItemNO from Amis_journal_DET where VoucherNO = '" & RTrim(LTrim(xVOUCHERNO)) & "' and Jtype = 'GJ' order by JitemNo desc")
    If Not rsGetItemNO.EOF And Not rsGetItemNO.BOF Then
        GetItemNO = Format(NumericVal(rsGetItemNO!jitemno) + 1, "0000")
    Else
        GetItemNO = "0001"
    End If
    Set rsGetItemNO = Nothing
End Function

Private Sub cmdGJSave_Click()
    On Error GoTo ErrorCode

    Dim str_MSG                                   As String

    str_MSG = "Error Appear In During @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If DetailPosting = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "GJ Details")
        MsgBox str_MSG, vbCritical, "GJ Detail Error "
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0

ErrorCode:
'    SaveLogFile
    ShowVBError
End Sub

Private Sub cmdVendor_Click()
    xJOURNALTYPE = "GJ"
    SelectEntity = "Vendor"
    frmEntity.Caption = "SEARCH VENDOR"
    ENTITYCODE = "V"
    Call frmEntity.LoadJournal("GJ")
    frmEntity.Show
    frmEntity.txtSearch.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Public Sub MoveKeyPress(KeyCode As Integer)
    Dim First3Letters                             As String
    On Error Resume Next
    If Screen.ActiveForm.ActiveControl Is Nothing Then
        Exit Sub
    End If
    First3Letters = Mid(Screen.ActiveForm.ActiveControl.Name, 1, 3)
    '''''BUGLIST: CHECK NOTHING FOR FORM
    Select Case KeyCode
    Case 13
        If First3Letters = "cbo" Then
            If Screen.ActiveForm.ActiveControl.Text = "" Then Call VBComBoBoxDroppedDown(Screen.ActiveForm.ActiveControl) Else SendKeys MOVEDOWN
        Else
            If First3Letters = "txt" Or First3Letters = "opt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
        End If
    Case 40
        If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
    Case 38
        If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEUP
    End Select
End Sub

Private Sub Form_Load()
    xJOURNALTYPE = "GJ"
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    InitCbo
    InitChart
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
End Sub

Private Sub cboGJAccountNo_Click()
'DESCRIPTION: DISPLAY THE ACCOUNT DESCRIPTION
    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String

    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")

    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

    GettheTaxBaseAmnt
    If cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        On Error Resume Next
        cboATC2.SetFocus
    Else
        fraATC.Visible = False
    End If

    If cboGJAccountNo.Text <> "" Then
        Dim rsdesc                                As ADODB.Recordset
        Set rsdesc = New ADODB.Recordset
        rsdesc.Open "Select Description from Amis_ChartAccount  where AcctCode = '" & RTrim(LTrim(cboGJAccountNo.Text)) & "'", gconDMIS, adOpenKeyset
        If Not rsdesc.EOF And Not rsdesc.BOF Then
            txtGJAccountName.Text = Null2String(rsdesc!Description)
        Else
            MessagePop InfoFriend, "INFORMATION", "Chart Account has no description"
        End If
        Set rsdesc = Nothing
    Else
    End If
End Sub

Private Sub Label2_Click()
    picSearchInvoice.Visible = False
End Sub

Private Sub Label4_Click()
    picChart.Visible = False
End Sub

Private Sub lstAccounts_DblClick()
    cboGJAccountNo.Text = lstAccounts.SelectedItem.Text
    txtGJAccountName.Text = lstAccounts.SelectedItem.SubItems(1)
    txtGJDebit.SetFocus
    picChart.Visible = False
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstAccounts_DblClick
    End If
End Sub
Private Sub lvwInvNo_DblClick()
    txtINVOICE_DETAIL.Text = lvwInvNo.SelectedItem.Text
    txtINVOICE_TYPE.Text = lvwInvNo.SelectedItem.SubItems(1)
    txtINVOICE_TYPE.SetFocus
    picSearchInvoice.Visible = False
End Sub

Private Sub lvwInvNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lvwInvNo_DblClick
    End If
End Sub

Private Sub txtCode_Change()
'InitCboCDJ_No
    labClass.Caption = ENTITYCODE & txtCode
End Sub

Private Sub txtGJCredit_GotFocus()
    If txtGJCredit.Text = "0.00" Then
        txtGJCredit.Text = ""
    Else
        txtGJCredit.Text = NumericVal(txtGJCredit.Text)
    End If
End Sub

Private Sub txtGJCredit_LostFocus()
    If txtGJCredit.Text = "" Then
        txtGJCredit.Text = "0.00"
    Else
        txtGJCredit.Text = ToDoubleNumber(txtGJCredit.Text)
    End If
End Sub

Private Sub txtGJDebit_GotFocus()
    If txtGJDebit.Text = "0.00" Then
        txtGJDebit.Text = ""
    Else
        txtGJDebit.Text = NumericVal(txtGJDebit.Text)
    End If
End Sub

Private Sub txtGJDebit_LostFocus()
    If txtGJDebit.Text = "" Then
        txtGJDebit.Text = "0.00"
    Else
        txtGJDebit.Text = ToDoubleNumber(txtGJDebit.Text)
    End If
End Sub

Private Sub txtInvoiceNo_Change()
    If txtInvoiceNo.Text = "" Then
        txtInvoiceType.Text = ""
        cboCDJNo.Enabled = True
        cboJTYPE.Enabled = True
        chkOther.Enabled = True
    Else
        cboCDJNo.Enabled = False
        cboJTYPE.Enabled = False
        chkOther.Enabled = False
    End If
End Sub

Function CHECK_IF_AR_SCHED(xACCT_CODE As String) As Boolean
'DESCRIPITON:CHECK AR ACCOUNT SCHEDULE THEN IF AR ACCOUNT RETURN TRUE
    Dim rsCHECK_IF_AR_SCHED                       As ADODB.Recordset
    Set rsCHECK_IF_AR_SCHED = New ADODB.Recordset
    rsCHECK_IF_AR_SCHED.Open "Select AcctCode from Amis_ChartAccount where AcctCode = '" & xACCT_CODE & "' and Is_Schedule_Accnt = 1", gconDMIS, adOpenKeyset
    If Not rsCHECK_IF_AR_SCHED.EOF And Not rsCHECK_IF_AR_SCHED.BOF Then
        CHECK_IF_AR_SCHED = True
    Else
        CHECK_IF_AR_SCHED = False
    End If
    Set rsCHECK_IF_AR_SCHED = Nothing
End Function

Sub FILL_CHARTACCOUNT()
'DESCRIPTION: DISPLAY THE CHART OF ACCOUNT

    Dim rsFILL_CHARTACCOUNT                       As ADODB.Recordset
    Dim Item                                      As ListItem
    Set rsFILL_CHARTACCOUNT = New ADODB.Recordset
    rsFILL_CHARTACCOUNT.Open "SELECT ACCTCODE,DESCRIPTION,ACCTTYPE FROM AMIS_CHARTACCOUNT ORDER BY DESCRIPTION ASC", gconDMIS, adOpenKeyset
    lstAccounts.ListItems.Clear
    If Not rsFILL_CHARTACCOUNT.EOF And Not rsFILL_CHARTACCOUNT.BOF Then
        Do While Not rsFILL_CHARTACCOUNT.EOF
            Set Item = lstAccounts.ListItems.Add(, , Null2String(rsFILL_CHARTACCOUNT!ACCTCODE))
            Item.SubItems(1) = Null2String(rsFILL_CHARTACCOUNT!Description)
            Item.SubItems(2) = Null2String(rsFILL_CHARTACCOUNT!ACCTTYPE)
            rsFILL_CHARTACCOUNT.MoveNext
        Loop
    End If
    Set rsFILL_CHARTACCOUNT = Nothing
End Sub


Private Sub txtInvoiceNo_LostFocus()
    If txtInvoiceNo.Text <> "" Then
        cboCDJNo.Enabled = False
        chkOther.Enabled = False
    Else
        cboCDJNo.Enabled = True
        chkOther.Enabled = True
    End If
End Sub

Private Sub txtSearchAccount_Change()
'DESCRIPTION: SEARCHING FOR ACCOUNT CODE
    Dim rssearch                                  As ADODB.Recordset
    Dim Item                                      As ListItem
    Set rssearch = New ADODB.Recordset
    rssearch.Open "SELECT ACCTCODE,DESCRIPTION,ACCTTYPE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION LIKE '" & RTrim(LTrim(txtSearchAccount.Text)) & "%' ORDER BY DESCRIPTION ASC", gconDMIS, adOpenKeyset
    lstAccounts.ListItems.Clear
    If Not rssearch.EOF And Not rssearch.BOF Then
        Do While Not rssearch.EOF
            Set Item = lstAccounts.ListItems.Add(, , Null2String(rssearch!ACCTCODE))
            Item.SubItems(1) = Null2String(rssearch!Description)
            Item.SubItems(2) = Null2String(rssearch!ACCTTYPE)
            rssearch.MoveNext
        Loop
    End If
    Set rssearch = Nothing
End Sub

Sub InitCboCDJ_No()
'DESCRIPTION: SELECT ALL THE CASH DISBURSEMENT NO OF THIS VENDOR CODE ACCOUNT
    Dim rsInitCboCDJ_No                           As ADODB.Recordset
    Set rsInitCboCDJ_No = New ADODB.Recordset
    rsInitCboCDJ_No.Open "Select VoucherNo from Amis_Journal_hd where VendorCode = '" & RTrim(LTrim(txtCode.Text)) & "'", gconDMIS, adOpenKeyset
    cboCDJNo.Clear
    If Not rsInitCboCDJ_No.EOF And Not rsInitCboCDJ_No.BOF Then
        Do While Not rsInitCboCDJ_No.EOF
            cboCDJNo.AddItem Null2String(rsInitCboCDJ_No!VOUCHERNO)
            rsInitCboCDJ_No.MoveNext
        Loop
    End If
    Set rsInitCboCDJ_No = Nothing
End Sub

Sub InitCbo()
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
    If Not rsATC.EOF And Not rsATC.BOF Then
        'Combo_Loadval cboATC, rsATC
        rsATC.MoveFirst: cboATC2.AddItem ""
        Do While Not rsATC.EOF
            cboATC2.AddItem Null2String(rsATC!ATC)
            rsATC.MoveNext
        Loop
    End If
    Set rsATC = Nothing
End Sub

Private Sub txtSearchAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtSearchAccount.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstAccounts.ListItems.Count > 0 And lstAccounts.Enabled = True Then: lstAccounts.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Sub FIND_JOURNAL_NO(xJType As String)
    Dim rsFIND_JOURNAL_NO                         As ADODB.Recordset
    Set rsFIND_JOURNAL_NO = New ADODB.Recordset

    If xJType = "CDJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "APJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "CRJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "GJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT DISTINCT VOUCHERNO FROM AMIS_JOURNAL_DET WHERE JTYPE = 'GJ' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "SJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "COB" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    ElseIf xJType = "OTH" Then
        rsFIND_JOURNAL_NO.Open "SELECT DISTINCT ADJ_VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ADJ_JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P' AND ADJ_VOUCHERNO IS NOT NULL", gconDMIS, adOpenKeyset
    ElseIf xJType = "VPJ" Then
        rsFIND_JOURNAL_NO.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE = " & N2Str2Null(xJType) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    Else
        Exit Sub
    End If

    If Not rsFIND_JOURNAL_NO.EOF And Not rsFIND_JOURNAL_NO.BOF Then
        Do While Not rsFIND_JOURNAL_NO.EOF
            If xJType = "OTH" Then
                cboCDJNo.AddItem Null2String(rsFIND_JOURNAL_NO!ADJ_VOUCHERNO)
            Else
                cboCDJNo.AddItem Null2String(rsFIND_JOURNAL_NO!VOUCHERNO)
            End If
            rsFIND_JOURNAL_NO.MoveNext
        Loop
    End If
    Set rsFIND_JOURNAL_NO = Nothing
End Sub

Private Sub txtSearchInvoice_Change()
    Dim rssearch                                  As ADODB.Recordset
    Dim rsDetails                                 As ADODB.Recordset
    Dim Item                                      As ListItem
    Set rssearch = New ADODB.Recordset
    lvwInvNo.ListItems.Clear

    If txtSearchInvoice.Text <> "" Then
        If RTrim(LTrim(cboJTYPE.Text)) = "CRJ" Then
            Set rsDetails = New ADODB.Recordset
            rsDetails.Open "SELECT INVOICENO,INVOICETYPE FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ' AND VOUCHERNO = '" & cboCDJNo.Text & "' AND (ABS(INVOICENO)  LIKE '" & txtSearchInvoice.Text & "%' or INVOICENO  LIKE '" & txtSearchInvoice.Text & "%' )", gconDMIS, adOpenKeyset
            If Not rsDetails.EOF And Not rsDetails.BOF Then
                Do While Not rsDetails.EOF
                    Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!INVOICENO))
                    Item.SubItems(1) = Null2String(rsDetails!InvoiceType)
                    rsDetails.MoveNext
                Loop
            End If
            Set rsDetails = Nothing
        ElseIf RTrim(LTrim(cboJTYPE.Text)) = "CDJ" Then
            Set rsDetails = New ADODB.Recordset
            rsDetails.Open "SELECT PV_VOUCHERNO,JTYPE FROM AMIS_CV_DETAIL WHERE JTYPE = 'APJ' AND VOUCHERNO = '" & cboCDJNo.Text & "' AND (ABS(PV_VOUCHERNO) LIKE '" & txtSearchInvoice.Text & "%' or PV_VOUCHERNO LIKE '" & txtSearchInvoice.Text & "%') ", gconDMIS, adOpenKeyset
            If Not rsDetails.EOF And Not rsDetails.BOF Then
                Do While Not rsDetails.EOF
                    Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!pv_voucherno))
                    Item.SubItems(1) = Null2String(rsDetails!jtype)
                    rsDetails.MoveNext
                Loop
            End If
            Set rsDetails = Nothing
        End If
    Else
        If RTrim(LTrim(cboJTYPE.Text)) = "CRJ" Then
            Set rsDetails = New ADODB.Recordset
            rsDetails.Open "SELECT INVOICENO,INVOICETYPE FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ' and VOUCHERNO LIKE '" & cboCDJNo.Text & "'", gconDMIS, adOpenKeyset
            If Not rsDetails.EOF And Not rsDetails.BOF Then
                Do While Not rsDetails.EOF
                    Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!INVOICENO))
                    Item.SubItems(1) = Null2String(rsDetails!InvoiceType)
                    rsDetails.MoveNext
                Loop
            End If
            Set rsDetails = Nothing
        ElseIf RTrim(LTrim(cboJTYPE.Text)) = "CDJ" Then
            Set rsDetails = New ADODB.Recordset
            rsDetails.Open "SELECT PV_VOUCHERNO,JTYPE FROM AMIS_CV_DETAIL WHERE JTYPE = 'APJ' AND VOUCHERNO LIKE '" & cboCDJNo.Text & "' ", gconDMIS, adOpenKeyset
            If Not rsDetails.EOF And Not rsDetails.BOF Then
                Do While Not rsDetails.EOF
                    Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!pv_voucherno))
                    Item.SubItems(1) = Null2String(rsDetails!jtype)
                    rsDetails.MoveNext
                Loop
            End If
            Set rsDetails = Nothing
        End If
    End If
    Set rssearch = Nothing
End Sub

Private Sub txtSearchInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtSearchInvoice.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then KeyCode = 0
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lvwInvNo.ListItems.Count > 0 And lvwInvNo.Enabled = True Then: lvwInvNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub GET_OTH_MAX_NO()
    Dim rsGET_OTH_MAX_NO                          As ADODB.Recordset
    Set rsGET_OTH_MAX_NO = New ADODB.Recordset
    'rsGET_OTH_MAX_NO.Open "SELECT ADJ_VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ADJ_JTYPE = 'OTH' AND ADJ_VOUCHERNO IS NOT NULL ORDER BY INVOICENO DESC ", gconDMIS, adOpenKeyset
    rsGET_OTH_MAX_NO.Open "SELECT MAX(ADJ_VOUCHERNO) AS MAX_ADJ_VOUCHERNO FROM AMIS_JOURNAL_DET WHERE ADJ_JTYPE = 'OTH' AND IS_OTHERS = 1", gconDMIS, adOpenKeyset
    If Not rsGET_OTH_MAX_NO.EOF And Not rsGET_OTH_MAX_NO.BOF Then
        txtOTH_NO.Text = Format((NumericVal(rsGET_OTH_MAX_NO!MAX_ADJ_VOUCHERNO) + 1), "000000")
    Else
        txtOTH_NO.Text = "000001"
    End If
    Set rsGET_OTH_MAX_NO = Nothing
End Sub

Sub FIND_SUB_DETAIL_REFERENCE()
    Dim rsCHECK_MULTI_DETAIL                      As ADODB.Recordset
    Dim rsDetails                                 As ADODB.Recordset
    Dim Item                                      As ListItem
    Set rsCHECK_MULTI_DETAIL = New ADODB.Recordset

    lvwInvNo.ListItems.Clear

    If RTrim(LTrim(cboJTYPE.Text)) = "CRJ" Then
        rsCHECK_MULTI_DETAIL.Open "SELECT COUNT(VOUCHERNO) AS XXX_COUNT FROM AMIS_CRJ_DETAIL ACD INNER JOIN AMIS_CHARTACCOUNT AC ON ACD.J_CLASS=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO = '" & cboCDJNo.Text & "' AND CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
    ElseIf RTrim(LTrim(cboJTYPE.Text)) = "CDJ" Then
        rsCHECK_MULTI_DETAIL.Open "SELECT COUNT(VOUCHERNO) AS XXX_COUNT FROM AMIS_CV_DETAIL ACD INNER JOIN AMIS_CHARTACCOUNT AC ON ACD.J_CLASS=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = 'APJ'", gconDMIS, adOpenKeyset
    ElseIf RTrim(LTrim(cboJTYPE.Text)) = "GJ" Then
        rsCHECK_MULTI_DETAIL.Open "SELECT COUNT(VOUCHERNO) AS XXX_COUNT FROM AMIS_JOURNAL_DET DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = 'GJ'", gconDMIS, adOpenKeyset
    Else
        Exit Sub
    End If

    If RTrim(LTrim(cboJTYPE.Text)) = "CRJ" Or RTrim(LTrim(cboJTYPE.Text)) = "CDJ" Then
        If Not rsCHECK_MULTI_DETAIL.EOF And Not rsCHECK_MULTI_DETAIL.BOF Then
            If NumericVal(rsCHECK_MULTI_DETAIL!XXX_COUNT) > 1 Then

                If txtINVOICE_DETAIL.Text = "" And txtINVOICE_TYPE.Text = "" Then
                    MsgBox "This voucher no. contains multi details. You need to select the reference detail.", vbInformation + vbOKOnly, "INFORMATION"
                End If

                picSearchInvoice.Visible = True
                picSearchInvoice.ZOrder 0

                If RTrim(LTrim(cboJTYPE.Text)) = "CRJ" Then
                    Set rsDetails = New ADODB.Recordset
                    rsDetails.Open "SELECT INVOICENO,INVOICETYPE FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO = '" & cboCDJNo.Text & "' AND CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
                    If Not rsDetails.EOF And Not rsDetails.BOF Then
                        Do While Not rsDetails.EOF
                            Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!INVOICENO))
                            Item.SubItems(1) = Null2String(rsDetails!InvoiceType)
                            rsDetails.MoveNext
                        Loop
                    End If
                    Set rsDetails = Nothing
                ElseIf RTrim(LTrim(cboJTYPE.Text)) = "CDJ" Then
                    Set rsDetails = New ADODB.Recordset
                    rsDetails.Open "SELECT PV_VOUCHERNO,JTYPE FROM AMIS_CV_DETAIL WHERE VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = 'APJ'", gconDMIS, adOpenKeyset
                    If Not rsDetails.EOF And Not rsDetails.BOF Then
                        Do While Not rsDetails.EOF
                            Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!pv_voucherno))
                            Item.SubItems(1) = Null2String(rsDetails!jtype)
                            rsDetails.MoveNext
                        Loop
                    End If
                    Set rsDetails = Nothing
                ElseIf RTrim(LTrim(cboJTYPE.Text)) = "GJ" Then
                    Set rsDetails = New ADODB.Recordset
                    rsDetails.Open "SELECT ADJ_VOUCHERNO,ADJ_JTYPE FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & cboCDJNo.Text & "' AND JTYPE = 'GJ'", gconDMIS, adOpenKeyset
                    If Not rsDetails.EOF And Not rsDetails.BOF Then
                        Do While Not rsDetails.EOF
                            Set Item = lvwInvNo.ListItems.Add(, , Null2String(rsDetails!ADJ_VOUCHERNO))
                            Item.SubItems(1) = Null2String(rsDetails!ADJ_JTYPE)
                            rsDetails.MoveNext
                        Loop
                    End If
                    Set rsDetails = Nothing
                End If
            Else
                txtINVOICE_DETAIL.Text = ""
                txtINVOICE_TYPE.Text = ""
            End If
        End If
    End If
    Set rsCHECK_MULTI_DETAIL = Nothing
End Sub

Function ReturnWithholdingTax(XXX As String)
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnWithholdingTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInPutTax()
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Sub GettheTaxBaseAmnt()
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    If xJOURNALTYPE = "GJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & frmAMISJournalEntry_GJ.txtVoucherNo.Text & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'GJ'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTaxBase2.Text = N2Str2IntZero(RS!SumDebit)
    End If
    Set RS = Nothing
End Sub


