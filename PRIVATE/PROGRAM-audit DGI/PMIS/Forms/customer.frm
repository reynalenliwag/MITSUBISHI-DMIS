VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPMIOSCustomer 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master File"
   ClientHeight    =   4440
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   9615
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "customer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   9615
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   4275
      Left            =   60
      ScaleHeight     =   4215
      ScaleWidth      =   2535
      TabIndex        =   38
      Top             =   90
      Width           =   2595
      Begin VB.Image Image1 
         Height          =   11640
         Left            =   0
         Picture         =   "customer.frx":08CA
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   2700
      ScaleHeight     =   3315
      ScaleWidth      =   6825
      TabIndex        =   23
      Top             =   90
      Width           =   6855
      Begin VB.TextBox txtDisc_Surch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   5010
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Type customer's discount or surcharge. Do not inlcude % symbol (e.g. 10, 15, 5)"
         Top             =   2130
         Width           =   1545
      End
      Begin VB.TextBox txtCR_limit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Type customer's credit limit. Do not use comma as separator (e.g 4321, 65000)"
         Top             =   2910
         Width           =   1395
      End
      Begin VB.TextBox txtCR_Amount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type customer's credit amount. Do not use comma as separator (e.g. 12345, 2000)"
         Top             =   2520
         Width           =   1395
      End
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type customer's whole name (e.g. BALIWAG TRANSIT, INC., BAYER, PHILS.)"
         Top             =   450
         Width           =   4875
      End
      Begin VB.TextBox txtPhoneNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Input customer's telephone number. Include area codes if possible (e.g. 0544750000)."
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txtCustCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "The code is system generated."
         Top             =   60
         Width           =   1065
      End
      Begin VB.TextBox txtPartsPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   5010
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2640
         Width           =   285
      End
      Begin RichTextLib.RichTextBox txtCustadrs 
         Height          =   885
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Type customer's complete address (e.g. LIBIS, QUEZON CITY)"
         Top             =   1140
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   1561
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         MaxLength       =   120
         Appearance      =   0
         TextRTF         =   $"customer.frx":15242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   6570
         TabIndex        =   32
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   510
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   870
         Width           =   1875
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   2550
         Width           =   1605
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Surcharge"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3840
         TabIndex        =   27
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   26
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   2970
         Width           =   1665
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Parts Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3840
         TabIndex        =   24
         Top             =   2700
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2700
      ScaleHeight     =   855
      ScaleWidth      =   6825
      TabIndex        =   21
      Top             =   3480
      Width           =   6855
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6030
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":152CA
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":155D4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Close window"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":158DE
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":15BE8
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Delete current record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4530
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":164B2
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":167BC
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Edit current record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3780
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":17086
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":17390
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Add a record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3030
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":17C5A
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":17F64
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Search for a file"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":1882E
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":18B38
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "View last record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1530
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":18F7A
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":19284
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "View first record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   780
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":196C6
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":199D0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "View next record"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   30
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":19E12
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1A11C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View previous record"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2700
      ScaleHeight     =   855
      ScaleWidth      =   6825
      TabIndex        =   22
      Top             =   3480
      Width           =   6855
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6030
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":1A55E
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1A868
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Discard changes"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "customer.frx":1B8AA
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1BBB4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Save changes"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   60
      TabIndex        =   35
      Top             =   0
      Width           =   2595
      Begin VB.TextBox textSearch 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   150
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   3705
         Left            =   30
         TabIndex        =   37
         Top             =   540
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   6535
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
         MouseIcon       =   "customer.frx":1BFF6
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
   Begin VB.Label Label2 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   8
      Left            =   10530
      TabIndex        =   34
      Top             =   5160
      Width           =   225
   End
   Begin VB.Label Label3 
      Caption         =   "-- required field"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   10710
      TabIndex        =   33
      Top             =   5130
      Width           =   1485
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   4410
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3570
      TabIndex        =   19
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMIOSCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer, rsRepor As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
Dim rsAddCust As ADODB.Recordset
AddorEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
initMemvars
Set rsAddCust = New ADODB.Recordset
    rsAddCust.Open "select custcode from customer order by custcode asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsAddCust.EOF And Not rsAddCust.BOF Then
   rsAddCust.MoveLast
   If IsNumeric(rsAddCust!custcode) = False Then
      Do While Not rsAddCust.BOF
         If IsNumeric(rsAddCust!custcode) = False Then
            rsAddCust.MovePrevious
         Else
            txtCustCode.Text = Format(NumericVal(rsAddCust!custcode) + 1, "00000")
            Exit Do
         End If
      Loop
   End If
Else
   txtCustCode.Text = "00001"
End If
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
AddorEdit = ""
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsCustomer.BOF Or Not rsCustomer.EOF Then
   If ShowConfirmDelete = True Then
      gconPMIOS.Execute "delete from Customer where id = " & labid.Caption
      ShowDeletedMsg
   End If
Else
   ShowNothingToDeleteMsg
End If
rsRefresh
StoreMemvars
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Picture3.Visible = False
textSearch.SetFocus

'Dim findStr As String
'findStr = InputSpeechBox("Please Input Code or Name ...", txtCustName.Text)
'If findStr <> "" Then
'   On Error Resume Next
'   rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "custcode", findStr).Bookmark
'   If Err.Number = 3021 Then
'      On Error GoTo ErrorCode
'      rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "custname", findStr).Bookmark
'   End If
'End If
'StoreMemvars
'Exit Sub

'ErrorCode:
'If Err.Number = 3021 Then
'   ShowCantFind findStr
'   Resume Next
'End If
End Sub

Private Sub cmdFirst_Click()
rsCustomer.MoveFirst
StoreMemvars
End Sub

Private Sub cmdLast_Click()
rsCustomer.MoveLast
StoreMemvars
End Sub

Private Sub cmdNext_Click()
rsCustomer.MoveNext
If rsCustomer.EOF Then
   rsCustomer.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsCustomer.MovePrevious
If rsCustomer.BOF Then
   rsCustomer.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode

Dim VTXTCustCode, VTXTCustName, VTXTCustAdrs, VTXTPhoneNo As String
Dim VTXTCR_Amount, VTXTCR_limit, VTXTDisc_Surch As Double
Dim VTXTPartsPrice As String

If IsNull(txtCustCode.Text) = True Or txtCustCode.Text = "" Then
   ShowIsRequiredMsg "Code"
   On Error Resume Next
   txtCustCode.SetFocus
   Exit Sub
Else
   If AddorEdit = "ADD" Then
      Dim rsfindDup As ADODB.Recordset
      Set rsfindDup = New ADODB.Recordset
          rsfindDup.Open "select custcode from Customer where custcode = " & N2Str2Null(txtCustCode.Text), gconPMIOS, adOpenForwardOnly, adLockReadOnly
      If Not rsfindDup.EOF And Not rsfindDup.BOF Then
         MsgSpeechBox "Code already exist!"
         On Error Resume Next
         txtCustCode.SetFocus
         Exit Sub
      End If
      Set rsfindDup = Nothing
   End If
End If
If txtCustName.Text = "" Then
   ShowIsRequiredMsg "Name"
   On Error Resume Next
   txtCustName.SetFocus
   Exit Sub
End If

VTXTCustCode = N2Str2Null(txtCustCode.Text)
VTXTCustName = N2Str2Null(txtCustName.Text)
VTXTCustAdrs = N2Str2Null(txtCustadrs.Text)
VTXTPhoneNo = N2Str2Null(txtPhoneNo.Text)
VTXTCR_Amount = NumericVal(txtCR_Amount.Text)
VTXTCR_limit = NumericVal(txtCR_limit.Text)
VTXTDisc_Surch = NumericVal(txtDisc_Surch.Text)
VTXTPartsPrice = N2Str2Null(txtPartsPrice.Text)
If AddorEdit = "ADD" Then
   gconPMIOS.Execute "Insert into Customer" & _
                    " (custcode,custname,Custadrs,phoneno,cr_amount,cr_limit,disc_surch,partsprice,lastupdate,usercode)" & _
                    " values (" & VTXTCustCode & ", " & VTXTCustName & ", " & VTXTCustAdrs & "," & _
                    " " & VTXTPhoneNo & ", " & VTXTCR_Amount & ", " & VTXTCR_limit & ", " & VTXTDisc_Surch & "," & _
                    " " & VTXTPartsPrice & ", " & "'" & LOGDATE & "'" & ", " & "" & N2Str2Null(LOGCODE) & "" & ")"
   ShowSuccessFullyAdded
Else
   gconPMIOS.Execute "update Customer set" & _
                    " Custcode = " & VTXTCustCode & "," & _
                    " Custname = " & VTXTCustName & "," & _
                    " Custadrs = " & VTXTCustAdrs & "," & _
                    " phoneno = " & VTXTPhoneNo & "," & _
                    " cr_amount = " & VTXTCR_Amount & "," & _
                    " cr_limit = " & VTXTCR_limit & "," & _
                    " disc_surch = " & VTXTDisc_Surch & "," & _
                    " partsprice = " & VTXTPartsPrice & "," & _
                    " lastupdate = " & "'" & LOGDATE & "'" & "," & _
                    " usercode = " & "" & N2Str2Null(LOGCODE) & "" & _
                    " where id = " & labid.Caption
   ShowSuccessFullyUpdated
End If
rsRefresh
On Error Resume Next
rsCustomer.Find "Custcode =" & VTXTCustCode
cmdCancel.Value = True
Exit Sub

ErrorCode:
ShowVBError
cmdCancel.Value = True
Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
rsRefresh
Frame1.Enabled = False
SetFormSettings Me
textSearch.Text = "": Picture3.ZOrder 0
initMemvars
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub initMemvars()
txtCustCode.Text = ""
txtCustName.Text = ""
txtCustadrs.Text = ""
txtPhoneNo.Text = ""
txtCR_Amount.Text = ""
txtCR_limit.Text = ""
txtDisc_Surch.Text = ""
txtPartsPrice.Text = ""
End Sub

Sub StoreMemvars()
If Not rsCustomer.EOF And Not rsCustomer.BOF Then
   labid.Caption = rsCustomer!ID
   txtCustCode.Text = Null2String(rsCustomer!custcode)
   txtCustName.Text = Null2String(rsCustomer!custname)
   txtCustadrs.Text = Null2String(rsCustomer!custadrs)
   txtPhoneNo.Text = Null2String(rsCustomer!phoneno)
   txtCR_Amount.Text = N2Str2IntZero(rsCustomer!cr_amount)
   txtCR_limit.Text = N2Str2IntZero(rsCustomer!cr_limit)
   txtDisc_Surch.Text = N2Str2IntZero(rsCustomer!disc_surch)
   txtPartsPrice.Text = Null2String(rsCustomer!partsprice)
Else
   ShowNoRecord
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from Customer order by id asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSCustomer = Nothing
UnloadForm Me
End Sub

Private Sub lstCustomer_GotFocus()
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstCustomer
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

Private Sub lstCustomer_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
If Trim(textSearch.Text) = "" Then
   FillGrid
Else
   FillSearchGrid (textSearch.Text)
End If
End Sub

Sub FillGrid()
Dim rsCust As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCust = New ADODB.Recordset
Set rsCust = gconPMIOS.Execute("select CustName,ID from Customer order by CustName asc")
If Not (rsCust.EOF And rsCust.BOF) Then
   lstCustomer.Enabled = True
   Listview_Loadval Me.lstCustomer.ListItems, rsCust
   lstCustomer.Refresh
Else
   lstCustomer.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsCust As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCust = New ADODB.Recordset
Set rsCust = gconPMIOS.Execute("select CustName,ID from Customer where CustName like'" & XXX & "%'")
If Not (rsCust.EOF And rsCust.BOF) Then
   lstCustomer.Enabled = True
   Listview_Loadval Me.lstCustomer.ListItems, rsCust
   lstCustomer.Refresh
Else
   lstCustomer.Enabled = False
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstCustomer.SetFocus
End Sub

