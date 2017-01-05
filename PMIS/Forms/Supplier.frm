VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMIOSSupplier 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Master File"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Supplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   8865
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2700
      ScaleHeight     =   855
      ScaleWidth      =   6075
      TabIndex        =   26
      Top             =   3600
      Width           =   6105
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
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Close window"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P&rint"
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
         MouseIcon       =   "Supplier.frx":0EDE
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print record"
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
         Left            =   3780
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":1AB2
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   3030
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":2686
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":2990
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Edit selected record"
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
         Left            =   2280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":325A
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":3564
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add record"
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
         Left            =   1530
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":3E2E
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":4138
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Firnd a record"
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
         MouseIcon       =   "Supplier.frx":4A02
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":4D0C
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
         MouseIcon       =   "Supplier.frx":514E
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":5458
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View previous record"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      Caption         =   "Data Entry"
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
      Height          =   3495
      Left            =   2700
      TabIndex        =   18
      Top             =   60
      Width           =   6105
      Begin VB.TextBox txtVat_Percnt 
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
         Left            =   4500
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Type the VAT percentage levied by the company."
         Top             =   2250
         Width           =   855
      End
      Begin Crystal.CrystalReport rptPrintSupplier 
         Left            =   5580
         Top             =   2970
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "List of New Suppliers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.CheckBox chkNonVat 
         BackColor       =   &H00DEDFDE&
         Caption         =   "Non VAT Supplier"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   3120
         Width           =   2325
      End
      Begin VB.TextBox txtSup_Addrs 
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
         Height          =   885
         Left            =   150
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Supplier.frx":589A
         ToolTipText     =   "Input the complete address of the suplier."
         Top             =   1320
         Width           =   5865
      End
      Begin VB.TextBox txtSupCode 
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
         Left            =   1140
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1065
      End
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
         Left            =   1290
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Type the Discount Surch. of the supplier."
         Top             =   3030
         Width           =   1815
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
         Left            =   1290
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type the telephone number of the supplier, include area codes (e.g. 0544750000)"
         Top             =   2250
         Width           =   1815
      End
      Begin VB.TextBox txtContact 
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
         Left            =   1290
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the full name of the contact person in the company."
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtSupName 
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
         Left            =   1140
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type the complete name of the supplier."
         Top             =   630
         Width           =   4635
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vat Percent"
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
         Left            =   3330
         TabIndex        =   35
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "- required field"
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
         Left            =   4500
         TabIndex        =   31
         Top             =   150
         Width           =   1335
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
         Height          =   255
         Index           =   3
         Left            =   4350
         TabIndex        =   30
         Top             =   180
         Width           =   135
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
         Index           =   1
         Left            =   5790
         TabIndex        =   29
         Top             =   660
         Width           =   135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Left            =   150
         TabIndex        =   28
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc. Surch."
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
         Left            =   150
         TabIndex        =   23
         Top             =   3090
         Width           =   1125
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Left            =   150
         TabIndex        =   22
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Left            =   150
         TabIndex        =   21
         Top             =   2670
         Width           =   1125
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Address"
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
         Left            =   150
         TabIndex        =   20
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   150
         TabIndex        =   19
         Top             =   660
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   2700
      ScaleHeight     =   855
      ScaleWidth      =   6075
      TabIndex        =   27
      Top             =   3600
      Width           =   6105
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
         Left            =   5280
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":58A0
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":5BAA
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   4530
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Supplier.frx":6BEC
         MousePointer    =   99  'Custom
         Picture         =   "Supplier.frx":6EF6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save changes"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   60
      TabIndex        =   32
      Top             =   60
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
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   150
         Width           =   2505
      End
      Begin MSComctlLib.ListView lstSupplier 
         Height          =   3825
         Left            =   60
         TabIndex        =   34
         Top             =   540
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   6747
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
         MouseIcon       =   "Supplier.frx":7338
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SUPPLIER NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   25
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   24
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMIOSSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplier As ADODB.Recordset
Dim AddOrEdit As String

Private Sub cmdPrint_Click()
On Error GoTo ErrorCode
Screen.MousePointer = 11
rptPrintSupplier.ReportFileName = PMIOS_REPORT_PATH & "printsupplier.rpt"
rptPrintSupplier.Connect = PMIOS_REPORT_Connection
rptPrintSupplier.Action = 1
Screen.MousePointer = 0
Exit Sub

ErrorCode:
Screen.MousePointer = 11
ShowVBError
Exit Sub
End Sub

Private Sub cmdAdd_Click()
Dim rsAddCust As ADODB.Recordset
AddOrEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
InitMemvars
Set rsAddCust = New ADODB.Recordset
    rsAddCust.Open "select supcode from Supplier order by SupCode asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsAddCust.EOF And Not rsAddCust.BOF Then
   rsAddCust.MoveLast
   If IsNumeric(rsAddCust!SupCode) = False Then
      Do While Not rsAddCust.BOF
         If IsNumeric(rsAddCust!SupCode) = False Then
            rsAddCust.MovePrevious
         Else
            txtSupCode.Text = Format(NumericVal(rsAddCust!SupCode) + 1, "00000")
            Exit Do
         End If
      Loop
   End If
Else
   txtSupCode.Text = "00001"
End If
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
AddOrEdit = ""
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsSupplier.BOF Or Not rsSupplier.EOF Then
   If ShowConfirmDelete = True Then
      gconPMIOS.Execute "delete from Supplier where id = " & labID.Caption
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
AddOrEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
'Picture3.Visible = False
textSearch.SetFocus
'Dim findStr As String
'findStr = InputSpeechBox("Please Input Code or Name ...", txtSupName.Text)
'If findStr <> "" Then
'   On Error Resume Next
'   rsSupplier.Bookmark = rsFind(rsSupplier.Clone, "supcode", findStr).Bookmark
'   If Err.Number = 3021 Then
'      On Error GoTo ErrorCode
'      rsSupplier.Bookmark = rsFind(rsSupplier.Clone, "supname", findStr).Bookmark
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
rsSupplier.MoveFirst
StoreMemvars
End Sub

Private Sub cmdLast_Click()
rsSupplier.MoveLast
StoreMemvars
End Sub

Private Sub cmdNext_Click()
rsSupplier.MoveNext
If rsSupplier.EOF Then
   rsSupplier.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsSupplier.MovePrevious
If rsSupplier.BOF Then
   rsSupplier.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim rsfindDup As ADODB.Recordset
Dim NONVAT As String
If IsNull(txtSupCode.Text) = True Then
   MsgSpeechBox "Code must not be empty"
   On Error Resume Next
   txtSupCode.SetFocus
   Exit Sub
Else
If AddOrEdit = "ADD" Then
   Set rsfindDup = New ADODB.Recordset
       rsfindDup.Open "select supcode from Supplier where SupCode = '" & txtSupCode.Text & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsfindDup.EOF And Not rsfindDup.BOF Then
      MsgSpeechBox "Code already exist!"
      On Error Resume Next
      txtSupCode.SetFocus
      Exit Sub
   End If
End If
End If
If txtSupName.Text = "" Then
   ShowIsRequiredMsg "Supplier Name"
   On Error Resume Next
   txtSupName.SetFocus
   Exit Sub
End If

Dim VTXTSupCode, VTXTSupName, VTXTSup_Addrs As String
Dim VTXTPhoneNo, VTXTContact As String
Dim VTXTDisc_Surch, VTXTVat_Percnt As String

VTXTSupCode = N2Str2Null(txtSupCode.Text)
VTXTSupName = N2Str2Null(txtSupName.Text)
VTXTSup_Addrs = N2Str2Null(txtSup_Addrs.Text)
VTXTPhoneNo = N2Str2Null(txtPhoneNo.Text)
VTXTContact = N2Str2Null(txtContact.Text)
VTXTDisc_Surch = NumericVal(txtDisc_Surch.Text)
VTXTVat_Percnt = NumericVal(txtVat_Percnt.Text)
If chkNonVat.Value = 1 Then NONVAT = "'Y'" Else NONVAT = "'N'"
If AddOrEdit = "ADD" Then
   If Not rsSupplier.EOF And Not rsSupplier.BOF Then
      rsSupplier.MoveLast
      labID.Caption = NumericVal(rsSupplier!ID) + 1
   End If
   gconPMIOS.Execute "Insert into Supplier" & _
                    " (SupCode,SupName,Sup_Addrs,phoneno,Contact,disc_surch,Vat_Percnt,lastupdate,usercode,DATE_ENTERED,NONVAT)" & _
                    " values (" & VTXTSupCode & ", " & VTXTSupName & ", " & VTXTSup_Addrs & "," & _
                    " " & VTXTPhoneNo & ", " & VTXTContact & ", " & VTXTDisc_Surch & "," & _
                    " " & VTXTVat_Percnt & ", " & "'" & LOGDATE & "'" & ", " & "" & N2Str2Null(LOGCODE) & ",'" & LOGDATE & "'," & NONVAT & ")"
   ShowSuccessFullyAdded
Else
   gconPMIOS.Execute "update Supplier set" & _
                   " SupCode = " & VTXTSupCode & "," & _
                   " SupName = " & VTXTSupName & "," & _
                   " Sup_Addrs = " & VTXTSup_Addrs & "," & _
                   " phoneno = " & VTXTPhoneNo & "," & _
                   " Contact = " & VTXTContact & "," & _
                   " disc_surch = " & VTXTDisc_Surch & "," & _
                   " Vat_Percnt = " & VTXTVat_Percnt & "," & _
                   " lastupdate = " & "'" & LOGDATE & "'" & "," & _
                   " NONVAT = " & NONVAT & "," & _
                   " usercode = " & "" & N2Str2Null(LOGCODE) & "" & _
                   " where id = " & labID.Caption
   ShowSuccessFullyUpdated
End If
rsRefresh
On Error Resume Next
rsSupplier.Find "id =" & labID.Caption
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
textSearch.Text = "": 'Picture3.ZOrder 0
InitMemvars
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub InitMemvars()
txtSupCode.Text = ""
txtSupName.Text = ""
txtSup_Addrs.Text = ""
txtPhoneNo.Text = ""
txtContact.Text = ""
txtDisc_Surch.Text = ""
txtVat_Percnt.Text = ""
End Sub

Sub StoreMemvars()
If Not rsSupplier.EOF And Not rsSupplier.BOF Then
   labID.Caption = rsSupplier!ID
   txtSupCode.Text = Null2String(rsSupplier!SupCode)
   txtSupName.Text = Null2String(rsSupplier!supname)
   txtSup_Addrs.Text = Null2String(rsSupplier!sup_addrs)
   txtPhoneNo.Text = Null2String(rsSupplier!phoneno)
   txtContact.Text = Null2String(rsSupplier!contact)
   txtDisc_Surch.Text = Null2String(rsSupplier!disc_surch)
   txtVat_Percnt.Text = Null2String(rsSupplier!vat_percnt)
   If Null2String(rsSupplier!NONVAT) = "Y" Then
      chkNonVat.Value = 1
   Else
      chkNonVat.Value = 0
   End If
Else
   ShowNoRecord
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select * from Supplier order by id asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSSupplier = Nothing
UnloadForm Me
End Sub

Private Sub lstSupplier_GotFocus()
rsSupplier.Bookmark = rsFind(rsSupplier.Clone, "ID", lstSupplier.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstSupplier_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsSupplier.Bookmark = rsFind(rsSupplier.Clone, "ID", lstSupplier.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstSupplier_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstSupplier
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

Private Sub lstSupplier_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim rsSup As ADODB.Recordset
lstSupplier.Sorted = False: lstSupplier.ListItems.Clear
Set rsSup = New ADODB.Recordset
Set rsSup = gconPMIOS.Execute("select SupName,ID from Supplier order by SupName asc")
If Not (rsSup.EOF And rsSup.BOF) Then
   lstSupplier.Enabled = True
   Listview_Loadval Me.lstSupplier.ListItems, rsSup
   lstSupplier.Refresh
Else
   lstSupplier.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsSup As ADODB.Recordset
lstSupplier.Sorted = False: lstSupplier.ListItems.Clear
Set rsSup = New ADODB.Recordset
Set rsSup = gconPMIOS.Execute("select SupName,ID from Supplier where SupName like'" & XXX & "%'")
If Not (rsSup.EOF And rsSup.BOF) Then
   lstSupplier.Enabled = True
   Listview_Loadval Me.lstSupplier.ListItems, rsSup
   lstSupplier.Refresh
Else
   lstSupplier.Enabled = False
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstSupplier.SetFocus
End Sub


