VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISMASTERFILEVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Master List"
   ClientHeight    =   5835
   ClientLeft      =   720
   ClientTop       =   435
   ClientWidth     =   6975
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Vendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6975
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   1110
      ScaleHeight     =   885
      ScaleWidth      =   5805
      TabIndex        =   24
      Top             =   4845
      Width           =   5805
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
         Left            =   5040
         MouseIcon       =   "Vendor.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
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
         Left            =   4320
         MouseIcon       =   "Vendor.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
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
         Left            =   3600
         MouseIcon       =   "Vendor.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
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
         Left            =   2880
         MouseIcon       =   "Vendor.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
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
         Left            =   2160
         MouseIcon       =   "Vendor.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
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
         Left            =   1440
         MouseIcon       =   "Vendor.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
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
         Left            =   720
         MouseIcon       =   "Vendor.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
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
         Left            =   0
         MouseIcon       =   "Vendor.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptVendor 
      Left            =   6300
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtTerms 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   4290
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   1500
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "NON-VAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5430
         TabIndex        =   16
         Top             =   1560
         Width           =   1275
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   510
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Vendor.frx":2D71
         Top             =   960
         Width           =   5835
      End
      Begin VB.TextBox txtTIN 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   4290
         MaxLength       =   18
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1890
         Width           =   2505
      End
      Begin VB.TextBox txtFax 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   17
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1890
         Width           =   2565
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   17
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1500
         Width           =   2565
      End
      Begin VB.TextBox txtPosition 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   570
         Width           =   1785
      End
      Begin VB.TextBox txtContact 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   570
         Width           =   3165
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   960
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtNameofVendor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2760
         MaxLength       =   40
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   180
         Width           =   4035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1260
         TabIndex        =   3
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3570
         TabIndex        =   19
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3570
         TabIndex        =   15
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -570
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -570
         TabIndex        =   1
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   6
         Top             =   240
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   5
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2475
      Left            =   45
      TabIndex        =   21
      Top             =   2295
      Width           =   6855
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         MaxLength       =   19
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   150
         Width           =   6735
      End
      Begin MSComctlLib.ListView lstVendor 
         Height          =   1875
         Left            =   30
         TabIndex        =   23
         Top             =   540
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   3307
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
         MouseIcon       =   "Vendor.frx":2D77
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "VENDOR NAME"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5415
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   33
      Top             =   4875
      Width           =   1485
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
         Left            =   720
         MouseIcon       =   "Vendor.frx":2ED9
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":302B
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "Vendor.frx":3369
         MousePointer    =   99  'Custom
         Picture         =   "Vendor.frx":34BB
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSVENDOR                                           As ADODB.Recordset
Dim ADDOREDIT                                          As String

'Upating Code       : AXP-0713200714:00
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Add", "VENDORS") = False Then Exit Sub

    ADDOREDIT = "ADD"
    InitMemVars
    Picture1.Visible = False
    Picture2.Visible = True
    lstVendor.Enabled = False
    txtSearch.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstVendor.Enabled = True
    fraDetails.Enabled = True
    txtSearch.Enabled = True
    StoreMemvars
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "VENDORS") = False Then Exit Sub
    On Error GoTo Errorcode:
    Dim LNGCOUNT                                       As Long


    LNGCOUNT = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_vw_PO_HISTORY WHERE SUPCODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If LNGCOUNT > 0 Then
        MsgBox "Vendor Record Exists in Purchase Order." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If


    LNGCOUNT = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_HD  WHERE VendorCode=" & N2Str2Null(txtCode)).Fields(0).Value
    If LNGCOUNT > 0 Then
        MsgBox "Vendor Record Exists in Journal Books." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If
    LNGCOUNT = gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_vw_RR_Trans WHERE RECVD_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If LNGCOUNT > 0 Then
        MsgBox "Vendor Record Exists in Parts Transaction." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If


    LNGCOUNT = gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_Po_hd WHERE CONTRACTOR_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If LNGCOUNT > 0 Then
        MsgBox "Vendor Record Exists in Sublet Purchase Order." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If



    LNGCOUNT = gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_PO_RC_HD WHERE CONTRACTOR_CODE=" & N2Str2Null(txtCode)).Fields(0).Value
    If LNGCOUNT > 0 Then
        MsgBox "Vendor Record Exists in Sublet Receiving." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If




    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        gconDMIS.Execute "delete from ALL_Vendor where Code = '" & lstVendor.SelectedItem & "'"
        LogAudit "X", "VENDOR MASTER FILE", lstVendor.SelectedItem.SubItems(1) & txtNameofVendor
        NEW_LogAudit "X", "VENDOR MASTER FILE", SQL_STATEMENT, "", "", txtCode, "", ""
    End If
    rsRefresh
    StoreMemvars
    FillGrid

    Exit Sub
Errorcode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "VENDORS") = False Then Exit Sub

    ADDOREDIT = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry (lstVendor.SelectedItem.SubItems(2))
    lstVendor.Enabled = False
    lstVendor.Enabled = False
    txtSearch.Enabled = False
    txtSearch.Enabled = False
    lstVendor.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

'Upating Code       : AXP-0713200714:00
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    RSVENDOR.MoveNext
    If RSVENDOR.EOF Then
        RSVENDOR.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:59
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    RSVENDOR.MovePrevious
    If RSVENDOR.BOF Then
        RSVENDOR.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "VENDORS") = False Then Exit Sub
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:


    Screen.MousePointer = 11
    rptVendor.Reset
    rptVendor.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptVendor.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptVendor, AMIS_REPORT_PATH & "\Files\Suppliers.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "VENDOR MASTER FILE", txtCode
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0

End Sub

Private Sub cmdSave_Click()

    On Error GoTo Errorcode

    Dim VtxtCode, VtxtNameofVendor, VtxtContact        As String
    Dim VtxtPosition, VtxtAddress, VtxtPhone           As String
    Dim VtxtTerms, VtxtFax, VtxtTIN                    As String
    Dim VchkVAT                                        As String
    Dim rsVendorDup                                    As ADODB.Recordset
    'VtxtCode = fN2Str2Null(txtCode.Text)
    VtxtCode = N2Str2Null(txtCode.Text)
    VtxtNameofVendor = N2Str2Null(txtNameofVendor.Text)
    VtxtContact = N2Str2Null(txtContact.Text)
    VtxtPosition = N2Str2Null(txtPosition.Text)
    VtxtAddress = N2Str2Null(txtAddress.Text)
    VtxtPhone = N2Str2Null(txtPhone.Text)
    VtxtTerms = N2Str2Null(txtTerms.Text)
    VtxtFax = N2Str2Null(txtFax.Text)
    VtxtTIN = N2Str2Null(txtTIN.Text)
    If Check1.Value = 1 Then
        VchkVAT = "'Y'"
    Else
        VchkVAT = "'N'"
    End If

    If ADDOREDIT = "ADD" Then

        Set rsVendorDup = New ADODB.Recordset
        rsVendorDup.Open "select code from ALL_Vendor where code = " & VtxtCode, gconDMIS
        If Not rsVendorDup.EOF And Not rsVendorDup.BOF Then
            MsgBox "Vendor Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
            Exit Sub
        End If
        SQL_STATEMENT = "Insert into ALL_Vendor " & _
                        "(Code,NameofVendor,contact,[position],address,phone,[Terms],fax,tin,NONVAT) " & _
                      " values (" & VtxtCode & _
                        ", " & VtxtNameofVendor & ", " & VtxtContact & _
                        ", " & VtxtPosition & ", " & VtxtAddress & _
                        ", " & VtxtPhone & ", " & VtxtTerms & _
                        ", " & VtxtFax & ", " & VtxtTIN & "," & VchkVAT & ")"
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(txtCode), "Code", "ALL_Vendor", ""))
        NEW_LogAudit "A", "VENDOR MASTER FILE", SQL_STATEMENT, TransactionID, "", txtCode, "", ""
    Else

        If txtCode <> Null2String(RSVENDOR!code) Then
            Set rsVendorDup = New ADODB.Recordset
            rsVendorDup.Open "select code from ALL_Vendor where code = " & VtxtCode, gconDMIS
            If Not rsVendorDup.EOF And Not rsVendorDup.BOF Then
                MsgBox "Vendor Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
                Exit Sub
            End If
        End If
        SQL_STATEMENT = "update ALL_Vendor set" & _
                      " NameofVendor = " & VtxtNameofVendor & "," & _
                      " contact = " & VtxtContact & "," & _
                      " [position] = " & VtxtPosition & "," & _
                      " address = " & VtxtAddress & "," & _
                      " phone = " & VtxtPhone & "," & _
                      " [Terms] = " & VtxtTerms & "," & _
                      " fax = " & VtxtFax & "," & _
                      " NONVAT = " & VchkVAT & "," & _
                      " tin = " & VtxtTIN & _
                      " where Code = " & VtxtCode
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(txtCode), "Code", "ALL_Vendor", ""))
        NEW_LogAudit "E", "VENDOR MASTER FILE", SQL_STATEMENT, TransactionID, "", txtCode, "", ""
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    RSVENDOR.Find "code = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    MsgBox "Error:" & err & " " & Error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "VENDOR MASTER FILE"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "VENDOR MASTER FILE")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    InitMemVars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set RSVENDOR = New ADODB.Recordset
    RSVENDOR.Open "select * from ALL_Vendor where Code <> '999999' order by NameofVendor asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemVars()
    Frame1.Enabled = True
    txtCode.Text = ""
    txtNameofVendor.Text = ""
    txtContact.Text = ""
    txtPosition.Text = ""
    txtAddress.Text = ""
    txtPhone.Text = ""
    txtTerms.Text = ""
    txtFax.Text = ""
    txtTIN.Text = ""
    txtSearch.Text = ""
End Sub

Sub StoreMemvars()
    If Not RSVENDOR.EOF And Not RSVENDOR.BOF Then
        Frame1.Enabled = False
        labID.Caption = RSVENDOR!ID
        txtCode.Text = Null2String(RSVENDOR!code)
        txtNameofVendor.Text = Null2String(RSVENDOR!Nameofvendor)
        txtContact.Text = Null2String(RSVENDOR!CONTACT)
        txtPosition.Text = Null2String(RSVENDOR!Position)
        txtAddress.Text = Null2String(RSVENDOR!Address)
        txtPhone.Text = Null2String(RSVENDOR!Phone)
        txtTerms.Text = Null2String(RSVENDOR!TERMS)
        txtFax.Text = Null2String(RSVENDOR!Fax)
        txtTIN.Text = Null2String(RSVENDOR!TIN)
        If Null2String(RSVENDOR!NONVAT) = "Y" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(xxx As Variant)
    Dim rsVendor2                                      As ADODB.Recordset
    Set rsVendor2 = New ADODB.Recordset
    rsVendor2.Open "select * from ALL_Vendor where ID = " & xxx, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVendor2.EOF And Not rsVendor2.BOF Then
        fraDetails.Enabled = False
        labID.Caption = rsVendor2!ID
        txtCode.Text = Null2String(rsVendor2!code)
        txtNameofVendor.Text = Null2String(rsVendor2!Nameofvendor)
        txtContact.Text = Null2String(rsVendor2!CONTACT)
        txtPosition.Text = Null2String(rsVendor2!Position)
        txtAddress.Text = Null2String(rsVendor2!Address)
        txtPhone.Text = Null2String(rsVendor2!Phone)
        txtTerms.Text = Null2String(rsVendor2!TERMS)
        txtFax.Text = Null2String(rsVendor2!Fax)
        txtTIN.Text = Null2String(rsVendor2!TIN)
        If Null2String(RSVENDOR!NONVAT) = "Y" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
    End If
End Sub

Private Sub FillGrid()
    Dim rsVendor2                                      As ADODB.Recordset
    lstVendor.Enabled = False
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendor2 = New ADODB.Recordset
    Set rsVendor2 = gconDMIS.Execute("select code,nameofvendor,ID from ALL_Vendor where code <> '999999' order by NameofVendor asc")
    If Not (rsVendor2.EOF And rsVendor2.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendor2
        lstVendor.Refresh
        lstVendor.Enabled = True
    End If

End Sub

Private Sub FillSearchGrid(xxx As String)
    Dim rsVendor2                                      As ADODB.Recordset

    lstVendor.Enabled = False: lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendor2 = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsVendor2 = gconDMIS.Execute("select code,nameofvendor,ID from ALL_Vendor where nameofvendor like '" & Repleys(xxx) & "%'")
    If Not (rsVendor2.EOF And rsVendor2.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendor2
        lstVendor.Refresh
        lstVendor.Enabled = True
    End If
End Sub

Private Sub lstVendor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstVendor
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

Private Sub lstVendor_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstVendor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    RSVENDOR.Bookmark = rsFind(RSVENDOR.Clone, "code", lstVendor.SelectedItem).Bookmark
    StoreMemvars
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtNameofVendor_Change()
    If ADDOREDIT = "ADD" Then
        If Len(txtNameofVendor.Text) = 1 Then
            Dim rsVendorCodes                          As ADODB.Recordset
            Set rsVendorCodes = New ADODB.Recordset
            rsVendorCodes.Open "select code from ALL_Vendor where LEFT(code,1) = '" & Left(txtNameofVendor.Text, 1) & "' ORDER BY CODE", gconDMIS
            If Not rsVendorCodes.EOF And Not rsVendorCodes.BOF Then
                rsVendorCodes.MoveLast
                txtCode.Text = Left(txtNameofVendor.Text, 1) & Format(N2Str2Zero(Mid(rsVendorCodes!code, 2, Len(rsVendorCodes!code) - 1)) + 1, "00000")
            Else
                txtCode.Text = Left(txtNameofVendor.Text, 1) & "00001"
            End If
        End If
    End If
End Sub

Private Sub txtNameofVendor_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstVendor.ListItems.Count > 0 And lstVendor.Enabled = True Then
            lstVendor.SetFocus
        End If
    End If
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub
