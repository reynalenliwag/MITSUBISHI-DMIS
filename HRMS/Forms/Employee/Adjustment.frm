VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Taxable/Non-Taxable Adjustment Entry "
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10470
   ForeColor       =   &H00000000&
   Icon            =   "Adjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   10470
   Begin VB.PictureBox picAdjustment 
      BackColor       =   &H80000009&
      Height          =   2355
      Left            =   4860
      Picture         =   "Adjustment.frx":0442
      ScaleHeight     =   2295
      ScaleWidth      =   3465
      TabIndex        =   10
      Top             =   1980
      Width           =   3525
      Begin VB.ComboBox cboQuensina 
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
         Height          =   330
         ItemData        =   "Adjustment.frx":26FB
         Left            =   60
         List            =   "Adjustment.frx":26FD
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Text            =   "cboQuensina"
         Top             =   90
         Width           =   3195
      End
      Begin VB.ComboBox cboAdjust 
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
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1170
         Width           =   2445
      End
      Begin VB.CheckBox chkTaxable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Taxable"
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
         Height          =   285
         Left            =   990
         Picture         =   "Adjustment.frx":26FF
         TabIndex        =   34
         Top             =   1560
         Width           =   1155
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   990
         TabIndex        =   35
         Top             =   1890
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dt_adjust 
         Height          =   345
         Left            =   60
         TabIndex        =   30
         Top             =   750
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   54460417
         CurrentDate     =   40179
      End
      Begin MSMask.MaskEdBox txtParticular 
         Height          =   315
         Left            =   2130
         TabIndex        =   33
         Top             =   1890
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   60
         TabIndex        =   15
         Top             =   510
         Width           =   645
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   -180
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         Height          =   285
         Left            =   30
         TabIndex        =   11
         Top             =   1230
         Width           =   915
      End
   End
   Begin wizButton.cmd cmdAdjustment 
      Height          =   2505
      Left            =   4830
      TabIndex        =   0
      Top             =   1950
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   4419
      TX              =   "cmd1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Adjustment.frx":49B8
   End
   Begin MSFlexGridLib.MSFlexGrid grdAdjustment 
      Height          =   4335
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   5
      ForeColor       =   0
      ForeColorFixed  =   0
      BackColorSel    =   14606302
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483633
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Adjustment.frx":49D4
   End
   Begin Crystal.CrystalReport rptAdjustment 
      Left            =   10980
      Top             =   540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   915
      Left            =   2610
      ScaleHeight     =   915
      ScaleWidth      =   7815
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtPosition 
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
         Left            =   870
         TabIndex        =   3
         Top             =   480
         Width           =   3765
      End
      Begin VB.TextBox txtName 
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
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   4605
      End
      Begin VB.TextBox txtYTDTaxableAdj 
         Alignment       =   1  'Right Justify
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
         Left            =   6390
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
      Begin VB.TextBox txtYTDNonTaxableAdj 
         Alignment       =   1  'Right Justify
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
         Left            =   6390
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
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
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Taxable Adj."
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
         Left            =   4710
         TabIndex        =   6
         Top             =   120
         Width           =   1785
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Non-Tax. Adj."
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
         Left            =   4710
         TabIndex        =   4
         Top             =   540
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   6105
      Left            =   0
      Picture         =   "Adjustment.frx":4CEE
      ScaleHeight     =   6045
      ScaleWidth      =   2445
      TabIndex        =   14
      Top             =   7560
      Width           =   2505
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   60
      Picture         =   "Adjustment.frx":18A4B
      ScaleHeight     =   6015
      ScaleWidth      =   2475
      TabIndex        =   16
      Top             =   120
      Width           =   2505
      Begin VB.TextBox txtSearch 
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   60
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   5505
         Left            =   0
         TabIndex        =   18
         Top             =   450
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   9710
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
         MouseIcon       =   "Adjustment.frx":1B787
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
         Picture         =   "Adjustment.frx":1B8E9
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4815
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   19
      Top             =   5445
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "Adjustment.frx":2F656
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":2F7A8
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Exit Window"
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
         Left            =   4170
         MouseIcon       =   "Adjustment.frx":2FB0E
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":2FC60
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Print this Record"
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
         Left            =   3480
         MouseIcon       =   "Adjustment.frx":2FFC6
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":30118
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Delete Selected Record"
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
         Left            =   2790
         MouseIcon       =   "Adjustment.frx":30443
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":30595
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   2100
         MouseIcon       =   "Adjustment.frx":308F1
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":30A43
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   1410
         MouseIcon       =   "Adjustment.frx":30D56
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":30EA8
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Find a Record"
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
         Left            =   720
         MouseIcon       =   "Adjustment.frx":311A2
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":312F4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "Adjustment.frx":3164C
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":3179E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   8955
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   28
      Top             =   5445
      Width           =   1440
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
         MouseIcon       =   "Adjustment.frx":31AFD
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":31C4F
         Style           =   1  'Graphical
         TabIndex        =   37
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
         MouseIcon       =   "Adjustment.frx":31F8D
         MousePointer    =   99  'Custom
         Picture         =   "Adjustment.frx":320DF
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label LABMONTH 
      BackColor       =   &H000000FF&
      Height          =   225
      Left            =   3840
      TabIndex        =   31
      Top             =   6630
      Width           =   3825
   End
End
Attribute VB_Name = "frmHRMSAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo, rsAdjustment                                           As ADODB.Recordset
Attribute rsAdjustment.VB_VarUserMemId = 1073938432
Dim ADDOREDIT, Diyt                                                   As String
Attribute ADDOREDIT.VB_VarUserMemId = 1073938434
Attribute Diyt.VB_VarUserMemId = 1073938434
Dim EMPLIVIL                                                          As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938436

Function StoreEntry(ByVal ID As Variant)
    Dim MM, DD, YY, TheDeyt                                           As String
    Dim ZEROS                                                         As String
    Set rsAdjustment = New ADODB.Recordset
    rsAdjustment.Open "select * from HRMS_Adjustment where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        LABID.Caption = rsAdjustment!ID
        If rsAdjustment!CUT_OFF = 1 Then cboQuensina.Text = "1st Cut-Off"
        If rsAdjustment!CUT_OFF = 2 Then cboQuensina.Text = "2nd Cut-Off"
        TheDeyt = Null2Date(rsAdjustment!DEYT)
        DD = Day(rsAdjustment!DEYT): MM = The_month(MONTH(TheDeyt)): YY = YEAR(TheDeyt)
        
        If rsAdjustment!Type = "NT" Then chkTaxable.Value = 0 Else chkTaxable.Value = 1
        
        'cboDay.Text = DD: cboMonth.Text = MM: cboYear.Text = YY
        dt_adjust = (rsAdjustment!DEYT)
        
        If Len(rsAdjustment!PARTICULAR) = 1 Then ZEROS = "0"
        cboAdjust.Text = GetAdjustmentDescription(rsAdjustment!PARTICULAR) & " - " & ZEROS & rsAdjustment!PARTICULAR
        txtParticular.Text = Null2String(rsAdjustment!PARTICULAR)
        txtAmount.Text = N2Str2Zero(rsAdjustment!AMOUNT)
    End If
End Function

Function GetAdjustmentDescription(CODE_ADJ As Integer) As String
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Description From HRMS_Codes_Adjustment Where Codes = " & CODE_ADJ & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetAdjustmentDescription = RSTMP!Description
    End If
    Set RSTMP = Nothing
End Function

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & EMPINFOEMPNO.Caption & "'", gconDMIS
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel from HRMS_EmpInfo where EMPLEVEL = " & EMPLIVIL & " AND empno = '" & frmHRMSEmpInfo.LABID.Caption & "'", gconDMIS
    Else
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel,resigned from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname asc", gconDMIS
        'rsEmpInfo.Open "select empno,[position],lastname,firstname,middlename,emplevel,resigned from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname asc", gconDMIS
    End If
End Sub

Sub InitGrid()
    With grdAdjustment
        .Rows = 2
        .ColWidth(0) = 1300: .ColWidth(1) = 1400: .ColWidth(2) = 3500: .ColWidth(3) = 1400: .ColWidth(4) = 1
        .Row = 0
        .Col = 0: .Text = "Date"
        .Col = 1: .Text = "Type"
        .Col = 2: .Text = "Particular"
        .Col = 3: .Text = "Amount"
        .Col = 4: .Text = "ID"
    End With
End Sub

Sub FillCboAdjustment()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ZEROS                                                         As String
    Set RSTMP = gconDMIS.Execute("Select * From HRMS_Codes_Adjustment Order By Description ASC")
    cboAdjust.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!codes) = 1 Then ZEROS = "0"
            cboAdjust.AddItem Null2String(RSTMP!Description) & " - " & ZEROS & RSTMP!codes
            RSTMP.MoveNext
        Loop
        cboAdjust.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

'Sub AddMonthName()
'    Dim X As Integer
'    cboMOnth.Clear
'    For X = 1 To 12
'        cboMOnth.AddItem MonthName(X)
'    Next
'End Sub

Sub InitMemvars()
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        
        '*****************************
        'Call AddMonthName
        'cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        LABMONTH.Caption = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    'fillcboDay cboDay
    FillCboAdjustment
    'cboDay.Text = Day(LOGDATE)
    txtParticular.Text = ""
    chkTaxable.Value = 1
    txtAmount.Text = 0
End Sub

Sub StoreMemVars()
    On Error GoTo Errorcode
    Dim CNT                                                           As Integer
    Dim Taayp                                                         As String
    Dim VYTDTaxableAdj, VYTDNonTaxableAdj                             As Double
    Dim DESC_ADJ                                                      As String

    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set rsAdjustment = New ADODB.Recordset
        rsAdjustment.Open "SELECT * FROM HRMS_ADJUSTMENT WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        CNT = 0: VYTDTaxableAdj = 0: VYTDNonTaxableAdj = 0
        If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
            rsAdjustment.MoveFirst
            cleargrid grdAdjustment
            grdAdjustment.Rows = grdAdjustment.Rows
            Do While Not rsAdjustment.EOF
                CNT = CNT + 1
                LABID.Caption = rsAdjustment!ID
                DESC_ADJ = GetAdjustmentDescription(rsAdjustment!PARTICULAR)
                If Null2String(rsAdjustment!Type) = "NT" Then Taayp = "Non-Taxable" Else Taayp = "Taxable"
                grdAdjustment.AddItem Null2Date(rsAdjustment!DEYT) & Chr(9) & Taayp & Chr(9) & UCase(DESC_ADJ) & Chr(9) & N2Str2Zero(rsAdjustment!AMOUNT) & Chr(9) & rsAdjustment!ID
                If YEAR(Null2Date(rsAdjustment!DEYT)) = YEAR(LOGDATE) Then
                    If Null2String(rsAdjustment!Type) = "T" Then
                        VYTDTaxableAdj = VYTDTaxableAdj + N2Str2Zero(rsAdjustment!AMOUNT)
                    End If
                    If Null2String(rsAdjustment!Type) = "NT" Then
                        VYTDNonTaxableAdj = VYTDNonTaxableAdj + N2Str2Zero(rsAdjustment!AMOUNT)
                    End If
                End If
                rsAdjustment.MoveNext
            Loop
            grdAdjustment.RemoveItem 1
        Else
            cleargrid grdAdjustment
        End If
        txtYTDTaxableAdj.Text = N2Str2Zero(VYTDTaxableAdj)
        txtYTDNonTaxableAdj.Text = N2Str2Zero(VYTDNonTaxableAdj)
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL order by lastname+', '+firstname asc")
    'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False: lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo  where EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    'Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno from HRMS_EmpInfo where RESIGNED IS NULL and lastname+', '+firstname like'" & XXX & "%' order by lastname+', '+firstname asc")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Private Sub cboAdjust_Change()
    If Right(cboAdjust, 3) = "003" Then
        chkTaxable.Value = 0
        txtAmount.Text = "100"
    ElseIf Right(cboAdjust, 3) = "004" Then
        txtAmount.Text = "150"
        chkTaxable.Value = 0
    Else
        txtAmount.Text = "0.00"
        chkTaxable.Value = 1
    End If
End Sub

Private Sub cboAdjust_Click()
    If Right(cboAdjust, 3) = "003" Then
        txtAmount.Text = "100"
        chkTaxable.Value = 0
    ElseIf Right(cboAdjust, 3) = "004" Then
        txtAmount.Text = "150"
        chkTaxable.Value = 0
    Else
        txtAmount.Text = "0.00"
        chkTaxable.Value = 1
    End If
End Sub

Private Sub cboAdjust_LostFocus()
    If Right(cboAdjust, 3) = "003" Then
        txtAmount.Text = "100"
        chkTaxable.Value = 0
    ElseIf Right(cboAdjust, 3) = "004" Then
        txtAmount.Text = "150"
        chkTaxable.Value = 0
    Else
        txtAmount.Text = "0.00"
        chkTaxable.Value = 1
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN ADJUSTMENTS") = False Then Exit Sub
    ADDOREDIT = "ADD"
    picAdjustment.Visible = True
    cmdAdjustment.Visible = True
    picSearch.Enabled = False
    picAdjustment.Enabled = True
    cmdAdjustment.ZOrder 0
    picAdjustment.ZOrder 0
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
    
    dt_adjust = LOGDATE
    On Error Resume Next
    txtParticular.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:
    ADDOREDIT = ""
    picAdjustment.Visible = False
    cmdAdjustment.Visible = False
    picSearch.Enabled = True
    picAdjustment.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    cmdAdjustment.ZOrder 1
    picAdjustment.ZOrder 1
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN ADJUSTMENTS") = False Then Exit Sub
    On Error GoTo Errorcode
    grdAdjustment.Col = 4
    If grdAdjustment.Text <> "" Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from HRMS_Adjustment where id = " & grdAdjustment.Text
            LogAudit "X", "DELETE EMPLOYEE ADJUSTMENT FILE", grdAdjustment.Text
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN ADJUSTMENTS") = False Then Exit Sub
    Dim fild                                                          As String
    grdAdjustment.Row = grdAdjustment.Row
    grdAdjustment.Col = 4
    fild = grdAdjustment.Text
    If fild <> "" Then
        ADDOREDIT = "EDIT"
        cmdAdjustment.Visible = True
        cmdAdjustment.ZOrder 0
        picSearch.Enabled = False
        picAdjustment.Visible = True
        picAdjustment.ZOrder 0
        picAdjustment.Enabled = True
        Picture1.Visible = False
        Picture2.Visible = True
        StoreEntry fild
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtsearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN ADJUSTMENTS") = False Then Exit Sub
    Screen.MousePointer = 11
    rptAdjustment.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptAdjustment.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptAdjustment.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptAdjustment.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
    PrintSQLReport rptAdjustment, HRMS_REPORT_PATH & "Adjustment.rpt", "{HRMS_Adjustment.empno} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {HRMS_Adjustment.CUT_OFF} = '" & CUTTOFF_CODE & "' AND {HRMS_Adjustment.PAY_MONTH} = " & PAY_MONTH & " AND {HRMS_Adjustment.PAY_YEAR} = " & PAY_YEAR & " ", DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT EMPLOYEE ADJUSTMENT FILE", rsEmpInfo!EMPNO
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim MM, DD, YY, TAYP                                              As String
    Dim ADJ_CODE                                                      As String
    Dim vCUTOFF                                                       As Integer
    If cboQuensina.Text = "" Then
        ShowIsRequiredMsg "Choose a Cut-Off"
        cboQuensina.SetFocus
        Exit Sub
    End If
    If txtAmount.Text = "" Then
        MsgBox "Enter the Amount to be Adjust", vbInformation, "Required"
        txtAmount.SetFocus
        Exit Sub
    End If
    ADJ_CODE = Right(cboAdjust, 3)
    If cboQuensina.Text = "1st Cut-Off" Then vCUTOFF = 1
    If cboQuensina.Text = "2nd Cut-Off" Then vCUTOFF = 2
    
    'MM = What_month(cboMonth)
    'YY = cboYear
    'MM = What_month(cboMonth): YY = cboYear.Text: DD = cboDay.Text
    
    MM = MONTH(dt_adjust)
    YY = YEAR(dt_adjust)
    DD = Day(dt_adjust)
    
    
    
'
'    Dim rs                                                   As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Set rs = gconDMIS.Execute("select emplevel from hrms_empinfo where empno = '" & rsEmpInfo!EMPNO & "'")
'    If Not (rs.EOF And rs.BOF) Then
'        EMPLIVIL = N2Str2Null(rsEmpInfo!EMPLEVEL)
'    End If
    
    
    
    
    
    Diyt = DateSerial(YY, MM, DD)
    If chkTaxable.Value = 1 Then TAYP = "T" Else TAYP = "NT"
    
    If ADDOREDIT = "ADD" Then
        'COMMENT BY  : MJP 010908 1017AM
        'DESCRIPTION :
            'gconDMIS.Execute "Insert into HRMS_Adjustment " & _
            '                 "(EMPLEVEL, Empno, Deyt, Type, Particular, Amount, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
            '                 "Values (" & EMPLIVIL & _
            '                 "," & N2Str2Null(RSEMPINFO!EMPNO) & _
            '                 "," & N2Date2Null(Diyt) & _
            '                 "," & N2Str2Null(TAYP) & _
            '                 ", '" & ADJ_CODE & _
            '                 "'," & txtAmount.Text & _
            '                 "," & vCUTOFF & _
            '                 "," & What_month(LABMONTH) & _
            '                 "," & YY & ")"
        'COMMENT BY  : MJP 010908 1017AM
        
        'UPDATE BY   : MJP 010908 1017AM
        'DESCRIPTION :
            gconDMIS.Execute "Insert into HRMS_Adjustment " & _
                             "(EMPLEVEL, Empno, Deyt, Type, Particular, Amount, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                             "Values (" & EMPLIVIL & _
                             "," & N2Str2Null(rsEmpInfo!EMPNO) & _
                             "," & N2Date2Null(Diyt) & _
                             "," & N2Str2Null(TAYP) & _
                             ", '" & ADJ_CODE & _
                             "'," & txtAmount.Text & _
                             "," & vCUTOFF & _
                             "," & What_month(LABMONTH) & _
                             "," & PAY_YEAR & ")"
        'UPDATE BY   : MJP 010908 1017AM
        
        LogAudit "A", "ADD EMPLOYEE ADJUSTMENT FILE", rsEmpInfo!EMPNO
        ShowSuccessFullyAdded
    Else
        'COMMENT BY  : MJP 010908 1017AM
        'DESCRIPTION :
            'gconDMIS.Execute "update HRMS_Adjustment set" & _
            '               " Empno = " & N2Str2Null(RSEMPINFO!EMPNO) & "," & _
            '               " Deyt = " & N2Date2Null(Diyt) & "," & _
            '               " Type = " & N2Str2Null(TAYP) & "," & _
            '               " Particular = '" & LTrim(RTrim(ADJ_CODE)) & "'," & _
            '               " Amount = " & txtAmount.Text & "," & _
            '               " CUT_OFF = " & vCUTOFF & "," & _
            '               " PAY_MONTH = " & What_month(LABMONTH) & "," & _
            '               " PAY_YEAR = " & YY & _
            '               " Where id = " & labID.Caption
        'COMMENT BY  : MJP 010908 1017AM
        
        'UPDATE BY   : MJP 010908 1017AM
        'DESCRIPTION :
            gconDMIS.Execute "update HRMS_Adjustment set" & _
                           " Empno = " & N2Str2Null(rsEmpInfo!EMPNO) & "," & _
                           " Deyt = " & N2Date2Null(Diyt) & "," & _
                           " Type = " & N2Str2Null(TAYP) & "," & _
                           " Particular = '" & LTrim(RTrim(ADJ_CODE)) & "'," & _
                           " Amount = " & txtAmount.Text & "," & _
                           " CUT_OFF = " & vCUTOFF & "," & _
                           " PAY_MONTH = " & What_month(LABMONTH) & "," & _
                           " PAY_YEAR = " & PAY_YEAR & _
                           " Where id = " & LABID.Caption
        'UPDATE BY   : MJP 010908 1017AM
        
        LogAudit "E", "UPDATE EMPLOYEE ADJUSTMENT FILE", rsEmpInfo!EMPNO
        ShowSuccessFullyUpdated
    End If
    cmdCancel.Value = True
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyEscape
            cmdCancel.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    txtsearch.Text = ""
    rsrefresh
    InitGrid
    InitMemvars
    cmdCancel_Click
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

'Private Sub txtAmount_KeyPress(KeyAscii As Integer)
'KeyAscii = OnlyNumeric(KeyAscii)
'End Sub

Private Sub txtParticular_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    'On Error Resume Next
    'rsEmpInfo.Bookmark = rsFIND(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    'StoreMemVars
    
    rsEmpInfo.Requery
    rsEmpInfo.Find ("EMPNO=" & ITEM.ListSubItems(1).Text)
    StoreMemVars
    
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
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

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then FillGrid Else FillSearchGrid (txtsearch.Text)
End Sub

