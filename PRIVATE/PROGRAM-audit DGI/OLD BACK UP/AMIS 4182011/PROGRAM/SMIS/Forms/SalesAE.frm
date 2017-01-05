VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Files_SalesAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Account Executives"
   ClientHeight    =   8565
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "SalesAE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update SAE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   4290
      MaskColor       =   &H0000FFFF&
      Picture         =   "SalesAE.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Update List"
      Top             =   870
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   4620
      Left            =   30
      TabIndex        =   11
      Top             =   -60
      Width           =   5820
      Begin VB.TextBox txtquota 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   4170
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker DTHIRED 
         Height          =   315
         Left            =   1350
         TabIndex        =   6
         Top             =   2730
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   50593793
         CurrentDate     =   39709
      End
      Begin MSComCtl2.DTPicker DTRESIGNED 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   50593793
         CurrentDate     =   39709
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contract"
         Height          =   615
         Left            =   1320
         TabIndex        =   39
         Top             =   3480
         Width           =   3255
         Begin VB.OptionButton OPT_CON_COMMISSIONED 
            Caption         =   "Commissioned"
            Height          =   255
            Left            =   1470
            TabIndex        =   10
            Top             =   270
            Width           =   1605
         End
         Begin VB.OptionButton OPT_CON_SAL 
            Caption         =   "Salaried"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.TextBox txtTINNO 
         Height          =   345
         Left            =   1350
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2310
         Width           =   2475
      End
      Begin VB.CheckBox CHK_HARICERTIFIED 
         Caption         =   "Hari Certified"
         Height          =   225
         Left            =   2820
         TabIndex        =   8
         Top             =   3180
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   10
         PasswordChar    =   "l"
         TabIndex        =   4
         Top             =   1860
         Width           =   2475
      End
      Begin VB.TextBox txtSAECODE 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   3930
         TabIndex        =   34
         Top             =   480
         Width           =   1605
      End
      Begin Crystal.CrystalReport rptSAE 
         Left            =   4365
         Top             =   1215
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2445
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   630
         Width           =   2475
      End
      Begin VB.TextBox txtMiddleName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   2
         Top             =   1050
         Width           =   2475
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quota"
         Height          =   225
         Index           =   3
         Left            =   750
         TabIndex        =   44
         Top             =   4230
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Date Resigned"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   42
         Top             =   3150
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Date Hired"
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Tin Number"
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   390
         TabIndex        =   36
         Top             =   1890
         Width           =   840
      End
      Begin VB.Label labEclass 
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
         Left            =   4320
         TabIndex        =   33
         Top             =   2550
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SAE Code"
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
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Team"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   225
         TabIndex        =   20
         Top             =   1485
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   270
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   300
         TabIndex        =   15
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   30
      TabIndex        =   17
      Top             =   4500
      Width           =   5775
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
         Left            =   1350
         MaxLength       =   35
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   150
         Width           =   4305
      End
      Begin MSComctlLib.ListView lstExecutive 
         Height          =   2505
         Left            =   60
         TabIndex        =   19
         Top             =   540
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   4419
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
         MouseIcon       =   "SalesAE.frx":0BD4
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Team"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Search Name "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   38
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   885
      Left            =   4335
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   30
      Top             =   7650
      Width           =   2580
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
         Left            =   750
         MouseIcon       =   "SalesAE.frx":0D36
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":0E88
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cancel"
         Top             =   45
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
         Left            =   60
         MouseIcon       =   "SalesAE.frx":11C6
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":1318
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   900
      Left            =   -1380
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   21
      Top             =   7650
      Width           =   9225
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
         Left            =   6480
         MouseIcon       =   "SalesAE.frx":1668
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   5790
         MouseIcon       =   "SalesAE.frx":1B20
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":1C72
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   5100
         MouseIcon       =   "SalesAE.frx":1FD8
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":212A
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   4410
         MouseIcon       =   "SalesAE.frx":2455
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":25A7
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   3720
         MouseIcon       =   "SalesAE.frx":2903
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":2A55
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   3030
         MouseIcon       =   "SalesAE.frx":2D68
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":2EBA
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   2340
         MouseIcon       =   "SalesAE.frx":31B4
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":3306
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   1650
         MouseIcon       =   "SalesAE.frx":365E
         MousePointer    =   99  'Custom
         Picture         =   "SalesAE.frx":37B0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labSAECODE 
      Caption         =   "Label4"
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
      Left            =   6060
      TabIndex        =   13
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6060
      TabIndex        =   12
      Top             =   510
      Width           =   1815
   End
End
Attribute VB_Name = "frmSMIS_Files_SalesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSrep                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub FillGrid()
    Dim rsSrep                                                        As ADODB.Recordset
    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
    Set rsSrep = New ADODB.Recordset
    Set rsSrep = gconDMIS.Execute("select Name,TeamName,ID from SMIS_vw_Srep order by name asc")
    If Not (rsSrep.EOF And rsSrep.BOF) Then
        Listview_Loadval Me.lstExecutive.ListItems, rsSrep
        lstExecutive.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsSrep                                                        As ADODB.Recordset
    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
    Set rsSrep = New ADODB.Recordset
    Set rsSrep = gconDMIS.Execute("select Name,TeamName,ID from SMIS_vw_Srep WHERE name like'" & ReplaceQuote(XXX) & "%' order by name asc")
    If Not (rsSrep.EOF And rsSrep.BOF) Then
        Listview_Loadval Me.lstExecutive.ListItems, rsSrep
        lstExecutive.Refresh
    End If
End Sub

Sub InitMemVars()
    txtSAECODE = ""
    Combo1 = ""
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtMiddleName.Text = ""
    txtTINNO = ""
    DTHIRED = LOGDATE
    CHK_HARICERTIFIED.Value = 0
    Text2 = ""
    txtquota = "0"
End Sub

Sub rsRefresh()
    Set rsSrep = New ADODB.Recordset
    rsSrep.Open "select * from SMIS_SalesTeam ORDER BY ID DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
On Error GoTo ErrorCode:
    If Not rsSrep.EOF And Not rsSrep.BOF Then
        labid.Caption = rsSrep!ID
        txtFirstName.Text = Null2String(rsSrep!fname)
        txtLastName.Text = Null2String(rsSrep!lname)
        txtMiddleName.Text = Null2String(rsSrep!MIDDLE)
        txtSAECODE = Null2String(rsSrep!SAECODE)
        labSAECODE = Null2String(rsSrep!SAECODE)
        Combo1.Text = Null2String(rsSrep!TeamName)
        labEclass = Null2String(rsSrep!eclass)
        Text2 = Null2String(rsSrep!Password)
        txtTINNO = Null2String(rsSrep!TINNO)
        DTHIRED = Null2String(rsSrep!Date_Hired)
        DTRESIGNED = Null2String(rsSrep!Date_Resigned)
        
        If Null2String(rsSrep!Contract) = "" Then
            OPT_CON_SAL.Value = True
        Else
            OPT_CON_COMMISSIONED.Value = True
        End If
        If Null2String(rsSrep!HARI_Certified) = "Y" Then
            CHK_HARICERTIFIED.Value = 1
        Else
            CHK_HARICERTIFIED.Value = 0
        End If
        
     
        If Null2String(labEclass) = "HR" Then
            txtSAECODE.Enabled = False
        Else
            txtSAECODE.Enabled = True
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
Exit Sub

ErrorCode:

'    If Err.Number = 3265 Then
'        gconDMIS.Execute ("alter table SMIS_SALESTEAM " & vbCrLf & "ADD HARI_Certified nvarchar(1)")
'        gconDMIS.Execute "Alter table SMIS_SALESTEAM " & vbCrLf & " ADD Contract  nvarchar(1)"
'        gconDMIS.Execute "alter table SMIS_SALESTEAM " & vbCrLf & " ADD TINNO nvarchar(50)"
'        gconDMIS.Execute "alter table SMIS_SALESTEAM " & vbCrLf & " ADD Date_Resigned    smalldatetime"
'        gconDMIS.Execute "alter table SMIS_SALESTEAM " & vbCrLf & " ADD Date_Hired smalldatetime    "
'        gconDMIS.Execute "alter table SMIS_SALESTEAM " & vbCrLf & " ADD quota integer "
'    End If
Err.Clear
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True

    txtSAECODE.Enabled = True
    InitMemVars
    lstExecutive.Enabled = False
    txtSEARCH.Enabled = False
    txtSAECODE = GenerateCode("SMIS_SALESTEAM", "SAECODE", "0000")
    On Error Resume Next
    txtLastName.SetFocus

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstExecutive.Enabled = True
    txtSEARCH.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode
    Dim TEMPRS                                                        As ADODB.Recordset
    If Not rsSrep.BOF Or Not rsSrep.EOF Then
        If ShowConfirmDelete = True Then
            Set TEMPRS = gconDMIS.Execute("SELECT COUNT(*) FROM CRIS_PROSPECTS WHERE usercode='" & labSAECODE & "'")
            If TEMPRS(0).Value > 0 Then
                MsgSpeechBox ("Sales Account Executive Record Exist in Database." & vbCrLf & "Record cannot be deleted")
                Exit Sub
            Else
                gconDMIS.Execute "delete from SMIS_vw_Srep where id = " & labid.Caption
                LogAudit "X", "SALES AGENT INFORMATION", txtLastName & " " & txtFirstName
            End If

        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars

    FillSearchGrid ""
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    txtLastName.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsSrep.MoveNext
    If rsSrep.EOF Then
        rsSrep.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSrep.MovePrevious
    If rsSrep.BOF Then
        rsSrep.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()

    If Function_Access(LOGID, "Acess_PRINT", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub

    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    With frmMain.rptMain
        .ReportTitle = "SALES EXECITIVE LISTING"
        .Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        .Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        .Connect = DMIS_REPORT_Connection
        .WindowTitle = "SALES PERSONNEL LIST"
        .ReportFileName = SMIS_REPORT_PATH & "listing\sae.rpt"
        .Action = 1
        Screen.MousePointer = 0
        LogAudit "G", "SALES AGENT LISTING"
    End With





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorCode:

    If txtLastName.Text = "" Or txtFirstName.Text = "" Then
        ShowIsRequiredMsg "Code and Description"
        Exit Sub
    End If
    Dim VTXTName                                                      As String
    Dim VTXTLASTNAME                                                  As String
    Dim VTXTFIRSTNAME                                                 As String
    Dim VTXTMIDDLENAME                                                As String
    Dim VTXTTEAMNAME                                                  As String
    Dim VTXTSAECODE                                                   As String
    Dim VTXTPASSWORD                                                  As String
    Dim QUOTA As String
    Dim lng                                                           As Integer

    VTXTLASTNAME = N2Str2Null(txtLastName.Text)
    VTXTFIRSTNAME = N2Str2Null(txtFirstName.Text)
    VTXTMIDDLENAME = N2Str2Null(txtMiddleName.Text)
    VTXTTEAMNAME = N2Str2Null(Combo1)
    VTXTSAECODE = N2Str2Null(txtSAECODE)
    VTXTPASSWORD = N2Str2Null(Text2)
    QUOTA = NumericVal(txtquota)
Dim HARI_Certified As String
Dim Contract As String
Dim TINNO As String
Dim Date_Resigned As String
Dim Date_Hired As String

    If CHK_HARICERTIFIED.Value = 1 Then
        HARI_Certified = "'Y'"
    Else
        HARI_Certified = "'N'"
    End If
    
    If OPT_CON_COMMISSIONED.Value = True Then
        Contract = "'C'"
    Else
        Contract = "'S'"
    End If
    TINNO = N2Str2Null(txtTINNO)
    Date_Resigned = N2Str2Null(DTRESIGNED)
    Date_Hired = N2Str2Null(DTHIRED)
        
    
    If RTrim(RTrim(txtSAECODE)) = "" Then
        ShowIsRequiredMsg " SAE Code"
        On Error Resume Next
        txtSAECODE.SetFocus
        Exit Sub
    End If
    ''''''
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESTEAM WHERE SAECODE=" & VTXTSAECODE).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "SAE Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsSrep!SAECODE)) <> UCase(txtSAECODE) Then
            MessagePop RecSaveWarning, "Duplicate Record", "SAE Code Already Exist"
            Exit Sub
        End If
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into SMIS_SalesTeam" & _
                      " (QUOTA, SAECODE, fname,lname,middle,teamname,eclass,PASSWORD,HARI_Certified,Contract,TINNO,Date_Resigned,Date_Hired)" & _
                      " values (" & QUOTA & "," & VTXTSAECODE & "," & VTXTFIRSTNAME & ", " & VTXTLASTNAME & ", " & VTXTMIDDLENAME & ", " & VTXTTEAMNAME & ",'SM'," & VTXTPASSWORD & "," & HARI_Certified & "," & Contract & "," & TINNO & "," & Date_Resigned & "," & Date_Hired & " )"
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "SALES ACCOUNT EXECUTIVE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtSAECODE), "SAECODE", "SMIS_SALESTEAM"), "", "SAECODE:" & txtSAECODE, "", ""
        LogAudit "A", "SALES AGENT INFORMATION", txtLastName & " " & txtFirstName


        '***********RESET THE SQL_STATEMENT VARIABLE**********
        SQL_STATEMENT = ""
        '***********RESET THE SQL_STATEMENT VARIABLE**********
    Else
        If labEclass = "HR" Then
            SQL_STATEMENT = " update HRMS_EmpInfo set" & _
                          " Firstname = " & VTXTFIRSTNAME & "," & _
                          " Lastname = " & VTXTLASTNAME & "," & _
                          " Middlename = " & VTXTMIDDLENAME & _
                          " where EMPNO = " & N2Str2Null(labSAECODE)
            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "EE", "SALES ACCOUNT EXECUTIVE", SQL_STATEMENT, FindTransactionID(N2Str2Null(labSAECODE), "EMPNO", "HRMS_EmpInfo"), "", "SAECODE:" & txtSAECODE, "", ""

            '***********RESET THE SQL_STATEMENT VARIABLE**********
            SQL_STATEMENT = ""
            '***********RESET THE SQL_STATEMENT VARIABLE**********


            SQL_STATEMENT = " update SMIS_SalesTeam set" & _
                          " FNAME = " & VTXTFIRSTNAME & "," & _
                          " LNAME= " & VTXTLASTNAME & "," & _
                          " Middle= " & VTXTMIDDLENAME & "," & _
                          " TEAMNAME= " & VTXTTEAMNAME & "," & _
                          " HARI_Certified= " & HARI_Certified & "," & _
                          " Contract= " & Contract & "," & _
                            " quota= " & QUOTA & "," & _
                          " TINNO= " & TINNO & "," & _
                          " Date_Resigned= " & Date_Resigned & "," & _
                          " Date_Hired= " & Date_Hired & _
                          " where id = " & labid.Caption

            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "E", "SALES ACCOUNT EXECUTIVE", SQL_STATEMENT, N2Str2Null(labid), "", "SAECODE:" & txtSAECODE, "", ""
            '***********RESET THE SQL_STATEMENT VARIABLE**********
            SQL_STATEMENT = ""
            '***********RESET THE SQL_STATEMENT VARIABLE**********
        Else
            SQL_STATEMENT = " update SMIS_SalesTeam set" & _
                          " FNAME = " & VTXTFIRSTNAME & "," & _
                          " LNAME= " & VTXTLASTNAME & "," & _
                          " Middle= " & VTXTMIDDLENAME & "," & _
                          " TEAMNAME= " & VTXTTEAMNAME & "," & _
                          " PASSWORD= " & VTXTPASSWORD & "," & _
                          " quota= " & QUOTA & "," & _
                          " SAECODE= " & VTXTSAECODE & "," & _
                          " HARI_Certified= " & HARI_Certified & "," & _
                          " Contract= " & Contract & "," & _
                          " TINNO= " & TINNO & "," & _
                          " Date_Resigned= " & Date_Resigned & "," & _
                          " Date_Hired= " & Date_Hired & _
                          " where id = " & labid.Caption

            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "E", "SALES ACCOUNT EXECUTIVE", SQL_STATEMENT, N2Str2Null(labid), "", "SAECODE:" & txtSAECODE, "", ""
            '***********RESET THE SQL_STATEMENT VARIABLE**********
            SQL_STATEMENT = ""
            '***********RESET THE SQL_STATEMENT VARIABLE**********
        End If
    End If
    rsRefresh
    Call FillCombo("SELECT DISTINCT TEAMNAME  from SMIS_vw_Srep WHERE LEN(TEAMNAME)>0", -1, 0, Combo1)
    If AddorEdit = "EDIT" Then
        rsSrep.Find "ID =" & labid.Caption
    End If
    cmdCancel.Value = True
    FillGrid
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdUpdate_Click()
    Dim RsSAE                                                         As ADODB.Recordset
    Dim cnt                                                           As Integer
    Set RsSAE = gconDMIS.Execute("SELECT FIRSTNAME, LASTNAME,MIDDLENAME, EMPNO AS SAECODE ,TINNO,DATEHIRED,RESIGNED FROM HRMS_EMPINFO WHERE EMPNO NOT  IN (SELECT SAECODE FROM SMIS_SALESTEAM WHERE SAECODE IS NOT NULL) AND IS_SAE=1")
    cnt = 0
    While Not RsSAE.EOF
        If MsgBox(" Employee  Name : " & Null2String(RsSAE!Firstname) & " " & Null2String(RsSAE!lastname) & "  " & Null2String(Left(RsSAE!MIDDLENAME, 1)) & vbCrLf & " Employee No: " & Null2String(RsSAE("SAECODE")) & vbCrLf & " Is not In Your Sales Data Base !" & vbCrLf & " Do You Want to Add", vbYesNo + vbExclamation) = vbYes Then
            SQL_STATEMENT = ("INSERT INTO SMIS_SALESTEAM (SAECODE,FNAME,LNAME,MIDDLE,ECLASS,TINNO,Date_Resigned,Date_Hired) VALUES(" & N2Str2Null(RsSAE!SAECODE) & " ," & N2Str2Null(RsSAE!Firstname) & " ," & N2Str2Null(RsSAE!lastname) & " ," & N2Str2Null(RsSAE!MIDDLENAME) & ",'HR'," & N2Str2Null(RsSAE!TINNO) & "," & N2Str2Null(RsSAE!DATEHIRED) & "," & N2Str2Null(RsSAE!RESIGNED) & ")")
            gconDMIS.Execute (SQL_STATEMENT)
            '****************NEW LOG AUDIT***********************
            NEW_LogAudit "A", "SALES AGENT INFORMATION", SQL_STATEMENT, FindTransactionID(N2Str2Null(RsSAE!SAECODE), "SAECODE", "SMIS_SALESTEAM"), "", "SAE CODE" & N2Str2Null(RsSAE!SAECODE), "", ""
            '****************NEW LOG AUDIT***********************
            LogAudit "A", "SALES AGENT INFORMATION FROM HR:INFO ", Null2String(RsSAE!lastname) & "  " & Null2String(Left(RsSAE!MIDDLENAME, 1)) & " Employee No: " & Null2String(RsSAE("SAECODE"))

            cnt = 1
        End If
        RsSAE.MoveNext
    Wend

    Set RsSAE = Nothing
    If cnt = 0 Then
        MessagePop InfoFriend, "UPDATE", "THE DATABASE LIST IS UP TO DATE"
    Else
        MessagePop RecSaveInfo, "UPDATED", "DATABASE LIST UPDATED"
        gconDMIS.Execute (" Update cris_prospects Set SAE = b.Name FROM cris_prospects A , SMIS_VW_SREP B Where A.USERCODE = b.SAECODE")
        rsRefresh
        StoreMemVars
        FillSearchGrid ""
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES ACCOUNT EXECUTIVE)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "SALES ACCOUNT EXECUTIVE")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False
    InitMemVars
    Picture1.Visible = True
    Picture2.Visible = False

    rsRefresh
    StoreMemVars
    Call ResizeColumnHeader(lstExecutive, "60,32")
    txtSEARCH = ""
    Call FillCombo("SELECT DISTINCT TEAMNAME  from SMIS_SalesTeam WHERE LEN(TEAMNAME)>0", -1, 0, Combo1)
End Sub

Private Sub lstExecutive_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstExecutive
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

Private Sub lstExecutive_DblClick()

    If Not lstExecutive.ListItems.Count = 0 Then
        cmdEdit.Value = True
    End If
End Sub

Private Sub lstExecutive_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsSrep.MoveFirst
    rsSrep.Find ("id=" & lstExecutive.SelectedItem.SubItems(2))
    StoreMemVars
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub txtquota_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSEARCH.Text) = "" Then FillGrid Else FillSearchGrid (txtSEARCH.Text)
End Sub

