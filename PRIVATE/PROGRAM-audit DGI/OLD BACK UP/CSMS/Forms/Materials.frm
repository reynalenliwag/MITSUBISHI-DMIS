VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSMaterials 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Data Entry"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Materials.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8895
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   2775
      ScaleHeight     =   945
      ScaleWidth      =   6315
      TabIndex        =   22
      Top             =   2760
      Width           =   6315
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
         Left            =   5280
         MouseIcon       =   "Materials.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Left            =   4560
         MouseIcon       =   "Materials.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   3840
         MouseIcon       =   "Materials.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":138C
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
         Left            =   3120
         MouseIcon       =   "Materials.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   2400
         MouseIcon       =   "Materials.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   1680
         MouseIcon       =   "Materials.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   960
         MouseIcon       =   "Materials.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   240
         MouseIcon       =   "Materials.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   2670
      TabIndex        =   8
      Top             =   30
      Width           =   6135
      Begin VB.TextBox txtLocation 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2190
         Width           =   2115
      End
      Begin VB.TextBox txtOnHand 
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
         Height          =   345
         Left            =   3720
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtSStock 
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
         Height          =   345
         Left            =   3720
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1410
         Width           =   1065
      End
      Begin Crystal.CrystalReport rptMatMas 
         Left            =   150
         Top             =   1050
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Material Master List"
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
      Begin VB.TextBox txtMatCde 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtPOCode 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtCost 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1410
         Width           =   1065
      End
      Begin VB.TextBox txtS_Price 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtMatDsc 
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
         Height          =   345
         Left            =   1410
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   630
         Width           =   4665
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   255
         Left            =   660
         TabIndex        =   18
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "On-Hand"
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
         Height          =   255
         Left            =   2910
         TabIndex        =   17
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Stock"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
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
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO Code"
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
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
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
         Height          =   255
         Left            =   930
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Height          =   255
         Left            =   930
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   690
         Width           =   975
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3555
      Left            =   60
      TabIndex        =   19
      Top             =   30
      Width           =   2565
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   150
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstMaterials 
         Height          =   2955
         Left            =   60
         TabIndex        =   21
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
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
         MouseIcon       =   "Materials.frx":2D71
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DESCRIPTION"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7290
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   31
      Top             =   2745
      Width           =   1800
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
         Left            =   780
         MouseIcon       =   "Materials.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         MouseIcon       =   "Materials.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "Materials.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   330
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   660
      TabIndex        =   13
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMatMas                                           As ADODB.Recordset
Dim AddorEdit                                          As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "MATERIALS") = False Then Exit Sub

    Screen.MousePointer = 11
    rptMatMas.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMatMas.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptMatMas, CSMS_REPORT_PATH & "materials.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "MATERIALS") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    txtMatCde.Enabled = True
    'txtMatCde.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtMatCde.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "MATERIALS") = False Then Exit Sub

    On Error GoTo Errorcode
    If Not rsMatMas.BOF Or Not rsMatMas.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from CSMS_MatMas where id = " & labid.Caption
            LogAudit "X", "MATERIAL DATA ENTRY", txtMatCde
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "MATERIALS") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtMatCde.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
    'Picture3.Visible = False
    'Dim findStr As String
    'findStr = InputSpeechBox("Please Input Material Code or Description ...", txtMatCde.Text)
    'If findStr <> "" Then
    '   On Error Resume Next
    '   rsMatMas.Bookmark = rsFind(rsMatMas.Clone, "matcde", findStr).Bookmark
    '   If Err.Number = 3021 Then
    '      On Error GoTo ErrorCode
    '      rsMatMas.Bookmark = rsFind(rsMatMas.Clone, "matdsc", findStr).Bookmark
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

Private Sub cmdNext_Click()
    On Error Resume Next
    rsMatMas.MoveNext
    If rsMatMas.EOF Then
        rsMatMas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsMatMas.MovePrevious
    If rsMatMas.BOF Then
        rsMatMas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    If txtMatDsc.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtMatDsc.SetFocus
        Exit Sub
    End If
    If IsNull(txtMatCde.Text) = True Or txtMatCde.Text = "" Then
        MsgSpeechBox "Code must not be empty"
        On Error Resume Next
        txtMatCde.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                              As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select matcde from CSMS_MatMas where matcde = '" & txtMatCde.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Code already exist!"
                On Error Resume Next
                txtMatCde.SetFocus
                Exit Sub
            End If
        End If
    End If

    Dim VTXTMatCde, VTXTMatDsc                         As String
    Dim VTXTS_Price, VTXTCost                          As Double
    Dim VTXTPOCode                                     As String
    Dim VTXTOnHand, VTXTSStock                         As Long
    Dim VTXTLocation                                   As String

    VTXTMatCde = N2Str2Null(txtMatCde.Text)
    VTXTMatDsc = N2Str2Null(txtMatDsc.Text)
    VTXTS_Price = NumericVal(txtS_Price.Text)
    VTXTCost = NumericVal(txtCost.Text)
    VTXTPOCode = N2Str2Null(txtPOCode.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    VTXTSStock = NumericVal(txtSStock.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)

    If AddorEdit = "ADD" Then
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            rsMatMas.MoveLast
            labid.Caption = NumericVal(rsMatMas!ID) + 1
        End If
        gconDMIS.Execute "Insert into CSMS_MatMas" & _
                       " (matcde,matdsc,s_price,cost,pocode,onhand,sstock,location)" & _
                       " values (" & VTXTMatCde & ", " & VTXTMatDsc & ", " & VTXTS_Price & ", " & _
                       " " & VTXTCost & ", " & VTXTPOCode & ", " & VTXTOnHand & ", " & VTXTSStock & ", " & VTXTLocation & ")"
        ShowSuccessFullyAdded
        LogAudit "A", "MATERIAL DATA ENTRY", txtMatDsc
    Else
        gconDMIS.Execute "update CSMS_MatMas set" & _
                       " matcde = " & VTXTMatCde & "," & _
                       " matdsc = " & VTXTMatDsc & "," & _
                       " s_price = " & VTXTS_Price & "," & _
                       " cost = " & VTXTCost & "," & _
                       " onhand = " & VTXTOnHand & "," & _
                       " sstock = " & VTXTSStock & "," & _
                       " location = " & VTXTLocation & "," & _
                       " pocode = " & VTXTPOCode & _
                       " where id = " & labid.Caption
        LogAudit "E", "MATERIAL DATA ENTRY", txtMatDsc
        ShowSuccessFullyUpdated
    End If
    rsRefresh
    On Error Resume Next
    rsMatMas.Find "id =" & labid.Caption
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
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
    textSearch.Text = "":    'Picture3.ZOrder 0
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtMatCde.Text = ""
    txtMatDsc.Text = ""
    txtS_Price.Text = ""
    txtCost.Text = ""
    txtPOCode.Text = ""
    txtOnHand.Text = 0
    txtSStock.Text = 0
    txtLocation.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        labid.Caption = rsMatMas!ID
        txtMatCde.Text = Null2String(rsMatMas!MATCDE)
        txtMatDsc.Text = Null2String(rsMatMas!MatDsc)
        txtS_Price.Text = N2Str2Zero(rsMatMas!s_price)
        txtCost.Text = N2Str2Zero(rsMatMas!COST)
        txtPOCode.Text = Null2String(rsMatMas!ModelCode)
        txtOnHand.Text = N2Str2IntZero(rsMatMas!ONHAND)
        txtSStock.Text = N2Str2IntZero(rsMatMas!SSTOCK)
        txtLocation.Text = Null2String(rsMatMas!Location)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select * from CSMS_MatMas order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSMaterials = Nothing
    UnloadForm Me
End Sub


Private Sub lstMaterials_GotFocus()
    rsMatMas.Bookmark = rsFind(rsMatMas.Clone, "MATCDE", lstMaterials.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstMaterials_ItemClick(ByVal item As MSComctlLib.ListItem)
    rsMatMas.Bookmark = rsFind(rsMatMas.Clone, "MATCDE", lstMaterials.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstMaterials_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstMaterials
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

Private Sub lstMaterials_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstMaterials_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then: On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsMaterials                                    As ADODB.Recordset
    lstMaterials.Enabled = False
    lstMaterials.Sorted = False: lstMaterials.ListItems.Clear
    Set rsMaterials = New ADODB.Recordset
    Set rsMaterials = gconDMIS.Execute("select MatDsc,MATCDE from CSMS_MatMas order by MatDsc asc")
    If Not (rsMaterials.EOF And rsMaterials.BOF) Then
        lstMaterials.Enabled = True
        Listview_Loadval Me.lstMaterials.ListItems, rsMaterials
        lstMaterials.Refresh
        lstMaterials.Enabled = True
    Else
        lstMaterials.Enabled = False
    End If
    
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsMaterials                                    As ADODB.Recordset
    lstMaterials.Sorted = False: lstMaterials.ListItems.Clear
    lstMaterials.Enabled = False
    Set rsMaterials = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsMaterials = gconDMIS.Execute("select MatDsc, MATCDE from CSMS_MatMas where MatDsc like'" & xxx & "%'")
    If Not (rsMaterials.EOF And rsMaterials.BOF) Then
        lstMaterials.Enabled = True
        Listview_Loadval Me.lstMaterials.ListItems, rsMaterials
        lstMaterials.Refresh
        lstMaterials.Enabled = True
    Else
        lstMaterials.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstMaterials.ListItems.Count > 0 And lstMaterials.Enabled = True Then
            lstMaterials.SetFocus
        End If
    End If
End Sub


