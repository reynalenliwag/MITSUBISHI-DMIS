VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISAC_Counter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Counter Master File"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_Counter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6240
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   5790
      TabIndex        =   15
      Top             =   3900
      Width           =   5790
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
         MouseIcon       =   "AC_Counter.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit Window"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print this Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Add Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Find a Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move to Next Record"
         Top             =   0
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
         MouseIcon       =   "AC_Counter.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      FillColor       =   &H8000000D&
      Height          =   1275
      Left            =   60
      ScaleHeight     =   1275
      ScaleWidth      =   6105
      TabIndex        =   5
      Top             =   90
      Width           =   6105
      Begin VB.TextBox txtDescription 
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
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type module description"
         Top             =   450
         Width           =   4365
      End
      Begin VB.TextBox txtModule 
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Type module type (e.g.  CHG, CSH, RIV, DR, MRR, PO, RR, ADB)"
         Top             =   60
         Width           =   1065
      End
      Begin MSMask.MaskEdBox txtNextNumber 
         Height          =   345
         Left            =   1470
         TabIndex        =   2
         ToolTipText     =   "Type the next number of the particular master file (e.g. 1705, 205 depending on the last number of the masterfile)"
         Top             =   840
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
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
         Index           =   3
         Left            =   2550
         TabIndex        =   14
         Top             =   870
         Width           =   225
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
         Left            =   5850
         TabIndex        =   13
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   480
         Width           =   1365
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
         Left            =   2550
         TabIndex        =   8
         Top             =   120
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Number"
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
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Module Type"
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
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   90
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2505
      Left            =   60
      TabIndex        =   9
      Top             =   1320
      Width           =   6105
      Begin VB.TextBox textSearch 
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
         Left            =   75
         MaxLength       =   35
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   150
         Width           =   5925
      End
      Begin MSComctlLib.ListView lstCounter 
         Height          =   1875
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   5955
         _ExtentX        =   10504
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "AC_Counter.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MODULE TYPE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6526
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
      Left            =   4650
      ScaleHeight     =   885
      ScaleWidth      =   1620
      TabIndex        =   24
      Top             =   3870
      Width           =   1620
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
         MouseIcon       =   "AC_Counter.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "AC_Counter.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "AC_Counter.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   330
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   660
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMISAC_Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCunter                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub initMemvars()
    txtModule.Text = ""
    txtDescription.Text = ""
    txtNextNumber.Text = ""
End Sub

Sub StoreMemvars()
    If Not rsCunter.EOF And Not rsCunter.BOF Then
        labid.Caption = rsCunter!ID
        txtModule.Text = Null2String(rsCunter!modul)
        txtDescription.Text = Null2String(rsCunter!Description)
        txtNextNumber.Text = N2Str2IntZero(rsCunter!nextnumber)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsCunter = New ADODB.Recordset
    rsCunter.Open "select * from PMIS_Counter WHERE [TYPE] = 'A' order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsCounter                                                     As ADODB.Recordset
    lstCounter.Enabled = False
    lstCounter.Sorted = False: lstCounter.ListItems.Clear
    Set rsCounter = New ADODB.Recordset
    Set rsCounter = gconDMIS.Execute("select Modul,Description,ID from PMIS_Counter WHERE [TYPE] = 'A' order by Modul asc")
    If Not (rsCounter.EOF And rsCounter.BOF) Then
        lstCounter.Enabled = True
        Listview_Loadval Me.lstCounter.ListItems, rsCounter
        lstCounter.Refresh
    Else
        lstCounter.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCounter                                                     As ADODB.Recordset
    lstCounter.Enabled = False
    lstCounter.Sorted = False: lstCounter.ListItems.Clear

    Set rsCounter = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsCounter = gconDMIS.Execute("select Modul,Description, ID from PMIS_Counter where [TYPE] = 'A' AND Modul like'" & XXX & "%'")
    If Not (rsCounter.EOF And rsCounter.BOF) Then
        lstCounter.Enabled = True
        Listview_Loadval Me.lstCounter.ListItems, rsCounter
        lstCounter.Refresh
    Else
        lstCounter.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "ACCESSORIES COUNTER") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next

    txtModule.SetFocus
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "ACCESSORIES COUNTER") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtModule.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "ACCESSORIES COUNTER") = False Then Exit Sub
    On Error GoTo ERRORCODE
    If Not rsCunter.BOF Or Not rsCunter.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from PMIS_Counter where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            Call NEW_LogAudit("X", "ACCESSORIES COUNTER", SQL_STATEMENT, labid, "", "TYPE: " & txtModule, "", "")
            
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemvars
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsCunter.MoveNext
    If rsCunter.EOF Then
        rsCunter.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsCunter.MovePrevious
    If rsCunter.BOF Then
        rsCunter.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "ACCESSORIES COUNTER") = False Then Exit Sub
    'LogAudit "V", "ACCESSORIES COUNTER MASTER FILE", txtModule & "-" & txtDescription
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERRORCODE
    Dim rsfindDup                                                     As ADODB.Recordset
    Dim VTXTModule                                                    As String
    Dim vtxtDescription                                               As String
    Dim VTXTNextNumber                                                As Long

    If IsNull(txtModule.Text) = True Then
        ShowIsRequiredMsg "Code"
        On Error Resume Next
        txtModule.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select modul from PMIS_Counter where [TYPE] = 'A' AND Modul = '" & txtModule.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Code already exist!"
                On Error Resume Next
                txtModule.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtDescription.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtDescription.SetFocus
        Exit Sub
    End If
    If txtNextNumber.Text = "" Then
        ShowIsRequiredMsg "Next Number"
        On Error Resume Next
        txtNextNumber.SetFocus
        Exit Sub
    End If

    VTXTModule = N2Str2Null(txtModule.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VTXTNextNumber = NumericVal(txtNextNumber.Text)

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_Counter" & _
                       " ([TYPE],Modul,Description,NextNumber,LastUpdate,UserCode)" & _
                       " values ('A'," & VTXTModule & "," & vtxtDescription & ", " & VTXTNextNumber & ", " & "'" & LOGDATE & "'" & ", " & _
                       " " & "" & N2Str2Null(LOGCODE) & "" & ")"
        
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", "ACCESSORIES COUNTER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtModule), "MODUL", "PMIS_Counter", "DETAILS", N2Str2Null("A"), "TYPE"), "", "CODE: " & txtModule, "", "")
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update PMIS_Counter set" & _
                       " Modul = " & VTXTModule & "," & _
                       " Description = " & vtxtDescription & "," & _
                       " NextNumber = " & VTXTNextNumber & "," & _
                       " LastUpdate = " & "'" & LOGDATE & "'" & "," & _
                       " UserCode = " & "" & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labid.Caption
                       
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", "ACCESSORIES COUNTER", SQL_STATEMENT, labid, "", "CODE: " & txtModule, "", "")
        ShowSuccessFullyUpdated
    End If
    
    rsRefresh
    On Error Resume Next
    rsCunter.Find "Modul = " & VTXTModule
    cmdCancel.Value = True
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = True Then
                Unload frmALL_AuditInquiry
                 
                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCESSORIES COUNTER)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "ACCESSORIES COUNTER")
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = ""
    initMemvars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISMaster_Counter = Nothing
    UnloadForm Me
End Sub

Private Sub lstCounter_GotFocus()
    rsCunter.Bookmark = rsFind(rsCunter.Clone, "ID", lstCounter.SelectedItem.SubItems(2)).Bookmark
    StoreMemvars
End Sub

Private Sub lstCounter_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsCunter.Bookmark = rsFind(rsCunter.Clone, "ID", lstCounter.SelectedItem.SubItems(2)).Bookmark
    StoreMemvars
End Sub

Private Sub lstCounter_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCounter
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

Private Sub lstCounter_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCounter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstCounter.ListItems.Count > 0 And lstCounter.Enabled = True Then: lstCounter.SetFocus
    End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub



