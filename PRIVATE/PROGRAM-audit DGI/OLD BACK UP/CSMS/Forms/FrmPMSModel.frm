VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSPMSModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PMS Model"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   ForeColor       =   &H8000000F&
   Icon            =   "FrmPMSModel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   60
      TabIndex        =   4
      Top             =   -60
      Width           =   5085
      Begin VB.TextBox txtFlatrate 
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   870
         Width           =   1125
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   990
         TabIndex        =   1
         Top             =   540
         Width           =   4005
      End
      Begin VB.TextBox txtModel 
         Height          =   285
         Left            =   990
         MaxLength       =   20
         TabIndex        =   0
         Top             =   210
         Width           =   4005
      End
      Begin VB.Label labID 
         Height          =   315
         Left            =   2190
         TabIndex        =   8
         Top             =   870
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Flat Rate"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   570
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   435
      End
   End
   Begin MSComctlLib.ListView lstCons 
      Height          =   1785
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3149
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmPMSModel.frx":014A
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Flat Rate"
         Object.Width           =   2293
      EndProperty
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
      Left            =   4440
      MouseIcon       =   "FrmPMSModel.frx":02AC
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":03FE
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exit Window"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
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
      MouseIcon       =   "FrmPMSModel.frx":0764
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":08B6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Refresh"
      Top             =   3240
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
      Left            =   3000
      MouseIcon       =   "FrmPMSModel.frx":0BE1
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":0D33
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Save PMS Model"
      Top             =   3240
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
      Left            =   2280
      MouseIcon       =   "FrmPMSModel.frx":1083
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":11D5
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Delete PMS Model"
      Top             =   3240
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
      Left            =   1560
      MouseIcon       =   "FrmPMSModel.frx":1500
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":1652
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Edit PMS Model"
      Top             =   3240
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
      Left            =   840
      MouseIcon       =   "FrmPMSModel.frx":19AE
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMSModel.frx":1B00
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add PMS Model"
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSPMSModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ShowModel()
    Dim RSMODEL                                        As ADODB.Recordset
    Set RSMODEL = New ADODB.Recordset
    lstCons.Enabled = False
    Set RSMODEL = gconDMIS.Execute("select ID,Model,Description,FlatAmt from CSMS_PMS_Hd order by Model asc")
    lstCons.Sorted = False: lstCons.ListItems.Clear
    If Not (RSMODEL.EOF And RSMODEL.BOF) Then
        Listview_Loadval Me.lstCons.ListItems, RSMODEL
        lstCons.Enabled = True
        lstCons.Refresh
        lstCons.Enabled = True
    Else
        lstCons.Enabled = False
    End If

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "MODEL") = False Then Exit Sub

    Frame1.Enabled = True
    cmdAdd.Enabled = False
    cmdSave.Enabled = True
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "MODEL") = False Then Exit Sub

    Dim rsGetmodel                                     As ADODB.Recordset
    Set rsGetmodel = New ADODB.Recordset
    Set rsGetmodel = gconDMIS.Execute("select PSM_Description from CSMS_Psm_Det")
    If Not (rsGetmodel.EOF And rsGetmodel.BOF) Then
        MsgBox "Cannot be delete! (with existing sub data)..."
        Exit Sub
    End If
    If MsgBox("Delete this item..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete * from [CSMS_PMS_Hd] where id = " & Val(labid.Caption) & ""
    LogAudit "X", "PMS MODEL", txtModel
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "MODEL") = False Then Exit Sub

    Frame1.Enabled = True
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefres_Click()
    txtModel = "": txtDesc = "": TXTFLATRATE = ""
    cmdAdd.Enabled = True: cmdEdit.Enabled = False
    cmdDelete.Enabled = False: cmdSave.Enabled = False
    Frame1.Enabled = False: labid.Caption = ""
    ShowModel
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    If NumericVal(TXTFLATRATE) <= 0 Then
        MsgBox "No Flat Rate..."
        Exit Sub
    End If
    If txtModel = "" Then
        MsgBox "Model name please..."
        Exit Sub
    End If
    If MsgBox("Save all entries..." & vbCrLf & vbCrLf & "Are you sure ?", vbQuestion + vbYesNo, "Message Box") = vbNo Then
        Exit Sub
    End If
    Dim xMode, xDescription                            As String
    Dim xFlatAmt                                       As Double
    xMode = N2Str2Null(txtModel)
    xDescription = N2Str2Null(txtDesc)
    xFlatAmt = NumericVal(TXTFLATRATE)
    gconDMIS.Execute "delete from [CSMS_PMS_Hd] where id = " & Val(labid.Caption) & ""
    gconDMIS.Execute "Insert into CSMS_PMS_Hd " & _
                   " (Model,Description,FlatAmt)" & _
                   " values(" & xMode & "," & xDescription & "," & xFlatAmt & ")"
    LogAudit "A", "PMS MODEL", txtModel
    cmdRefres_Click
    cmdAdd.Value = True
    Exit Sub

ErrorCode:

    ShowVBError
    Exit Sub

End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cmdRefres_Click
    ShowModel
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub lstCons_DblClick()
    labid.Caption = lstCons.SelectedItem
    txtModel = lstCons.SelectedItem.SubItems(1)
    txtDesc = lstCons.SelectedItem.SubItems(2)
    TXTFLATRATE = lstCons.SelectedItem.SubItems(3)
    cmdAdd.Enabled = False: cmdEdit.Enabled = True
    cmdDelete.Enabled = True: cmdSave.Enabled = False
    Frame1.Enabled = False
End Sub

