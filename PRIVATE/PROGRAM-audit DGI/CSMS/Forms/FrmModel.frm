VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6165
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmEntry 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   6105
      Begin VB.TextBox txtMake 
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
         Left            =   1260
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1020
         Width           =   4635
      End
      Begin VB.TextBox txtCode 
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
         Left            =   1260
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   1260
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   630
         Width           =   4635
      End
      Begin Crystal.CrystalReport rptROJOBS 
         Left            =   5460
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Jobs Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
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
         Height          =   225
         Left            =   1410
         TabIndex        =   5
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         TabIndex        =   3
         Top             =   690
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2025
      Left            =   0
      TabIndex        =   6
      Top             =   1650
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3572
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
      MouseIcon       =   "FrmModel.frx":0000
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Make"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "&Refresh"
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "FrmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveEdit             As String
Private Sub Form_Load()
    mnuRefresh_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCode = ListView1.SelectedItem
    txtDesc = ListView1.SelectedItem.SubItems(1)
    txtMake = ListView1.SelectedItem.SubItems(2)
    mnuEdit.Enabled = True
    mnuDelete.Enabled = True
    mnuAdd.Enabled = False
    mnuSave.Enabled = False
End Sub

Private Sub mnuAdd_Click()
    FrmEntry.Enabled = True: ListView1.Enabled = True
    mnuSave.Enabled = True: mnuAdd.Enabled = False
    SaveEdit = "Add"
    CreateCode
    txtDesc.SetFocus
End Sub
Sub CreateCode()
    Dim rsCreate         As ADODB.Recordset
    Set rsCreate = New ADODB.Recordset
    Set rsCreate = gconDMIS.Execute("Select jModel from CSMS_JobModel Order by jModel desc")
    If Not rsCreate.EOF And Not rsCreate.BOF Then
        txtCode = Format(Val(rsCreate![jmodel]) + 1, "000")
    Else
        txtCode = Format(1, "000")
    End If
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("DELETE this item..." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from CSMS_JobModel Where jModel = '" & txtCode & "'"
    ViewGrid
End Sub

Private Sub mnuEdit_Click()
    mnuEdit.Enabled = False
    mnuAdd.Enabled = False
    mnuSave.Enabled = True
    FrmEntry.Enabled = True
    mnuDelete.Enabled = False
    SaveEdit = "Edit"
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub
Private Sub mnuRefresh_Click()
    FrmEntry.Enabled = False: mnuAdd.Enabled = True
    mnuSave.Enabled = False: mnuEdit.Enabled = False: mnuDelete.Enabled = False
    txtCode = "": txtDesc = "": txtMake = ""
    ViewGrid
End Sub

Private Sub mnuSave_Click()
    If MsgBox("Save entries..." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If
    Dim xtxtCode, xtxtDesc As String
    xtxtCode = N2Str2Null(txtCode)
    xtxtDesc = N2Str2Null(UCase(txtDesc))
    txtMake = N2Str2Null(txtMake)
    If SaveEdit = "Add" Then
        gconDMIS.Execute "Insert into CSMS_JobModel " & _
                       " (jmodel,[Desc],[Make])" & _
                       " values(" & xtxtCode & "," & xtxtDesc & "," & txtMake & ")"
    Else
        gconDMIS.Execute ("UPDATE CSMS_JobModel SET [desc]=" & xtxtDesc & ",[Make]=" & txtMake & " where  jmodel= " & xtxtCode & "")
    End If
    mnuRefresh_Click
    mnuAdd_Click
    ViewGrid
End Sub
Sub ViewGrid()
    ListView1.Sorted = False: ListView1.ListItems.Clear
    Dim rsCreate         As ADODB.Recordset
    Set rsCreate = New ADODB.Recordset
    Set rsCreate = gconDMIS.Execute("Select jModel,[Desc],[Make] from CSMS_JobModel Order by [jModel] asc")
    If Not rsCreate.EOF And Not rsCreate.BOF Then
        Listview_Loadval Me.ListView1.ListItems, rsCreate
    End If
End Sub
