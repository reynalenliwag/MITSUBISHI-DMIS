VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSJobCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Category"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6165
   ForeColor       =   &H8000000F&
   Icon            =   "FrmJobCategory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      TabIndex        =   1
      Top             =   60
      Width           =   6105
      Begin VB.TextBox txtFlatRate 
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
         Left            =   1380
         TabIndex        =   7
         Top             =   1020
         Width           =   1245
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
         Left            =   1380
         TabIndex        =   0
         Top             =   630
         Width           =   4605
      End
      Begin VB.TextBox txtCode 
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
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   1065
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
         Caption         =   "Flat Rate"
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
         Left            =   540
         TabIndex        =   8
         Top             =   1080
         Width           =   795
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
         Left            =   330
         TabIndex        =   5
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label9 
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
         Left            =   510
         TabIndex        =   4
         Top             =   270
         Width           =   825
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
         TabIndex        =   3
         Top             =   270
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2265
      Left            =   60
      TabIndex        =   6
      Top             =   1650
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   3995
      View            =   3
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Description"
         Object.Width           =   8378
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Flat Rate"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save!"
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
Attribute VB_Name = "frmCSMSJobCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaveEdit                                           As String

Sub CreateCode()
    Dim rsCreate                                       As ADODB.Recordset
    Set rsCreate = New ADODB.Recordset
    Set rsCreate = gconDMIS.Execute("Select jCat from CSMS_JobCategory Order by jCat desc")
    If Not rsCreate.EOF And Not rsCreate.BOF Then
        txtcode = Format(Val(rsCreate![jcat]) + 1, "000")
    Else
        txtcode = Format(1, "000")
    End If
End Sub

Sub ViewGrid()
    ListView1.Enabled = False
    ListView1.Sorted = False: ListView1.ListItems.Clear
    Dim rsCreate                                       As ADODB.Recordset
    Set rsCreate = New ADODB.Recordset
    Set rsCreate = gconDMIS.Execute("Select jCat,[Desc],flatrate from CSMS_JobCategory Order by [jCat] asc")
    If Not rsCreate.EOF And Not rsCreate.BOF Then
        Listview_Loadval Me.ListView1.ListItems, rsCreate
        ListView1.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    mnuRefresh_Click
    txtcode.Enabled = False
End Sub

Private Sub ListView1_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtcode = ListView1.SelectedItem
    txtDesc = ListView1.SelectedItem.SubItems(1)
    TXTFLATRATE = ListView1.SelectedItem.SubItems(2)
    mnuEdit.Enabled = True
    mnuDelete.Enabled = True
    mnuAdd.Enabled = False
End Sub

Private Sub mnuAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "JOBS") = False Then Exit Sub


    Frame1.Enabled = True: ListView1.Enabled = True
    mnuSave.Enabled = True: mnuAdd.Enabled = False
    SaveEdit = "Add"
    CreateCode
    On Error Resume Next
    txtDesc.SetFocus
End Sub

Private Sub mnuDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "JOBS") = False Then Exit Sub

    If MsgBox("DELETE this item..." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "delete  from CSMS_JobCategory Where jCat = '" & txtcode & "'"
    ViewGrid
End Sub

Private Sub mnuEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "JOBS") = False Then Exit Sub

    mnuEdit.Enabled = False
    mnuAdd.Enabled = False
    mnuSave.Enabled = True
    Frame1.Enabled = True
    mnuDelete.Enabled = False
    SaveEdit = "Edit"
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRefresh_Click()
    mnuSave.Enabled = False: mnuEdit.Enabled = False: mnuDelete.Enabled = False
    txtcode = "": txtDesc = "": TXTFLATRATE = ""
    Frame1.Enabled = False
    mnuAdd.Enabled = True
    SaveEdit = ""
    ViewGrid
End Sub

Private Sub mnuSave_Click()
    If MsgBox("Save entries..." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrorCode

    Dim xtxtCode, xtxtDesc                             As String
    Dim xtxtFlatRate                                   As Double
    xtxtCode = N2Str2Null(txtcode)
    xtxtDesc = N2Str2Null(UCase(txtDesc))
    xtxtFlatRate = NumericVal(TXTFLATRATE)
    If SaveEdit = "Add" Then
        gconDMIS.Execute "Insert into CSMS_JobCategory " & _
                       " (jCat,[Desc],flatrate)" & _
                       " values(" & xtxtCode & "," & xtxtDesc & "," & xtxtFlatRate & ")"
    Else
        gconDMIS.Execute ("update CSMS_JobCategory SET [desc]=" & xtxtDesc & ",flatrate = " & xtxtFlatRate & " where  jcat= " & xtxtCode & "")
    End If
    mnuRefresh_Click
    ViewGrid
    mnuAdd_Click
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub

End Sub

