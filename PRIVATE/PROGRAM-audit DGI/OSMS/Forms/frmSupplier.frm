VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "frmSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5325
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   30
      TabIndex        =   11
      Top             =   2100
      Width           =   5115
      Begin VB.OptionButton optName 
         Caption         =   "Supplier &Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3090
         TabIndex        =   13
         Top             =   150
         Width           =   1575
      End
      Begin VB.OptionButton optCode 
         Caption         =   "Supplier &Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   540
         Width           =   4905
      End
      Begin MSComctlLib.ListView lstSupplier 
         Height          =   2175
         Left            =   60
         TabIndex        =   16
         Top             =   930
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   3836
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
         MouseIcon       =   "frmSupplier.frx":030A
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier Data Entry"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5145
      Begin VB.TextBox txtConPerson 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtSupplierCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtSupplierName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtSupplierAdd 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtSupplierTelNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   1710
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   1380
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   120
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   17
      Top             =   5340
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
         Left            =   4320
         MouseIcon       =   "frmSupplier.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   60
         Width           =   675
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
         Left            =   3660
         MouseIcon       =   "frmSupplier.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   60
         Width           =   675
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
         Left            =   3000
         MouseIcon       =   "frmSupplier.frx":0DA1
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   60
         Width           =   675
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
         Left            =   2340
         MouseIcon       =   "frmSupplier.frx":124F
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":13A1
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   675
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
         MouseIcon       =   "frmSupplier.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   675
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
         Left            =   1020
         MouseIcon       =   "frmSupplier.frx":1B00
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   60
         Width           =   675
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
         Left            =   360
         MouseIcon       =   "frmSupplier.frx":1FAA
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3720
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   25
      Top             =   5340
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
         Left            =   720
         MouseIcon       =   "frmSupplier.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
         Width           =   675
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
         MouseIcon       =   "frmSupplier.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "frmSupplier.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplier As ADODB.Recordset
Dim AddorEdit As String
Dim PrevSCODE As String

Private Sub cmdAdd_Click()
    Frame1.Enabled = True
    AddorEdit = "ADD"
    Frame1.Caption = "Add A Record"
    Picture1.Visible = False
    On Error Resume Next
    txtSupplierCode.SetFocus
    initMemvars
End Sub

Sub initMemvars()
    txtSupplierCode.Text = ""
    txtSupplierName.Text = ""
    txtSupplierAdd.Text = ""
    txtSupplierTelNo.Text = ""
    txtConPerson.Text = ""
End Sub

Private Sub cmdCancel_Click()
    Frame1.Caption = "Supplier Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Frame1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute ("delete from  OSMS_Supplier where Supplier_Code = '" & txtSupplierCode.Text & "'")
        rsRefresh
        StoreMemVars
        FillSearchGrid txtSearch
    End If
End Sub

Sub rsRefresh()
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select * from  OSMS_Supplier order by Supplier_Code asc", gconDMIS
End Sub
Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Frame1.Caption = "Edit A Record"
    Picture1.Visible = False
    PrevSCODE = txtSupplierCode.Text
    On Error Resume Next
    txtSupplierCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    txtSearch.SetFocus
End Sub

Function RecordFound(AAA As Variant) As Boolean
    If AAA <> "" Then
        Dim rsRecordFound As ADODB.Recordset
        Set rsRecordFound = New Recordset
        rsRecordFound.Open "Select Supplier_Name from  OSMS_Supplier order by Supplier_Code asc", gconDMIS
        rsRecordFound.Find "Supplier_Name like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsSupplier.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from  OSMS_Supplier order by Supplier_Code asc", gconDMIS
            rsRecordFound.Find "Supplier_Code = '" & AAA & "'"
            If Not rsRecordFound.EOF Then
                rsSupplier.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                RecordFound = False
            End If
        End If
    End If
End Function

Private Sub cmdNext_Click()
    On Error Resume Next
    rsSupplier.MoveNext
    If rsSupplier.EOF Then
        ShowLastRecordMsg
        rsSupplier.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsSupplier.MovePrevious
    If rsSupplier.BOF Then
        ShowFirstRecordMsg
        rsSupplier.MoveFirst
    End If
    StoreMemVars
End Sub

Sub StoreMemVars()
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        txtSupplierCode.Text = Null2String(rsSupplier!Supplier_code)
        txtSupplierName.Text = Null2String(rsSupplier!SUPPLIER_NAME)
        txtSupplierAdd.Text = Null2String(rsSupplier!Supplier_Address)
        txtSupplierTelNo.Text = Null2String(rsSupplier!Supplier_TelNo)
        txtConPerson.Text = rsSupplier!Contact_Person
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

'Upating Code       : AXP-0716200719:06
Private Sub cmdSave_Click()
    
On Error GoTo Errorcode:

    Screen.MousePointer = 11
    If txtSupplierCode.Text = "" Then
        MsgBoxXP "Supplier Code must not be empty!", "", XP_OKOnly, msg_Information
        On Error Resume Next
        txtSupplierCode.SetFocus
        Exit Sub
    End If

    If txtSupplierName.Text = "" Then
        MsgBoxXP "Supplier Name must not be empty!", "", XP_OKOnly, msg_Information
        On Error Resume Next
        txtSupplierName.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsSupplier.Find "Supplier_Code = '" & txtSupplierCode.Text & "'"
        If Not rsSupplier.EOF Then
            Screen.MousePointer = 0
            MsgBoxXP "Supplier Code already exists!", "", XP_OKOnly, msg_Information
            txtSupplierCode.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "Insert into OSMS_Supplier" & _
                         "(Supplier_Code, Supplier_Name, Supplier_Address, Supplier_TelNo, Contact_Person) values ('" & txtSupplierCode.Text & "','" & txtSupplierName.Text & "','" & txtSupplierAdd.Text & "','" & txtSupplierTelNo.Text & "', '" & txtConPerson.Text & "')"
    Else
        If PrevSCODE <> txtSupplierCode Then
            rsSupplier.Find "Supplier_Code = '" & txtSupplierCode.Text & "'"
            If Not rsSupplier.EOF Then
                Screen.MousePointer = 0
                MsgBoxXP "Supplier Code already exists!", "", XP_OKOnly, msg_Information
                txtSupplierCode.SetFocus
                Exit Sub
            End If
        End If

        gconDMIS.Execute "update OSMS_Supplier set " & _
                         "Supplier_Code = " & N2Str2Null(txtSupplierCode) & "," & _
                         "Supplier_Name = " & N2Str2Null(txtSupplierName) & "," & _
                         "Supplier_Address = " & N2Str2Null(txtSupplierAdd) & "," & _
                         "Supplier_TelNo = " & N2Str2Null(txtSupplierTelNo) & "," & _
                         "Contact_person = " & N2Str2Null(txtConPerson) & "" & _
                       " where Supplier_Code = " & N2Str2Null(PrevSCODE)
    End If
    rsRefresh
    cmdCancel.Value = True
    FillSearchGrid txtSearch
    Screen.MousePointer = 0
Exit Sub
Errorcode:
ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    rsRefresh
    txtSearch.Text = ""
    StoreMemVars
    Call AddColumnHeader("SupplierCode, SupplierName", lstSupplier)
    Call ResizeColumnHeader(lstSupplier, "28,65")
End Sub



Private Sub lstSupplier_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSupplier.Bookmark = rsFind(rsSupplier.Clone, "Supplier_Code", lstSupplier.SelectedItem.Text).Bookmark
    StoreMemVars
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

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsSupplier2 As ADODB.Recordset
    lstSupplier.Sorted = False: lstSupplier.ListItems.Clear
    lstSupplier.Enabled = False
    Set rsSupplier2 = New ADODB.Recordset
    If optCode.Value = True Then
        Set rsSupplier2 = gconDMIS.Execute("select Supplier_Code,Supplier_name from OSMS_Supplier where Supplier_Code like'" & XXX & "%' order by Supplier_Code asc")
    Else
        Set rsSupplier2 = gconDMIS.Execute("select Supplier_Code,Supplier_name from OSMS_Supplier where Supplier_name like'" & XXX & "%' order by Supplier_Code asc")
    End If
    If Not (rsSupplier2.EOF And rsSupplier2.BOF) Then
        Listview_Loadval Me.lstSupplier.ListItems, rsSupplier2
        lstSupplier.Refresh
          lstSupplier.Enabled = True
    End If
  
End Sub


Private Sub optCode_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub
Private Sub optName_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub



