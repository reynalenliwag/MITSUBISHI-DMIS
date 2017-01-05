VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Files_FreeBeeies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Free Beeies"
   ClientHeight    =   5370
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FreeBeeies.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5835
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   11
      Top             =   4410
      Width           =   6075
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
         Left            =   5010
         MouseIcon       =   "FreeBeeies.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   4320
         MouseIcon       =   "FreeBeeies.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
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
         Left            =   3630
         MouseIcon       =   "FreeBeeies.frx":11FF
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
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
         Left            =   2940
         MouseIcon       =   "FreeBeeies.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Add Record"
         Top             =   60
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
         Left            =   2250
         MouseIcon       =   "FreeBeeies.frx":1B12
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Find a Record"
         Top             =   60
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
         Left            =   1560
         MouseIcon       =   "FreeBeeies.frx":1F5E
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":20B0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move to Next Record"
         Top             =   60
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
         Left            =   870
         MouseIcon       =   "FreeBeeies.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":255A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   1
      Top             =   -60
      Width           =   5715
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "FreeBeeies.frx":28B9
         Left            =   1200
         List            =   "FreeBeeies.frx":28C6
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   210
         Width           =   2655
      End
      Begin VB.TextBox txtColor_desc 
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
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   4440
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
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
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Width           =   1425
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3375
      Left            =   30
      TabIndex        =   5
      Top             =   990
      Width           =   5715
      Begin VB.OptionButton optCode 
         Caption         =   "&Type"
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
         Left            =   3030
         TabIndex        =   10
         Top             =   180
         Width           =   1245
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Description"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
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
         Height          =   345
         Left            =   90
         MaxLength       =   35
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lstColor 
         Height          =   2325
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4101
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
         MouseIcon       =   "FreeBeeies.frx":28FE
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label3 
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
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4260
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   19
      Top             =   4455
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
         Left            =   750
         MouseIcon       =   "FreeBeeies.frx":2A60
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":2BB2
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   60
         MouseIcon       =   "FreeBeeies.frx":2EF0
         MousePointer    =   99  'Custom
         Picture         =   "FreeBeeies.frx":3042
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   8160
      TabIndex        =   4
      Top             =   570
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmSMIS_Files_FreeBeeies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :cmdSave_Click:Update for Double Entry Code For Add and Edit Both
'DATE STARTED       :5/11/200715:02
'LAST UPDATED       :5/11/200715:02
'DATABASE UPDATES   :NONE
'WHO UPDATED        :AXP5/11/200715:02
'==========================================================================================


Option Explicit
Dim rsColor                             As ADODB.Recordset
Dim AddorEdit                           As String

'Upating Code       : AXP-0707200712:19
Private Sub cmdADD_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "FREE BEEIES") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    initMemvars
    lstColor.Enabled = False
    txtSearch.Enabled = False
    optDesc.Enabled = False
    optCode.Enabled = False
    On Error Resume Next
    'txtColor_code.SetFocus
    Combo1.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lstColor.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True

    optDesc.Enabled = True
    optCode.Enabled = True

    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "FREE BEEIES") = False Then Exit Sub
    On Error GoTo ErrorCode
    '''AXPREDH
    If Not rsColor.BOF Or Not rsColor.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from smis_vacc where id = " & LABID.Caption
            ShowDeletedMsg
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

'Upating Code       : AXP-0707200712:19
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_EDIT", "FREE BEEIES") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraDetails.Enabled = False




    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:19
Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsColor.MoveNext
    If rsColor.EOF Then
        rsColor.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsColor.MovePrevious
    If rsColor.BOF Then
        rsColor.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()

End Sub

'Upating Code       : AXP-0707200712:20
Private Sub cmdSave_Click()
    Dim lng                             As Integer
    On Error GoTo ErrorCode:
'
'    If txtColor_code.Text = "" Or txtColor_desc.Text = "" Then
'        ShowIsRequiredMsg "Color Code and Description"
'        On Error Resume Next
'        txtColor_code.SetFocus
'        Exit Sub
'    End If
'    '''''''AXP5/11/200715:02
'    lng = gconDMIS.Execute("select Count(*) from ALL_Color WHERE color_code=" & N2Str2Null(txtColor_code)).Fields(0).Value
'    If AddorEdit = "ADD" Then
'        If lng >= 1 Then
'            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
'            Exit Sub
'        End If
'    Else
'        If lng >= 1 And UCase(Null2String(rsColor!Color_code)) <> UCase(txtColor_code) Then
'            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
'            Exit Sub
'        End If
'    End If
'
'    If AddorEdit = "ADD" Then
'        gconDMIS.Execute "Insert into ALL_Color" & _
'                       " (color_code,color_desc)" & _
'                       " values (" & N2Str2Null(txtColor_code.Text) & ", " & N2Str2Null(txtColor_desc.Text) & ")"
'    Else
'        gconDMIS.Execute "update ALL_Color set" & _
'                       " color_code = " & N2Str2Null(txtColor_code.Text) & "," & _
'                       " color_desc = " & N2Str2Null(txtColor_desc.Text) & _
'                       " where id = " & labid.Caption
'    End If
'
'    rsRefresh
'    If AddorEdit = "EDIT" Then
'        rsColor.Find ("ID=" & labid)
'    End If
'    cmdCancel.Value = True
'    FillSearchGrid ""
'




    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsColor2                        As ADODB.Recordset
    lstColor.Sorted = False
    lstColor.ListItems.Clear
    lstColor.Enabled = False
    Set rsColor2 = New ADODB.Recordset

    
        Set rsColor2 = gconDMIS.Execute("select  ACCESSORIESNAME,TYPE,ID  from smis_vacc where ACCESSORIESNAME like'" & ReplaceQuote(xxx) & "%' order by color_desc asc")
    

    If Not (rsColor2.EOF And rsColor2.BOF) Then
        Listview_Loadval Me.lstColor.ListItems, rsColor2
        lstColor.Refresh
        lstColor.Enabled = True
    End If

End Sub

Private Sub Combo1_Change()
GetFreeBeeies Combo1.Text
End Sub

Private Sub Combo1_Click()
GetFreeBeeies Combo1.Text
End Sub
Function GetFreeBeeies(xxx)
'

    If xxx = "STANDARD FREEBEEIES" Then
        GetFreeBeeies = "ST"
    ElseIf xxx = "ADDITIONAL FREEBEEIES" Then
        GetFreeBeeies = "AF"
    Else
        GetFreeBeeies = "OT"
    End If

End Function


Function SetFreeBeeies(xxx)
'
    
    If xxx = "ST" Then
        SetFreeBeeies = "STANDARD FREEBEEIES"
    ElseIf xxx = "AF" Then
        SetFreeBeeies = "ADDITIONAL FREEBEEIES"
    Else
        SetFreeBeeies = "OTHERS"
    End If

End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picAdds.Visible = True And KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    


    txtSearch.Text = vbNullString
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    
End Sub

Private Sub lstColor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstColor
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

Private Sub lstColor_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstColor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsColor.MoveFirst
    rsColor.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub optCode_Click()
    If txtSearch = "" Then FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optDesc_Click()
    If txtSearch = "" Then FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Sub rsRefresh()
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select * from ALL_Color order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsColor.EOF And Not rsColor.BOF Then
        LABID.Caption = rsColor!ID
        Combo1.Text = SetFreeBeeies(Null2String(rsColor!Color_code))
        txtColor_desc.Text = Null2String(rsColor!COLOR_DESC)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub txtColor_code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

