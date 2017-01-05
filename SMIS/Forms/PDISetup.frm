VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmSMIS_Files_PDISetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDI CHECK LIST TO MODEL"
   ClientHeight    =   6240
   ClientLeft      =   315
   ClientTop       =   525
   ClientWidth     =   8625
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "PDISetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MouseIcon       =   "PDISetup.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "PDISetup.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Edit Record"
      Top             =   2550
      Width           =   435
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MouseIcon       =   "PDISetup.frx":0D78
      MousePointer    =   99  'Custom
      Picture         =   "PDISetup.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add Record"
      Top             =   2070
      Width           =   435
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Model And Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   90
      TabIndex        =   9
      Top             =   60
      Width           =   8505
      Begin VB.ComboBox cboPDI_Model 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   780
         TabIndex        =   11
         Top             =   360
         Width           =   2865
      End
      Begin VB.ComboBox cboPDI_Category 
         Height          =   345
         Left            =   5040
         TabIndex        =   10
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4020
         TabIndex        =   13
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "System That User Can Access"
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Added PDI Check List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   4560
      TabIndex        =   7
      Top             =   990
      Width           =   4035
      Begin MSComctlLib.ListView lvPDISelected 
         Height          =   4860
         Left            =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         MouseIcon       =   "PDISetup.frx":11DD
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MODULES"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of PDI Check List Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   90
      TabIndex        =   5
      Top             =   960
      Width           =   3975
      Begin MSComctlLib.ListView lvPDIList 
         Height          =   4860
         Left            =   30
         TabIndex        =   6
         Top             =   270
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         MouseIcon       =   "PDISetup.frx":133F
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MODULES"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton cmdRemovePDI 
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
      Height          =   495
      Left            =   4080
      MouseIcon       =   "PDISetup.frx":14A1
      MousePointer    =   99  'Custom
      Picture         =   "PDISetup.frx":15F3
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move from PDI Check List"
      Top             =   1590
      Width           =   435
   End
   Begin VB.CommandButton cmdAddPDI 
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
      Height          =   495
      Left            =   4080
      MouseIcon       =   "PDISetup.frx":17BD
      MousePointer    =   99  'Custom
      Picture         =   "PDISetup.frx":190F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Add to PDI Check List"
      Top             =   1110
      Width           =   435
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1440
      TabIndex        =   1
      Top             =   9360
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.Label labPDI_Category 
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   4110
      TabIndex        =   4
      Top             =   3180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labPDI_ModelCode 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4110
      TabIndex        =   0
      Top             =   3630
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmSMIS_Files_PDISetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Function GetModelCode(XXX As String) As String
    ''UDPATING CODE    :AXP-672007314
    Dim rsModelCode                                                   As ADODB.Recordset
    Set rsModelCode = gconDMIS.Execute("select CODE FROM ALL_ModelCode where description=" & N2Str2Null(XXX))
    If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
        GetModelCode = Null2String(rsModelCode!CODE)
    End If
    Set rsModelCode = Nothing
End Function

Private Function SETCATEGORY(ModelCode As String) As String
    'UDPATING CODE    :AXP-672007312
    ModelCode = UCase(RTrim(LTrim(ModelCode)))
    Select Case UCase(ModelCode)
        Case "VEHICLE EXTERIOR"
            SETCATEGORY = "VE"
        Case "VEHICLE INTERIOR"
            SETCATEGORY = "VI"
        Case "ENGINE COMPARTMENT"
            SETCATEGORY = "EC"
        Case "ELECTRICAL"
            SETCATEGORY = "EE"
        Case "TOOLS"
            SETCATEGORY = "TO"
    End Select
End Function

'Private Function SetModel(ModelCode As String) As String
'''UDPATING CODE    :AXP-672007314
'    Dim rsModelCode                    As ADODB.Recordset
'    Set rsModelCode = gconDMIS.Execute("select description FROM ALL_ModelCode where CODE=" & N2Str2Null(ModelCode))
'    If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
'        SetModel = Null2String(rsModelCode!Description)
'    End If
'    Set rsModelCode = Nothing
'End Function
'Private Function GetCategory(xxx As String) As String
'
'    xxx = UCase(RTrim(LTrim(xxx)))
'    Select Case xxx
'        Case "VE"
'            GetCategory = "VEHICLE EXTERIOR"
'        Case "VI"
'            GetCategory = "VEHICLE INTERIOR"
'        Case "EC"
'            GetCategory = "ENGINE COMPARTMENT"
'        Case "EE"
'            GetCategory = "ELECTRICAL"
'        Case "TO"
'            GetCategory = "TOOLS"
'    End Select
'End Function

Sub InitData()
    Dim TEMPRS                                                        As ADODB.Recordset

    AddColumnHeader "SN,DESCRIPTION,CATEGORY", lvPDIList
    ResizeColumnHeader lvPDIList, "10,70,15"

    AddColumnHeader "SN,DESCRIPTION,CATEGORY", lvPDISelected
    ResizeColumnHeader lvPDISelected, "10,70,15"

    Set TEMPRS = gconDMIS.Execute("SELECT Description FROM ALL_MODELCODE")
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        Combo_Loadval cboPDI_Model, TEMPRS
    End If
    With cboPDI_Category
        .AddItem "VEHICLE EXTERIOR"
        .AddItem "VEHICLE INTERIOR"
        .AddItem "ENGINE COMPARTMENT"
        .AddItem "ELECTRICAL"
        .AddItem "TOOLS"
        .AddItem "ALL"
    End With
End Sub

Private Sub AddPDI()
    On Error GoTo ErrorCode:

    If lvPDIList.SelectedItem Is Nothing Then
        MessagePop InfoVoid, "Selection Required", "There is Nothing To Select"
        Exit Sub
    End If
    Dim i                                                             As Integer
    Dim PDI_ID                                                        As Long
    Dim Category                                                      As String
    Dim ModelCode                                                     As String

    ' lvPDIList.Enabled = False

    If lvPDIList.ListItems.Count <= 0 Then: Exit Sub
    ModelCode = N2Str2Null(labPDI_ModelCode)
    For i = 1 To lvPDIList.ListItems.Count
        If lvPDIList.ListItems(i).Selected = True Then
            With lvPDISelected
                .ListItems.Add 1, , 1
                PDI_ID = lvPDIList.ListItems(i).ListSubItems(3).Text
                Category = lvPDIList.ListItems(i).ListSubItems(2).Text
                .ListItems(1).ListSubItems.Add , , lvPDIList.ListItems(i).ListSubItems(1).Text
                .ListItems(1).ListSubItems.Add , , Category
                .ListItems(1).ListSubItems.Add , , PDI_ID
                gconDMIS.Execute ("INSERT INTO SMIS_PDI_SETUP (PDI_ID,MODELCODE ,CATEGORY) VALUES ( " & PDI_ID & " ," & N2Str2Null(ModelCode) & "," & N2Str2Null(Category) & " )")
            End With
        End If
    Next
    For i = lvPDIList.ListItems.Count To 1 Step -1
        If lvPDIList.ListItems(i).Selected = True Then
            lvPDIList.ListItems.Remove (i)
        End If
    Next
    For i = 1 To lvPDISelected.ListItems.Count
        lvPDISelected.ListItems(i).Text = i
    Next
    cmdAddPDI.Enabled = False
    'lvPDIList.Enabled = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cboPDI_Category_Change()
    labPDI_Category = SETCATEGORY(cboPDI_Category)
    FillPDIList
    FillPDISetUpList
    cmdRemovePDI.Enabled = False
    cmdAddPDI.Enabled = False
End Sub

Private Sub cboPDI_Category_Click()
    cboPDI_Category_Change
End Sub

Private Sub cboPDI_Model_CLICK()
    cboPDI_Model_Change
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    frmSMIS_Files_PDICheckList.Show
    frmSMIS_Files_PDICheckList.cmdAdd.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdAddPDI_Click()
    AddPDI
End Sub

Private Sub cmdAddPDIList_Click()

End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If lvPDIList.SelectedItem Is Nothing Then: Exit Sub
    frmSMIS_Files_PDICheckList.Show
    frmSMIS_Files_PDICheckList.SearchID lvPDIList.SelectedItem.ListSubItems(3).Text
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdRemovePDI_Click()
    On Error GoTo ErrorCode:

    If lvPDISelected.SelectedItem Is Nothing Then
        MessagePop InfoVoid, "Selection Required", "Please Select From List"
        Exit Sub
    End If

    If RTrim(LTrim(labPDI_ModelCode)) = "" Then
        MessagePop InfoVoid, "Model Missing", "Please Select Model From The List"
        Exit Sub
    End If

    Dim PDI_ID                                                        As Long
    Dim ModelCode                                                     As String
    Dim i                                                             As Integer
    With lvPDISelected
        PDI_ID = lvPDISelected.SelectedItem.ListSubItems(3).Text
        ModelCode = N2Str2Null(labPDI_ModelCode)
        gconDMIS.Execute ("DELETE FROM SMIS_PDI_SETUP WHERE PDI_ID=" & PDI_ID & " AND MODELCODE=" & ModelCode)
        .ListItems.Remove .SelectedItem.Index
    End With
    For i = 1 To lvPDISelected.ListItems.Count
        lvPDISelected.ListItems(i).Text = i
    Next
    FillPDIList
    cmdRemovePDI.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub FillPDIList()
    'UDPATING CODE    :AXP-672007315
    On Error GoTo ErrorCode:

    If labPDI_ModelCode = "" Then: Exit Sub
    Dim MCODE                                                         As String
    Dim CCODE                                                         As String
    Dim SQL                                                           As String
    MCODE = N2Str2Null(labPDI_ModelCode)
    CCODE = N2Str2Null(labPDI_Category)

    lvPDIList.Enabled = False

    If labPDI_Category = "" Then
        SQL = " SELECT INSPECTIONNAME,CATEGORY,PDI_ID FROM SMIS_PDI_LIST WHERE PDI_ID NOT IN " & _
            " ( SELECT PDI_ID FROM SMIS_PDI_SETUP WHERE MODELCODE=" & MCODE & ") "
    Else
        SQL = " SELECT INSPECTIONNAME,CATEGORY,PDI_ID FROM SMIS_PDI_LIST WHERE PDI_ID NOT IN " & _
            " ( SELECT PDI_ID FROM SMIS_PDI_SETUP WHERE MODELCODE=" & MCODE & ") " & _
            " AND CATEGORY=" & CCODE
    End If

    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(SQL)

    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        'Listview_Loadval lvPDIList.ListItems, TEMPRS
        flex_FillListView TEMPRS, lvPDIList, True, False
        lvPDIList.Enabled = True
    End If





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub FillPDISetUpList()
    'UDPATING CODE    :AXP-672007316
    On Error GoTo ErrorCode:

    If labPDI_ModelCode = "" Then: Exit Sub
    Dim MCODE                                                         As String
    Dim CCODE                                                         As String
    Dim SQL                                                           As String
    lvPDISelected.Enabled = False
    MCODE = N2Str2Null(labPDI_ModelCode)
    CCODE = N2Str2Null(labPDI_Category)
    If labPDI_Category = "" Then
        SQL = " SELECT INSPECTIONNAME,CATEGORY,PDI_ID FROM SMIS_PDI_LIST WHERE PDI_ID IN " & _
            " ( SELECT PDI_ID FROM SMIS_PDI_SETUP WHERE MODELCODE=" & MCODE & ") "
    Else
        SQL = " SELECT INSPECTIONNAME,CATEGORY,PDI_ID FROM SMIS_PDI_LIST WHERE PDI_ID IN " & _
            " ( SELECT PDI_ID FROM SMIS_PDI_SETUP WHERE MODELCODE=" & MCODE & ") " & _
            " AND CATEGORY=" & CCODE
    End If
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute(SQL)

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        lvPDISelected.Enabled = True
    End If

    flex_FillListView TEMPRS, lvPDISelected, True, False





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
End Sub

Private Sub Label6_Click()

End Sub

Private Sub lvPDIList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvPDIList
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

Private Sub lvPDIList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdAddPDI.Enabled = True
    cmdRemovePDI.Enabled = False
End Sub

Private Sub lvPDISelected_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdRemovePDI.Enabled = True
    cmdAddPDI.Enabled = False
End Sub

Public Sub cboPDI_Model_Change()
    If LTrim(Trim(cboPDI_Model)) = "" Then: Exit Sub
    labPDI_ModelCode = GetModelCode(cboPDI_Model)
    FillPDIList
    FillPDISetUpList
    cmdRemovePDI.Enabled = False
    cmdAddPDI.Enabled = False
End Sub

