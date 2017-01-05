VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMS_Shift_Management 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Management"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Shift_Management.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   8775
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Text            =   " "
      Top             =   300
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   9015
      TabIndex        =   2
      Top             =   780
      Width           =   9015
      Begin MSComctlLib.ListView lstListofShiftEmp 
         Height          =   3615
         Left            =   4350
         TabIndex        =   6
         Top             =   420
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Shift_Management.frx":058A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "EMPNO"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstListofEmp 
         Height          =   3615
         Left            =   30
         TabIndex        =   5
         Top             =   420
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Shift_Management.frx":06EC
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "EMPNO"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.CommandButton cmdAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         Picture         =   "Shift_Management.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1230
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdRemove 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         Picture         =   "Shift_Management.frx":0C65
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1710
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Double Click Item to Exclude in Selected Shift"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4350
         TabIndex        =   11
         Top             =   4140
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Double Click Item to Include in Selected Shift"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   4140
         Width           =   3975
      End
      Begin VB.Label labTotalEmployee 
         Caption         =   "List of Employees"
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
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   3375
      End
      Begin VB.Label labListEmpinShift 
         Caption         =   "List of Employees"
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
         Left            =   4260
         TabIndex        =   7
         Top             =   0
         Width           =   3315
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Employee :"
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
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   5370
      Width           =   2145
   End
   Begin VB.Label labid 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2310
      TabIndex        =   12
      Top             =   5370
      Width           =   3255
   End
   Begin VB.Label labShiftDetails 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   360
      Width           =   3555
   End
   Begin VB.Label Label4 
      Caption         =   "Select Shift From The List"
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
      TabIndex        =   1
      Top             =   60
      Width           =   7185
   End
End
Attribute VB_Name = "frmHRMS_Shift_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsShift                                                           As ADODB.Recordset

Sub FillCombo()
    Set rsShift = New ADODB.Recordset
    rsShift.Open "Select * from HRMS_TIme_Shift_Code", gconDMIS, adOpenKeyset, adLockReadOnly
    While Not rsShift.EOF
        Combo1.AddItem Null2String(rsShift!shiftcode)
        rsShift.MoveNext
    Wend
End Sub

Sub rsrefresh1()
    lstListofEmp.ListItems.Clear
    Dim rsEmp1                                                        As ADODB.Recordset
    Set rsEmp1 = New ADODB.Recordset
    Set rsEmp1 = gconDMIS.Execute("Select upper(lastname) +', '+ upper(firstname), empno, shift from HRMS_EmpInfo where  activeinactive ='a' and isnull(shift,'')  <> '" & Repleys(Combo1.Text) & "' order by lastname")


    Listview_Loadval lstListofEmp.ListItems, rsEmp1
    lstListofEmp.Refresh

End Sub

Sub rsrefresh2()
    lstListofShiftEmp.ListItems.Clear
    Dim rsEmp2                                                        As ADODB.Recordset
    Set rsEmp2 = New ADODB.Recordset
    Set rsEmp2 = gconDMIS.Execute("Select upper(lastname) +', ' + upper(firstname), empno, shift, activeinactive from HRMS_EmpInfo where activeinactive = 'A' and SHIFT = " & N2Str2Null(Repleys(Combo1.Text)) & " ORder by lastname")


    Listview_Loadval lstListofShiftEmp.ListItems, rsEmp2
    lstListofShiftEmp.Refresh
End Sub

Sub DeleteShift()

    vcombo1 = N2Str2Null(Combo1.Text)
    If labid <> "" And Combo1.Text <> "" Then
        gconDMIS.Execute "Update HRMS_EmpInfo set " & _
                       " shift = '" & "'" & _
                       " where Empno = '" & labid & "'"
    End If
End Sub

Private Sub Combo1_Change()
    Combo1_Click
End Sub

Private Sub Combo1_Click()
    rsShift.MoveFirst
    rsShift.Find ("SHIFTCODE='" & Repleys(Combo1) & "'")
    If Not rsShift.EOF Or Not rsShift.BOF Then
        If IsDate(rsShift!FROM1) = True And IsDate(rsShift!TO1) = True Then
            labShiftDetails = TimeValue(rsShift!FROM1) & "-" & TimeValue(rsShift!TO1)
            labTotalEmployee = "List of Employees not in " & Combo1
            labListEmpinShift = "List of Employees Under " & Combo1
        End If
    End If
    rsrefresh1
    rsrefresh2
End Sub

Private Sub cmdAdd_Click()
    Dim vcombo1                                                       As String
    vcombo1 = N2Str2Null(Combo1.Text)
    If labid <> "" And Combo1.Text <> "" Then
        gconDMIS.Execute "Update HRMS_EmpInfo set " & _
                       " shift = " & vcombo1 & _
                       " where Empno = '" & labid & "'"
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillCombo
    rsrefresh1
    rsrefresh2
    Screen.MousePointer = 0
End Sub

Private Sub lstListofEmp_DblClick()
    cmdAdd_Click
    rsrefresh1
    rsrefresh2

    '    lstListofShiftEmp.ListItems.Add , , lstListofEmp.SelectedItem
    '    lstListofShiftEmp.ListItems(lstListofShiftEmp.ListItems.Count).ListSubItems.Add , , lstListofEmp.SelectedItem.ListSubItems(1)
    '    lstListofEmp.ListItems.Remove (lstListofEmp.SelectedItem.INDEX)
End Sub

Private Sub lstListofEmp_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    labid = lstListofEmp.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub lstListofShiftEmp_DblClick()
    DeleteShift
    rsrefresh1
    rsrefresh2
    '    lstListofEmp.ListItems.Add , , lstListofShiftEmp.SelectedItem
    '    lstListofEmp.ListItems(lstListofEmp.ListItems.Count).ListSubItems.Add , , lstListofShiftEmp.SelectedItem.ListSubItems(1)
    '    lstListofShiftEmp.ListItems.Remove (lstListofShiftEmp.SelectedItem.INDEX)
End Sub

Private Sub lstListofShiftEmp_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    cmdRemove.Enabled = True
    labid = lstListofShiftEmp.SelectedItem.ListSubItems(1).Text
End Sub

