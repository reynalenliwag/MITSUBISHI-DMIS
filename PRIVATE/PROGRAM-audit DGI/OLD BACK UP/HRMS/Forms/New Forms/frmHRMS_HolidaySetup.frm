VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMS_HolidaySetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Holiday Setup"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_HolidaySetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   7965
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2160
      ScaleHeight     =   855
      ScaleWidth      =   5730
      TabIndex        =   7
      Top             =   3420
      Width           =   5730
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
         Left            =   4950
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   4260
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   3570
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   2880
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   2190
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   1500
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   810
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   120
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4020
      Left            =   90
      ScaleHeight     =   3960
      ScaleWidth      =   1845
      TabIndex        =   6
      Top             =   90
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "frmHRMS_HolidaySetup.frx":2A31
         Top             =   0
         Width           =   9915
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   2010
      ScaleHeight     =   1995
      ScaleWidth      =   5865
      TabIndex        =   4
      Top             =   1350
      Width           =   5865
      Begin MSComctlLib.ListView lsvHol 
         Height          =   1965
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3466
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
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
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":1678E
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Month"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Day"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox picDed 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Left            =   2070
      ScaleHeight     =   1185
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      Begin VB.CheckBox Check1 
         Caption         =   "Special Holiday"
         Height          =   300
         Left            =   3000
         TabIndex        =   24
         Top             =   780
         Width           =   2265
      End
      Begin VB.TextBox txtday 
         Height          =   360
         Left            =   2100
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtMonth 
         Height          =   360
         Left            =   900
         TabIndex        =   20
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   1290
         TabIndex        =   1
         Top             =   300
         Width           =   4365
      End
      Begin Crystal.CrystalReport rptDeduction 
         Left            =   5400
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   -90
         TabIndex        =   23
         Top             =   0
         Width           =   5835
         _Version        =   655364
         _ExtentX        =   10292
         _ExtentY        =   503
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Height          =   240
         Index           =   2
         Left            =   1560
         TabIndex        =   22
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   810
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label lblID 
         Caption         =   "ID"
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   -60
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6450
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   16
      Top             =   3420
      Width           =   1440
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
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":168F0
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":16A42
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
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
         MouseIcon       =   "frmHRMS_HolidaySetup.frx":16D92
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_HolidaySetup.frx":16EE4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMS_HolidaySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsHOL                                                             As ADODB.Recordset
Dim ADD_EDIT                                                          As String

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Sub FillGrid()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * from HRMS_holiday_LIST order by description asc")
    lsvHol.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvHol.ListItems.Add(, , Null2String(RSTMP!Description))
            ITEM.SubItems(1) = Null2String(RSTMP!MANTH)
            ITEM.SubItems(2) = Null2String(RSTMP!DEYT)
            ITEM.SubItems(3) = Null2String(RSTMP!ID)

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub InitMemvars()
    txtDesc.Text = ""
    txtMonth.Text = ""
    txtday.Text = ""
End Sub

Sub rsrefresh()
    Set rsHOL = New ADODB.Recordset
    rsHOL.Open "select * from HRMS_HOLIDAY_LIST Order by description ASC", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    On Error Resume Next

    If Not (rsHOL.EOF And rsHOL.BOF) Then
        Check1.Value = N2Str2Zero(rsHOL!Type)
        lblID.Caption = rsHOL!ID
        txtDesc.Text = Null2String(rsHOL!Description)
        txtMonth.Text = Null2String(rsHOL!MANTH)
        txtday.Text = Null2String(rsHOL!DEYT)
    Else
        ShowNoRecord
        cmdAdd_Click
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_Add", "HOLIDAY SETUP") = False Then Exit Sub

    ADD_EDIT = "ADD"
    InitMemvars

    picDed.Enabled = True
    lsvHol.Enabled = False

    Picture1.Visible = False
    Picture2.Visible = True

    txtDesc.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picDed.Enabled = False
    lsvHol.Enabled = True

    Picture1.Visible = True
    Picture2.Visible = False
    lsvHol.Enabled = True

    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "HOLIDAY SETUP") = False Then Exit Sub
    If Not lsvHol.ListItems.count = 0 Then
        If MsgBox("Delete this Record", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("DELETE FROM HRMS_HOLIDAY_LIST WHERE ID = " & lblID.Caption & "")

            ShowDeletedMsg
            FillGrid
            rsrefresh
            StoreMemVars
        End If
    Else
        ShowNoRecord
    End If
End Sub

Private Sub cmdEdit_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_Edit", "HOLIDAY SETUP") = False Then Exit Sub
    ADD_EDIT = "EDIT"
    picDed.Enabled = True

    Picture1.Visible = False
    Picture2.Visible = True
    lsvHol.Enabled = False

    txtDesc.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    rsHOL.MoveNext
    If rsHOL.EOF Then
        rsHOL.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsHOL.MovePrevious
    If rsHOL.BOF Then
        rsHOL.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "HOLIDAY SETUP RECORD") = False Then Exit Sub

    Screen.MousePointer = 11
    rptDeduction.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptDeduction.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptDeduction.Formulas(3) = "PRINTedBY = '" & LOGNAME & "'"

    PrintSQLReport rptDeduction, HRMS_REPORT_PATH & "HolidaySetup.rpt", "", DMIS_REPORT_Connection, 1
    LogAudit "V", "HOLIDAY SETUP RECORD", ""
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim DEDMANTH                                                      As String
    Dim DEDDAY                                                        As String
    Dim DEDDESC                                                       As String
    If txtDesc.Text = "" Then
        ShowIsRequiredMsg "Holiday Description Cannot be Blank"
        txtDesc.SetFocus
        Exit Sub
    End If
    If txtMonth.Text = "" Then
        ShowIsRequiredMsg "MOnth Cannot be Blank"
        txtMonth.SetFocus
        Exit Sub
    End If
    If txtday.Text = "" Then
        ShowIsRequiredMsg "Day Cannot be Blank"
        txtday.SetFocus
        Exit Sub
    End If

    DEDDESC = N2Str2Null(txtDesc.Text)
    DEDDAY = N2Str2Null(txtday.Text)
    DEDMANTH = N2Str2Null(txtMonth.Text)

    If ADD_EDIT = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_HOLIDAY_LIST" & _
                         "(Description, MANTH, DEYT,TYPE) " & _
                       " values (" & DEDDESC & _
                         "," & DEDMANTH & _
                         "," & DEDDAY & _
                         "," & Check1.Value & ")"

        LogAudit "A", "HOLIDAY SETUP", txtDesc
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "Update HRMS_HOLIDAY_LIST set " & _
                       " Description = " & DEDDESC & _
                         ", Manth = " & DEDMANTH & _
                         ", Deyt = " & DEDDAY & _
                         ", TYPE = " & Check1.Value & _
                       " Where ID = " & lblID.Caption

        LogAudit "E", "HOLIDAY SETUP RECORD", txtDesc
        ShowSuccessFullyUpdated
    End If

    rsrefresh
    FillGrid
    cmdCancel_Click

    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    FillGrid
    rsrefresh
    StoreMemVars
End Sub

Private Sub lsvHol_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvHol
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lsvHol_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsHOL.MoveFirst
    rsHOL.Find "id=" & lsvHol.SelectedItem.ListSubItems(3)
    StoreMemVars
End Sub

Private Sub txtday_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtMonth_LostFocus()
    If txtMonth.Text = "" Then
        txtMonth.Text = "12"
        Exit Sub
    End If
    If Val(txtMonth) = 0 Then
        txtMonth = "1"
    End If
    If Val(txtMonth) > 12 Then
        txtMonth.Text = 12
    End If
End Sub

