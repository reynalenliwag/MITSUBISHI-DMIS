VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Time_Shift_Management 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Shift Management"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmHRMS_Time_Shift_Management.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1650
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   18
      Top             =   4290
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   4170
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   3480
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   2790
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   2100
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   1410
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   720
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   30
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5760
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   27
      Top             =   4320
      Width           =   1440
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
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   29
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
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Time_Shift_Management.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1695
      Left            =   3690
      ScaleHeight     =   1635
      ScaleWidth      =   2415
      TabIndex        =   12
      Top             =   2280
      Width           =   2475
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   630
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   1170
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   630
         TabIndex        =   13
         Top             =   390
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48824321
         CurrentDate     =   39463
      End
      Begin VB.Label Label6 
         Caption         =   "EFFECTIVITY DATE"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Shift"
         Height          =   345
         Left            =   150
         TabIndex        =   15
         Top             =   1230
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   345
         Left            =   150
         TabIndex        =   14
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4125
      Left            =   60
      Picture         =   "frmHRMS_Time_Shift_Management.frx":36A3
      ScaleHeight     =   4095
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   120
      Width           =   2385
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   30
         MaxLength       =   35
         TabIndex        =   6
         Top             =   60
         Width           =   2295
      End
      Begin MSComctlLib.ListView lsTimeShift 
         Height          =   3645
         Left            =   30
         TabIndex        =   7
         Top             =   450
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   6429
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
         MouseIcon       =   "frmHRMS_Time_Shift_Management.frx":63DF
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SHIFT CODE"
            Object.Width           =   3528
         EndProperty
         Picture         =   "frmHRMS_Time_Shift_Management.frx":6541
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1485
      Left            =   2550
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.TextBox txtNo 
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
         Left            =   1500
         TabIndex        =   10
         Top             =   210
         Width           =   3075
      End
      Begin VB.TextBox txtPosition 
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
         Left            =   1470
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtName 
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
         Left            =   1500
         TabIndex        =   8
         Top             =   630
         Width           =   3075
      End
      Begin VB.Label Label1 
         Caption         =   "Employee No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   270
         TabIndex        =   11
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   2
         Top             =   1170
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   270
         TabIndex        =   1
         Top             =   690
         Width           =   1305
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdShift 
      Height          =   2625
      Left            =   2610
      TabIndex        =   3
      Top             =   1650
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4630
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LABID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "frmHRMS_Time_Shift_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMPINFO                                                         As ADODB.Recordset
Dim rsShiftCode                                                       As ADODB.Recordset
Dim rsTimeShiftMgt                                                    As ADODB.Recordset

Sub rsRefresh()
    Set RSEMPINFO = New ADODB.Recordset
    Set RSEMPINFO = gconDMIS.Execute("Select EmpNo, LastName, FirstName, [Position] from HRMS_EmpInfo order by lastname")
    rsrefresh2 txtNo
End Sub

Sub rsrefresh2(XXX As String)

    Set rsTimeShiftMgt = New ADODB.Recordset
    Set rsTimeShiftMgt = gconDMIS.Execute("Select * from HRMS_Time_Shift_Management where EmplNo like '" & XXX & "'")
End Sub

Sub FillGrid()
    grdShift.Rows = 1
    If Not rsTimeShiftMgt.BOF And Not rsTimeShiftMgt.EOF Then
        rsTimeShiftMgt.MoveFirst
        Do While Not rsTimeShiftMgt.EOF
            grdShift.AddItem Null2String(rsTimeShiftMgt!EMPLNO) & Chr(9) & _
                             Null2String(rsTimeShiftMgt!Shift) & Chr(9) & _
                             Null2String(rsTimeShiftMgt!DateFrom) & Chr(9) & _
                             Null2String(rsTimeShiftMgt!DateTo) & Chr(9) & _
                             Null2String(rsTimeShiftMgt!ID)
            rsTimeShiftMgt.MoveNext
        Loop
    End If
End Sub

Sub FillListview()
    Listview_Loadval lsTimeShift.ListItems, gconDMIS.Execute("select EmpNo from HRMS_EmpInfo order by lastname ")
End Sub

Sub storeMemvars()
    If RSEMPINFO.RecordCount = 0 Then
        txtNo = ""
        txtName = ""
        txtPosition = ""
    End If
    If Not RSEMPINFO.BOF And Not RSEMPINFO.EOF Then
        txtNo = Null2String(RSEMPINFO!EMPNO)
        txtName = Null2String(RSEMPINFO!lastname) & ", " & Null2String(RSEMPINFO!FIRSTNAME)
        txtPosition = Null2String(RSEMPINFO!Position)
    End If
End Sub

Sub FillCombo()
    Set rsShiftCode = New ADODB.Recordset
    Set rsShiftCode = gconDMIS.Execute("Select ShiftCode from HRMS_Time_Shift_Code")
    If Not rsShiftCode.BOF And Not rsShiftCode.EOF Then
        Do While Not rsShiftCode.EOF
            Combo1.AddItem Null2String(rsShiftCode!shiftcode)
            rsShiftCode.MoveNext
        Loop
    End If
End Sub

Sub InitGrid()
    With grdShift
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1200: .ColWidth(2) = 1300: .ColWidth(3) = 1000: .ColWidth(4) = 1000
        .Row = 0
        .Col = 0: .Text = "EMP NO"
        .Col = 1: .Text = "Shift"
        .Col = 2: .Text = "Date From"
        .Col = 3: .Text = "Date To"
        .Col = 4: .Text = "ID"
    End With
End Sub

Private Sub cmdAdd_Click()
    Picture3.Visible = True
    Frame2.Enabled = True
    Frame2.Caption = "ADD"
    Picture1.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Picture3.Visible = False
    Frame2.Enabled = True
    Frame2.Enabled = False
    Picture1.Visible = True
End Sub

Private Sub cmdDelete_Click()
    gconDMIS.Execute "Delete from HRMS_Time_Shift_Management where ID =" & labID
    rsRefresh
    FillGrid
End Sub

Private Sub cmdEdit_Click()
    Picture3.Visible = True
    Frame2.Enabled = True
    Frame2.Caption = "EDIT"
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
    txtSearch = ""
End Sub

Private Sub cmdNext_Click()
    RSEMPINFO.MoveNext
    If RSEMPINFO.EOF Then
        RSEMPINFO.MoveLast
        ShowLastRecordMsg
    End If
    storeMemvars
End Sub

Private Sub cmdPrevious_Click()
    RSEMPINFO.MovePrevious
    If RSEMPINFO.BOF Then
        RSEMPINFO.MoveFirst
        ShowFirstRecordMsg
    End If
    storeMemvars
End Sub

Private Sub cmdSave_Click()
    Dim vtxtNo                                                        As String
    Dim vcombo1                                                       As String
    Dim vdtpicker1                                                    As String
    Dim vdateless1                                                    As String
    Dim dateLess1                                                     As Date
    dateLess1 = DTPicker1.Value
    vdateless1 = N2Str2Null(DateAdd("d", -1, dateLess1))

    vtxtNo = N2Str2Null(txtNo)
    vcombo1 = N2Str2Null(Combo1)
    vdtpicker1 = N2Str2Null(DTPicker1.Value)

    If Combo1 = "" Or Combo1 = "Combo1" Then
        MsgBox "Please select name of shift to follow"
        Exit Sub
    End If

    If Frame2.Caption = "ADD" Then

        If rsTimeShiftMgt.RecordCount <> 0 Then
            rsTimeShiftMgt.MoveLast
            gconDMIS.Execute "Update HRMS_Time_Shift_Management set " & _
                           " DateTo = " & vdateless1 & _
                           " where ID = " & Null2String(rsTimeShiftMgt!ID)


        End If
        gconDMIS.Execute "Insert into HRMS_Time_Shift_Management ( EmplNo, Shift, DateFrom )" & _
                       " values ( " & vtxtNo & "," & vcombo1 & "," & vdtpicker1 & ")"

    End If
    If Frame2.Caption = "EDIT" Then
        If IsNumeric(labID) Then
            If rsTimeShiftMgt.RecordCount > 1 Then
                rsTimeShiftMgt.MovePrevious
                If Null2String(rsTimeShiftMgt!DateFrom) > vdtpicker1 Then
                    MsgBox "No two time shifts may overlap on same date!...Please check the entered date"
                    Exit Sub
                Else
                    gconDMIS.Execute "Update HRMS_Time_Shift_Management set " & _
                                   " EmplNo = " & vtxtNo & "," & _
                                   " Shift = " & vcombo1 & "," & _
                                   " DateFrom = " & vdtpicker1 & _
                                   " where ID = " & labID
                End If
            End If
        Else
            MsgBox "Can't Save Without selecting an item to edit!"
            Exit Sub
        End If
    End If
    rsRefresh
    FillGrid
    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Picture3.Visible = False
    rsRefresh
    storeMemvars
    FillCombo
    FillListview
    InitGrid
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub grdShift_Click()
    labID = grdShift.TextMatrix(grdShift.RowSel, grdShift.ColSel)
    On Error Resume Next
    DTPicker1.Value = grdShift.TextMatrix(grdShift.RowSel, 2)
    On Error Resume Next
    Combo1 = grdShift.TextMatrix(grdShift.RowSel, 1)

End Sub

Private Sub lsTimeShift_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RSEMPINFO.Bookmark = rsFind(RSEMPINFO.Clone, "EmpNo", Me.lsTimeShift.SelectedItem).Bookmark
    storeMemvars
End Sub

Private Sub txtNo_Change()
    rsrefresh2 txtNo
    FillGrid
End Sub

Private Sub TXTSEARCH_Change()
    Listview_Loadval lsTimeShift.ListItems, gconDMIS.Execute("select EmpNo from HRMS_EmpInfo where Empno like '%" & Repleys(txtSearch) & "%'")
End Sub

