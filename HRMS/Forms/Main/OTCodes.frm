VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSOTCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Codes"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7920
   ForeColor       =   &H00D8E9EC&
   Icon            =   "OTCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   7920
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2250
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   12
      Top             =   3840
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
         MouseIcon       =   "OTCodes.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "OTCodes.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":0A4C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "OTCodes.frx":0DB2
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":0F04
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "OTCodes.frx":122F
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":1381
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "OTCodes.frx":16DD
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":182F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "OTCodes.frx":1B42
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":1C94
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "OTCodes.frx":1F8E
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":20E0
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
         Left            =   30
         MouseIcon       =   "OTCodes.frx":2438
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":258A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6390
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   21
      Top             =   3840
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
         Left            =   720
         MouseIcon       =   "OTCodes.frx":28E9
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":2A3B
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "OTCodes.frx":2D79
         MousePointer    =   99  'Custom
         Picture         =   "OTCodes.frx":2ECB
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4560
      Left            =   45
      ScaleHeight     =   4500
      ScaleWidth      =   1845
      TabIndex        =   11
      Top             =   90
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   -45
         Picture         =   "OTCodes.frx":321B
         Top             =   -315
         Width           =   9915
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2010
      ScaleHeight     =   2295
      ScaleWidth      =   5865
      TabIndex        =   9
      Top             =   1440
      Width           =   5865
      Begin MSComctlLib.ListView lstOTCodes 
         Height          =   2175
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
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
         MouseIcon       =   "OTCodes.frx":16F78
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODES"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESC"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RATE"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.PictureBox picOTCodes 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   2010
      ScaleHeight     =   1305
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   90
      Width           =   5865
      Begin VB.CheckBox Check1 
         Caption         =   "Holiday "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   24
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txtPay_Code 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   750
         TabIndex        =   5
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtPay_Rate 
         Alignment       =   1  'Right Justify
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
         Left            =   750
         TabIndex        =   4
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txtPay_Desc 
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
         Left            =   750
         TabIndex        =   3
         Top             =   450
         Width           =   4935
      End
      Begin Crystal.CrystalReport rptOTCodes 
         Left            =   5310
         Top             =   30
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Height          =   255
         Left            =   -60
         TabIndex        =   7
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
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
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   510
         Width           =   675
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         TabIndex        =   2
         Top             =   900
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1050
         TabIndex        =   1
         Top             =   900
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmHRMSOTCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOTCodes                                                         As ADODB.Recordset
Dim AddorEdit                                                         As String

'UPDATE BY : MJP 10-01-07 05:37 PM -----------------------------------------------------------------
Sub GenerateNewOTCode()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select Pay_Code From HRMS_OTCodes Order BY Pay_COde DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        RSTMP.MoveNext
        txtPay_Code.Text = Format(RSTMP!PAY_CODE + 1, "000")
    Else
        txtPay_Code.Text = Format(1, "000")
    End If
    Set RSTMP = Nothing
End Sub
'UPDATE BY : MJP 10-01-07 05:37 PM -----------------------------------------------------------------

Sub rsrefresh()
    Set rsOTCodes = New ADODB.Recordset
    rsOTCodes.Open "select * from HRMS_OTCodes order by Pay_Code", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemvars()
    picOTCodes.Enabled = True
    txtPay_Code.Text = ""
    txtPay_Rate.Text = ""
    txtPay_Desc.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsOTCodes.EOF And Not rsOTCodes.BOF Then
        picOTCodes.Enabled = False
        labID.Caption = rsOTCodes!ID
        txtPay_Code.Text = Null2String(rsOTCodes!PAY_CODE)
        txtPay_Rate.Text = Null2String(rsOTCodes!pay_rate)
        txtPay_Desc.Text = Null2String(rsOTCodes!PAY_DESC)
        If IsNull(rsOTCodes!IsHoliday) = False Then
            If rsOTCodes!IsHoliday = True Then
                Check1.Value = 1
            Else
                Check1.Value = 0
            End If
        Else
            Check1.Value = 0
        End If

    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsOTCodes2                                                    As ADODB.Recordset
    lstOTCodes.Enabled = False
    lstOTCodes.Sorted = False: lstOTCodes.ListItems.Clear
    Set rsOTCodes2 = New ADODB.Recordset
    Set rsOTCodes2 = gconDMIS.Execute("select Pay_Code,Pay_Desc,Pay_Rate from HRMS_OTCodes")
    If Not (rsOTCodes2.EOF And rsOTCodes2.BOF) Then
        Listview_Loadval Me.lstOTCodes.ListItems, rsOTCodes2
        lstOTCodes.Refresh
        lstOTCodes.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    'On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Add", "FILES OVERTIME CODES") = False Then Exit Sub

    AddorEdit = "ADD"
    InitMemvars

    'UPDATE BY : MJP 10-01-07 05:37 PM -----------------------------------------------------------------
    'DESCRIPTION : TO GENERATE NEW OT CODE
    GenerateNewOTCode
    'UPDATE BY : MJP 10-01-07 05:37 PM -----------------------------------------------------------------

    lstOTCodes.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picOTCodes.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstOTCodes.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    'On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Delete", "FILES OVERTIME CODES") = False Then Exit Sub

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_OTCodes where id = " & labID.Caption

        LogAudit "X", "DELETE OVERTIME CODE", LOGNAME & "-" & txtPay_Code.Text
        ShowDeletedMsg
    End If

    rsrefresh
    StoreMemVars

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    'On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "FILES OVERTIME CODES") = False Then Exit Sub

    AddorEdit = "EDIT"
    picOTCodes.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstOTCodes.Enabled = False

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Pls use the List view to find...", vbInformation, "Find"
End Sub

Private Sub cmdNext_Click()
    rsOTCodes.MoveNext
    If rsOTCodes.EOF Then
        rsOTCodes.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsOTCodes.MovePrevious
    If rsOTCodes.BOF Then
        rsOTCodes.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    'On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "FILES OVERTIME CODES") = False Then Exit Sub

    Screen.MousePointer = 11
    rptOTCodes.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptOTCodes.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptOTCodes.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptOTCodes.Formulas(3) = "PRINTEDBY = '" & LOGNAME & "'"

    PrintSQLReport rptOTCodes, HRMS_REPORT_PATH & "OT_MASTERFILE.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    LogAudit "V", "PRINT OVERTIME CODE", LOGNAME

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim vHoliday                                                      As Integer
    txtPay_Code.Text = N2Str2Null(txtPay_Code.Text)
    txtPay_Rate.Text = N2Str2Null(txtPay_Rate.Text)
    txtPay_Desc.Text = N2Str2Null(txtPay_Desc.Text)
    If Check1.Value = 0 Then
        vHoliday = 0
    Else
        vHoliday = 1
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_OTCodes " & _
                         "(Pay_Code,Pay_Rate,Pay_Desc,isholiday) " & _
                       " values (" & txtPay_Code.Text & ", " & _
                         "" & txtPay_Rate.Text & ", " & txtPay_Desc.Text & ", " & vHoliday & ")"

        LogAudit "A", "ADD OVERTIME CODE", LOGNAME & "-" & txtPay_Code
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update HRMS_OTCodes set" & _
                       " Pay_Code = " & txtPay_Code.Text & "," & _
                       " Pay_Desc = " & txtPay_Desc.Text & "," & _
                       " Pay_Rate = " & txtPay_Rate.Text & "," & _
                       " Isholiday= " & vHoliday & _
                       " where id = " & labID.Caption

        LogAudit "E", "EDIT OVERTIME CODE", LOGNAME & "-" & txtPay_Code
        ShowSuccessFullyUpdated
    End If

    rsrefresh
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsrefresh
    StoreMemVars
    FillGrid
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstOTCodes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOTCodes
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

Private Sub lstOTCodes_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstOTCodes_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsOTCodes.Bookmark = rsFind(rsOTCodes.Clone, "Pay_Code", Me.lstOTCodes.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtPay_Code_LostFocus()
    txtPay_Code.Text = UCase(txtPay_Code.Text)
End Sub

