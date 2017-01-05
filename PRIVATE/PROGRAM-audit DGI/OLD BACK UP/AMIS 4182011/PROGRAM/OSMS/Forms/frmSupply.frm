VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOSMSFilesSupply 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supply"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000F&
   Icon            =   "frmSupply.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   5835
   Begin VB.Frame Frame1 
      Caption         =   "Supply Data Entry"
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
      Height          =   2655
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5745
      Begin VB.TextBox txtCost 
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   4050
         TabIndex        =   14
         Top             =   2040
         Width           =   1125
      End
      Begin VB.TextBox txtOnHand 
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   1800
         TabIndex        =   13
         Top             =   2040
         Width           =   1125
      End
      Begin VB.TextBox txtLastIssued 
         Enabled         =   0   'False
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtLastReceived 
         Enabled         =   0   'False
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1320
         Width           =   1125
      End
      Begin VB.ComboBox cboSupplier 
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
         Left            =   1800
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtSupplyDesc 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtSupplyCode 
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
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3150
         TabIndex        =   15
         Top             =   2070
         Width           =   1065
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "On-Hand"
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
         TabIndex        =   12
         Top             =   2070
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Last Issued"
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
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Last Received"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   9
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier "
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
         TabIndex        =   7
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   5
         Top             =   660
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Code"
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
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label labSupID 
         Caption         =   "Label7"
         Height          =   375
         Left            =   5100
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   60
      TabIndex        =   16
      Top             =   2730
      Width           =   5745
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
         Height          =   225
         Left            =   1020
         TabIndex        =   18
         Top             =   150
         Width           =   1275
      End
      Begin VB.OptionButton optCode 
         Caption         =   "&Code"
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
         Left            =   120
         TabIndex        =   17
         Top             =   150
         Value           =   -1  'True
         Width           =   795
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
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   390
         Width           =   5625
      End
      Begin MSComctlLib.ListView lstSupply 
         Height          =   2745
         Left            =   30
         TabIndex        =   20
         Top             =   780
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4842
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
         MouseIcon       =   "frmSupply.frx":030A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SUPPLY DESCRIPTION"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   750
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   21
      Top             =   6360
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
         MouseIcon       =   "frmSupply.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   28
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
         MouseIcon       =   "frmSupply.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   27
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
         MouseIcon       =   "frmSupply.frx":0DA1
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":0EF3
         Style           =   1  'Graphical
         TabIndex        =   24
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
         MouseIcon       =   "frmSupply.frx":124F
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":13A1
         Style           =   1  'Graphical
         TabIndex        =   25
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
         MouseIcon       =   "frmSupply.frx":16B4
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":1806
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "frmSupply.frx":1B00
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "frmSupply.frx":1FAA
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4350
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   29
      Top             =   6360
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
         MouseIcon       =   "frmSupply.frx":245B
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":25AD
         Style           =   1  'Graphical
         TabIndex        =   30
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
         MouseIcon       =   "frmSupply.frx":28EB
         MousePointer    =   99  'Custom
         Picture         =   "frmSupply.frx":2A3D
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmOSMSFilesSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupply As ADODB.Recordset
Dim rsSupplier As ADODB.Recordset
Dim AddorEdit As String
Dim PrevSupCODE As String

Private Sub cmdAdd_Click()
    Frame1.Caption = "Add A Record"
    AddorEdit = "ADD"
    Picture1.Visible = False
    Frame1.Enabled = True
    On Error Resume Next
    txtSupplyCode.SetFocus
    initMemvars
End Sub

Sub initMemvars()
    txtSupplyCode.Text = ""
    txtSupplyDesc.Text = ""
    txtLastReceived.Text = ""
    txtLastIssued.Text = ""
    txtOnHand.Text = ""
    txtCost.Text = ""
    InitCBOSUPPLIER
End Sub

Private Sub cmdCancel_Click()
    Frame1.Caption = "Department Data Entry"
    AddorEdit = ""
    Picture1.Visible = True
    Frame1.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBoxXP("Are you sure you want to delete this record?", "Delete Current Record", XP_YesNo, msg_Question) = True Then
        gconDMIS.Execute "delete from OSMS_SUPPLY where ID = " & labSupID.Caption
        rsRefresh
        StoreMemVars
    End If
End Sub

Private Sub cmdEdit_Click()
    Frame1.Caption = "Edit Record"
    AddorEdit = "EDIT"
    PrevSupCODE = labSupID.Caption
    Frame1.Enabled = True
    On Error Resume Next
    txtSupplyCode.SetFocus
    Picture1.Visible = False
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
        Set rsRecordFound = rsSupply.Clone
        rsRecordFound.Find "Supply_Description like '" & AAA & "%'"
        If Not rsRecordFound.EOF Then
            rsSupply.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            Set rsRecordFound = New Recordset
            rsRecordFound.Open "Select * from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
            rsRecordFound.Find "Supply_Code = '" & AAA & "'"
            If Not rsRecordFound.EOF Then
                rsSupply.Bookmark = rsRecordFound.Bookmark
                RecordFound = True
            Else
                RecordFound = False
            End If
        End If
    End If
End Function

Private Sub cmdNext_Click()
    rsSupply.MoveNext
    If rsSupply.EOF Then
        MsgBoxXP "Last of Record!", "Last Record", XP_OKOnly, msg_Information
        rsSupply.MoveLast
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSupply.MovePrevious
    If rsSupply.BOF Then
        MsgBoxXP "Beginning of Record!", "Beggining of Record", XP_OKOnly, msg_Information
        rsSupply.MoveFirst
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0716200719:04
Private Sub cmdSave_Click()

On Error GoTo Errorcode:

    Screen.MousePointer = 11
    If txtSupplyCode.Text = "" Then
        MsgBoxXP "Supply Code must not be empty!", "Input Supply Code", XP_OKOnly, msg_Information
        On Error Resume Next
        txtSupplyCode.SetFocus
        Exit Sub
    End If

    If txtSupplyDesc.Text = "" Then
        MsgBoxXP "Supplier Name must not be empty!", "Input Supplier Name", XP_OKOnly, msg_Information
        On Error Resume Next
        txtSupplyDesc.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        rsSupply.Find " Supply_Code = '" & txtSupplyCode & "'"
        If Not rsSupply.EOF Then
            Screen.MousePointer = 0
            MsgBox "Supply Code already exists!"
            On Error Resume Next
            txtSupplyCode.SetFocus
            Exit Sub
        End If


        gconDMIS.Execute "insert into OSMS_Supply " & _
                         "(Supply_Code, Supply_Description, Supplier_Code, LastRRDate, LAstIssueDate, OnHand, Cost) values (" & N2Str2Null(txtSupplyCode.Text) & "," & N2Str2Null(txtSupplyDesc.Text) & "," & N2Str2Null(SETCBOSUPPLIER(cboSupplier.Text)) & "," & N2Date2Null(txtLastReceived.Text) & "," & N2Date2Null(txtLastIssued.Text) & "," & NumericVal(txtOnHand.Text) & "," & NumericVal(txtCost.Text) & ")"
    Else
        ' If PrevSupCODE <> txtSupplyCode.Text Then
        '     rsSupply.Find " Supply_Code = '" & txtSupplyCode & "'"
        '     If Not rsSupply.EOF Then
        '         Screen.MousePointer = 0
        '         MsgBox "Supply Code already exists!"
        '        txtSupplyCode.SetFocus
        '       Exit Sub
        '  End If
        'End If

        gconDMIS.Execute "update OSMS_SUPPLY set " & _
                         "Supply_Code = " & N2Str2Null(txtSupplyCode.Text) & "," & _
                         "Supply_Description = " & N2Str2Null(txtSupplyDesc.Text) & "," & _
                         "Supplier_Code = " & N2Str2Null(SETCBOSUPPLIER(cboSupplier)) & "," & _
                         "LastRRDate = " & N2Date2Null(txtLastReceived.Text) & "," & _
                         "LastIssueDate = " & N2Date2Null(txtLastIssued.Text) & "," & _
                         "OnHand = " & NumericVal(txtOnHand.Text) & "," & _
                         "Cost = " & NumericVal(txtCost.Text) & "" & _
                         "where ID = '" & PrevSupCODE & "'"
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
    Call AddColumnHeader("SUPPLYCODE, SUPPLYNAME", lstSupply)
    Call ResizeColumnHeader(lstSupply, "28,65")

End Sub

Sub rsRefresh()
    Set rsSupply = New ADODB.Recordset
    rsSupply.Open "select * from OSMS_SUPPLY order by supply_description asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        labSupID.Caption = Null2String(rsSupply!ID)
        txtSupplyCode.Text = Null2String(rsSupply!Supply_Code)
        txtSupplyDesc.Text = Null2String(rsSupply!Supply_Description)
        cboSupplier.Text = SETCBOSUPPLIER2(Null2String(rsSupply!Supplier_code))
        txtLastReceived.Text = Null2Date(rsSupply!lastrrdate)
        txtLastIssued.Text = Null2Date(rsSupply!LastIssueDate)
        txtOnHand.Text = NumericVal(rsSupply!ONHAND)
        txtCost.Text = NumericVal(rsSupply!Cost)
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
        cmdAdd.Value = True
    End If
End Sub

Sub InitCBOSUPPLIER()
    Set rsSupplier = New Recordset
    rsSupplier.Open "Select Supplier_Name from OSMS_Supplier order by Supplier_Name asc", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        cboSupplier.Clear
        Do While Not rsSupplier.EOF
            cboSupplier.AddItem Null2String(rsSupplier!SUPPLIER_NAME)
            rsSupplier.MoveNext
        Loop
    End If
End Sub

Function SETCBOSUPPLIER(XXX As Variant) As String
    Set rsSupplier = New Recordset
    rsSupplier.Open "Select Supplier_Name, SUPPLIER_CODE from   OSMS_Supplier WHERE Supplier_Name = '" & XXX & "'", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SETCBOSUPPLIER = Null2String(rsSupplier!Supplier_code)
    End If
End Function

Function SETCBOSUPPLIER2(XXX As Variant) As String
    Set rsSupplier = New Recordset
    rsSupplier.Open "Select Supplier_Name, SUPPLIER_CODE from  OSMS_Supplier WHERE Supplier_CODE = '" & XXX & "'", gconDMIS
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SETCBOSUPPLIER2 = Null2String(rsSupplier!SUPPLIER_NAME)
    End If
End Function




Private Sub lstSupply_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSupply.Bookmark = rsFind(rsSupply.Clone, "SUPPLY_CODE", lstSupply.SelectedItem.Text).Bookmark
    StoreMemVars
End Sub

Private Sub lstSupply_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSupply
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

Private Sub lstSupply_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsSupply2 As ADODB.Recordset
    lstSupply.Enabled = False
    lstSupply.Sorted = False
    lstSupply.ListItems.Clear
    Set rsSupply2 = New ADODB.Recordset

    If optCode.Value = True Then
        Set rsSupply2 = gconDMIS.Execute("select SUPPLY_CODE,SUPPLY_DESCRIPTION from OSMS_SUPPLY where SUPPLY_CODE like'" & XXX & "%' order by SUPPLY_CODE asc")
    Else
        Set rsSupply2 = gconDMIS.Execute("select SUPPLY_CODE,SUPPLY_DESCRIPTION from OSMS_SUPPLY where SUPPLY_DESCRIPTION like'" & XXX & "%' order by SUPPLY_CODE asc")
    End If

    If Not (rsSupply2.EOF And rsSupply2.BOF) Then
        Listview_Loadval Me.lstSupply.ListItems, rsSupply2
        lstSupply.Refresh
         lstSupply.Enabled = True
    End If
   
End Sub



Private Sub optCode_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub
Private Sub optDesc_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub







