VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMS_Files_SellingDealer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selling Dealer "
   ClientHeight    =   5370
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCSMS_Files_SellingDealer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   20
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":0A1C
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   11
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":11FF
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   10
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   9
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":1B12
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   8
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":1F5E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":20B0
         Style           =   1  'Graphical
         TabIndex        =   7
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":2408
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":255A
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   13
      Top             =   -60
      Width           =   5715
      Begin VB.TextBox txtSelling_Code 
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
         Left            =   1230
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   1200
      End
      Begin VB.TextBox txtSelling_Dealer 
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
         Left            =   1230
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4380
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Code"
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
         TabIndex        =   15
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   630
         TabIndex        =   14
         Top             =   660
         Width           =   1425
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3375
      Left            =   30
      TabIndex        =   18
      Top             =   990
      Width           =   5715
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
         Height          =   375
         Left            =   3030
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lstDealer 
         Height          =   2325
         Left            =   60
         TabIndex        =   5
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":28B9
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
         TabIndex        =   19
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
      TabIndex        =   21
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":2A1B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":2B6D
         Style           =   1  'Graphical
         TabIndex        =   22
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
         MouseIcon       =   "frmCSMS_Files_SellingDealer.frx":2EAB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Files_SellingDealer.frx":2FFD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   8160
      TabIndex        =   17
      Top             =   570
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmCSMS_Files_SellingDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :COPIED FROM CSMS_COLOR MASTER FILE
'DATE STARTED       :8/13/20079:13
'LAST UPDATED       :5/11/200715:02
'DATABASE UPDATES   :UPDATE CSMS_SellingDealer
'WHO UPDATED        :AXPBTT5/11/200715:02
'==========================================================================================


Option Explicit
Dim rsColor                                            As ADODB.Recordset
Dim AddorEdit                                          As String

Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "SELLING DEALER") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    initMemvars
    lstDealer.Enabled = False
    txtSearch.Enabled = False
    optDesc.Enabled = False
    optCode.Enabled = False
    On Error Resume Next
    txtSelling_Code.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lstDealer.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True

    optDesc.Enabled = True
    optCode.Enabled = True

    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SELLING DEALER") = False Then Exit Sub
    On Error GoTo ErrorCode
    '''AXPREDH
    If Not rsColor.BOF Or Not rsColor.EOF Then
        If ShowConfirmDelete = True Then
            If gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_CUSVEH WHERE SELLING_DEALER='" & txtSelling_Code & "'").Fields(0).Value > 0 Then
                MsgBox "Unable to Process Your Request. Selling Dealer Is Use"
                Exit Sub
            Else
                SQL_STATEMENT = "delete from CSMS_SellingDealer where id = " & labid.Caption
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT----------------------------------------------------
                Call NEW_LogAudit("X", "SELLING DEALER", SQL_STATEMENT, labid, "", "CODE: " & txtSelling_Code, "", "")
                'NEW LOG AUDIT----------------------------------------------------

                ShowDeletedMsg
                FillSearchGrid ""
            End If
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
    If Function_Access(LOGID, "Acess_EDIT", "SELLING DEALER") = False Then Exit Sub
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
    Dim lng                                            As Integer
    'On Error GoTo ErrorCode:

    If txtSelling_Code.Text = "" Or txtSelling_Dealer.Text = "" Then
        ShowIsRequiredMsg "Code and Name of Dealer "
        On Error Resume Next
        txtSelling_Code.SetFocus
        Exit Sub
    End If
    '''''''AXPBTT5/11/200715:02
    lng = gconDMIS.Execute("select Count(*) from CSMS_SellingDealer WHERE DEALERCODE=" & N2Str2Null(txtSelling_Code)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(LTrim(RTrim(Null2String(rsColor!DEALERCODE)))) <> UCase(LTrim(RTrim(txtSelling_Code))) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into CSMS_SellingDealer" & _
                      " (DEALERCode,DealerNAME)" & _
                      " values (" & N2Str2Null(txtSelling_Code.Text) & ", " & N2Str2Null(txtSelling_Dealer.Text) & ")"
        gconDMIS.Execute SQL_STATEMENT

        Call NEW_LogAudit("A", "SELLING DEALER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtSelling_Code), "DEALERCODE", "CSMS_SELLINGDEALER"), "", "CODE: " & txtSelling_Code, "", "")
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_SellingDealer set" & _
                      " DEALERcode = " & N2Str2Null(txtSelling_Code.Text) & "," & _
                      " DealerNAME = " & N2Str2Null(txtSelling_Dealer.Text) & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT

        Call NEW_LogAudit("E", "SELLING DEALER", SQL_STATEMENT, "", labid, "CODE: " & txtSelling_Code, "", "")
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    If AddorEdit = "EDIT" Then
        rsColor.Find ("ID=" & labid)
    End If

    cmdCancel.Value = True
    FillSearchGrid ""

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsColor2                                       As ADODB.Recordset
    lstDealer.Sorted = False
    lstDealer.ListItems.Clear
    lstDealer.Enabled = False
    Set rsColor2 = New ADODB.Recordset

    If optCode.Value = True Then
        Set rsColor2 = gconDMIS.Execute("select  DEALERCODE, DEALERNAME, ID from CSMS_SellingDealer where DEALERCODE like'" & ReplaceQuote(XXX) & "%' order by DEALERCODE asc")
    Else
        Set rsColor2 = gconDMIS.Execute("select  DEALERCODE, DEALERNAME, ID from CSMS_SellingDealer where DEALERNAME  like'" & ReplaceQuote(XXX) & "%' order by DEALERNAME asc")
    End If

    If Not (rsColor2.EOF And rsColor2.BOF) Then
        Listview_Loadval Me.lstDealer.ListItems, rsColor2
        lstDealer.Refresh
        lstDealer.Enabled = True
    End If

End Sub

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
    txtSelling_Code.Text = vbNullString
    txtSelling_Dealer.Text = vbNullString
End Sub

Private Sub lstDealer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDealer
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

Private Sub lstDealer_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstDealer_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsColor.MoveFirst
    rsColor.Find ("ID=" & ITEM.ListSubItems(2).Text)
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
    rsColor.Open "select * from CSMS_SellingDealer order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsColor.EOF And Not rsColor.BOF Then
        labid.Caption = rsColor!ID
        txtSelling_Code.Text = LTrim(RTrim(Null2String(rsColor!DEALERCODE)))
        txtSelling_Dealer.Text = LTrim(RTrim(Null2String(rsColor!dealerNAME)))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub txtSelling_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub
