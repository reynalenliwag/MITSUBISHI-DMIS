VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Files_Document 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Documents"
   ClientHeight    =   5940
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Documents.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   5835
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
      Height          =   1365
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   5715
      Begin VB.CheckBox chkBoth 
         Caption         =   "For Both"
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
         Left            =   3600
         TabIndex        =   25
         Top             =   1020
         Width           =   1785
      End
      Begin VB.CheckBox chkCorp 
         Caption         =   "For Corporate"
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
         Left            =   2010
         TabIndex        =   24
         Top             =   1020
         Width           =   1785
      End
      Begin VB.CheckBox chkInd 
         Caption         =   "For Individual"
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
         Left            =   390
         TabIndex        =   23
         Top             =   1020
         Width           =   1785
      End
      Begin VB.TextBox txtCode 
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
         Left            =   1650
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   3780
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   3810
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Document Code"
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
         Left            =   60
         TabIndex        =   4
         Top             =   240
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
         TabIndex        =   3
         Top             =   660
         Width           =   1425
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   30
      TabIndex        =   6
      Top             =   1410
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lst 
         Height          =   2325
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         MouseIcon       =   "Documents.frx":08CA
         NumItems        =   0
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
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4260
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   20
      Top             =   4980
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
         MouseIcon       =   "Documents.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancel"
         Top             =   45
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
         MouseIcon       =   "Documents.frx":0EBC
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   30
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   4980
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
         MouseIcon       =   "Documents.frx":135E
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "Documents.frx":1816
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":1968
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "Documents.frx":1C93
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":1DE5
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "Documents.frx":2141
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":2293
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "Documents.frx":25A6
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":26F8
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "Documents.frx":29F2
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":2B44
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "Documents.frx":2E9C
         MousePointer    =   99  'Custom
         Picture         =   "Documents.frx":2FEE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
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
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "frmSMIS_Files_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDocument                                                        As ADODB.Recordset
Dim AddorEdit                                                         As String

Function CheckExists() As Boolean
    If gconDMIS.Execute("select count(*)   from SMIS_LoanDocument where documentcode=" & N2Str2Null(TXTCODE)).Fields(0).Value > 0 Then

        CheckExists = True
    End If
End Function

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    lst.Sorted = False: lst.ListItems.Clear
    lst.Enabled = False
    Set TEMPRS = New ADODB.Recordset

    If optCode.Value = True Then
        Set TEMPRS = gconDMIS.Execute("select  Code, DocumentName, Case FormFor when 'I' then 'Individual ' when 'C' then 'Corporate ' when 'B' then 'For Both ' end FormFor , ID from SMIS_DOCUMENT where CODE like'" & ReplaceQuote(XXX) & "%' order by 1 asc")
    Else
        Set TEMPRS = gconDMIS.Execute("select  Code, DocumentName, Case FormFor when 'I' then 'Individual ' when 'C' then 'Corporate ' when 'B' then 'For Both ' end FormFor , ID from SMIS_DOCUMENT where DocumentName like'" & ReplaceQuote(XXX) & "%' order by 1 asc")
    End If

    If Not (TEMPRS.EOF And TEMPRS.BOF) Then
        Listview_Loadval Me.lst.ListItems, TEMPRS
        lst.Refresh
        lst.Enabled = True
    End If

End Sub

Sub InitMemVars()
    TXTCODE.Text = vbNullString
    txtDescription.Text = vbNullString
    chkInd.Value = 0
    chkCorp.Value = 0
    chkBoth.Value = 0
    chkBoth.Enabled = True
    chkCorp.Enabled = True
    chkInd.Enabled = True
    chkBoth.Enabled = True
End Sub

Sub rsRefresh()
    Set rsDocument = New ADODB.Recordset
    rsDocument.Open "select * from SMIS_DOCUMENT  order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    chkInd.Value = 0
    chkCorp.Value = 0
    chkBoth.Value = 0
    If Not (rsDocument.EOF Or rsDocument.BOF) Then
        labid.Caption = rsDocument!ID
        TXTCODE.Text = Null2String(rsDocument!CODE)
        txtDescription.Text = Null2String(rsDocument!DocumentName)

        If Null2String(rsDocument!FormFor) = "I" Then
            chkInd.Value = 1
        ElseIf Null2String(rsDocument!FormFor) = "C" Then
            chkCorp.Value = 1
        ElseIf Null2String(rsDocument!FormFor) = "B" Then
            chkCorp.Value = 1
            chkInd.Value = 1
            chkBoth.Value = 1
        End If

        If CheckExists Then
            TXTCODE.Enabled = False
            chkCorp.Enabled = False
            chkInd.Enabled = False
            chkBoth.Enabled = False
        Else
            TXTCODE.Enabled = True
            chkCorp.Enabled = True
            chkInd.Enabled = True
            chkBoth.Enabled = True
        End If

    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub chkBoth_Click()
    If chkBoth.Value = 1 Then
        chkInd.Value = 1
        chkCorp.Value = 1
    End If
End Sub

Private Sub chkCorp_Click()
    If chkCorp.Value = 1 And chkInd.Value = 1 Then
        chkBoth.Value = 1
    End If
End Sub

Private Sub chkInd_Click()
    If chkCorp.Value = 1 And chkInd.Value = 1 Then
        chkBoth.Value = 1
    End If
End Sub

Private Sub cmdAdd_Click()

    If Function_Access(LOGID, "Acess_ADD", "FINANCIAL DOCUMENTS") = False Then Exit Sub

    On Error GoTo ErrorCode:
    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    TXTCODE.Enabled = True
    picSaves.Visible = True
    InitMemVars
    lst.Enabled = False
    txtSEARCH.Enabled = False
    fraDetails.Enabled = False
    On Error Resume Next
    TXTCODE.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lst.Enabled = True
    txtSEARCH.Enabled = True
    fraDetails.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "FINANCIAL DOCUMENTS") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If CheckExists Then
        MessagePop RecLocekd, "Cannot Edit or Delete", "Current Document Code is in use Cannot Delete"
        Exit Sub
    End If
    On Error GoTo ErrorCode

    If Not rsDocument.BOF Or Not rsDocument.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from SMIS_DOCUMENT where id = " & labid.Caption

            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "X", "FINANCIAL DOCUMENTS", SQL_STATEMENT, Null2String(labid), "", "Code:" & TXTCODE, "", ""
            LogAudit "X", "FINANCIAL DOCUMENT", txtDescription
            ShowDeletedMsg
            rsRefresh
            FillSearchGrid ""

            If Not (rsDocument.EOF Or rsDocument.BOF) Then
                rsDocument.MoveNext

            End If
            StoreMemVars

        End If
    Else
        ShowNothingToDeleteMsg
    End If

    Exit Sub

ErrorCode:
    ShowVBError



End Sub

Private Sub cmdEdit_Click()

    If Function_Access(LOGID, "Acess_EDIT", "FINANCIAL DOCUMENTS") = False Then Exit Sub
    On Error GoTo ErrorCode:
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

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsDocument.MoveNext
    If rsDocument.EOF Then
        rsDocument.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    rsDocument.MovePrevious
    If rsDocument.BOF Then
        rsDocument.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
    Dim lng                                                           As Integer

    On Error GoTo ErrorCode:

    If Trim(TXTCODE.Text) = "" Or Trim(txtDescription.Text) = "" Then
        ShowIsRequiredMsg "Color Code and Description "
        On Error Resume Next
        TXTCODE.SetFocus
        Exit Sub
    End If

    If chkInd.Value = 0 And chkCorp.Value = 0 And chkBoth.Value = 0 Then
        MessagePop InfoVoid, "Required Information Missing", "Select At Least One Check Box Indicated Where Will Be Used"
        On Error Resume Next
        chkInd.SetFocus
        Exit Sub
    End If


    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_DOCUMENT WHERE CODE=" & N2Str2Null(TXTCODE)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exists In Your Database"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsDocument!CODE)) <> UCase(TXTCODE) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exists In Your Database"
            Exit Sub
        End If
    End If


    Dim ForCheck                                                      As String
    If chkBoth.Value = 1 Then
        ForCheck = "B"
    ElseIf chkCorp.Value = 1 And chkInd.Value = 0 Then
        ForCheck = "C"
    ElseIf chkCorp.Value = 0 And chkInd.Value = 1 Then
        ForCheck = "I"

    End If

    If AddorEdit = "ADD" Then
        If Not rsDocument.EOF And Not rsDocument.BOF Then
            rsDocument.MoveLast
            labid.Caption = NumericVal(rsDocument!ID) + 1
        End If
        SQL_STATEMENT = "Insert into SMIS_DOCUMENT " & _
                      " (CODE,DOCUMENTNAME, FormFor)" & _
                      " values (" & N2Str2Null(TXTCODE) & ", " & N2Str2Null(txtDescription) & "," & N2Str2Null(ForCheck) & ")"

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "FINANCIAL DOCUMENTS", SQL_STATEMENT, FindTransactionID(N2Str2Null(TXTCODE), "CODE", "SMIS_DOCUMENT"), "", "Code:" & TXTCODE, "", ""

        LogAudit "A", "FINANCIAL DOCUMENT", txtDescription
    Else
        SQL_STATEMENT = "update SMIS_DOCUMENT set" & _
                      " CODE = " & N2Str2Null(TXTCODE) & "," & _
                      " FormFor = " & N2Str2Null(ForCheck) & "," & _
                      " DOCUMENTNAME = " & N2Str2Null(txtDescription) & _
                      " where id = " & labid.Caption

        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "FINANCIAL DOCUMENTS", SQL_STATEMENT, Null2String(labid), "", "Code:" & TXTCODE, "", ""

        LogAudit "E", "FINANCIAL DOCUMENT", txtDescription
    End If

    rsRefresh
    If AddorEdit = "EDIT" Then
        rsDocument.Find ("ID=" & labid)
    End If

    FillSearchGrid ""
    cmdCancel.Value = True

    If FormExist("frmSMIS_Trans_ApplicationIndividual") Then
        frmSMIS_Trans_ApplicationIndividual.cmdDocumentCheckList_Click
        Unload Me
    End If

    If FormExist("frmSMIS_Trans_ApplicationCorporate") Then
        frmSMIS_Trans_ApplicationIndividual.cmdDocumentCheckList_Click
        Unload Me
    End If





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picAdds.Visible = True And KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (FINANCIAL DOCUMENTS)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "FINANCIAL DOCUMENTS")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    txtSEARCH.Text = vbNullString
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    InitMemVars
    StoreMemVars

    Call AddColumnHeader("CODE, DOCUMENTNAME, FOR ", lst)
    Call ResizeColumnHeader(lst, "20 ,55,20")

    Screen.MousePointer = 0

End Sub

Private Sub lst_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lst
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

Private Sub lst_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lst_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsDocument.MoveFirst
    rsDocument.Find ("ID=" & Item.ListSubItems(lst.ColumnHeaders.Count).Text)
    StoreMemVars
End Sub

Private Sub optCode_Click()
    FillSearchGrid (txtSEARCH.Text)
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optDesc_Click()
    FillSearchGrid (txtSEARCH.Text)
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSEARCH_Change()
    On Error GoTo ErrorCode:
    FillSearchGrid (txtSEARCH.Text)

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

