VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSSellingDealer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selling Dealer"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   FillColor       =   &H000080FF&
   ForeColor       =   &H8000000D&
   Icon            =   "frmCSMSSellingDealer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewRec 
      Caption         =   "View"
      Height          =   495
      Left            =   300
      TabIndex        =   15
      ToolTipText     =   "View Record"
      Top             =   4530
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   600
      Left            =   2040
      TabIndex        =   14
      ToolTipText     =   "Edit Selected Record"
      Top             =   4530
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   570
      Left            =   3210
      TabIndex        =   13
      ToolTipText     =   "Delete Selected Record"
      Top             =   4530
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   600
      Left            =   4470
      TabIndex        =   12
      ToolTipText     =   "Update Record"
      Top             =   4530
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   4470
      TabIndex        =   11
      Top             =   4530
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   90
      TabIndex        =   9
      Top             =   2130
      Width           =   6555
      Begin MSComctlLib.ListView DealerView 
         Height          =   1620
         Left            =   165
         TabIndex        =   10
         Top             =   240
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   2858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483647
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Selling Dealer"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   615
      Left            =   4470
      TabIndex        =   8
      Top             =   4530
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5670
      TabIndex        =   7
      ToolTipText     =   "Exit Window"
      Top             =   4530
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selling Dealer Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6540
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   6
         Top             =   1140
         Width           =   3105
      End
      Begin VB.TextBox txt 
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   5
         Top             =   780
         Width           =   3105
      End
      Begin VB.TextBox txt 
         Height          =   270
         Index           =   0
         Left            =   1635
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Selling Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   450
         TabIndex        =   3
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Selling Dealer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   780
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmCSMSSellingDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim error                                                             As Boolean
Dim UPDATE_MODE                                                       As Boolean
Dim TheSellingcode                                                    As String

Private Sub cmdAdd_Click()
    Call lockedField(True)
    cmdSave.Visible = True
    cmdAdd.Visible = False
    On Error Resume Next
    txt(1).SetFocus
    txt(0) = generate_SellingCode()
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SELLING DEALER") = False Then Exit Sub
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset
    Dim answer                                                        As Integer
    Dim thecode                                                       As String

    thecode = TheSellingcode

    If TheSellingcode <> "" Then
        answer = MsgBox("Are you sure Do you want To delete " & thecode & "?", vbYesNo + vbExclamation, "WARNING!")
        If answer = vbYes Then

            SQL = "DELETE FROM CSMS_sellingDealer where sellingCode='" & thecode & "'"

            gconDMIS.Execute (SQL)
            MsgBox "All Information Has Been Deleted!", vbInformation, "Confirm"
            Call fillListView
            Call initMemvars

        End If
    Else
        MsgBox "Please Select Information!", vbInformation, "Nothing To Delete"
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "SELLING DEALER") = False Then Exit Sub
    On Error Resume Next
    txt(1).SetFocus
    Call lockedField(True)
    cmdUpdate.Visible = True
    cmdDelete.Visible = True
    cmdSave.Visible = False

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Sub SaveData()
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset
    Dim code                                                          As String
    Dim sellingdealer                                                 As String
    Dim Description                                                   As String
    Dim Address                                                       As String
    Dim Contact                                                       As String

    On Error GoTo loaderror

    code = Trim(txt(0).Text)
    sellingdealer = Trim(txt(1).Text)
    Description = Trim(txt(2).Text)

    If (Len(sellingdealer) = 0) Then
        MsgBox "Missing Parameter!", vbInformation
        On Error Resume Next
        txt(1).SetFocus
        Exit Sub
    End If

    If (Len(Description) = 0) Then
        MsgBox "Missing Parameter!", vbInformation
        On Error Resume Next
        txt(2).SetFocus
        Exit Sub
    End If

    If UPDATE_MODE = False Then
        SQL = "INSERT INTO CSMS_sellingDealer VALUES('" & code & "','" & sellingdealer & "','" & Description & "')"

    Else
        SQL = "UPDATE CSMS_sellingDealer SET Sellingdealer='" & sellingdealer & "',description='" & Description & "' WHERE SellingCode ='" & code & "'"
    End If

    gconDMIS.Execute (SQL)
    Call initMemvars
    fillListView
    MsgBox "All information Has been Save!", vbInformation, "Save Complete"
    cmdUpdate.Visible = False
    Call lockedField(False)
    Exit Sub

loaderror:

    MsgBox Err.Description

End Sub
Private Sub cmdSave_Click()
    UPDATE_MODE = False
    Call SaveData
    Call initMemvars
    cmdAdd.Visible = True
    cmdSave.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    UPDATE_MODE = True
    Call SaveData
End Sub

Private Sub cmdViewRec_Click()
    cmdAdd.Visible = True
    cmdUpdate.Visible = False
     Call fillListView
End Sub

Private Sub DealerView_Click()

    On Error Resume Next
    cmdEdit.Visible = True
    cmdDelete.Visible = True
    TheSellingcode = DealerView.ListItems.ITEM(DealerView.SelectedItem.INDEX)
    TheSellingcode = DealerView.SelectedItem.SubItems(0)

     Call fillField

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call lockedField(False)
    txt(0).Locked = True
    cmdEdit.Visible = False
    cmdUpdate.Visible = False
    cmdDelete.Visible = False
    cmdSave.Visible = False
    UPDATE_MODE = False
End Sub

Sub fillListView()
    Dim ITEM                                                          As ListItem
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

    SQL = "select * from CSMS_sellingDealer"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If rs.EOF And rs.BOF Then
        MsgBox "No Record Avalable!", vbInformation, "INFO"
    End If
    DealerView.ListItems.Clear
    With rs
        Do While Not .EOF
            Set ITEM = DealerView.ListItems.Add(, , !sellingCode)
            ITEM.SubItems(1) = Null2String(!sellingdealer)
            ITEM.SubItems(2) = Null2String(!Description)

            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub
Private Function generate_SellingCode() As String
    Dim rs                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim temp                                                          As Integer

    SQL = "SELECT sellingCode FROM CSMS_sellingdealer"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    With rs
        If .BOF And .EOF Then
            temp = 0
        Else
            temp = !sellingCode
        End If
        generate_SellingCode = Format(temp + 1, "00000")
    End With
    Set rs = Nothing
End Function

Sub initMemvars()

    txt(0).Text = ""
    txt(1).Text = ""
    txt(2).Text = ""

End Sub

Sub lockedField(ByVal b As Boolean)

    txt(1).Locked = Not b
    txt(2).Locked = Not b

End Sub
Sub fillField()
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset
    Dim thecode                                                       As String

     thecode = TheSellingcode
    SQL = "SELECT * FROM CSMS_sellingDealer where sellingCode='" & thecode & "'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    With rs
        txt(0).Text = Null2String(!sellingCode)
        txt(1).Text = Null2String(!sellingdealer)
        txt(2).Text = Null2String(!Description)
    End With
    Set rs = Nothing
End Sub

