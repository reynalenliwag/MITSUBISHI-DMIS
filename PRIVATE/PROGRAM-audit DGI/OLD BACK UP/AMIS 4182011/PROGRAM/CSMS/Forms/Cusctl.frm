VERSION 5.00
Begin VB.Form frmCSMSCusctl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Control"
   ClientHeight    =   1995
   ClientLeft      =   390
   ClientTop       =   510
   ClientWidth     =   6270
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Cusctl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   6270
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   450
      ScaleHeight     =   795
      ScaleWidth      =   5685
      TabIndex        =   8
      Top             =   1050
      Width           =   5745
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
         Left            =   4980
         MouseIcon       =   "Cusctl.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   15
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
         MouseIcon       =   "Cusctl.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   15
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
         Left            =   3540
         MouseIcon       =   "Cusctl.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   15
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
         Left            =   2835
         MouseIcon       =   "Cusctl.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   15
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
         Left            =   2130
         MouseIcon       =   "Cusctl.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   15
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
         Left            =   1425
         MouseIcon       =   "Cusctl.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   15
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
         Left            =   705
         MouseIcon       =   "Cusctl.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   15
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
         Left            =   0
         MouseIcon       =   "Cusctl.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   15
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   2
      Top             =   -30
      Width           =   6105
      Begin VB.CommandButton cmdGenCtl 
         Caption         =   "Update Customer Controls"
         Height          =   735
         Left            =   5130
         TabIndex        =   7
         Top             =   210
         Width           =   885
      End
      Begin VB.TextBox txtCtlcde 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   840
      End
      Begin VB.TextBox txtCtldsc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   3870
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
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
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4725
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   17
      Top             =   1080
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
         Left            =   735
         MouseIcon       =   "Cusctl.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "Cusctl.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "Cusctl.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   5550
      TabIndex        =   6
      Top             =   600
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5490
      TabIndex        =   5
      Top             =   720
      Width           =   225
   End
End
Attribute VB_Name = "frmCSMSCusctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCusCtl                           As ADODB.Recordset
Dim AddorEdit                          As String

Private Sub cmdGenCtl_Click()
    Dim NewCtlCde                      As String
    Dim rsCUSTOMER                     As ADODB.Recordset
    Dim k                              As Integer
    Screen.MousePointer = 11
    gconDMIS.Execute "delete from ALL_Cusctl"
    gconDMIS.Execute "delete from ALL_Cusctl"
    gconDMIS.Execute "delete from ALL_Cusctl"
    gconDMIS.Execute "delete from ALL_Cusctl"
    For k = 65 To 90
        Set rsCUSTOMER = New ADODB.Recordset
        rsCUSTOMER.Open "select cuscde from ALL_CUSMAS where left(cuscde,1) = '" & Chr(k) & "' order by cuscde desc", gconDMIS
        If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCUSTOMER!Cuscde, 2, 5)) + 1, "00000")
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
            gconDMIS.Execute "insert into ALL_Cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
    Next
    Screen.MousePointer = 0
    rsRefresh
    StoreMemVars
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
    On Error Resume Next
    txtCtlcde.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode
    If Not rsCusCtl.BOF Or Not rsCusCtl.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from ALL_Cusctl where id = " & labid.Caption
            ShowDeletedMsg
            LogAudit "X", "CUSTOMER CONTROL", txtCtlcde
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

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtCtlcde.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                        As String
    findStr = InputSpeechBox("Please Input Customer Control Code or Description...", txtCtlcde.Text)
    If findStr <> "" Then
        On Error Resume Next
        rsCusCtl.Bookmark = rsFind(rsCusCtl.Clone, "ctlcde", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error GoTo ErrorCode
            rsCusCtl.Bookmark = rsFind(rsCusCtl.Clone, "ctldsc", findStr).Bookmark
        End If
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        ShowCantFind findStr
        Resume Next
    End If
End Sub

Private Sub cmdNext_Click()
    rsCusCtl.MoveNext
    If rsCusCtl.EOF Then
        rsCusCtl.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsCusCtl.MovePrevious
    If rsCusCtl.BOF Then
        rsCusCtl.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    MsgSpeechBox "Not Yet Implemented"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    If txtCtlcde.Text = "" Or txtCtldsc.Text = "" Then
        MsgSpeechBox "Code and Description is Required"
        On Error Resume Next
        txtCtlcde.SetFocus
        Exit Sub
    End If
    Dim vtxtCtlcde, vtxtCtldsc         As String
    vtxtCtlcde = N2Str2Null(txtCtlcde.Text)
    vtxtCtldsc = N2Str2Null(txtCtldsc.Text)
    If AddorEdit = "ADD" Then
        If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
            rsCusCtl.MoveLast
            labid.Caption = NumericVal(rsCusCtl!ID) + 1
        End If
        gconDMIS.Execute "Insert into ALL_Cusctl" & _
                       " (ctlcde,ctldsc)" & _
                       " values (" & vtxtCtlcde & ", " & vtxtCtldsc & ")"
        LogAudit "A", "CUSTOMER CONTROL", txtCtlcde
    Else
        gconDMIS.Execute "update ALL_Cusctl set" & _
                       " ctlcde = " & vtxtCtlcde & "," & _
                       " ctldsc = " & vtxtCtldsc & _
                       " where id = " & labid.Caption

    End If
    rsRefresh
    On Error Resume Next
    rsCusCtl.Find "id = " & labid.Caption
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    InitMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub InitMemvars()
    txtCtlcde.Text = ""
    txtCtldsc.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
        labid.Caption = rsCusCtl!ID
        txtCtlcde.Text = Null2String(rsCusCtl!ctlcde)
        txtCtldsc.Text = Null2String(rsCusCtl!ctldsc)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsCusCtl = New ADODB.Recordset
    rsCusCtl.Open "select * from ALL_Cusctl order by ctlcde asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSCusctl = Nothing
    UnloadForm Me
End Sub
