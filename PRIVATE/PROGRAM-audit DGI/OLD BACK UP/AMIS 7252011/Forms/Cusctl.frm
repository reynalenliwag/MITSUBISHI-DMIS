VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Begin VB.Form frmAMISDATACusctl 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Control"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Cusctl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   5985
   Begin wizButton.cmd cmdGenCtl 
      Height          =   765
      Left            =   4860
      TabIndex        =   19
      Top             =   180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1349
      TX              =   "Update Customer Controls"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Cusctl.frx":08CA
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      ScaleHeight     =   825
      ScaleWidth      =   5835
      TabIndex        =   17
      Top             =   1050
      Width           =   5865
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5070
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":08E6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P&rint"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4350
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3630
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":14BA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2910
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":1D84
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2190
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":264E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1470
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":2F18
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   750
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":37E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Cusctl.frx":3C24
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   60
      TabIndex        =   12
      Top             =   -30
      Width           =   5865
      Begin VB.TextBox txtCtlcde 
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Width           =   3540
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      ScaleHeight     =   825
      ScaleWidth      =   5835
      TabIndex        =   18
      Top             =   1050
      Width           =   5865
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5070
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":4066
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4350
         MaskColor       =   &H0000FFFF&
         Picture         =   "Cusctl.frx":50A8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   5550
      TabIndex        =   16
      Top             =   600
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5490
      TabIndex        =   15
      Top             =   720
      Width           =   225
   End
End
Attribute VB_Name = "frmAMISDATACusctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCusCtl As ADODB.Recordset
Dim AddorEdit As String

Private Sub cmdGenCtl_Click()
Dim NewCtlCde As String
Dim rsCustomer As ADODB.Recordset
Dim k As Integer
Screen.MousePointer = 11
gconAmis.Execute "delete from cusctl"
For k = 65 To 90
    Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select custcode from customer where left(custcode,1) = '" & Chr(k) & "' order by custcode desc", gconAmis
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
       NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!custcode, 2, 5)) + 1, "00000")
       gconAmis.Execute "insert into cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
    Else
       gconAmis.Execute "insert into cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
    End If
Next
Screen.MousePointer = 0
rsRefresh
StoreMemvars
End Sub

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
initMemvars
On Error Resume Next
txtCtlcde.SetFocus
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsCusCtl.BOF Or Not rsCusCtl.EOF Then
   If ShowConfirmDelete = True Then
      gconAmis.Execute "delete from cusctl where id = " & labID.Caption
      ShowDeletedMsg
   End If
Else
   ShowNothingToDeleteMsg
End If
rsRefresh
StoreMemvars
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
Dim findStr As String
findStr = InputSpeechBox("Please Input Customer Control Code or Description...", txtCtlcde.Text)
If findStr <> "" Then
   On Error Resume Next
   rsCusCtl.Bookmark = rsFind(rsCusCtl.Clone, "ctlcde", findStr).Bookmark
   If Err.Number = 3021 Then
      On Error GoTo ErrorCode
      rsCusCtl.Bookmark = rsFind(rsCusCtl.Clone, "ctldsc", findStr).Bookmark
   End If
End If
StoreMemvars
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
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsCusCtl.MovePrevious
If rsCusCtl.BOF Then
   rsCusCtl.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
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
Dim vtxtCtlcde, vtxtCtldsc As String
vtxtCtlcde = N2Str2Null(txtCtlcde.Text)
vtxtCtldsc = N2Str2Null(txtCtldsc.Text)
If AddorEdit = "ADD" Then
   If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
      rsCusCtl.MoveLast
      labID.Caption = NumericVal(rsCusCtl!ID) + 1
   End If
   gconAmis.Execute "Insert into cusctl" & _
                      " (ctlcde,ctldsc)" & _
                      " values (" & vtxtCtlcde & ", " & vtxtCtldsc & ")"
Else
   gconAmis.Execute "update cusctl set" & _
                    " ctlcde = " & vtxtCtlcde & "," & _
                    " ctldsc = " & vtxtCtldsc & _
                    " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsCusCtl.Find "id = " & labID.Caption
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
initMemvars
StoreMemvars
DrawXPCtl Me
Screen.MousePointer = 0
End Sub

Sub initMemvars()
txtCtlcde.Text = ""
txtCtldsc.Text = ""
End Sub

Sub StoreMemvars()
If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
   labID.Caption = rsCusCtl!ID
   txtCtlcde.Text = Null2String(rsCusCtl!ctlcde)
   txtCtldsc.Text = Null2String(rsCusCtl!ctldsc)
Else
   ShowNoRecord
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsCusCtl = New ADODB.Recordset
    rsCusCtl.Open "select * from cusctl order by ctlcde asc", gconAmis, adOpenForwardOnly, adLockReadOnly
End Sub
