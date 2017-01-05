VERSION 5.00
Begin VB.Form frmSignatories 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signatories"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "Signatories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   5805
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5685
      TabIndex        =   27
      Top             =   3930
      Width           =   5715
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
         Left            =   4920
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   705
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
         Left            =   4230
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   705
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
         Left            =   3480
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   765
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
         Left            =   2790
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   705
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
         Left            =   2100
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   705
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
         Left            =   1410
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   705
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
         Left            =   720
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
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
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   30
      TabIndex        =   19
      Top             =   30
      Width           =   5715
      Begin VB.TextBox txtNotedBy2 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1770
         Width           =   3645
      End
      Begin VB.TextBox txtCorpSec 
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
         Left            =   1920
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3390
         Width           =   3645
      End
      Begin VB.TextBox txtAccountNo 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2580
         Width           =   3645
      End
      Begin VB.TextBox txtSBManager 
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
         Left            =   1920
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2970
         Width           =   3645
      End
      Begin VB.TextBox txtGeneralManager 
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
         Left            =   1920
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2160
         Width           =   3645
      End
      Begin VB.TextBox txtNotedBy1 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1380
         Width           =   3645
      End
      Begin VB.TextBox txtApprovedBy 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   990
         Width           =   3645
      End
      Begin VB.TextBox txtCheckedBy 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtPreparedBy 
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
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   210
         Width           =   3645
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "2nd Noted By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   32
         Top             =   1830
         Width           =   1725
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "Corporate Sec."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   3450
         Width           =   1725
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Bank Account No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   30
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Bank Manager"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   3030
         Width           =   1725
      End
      Begin VB.Label labPrev 
         BackColor       =   &H8000000D&
         Caption         =   "Label9"
         Height          =   345
         Left            =   150
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label labid 
         BackColor       =   &H8000000D&
         Caption         =   "Label9"
         Height          =   315
         Left            =   180
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "General Manager"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   2220
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "1st Noted By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Checked By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   660
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Prepared By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   270
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   4200
      ScaleHeight     =   855
      ScaleWidth      =   1515
      TabIndex        =   28
      Top             =   3930
      Width           =   1545
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
         Left            =   750
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":1DCA
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   30
         MaskColor       =   &H0000FFFF&
         Picture         =   "Signatories.frx":20DC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSignatories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSignatories As Recordset
Dim AddorEdit As String

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
initMemvars
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.SetFocus
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdDelete.SetFocus
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdit.SetFocus
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdExit.SetFocus
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFind.SetFocus
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdNext.SetFocus
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPrevious.SetFocus
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPrint.SetFocus
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSave.SetFocus
End Sub
'-----------------------------------

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsSignatories.BOF Or Not rsSignatories.EOF Then
   If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = 6 Then
      gconAMIS.Execute "delete * from Signatories where id = " & labID.Caption
   End If
Else
   MsgBox "Nothing to delete!", vbExclamation, "Warning"
End If
rsRefresh
StoreMemvars
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Exit Sub
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode

txtPreparedBy.Text = N2Str2Null(txtPreparedBy.Text)
txtCheckedBy.Text = N2Str2Null(txtCheckedBy.Text)
txtApprovedBy.Text = N2Str2Null(txtApprovedBy.Text)
txtNotedBy1.Text = N2Str2Null(txtNotedBy1.Text)
txtNotedBy2.Text = N2Str2Null(txtNotedBy2.Text)
txtGeneralManager.Text = N2Str2Null(txtGeneralManager.Text)
txtAccountNo.Text = N2Str2Null(txtAccountNo.Text)
txtSBManager.Text = N2Str2Null(txtSBManager.Text)
txtCorpSec.Text = N2Str2Null(txtCorpSec.Text)
If AddorEdit = "ADD" Then
   gconAMIS.Execute "insert into signatories " & _
                    "(preparedby,checkedby,approvedby,notedby1,notedby2,generalmanager,accountno,sbmanager,corpsec)" & _
                    " values (" & txtPreparedBy.Text & ", " & txtCheckedBy.Text & ", " & txtApprovedBy.Text & _
                    ", " & txtNotedBy1.Text & ", " & txtNotedBy2.Text & ", " & txtGeneralManager.Text & ", " & txtAccountNo.Text & _
                    ", " & txtSBManager.Text & ", " & txtCorpSec.Text & ")"
Else
   gconAMIS.Execute "update Signatories set" & _
                    " preparedby = " & txtPreparedBy.Text & "," & _
                    " checkedby = " & txtCheckedBy.Text & "," & _
                    " approvedby = " & txtApprovedBy.Text & "," & _
                    " notedby1 = " & txtNotedBy1.Text & "," & _
                    " notedby2 = " & txtNotedBy2.Text & "," & _
                    " generalmanager = " & txtGeneralManager.Text & "," & _
                    " accountno = " & txtAccountNo.Text & "," & _
                    " SBmanager = " & txtSBManager.Text & "," & _
                    " corpsec = " & txtCorpSec.Text & _
                    " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsSignatories.Find "id = " & labID.Caption
cmdCancel.Value = True
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "" Then
   KeyAscii = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyReturn
            If Mid(Me.ActiveControl.Name, 1, 3) = "txt" Or Mid(Me.ActiveControl.Name, 1, 3) = "opt" Then
               SendKeys "{TAB}"
            End If
End Select
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
Set rsSignatories = New Recordset
    rsSignatories.Open "select * from Signatories", gconAMIS, adOpenForwardOnly, adLockReadOnly
Frame1.Enabled = False
initMemvars
StoreMemvars
If Not rsSignatories.EOF Or Not rsSignatories.BOF Then labPrev.Caption = labID.Caption
Screen.MousePointer = 0
End Sub

Sub initMemvars()
txtPreparedBy.Text = ""
txtCheckedBy.Text = ""
txtApprovedBy.Text = ""
txtNotedBy1.Text = ""
txtNotedBy2.Text = ""
txtGeneralManager.Text = ""
txtAccountNo.Text = ""
txtSBManager.Text = ""
txtCorpSec.Text = ""
End Sub

Sub StoreMemvars()
If Not rsSignatories.EOF And Not rsSignatories.BOF Then
   labID.Caption = rsSignatories!ID
   txtPreparedBy.Text = Null2String(rsSignatories!preparedby)
   txtCheckedBy.Text = Null2String(rsSignatories!checkedby)
   txtApprovedBy.Text = Null2String(rsSignatories!approvedby)
   txtNotedBy1.Text = Null2String(rsSignatories!notedby1)
   txtNotedBy2.Text = Null2String(rsSignatories!notedby2)
   txtGeneralManager.Text = Null2String(rsSignatories!generalmanager)
   txtAccountNo.Text = Null2String(rsSignatories!accountno)
   txtSBManager.Text = Null2String(rsSignatories!sbmanager)
   txtCorpSec.Text = Null2String(rsSignatories!CorpSec)
   cmdAdd.Enabled = False
Else
   MsgBox "No Such Record!", vbCritical, "Warning"
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsSignatories = New Recordset
    rsSignatories.Open "select * from Signatories", gconAMIS, adOpenForwardOnly, adLockReadOnly
End Sub
