VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmJobs 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jobs Data Entry"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ForeColor       =   &H8000000D&
   Icon            =   "Jobs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   6855
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   30
      ScaleHeight     =   855
      ScaleWidth      =   5865
      TabIndex        =   16
      Top             =   4770
      Width           =   5895
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
         Left            =   5040
         MaskColor       =   &H0000FFFF&
         Picture         =   "Jobs.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
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
         Picture         =   "Jobs.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   825
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
         Picture         =   "Jobs.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "Jobs.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "Jobs.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "Jobs.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "Jobs.frx":17B6
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "Jobs.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   6765
      Begin VB.ComboBox cboCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   270
         Width           =   5025
      End
      Begin Crystal.CrystalReport rptROJOBS 
         Left            =   6120
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Jobs Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.TextBox txtJCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtDesc1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1050
         Width           =   5325
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   3105
         Left            =   120
         TabIndex        =   21
         Top             =   1470
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   5477
         _Version        =   393216
         Cols            =   5
         ForeColor       =   0
         BackColorFixed  =   -2147483635
         ForeColorFixed  =   16777215
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483635
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
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "JOB CATEGORY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   180
         TabIndex        =   19
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   780
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   4230
      ScaleHeight     =   855
      ScaleWidth      =   1665
      TabIndex        =   17
      Top             =   4770
      Width           =   1695
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
         Left            =   840
         MaskColor       =   &H0000FFFF&
         Picture         =   "Jobs.frx":1DCA
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
         Left            =   120
         MaskColor       =   &H0000FFFF&
         Picture         =   "Jobs.frx":20DC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   15
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJOBMAST As Recordset
Dim rsJOB_CATEGORY As Recordset
Dim AddorEdit As String

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdCancel.SetFocus
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdDelete.SetFocus
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdEdit.SetFocus
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdExit.SetFocus
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdFind.SetFocus
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdNext.SetFocus
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPrevious.SetFocus
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
PrintSQLReport rptJOBMAST, CSMIOS_REPORT_PATH & "JOBMAST.rpt", "", CSMIOS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPrint.SetFocus
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdSave.SetFocus
End Sub

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
If Not rsJOBMAST.EOF And Not rsJOBMAST.BOF Then txtJCode.SetFocus
initMemvars
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorCode
If Not rsJOBMAST.BOF Or Not rsJOBMAST.EOF Then
   If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = 6 Then
      gconCSMIOS.Execute "delete from JOBMAST where id = " & labID.Caption
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

Private Sub cmdFind_Click()
Dim findStr As String
findStr = InputBox("Please Input Job Code or Description ...", "Find")
If findStr <> "" Then
   On Error Resume Next
   rsJOBMAST.Bookmark = rsFind(rsJOBMAST.Clone, "jcode", findStr).Bookmark
   If Err.Number = 3021 Then
      On Error GoTo ErrorCode
      rsJOBMAST.Bookmark = rsFind(rsJOBMAST.Clone, "desc1", findStr).Bookmark
   End If
End If
StoreMemvars
Exit Sub

ErrorCode:
If Err.Number = 3021 Then
   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
   Resume Next
End If
End Sub

Private Sub cmdNext_Click()
If Not rsJOBMAST.EOF Or Not rsJOBMAST.BOF Then labPrev.Caption = labID.Caption
rsJOBMAST.MoveNext
If rsJOBMAST.EOF Then
   rsJOBMAST.MoveLast
   MsgBox "Last of Record"
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
If Not rsJOBMAST.EOF Or Not rsJOBMAST.BOF Then labPrev.Caption = labID.Caption
rsJOBMAST.MovePrevious
If rsJOBMAST.BOF Then
   rsJOBMAST.MoveFirst
   MsgBox "Beginning of Record"
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
If IsNull(txtJCode.Text) = True Then
   MsgBox "Code must not be empty"
   Exit Sub
Else
End If
If txtDesc1.Text = "" Then
   MsgBox "Description is Required", vbInformation, "Error"
   Exit Sub
End If
If AddorEdit = "ADD" Then
   Dim rsfindDup As Recordset
   Set rsfindDup = New Recordset
       rsfindDup.Open "select jcode from JOBMAST where jcode = '" & txtJCode.Text & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsfindDup.EOF And Not rsfindDup.BOF Then
      MsgBox "Code already exist!"
      Exit Sub
   End If
End If

Dim VTXTJCode, VTXTDesc1 As String
Dim VTXTStd_mHrs, VTXTFlatrate As Double
Dim VTXTPOCode, VTXTValidate As String

VTXTJCode = N2Str2Null(txtJCode.Text)
VTXTDesc1 = N2Str2Null(txtDesc1.Text)
VTXTStd_mHrs = N2Str2IntZero(txtStd_mHrs.Text)
VTXTFlatrate = N2Str2IntZero(txtFlatrate.Text)
VTXTPOCode = N2Str2Null(txtPOCode.Text)
VTXTValidate = N2Str2Null(txtValidate.Text)

If AddorEdit = "ADD" Then
   If Not rsJOBMAST.EOF And Not rsJOBMAST.BOF Then
      rsJOBMAST.MoveLast
      labID.Caption = Val(rsJOBMAST!ID) + 1
   End If
   gconCSMIOS.Execute "Insert into JOBMAST" & _
                    " (jcode,desc1,std_mhrs,flatrate,pocode,validate)" & _
                    " values (" & VTXTJCode & ", " & VTXTDesc1 & ", " & VTXTStd_mHrs & ", " & _
                    " " & VTXTFlatrate & ", " & VTXTPOCode & ", " & VTXTValidate & ")"
Else
   gconCSMIOS.Execute "update JOBMAST set" & _
                    " jcode = " & VTXTJCode & "," & _
                    " desc1 = " & VTXTDesc1 & "," & _
                    " std_mhrs = " & VTXTStd_mHrs & "," & _
                    " flatrate = " & VTXTFlatrate & "," & _
                    " pocode = " & VTXTPOCode & "," & _
                    " validate = " & VTXTValidate & _
                    " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsJOBMAST.Find "id =" & labID.Caption
cmdCancel.Value = True
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
cmdCancel.Value = True
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
Set rsJOBMAST = New Recordset
    rsJOBMAST.Open "select * from JOBMAST order by main_cat asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
rsJOBMAST.MoveFirst
Do While Not rsJOBMAST.EOF
      If Null2String(rsJOBMAST!desc2) <> "" Then
         gconCSMIOS.Execute "update jobmast set" & _
                            " desc1 = " & N2Str2Null(Null2String(rsJOBMAST!desc1) & " " & Null2String(rsJOBMAST!desc2)) & _
                            " where id = " & rsJOBMAST!ID
      End If
   rsJOBMAST.MoveNext
Loop
Frame1.Enabled = False
initMemvars
StoreMemvars
If Not rsJOBMAST.EOF Or Not rsJOBMAST.BOF Then labPrev.Caption = labID.Caption
Screen.MousePointer = 0
End Sub

Sub initMemvars()
txtJCode.Text = ""
txtDesc1.Text = ""
txtStd_mHrs.Text = ""
txtFlatrate.Text = ""
txtPOCode.Text = ""
txtValidate.Text = ""
End Sub

Sub StoreMemvars()
If Not rsJOBMAST.EOF And Not rsJOBMAST.BOF Then
   labID.Caption = rsJOBMAST!ID
   txtJCode.Text = Null2String(rsJOBMAST!jcode)
   txtDesc1.Text = Null2String(rsJOBMAST!desc1)
   txtStd_mHrs.Text = N2Str2Zero(rsJOBMAST!std_mhrs)
   txtFlatrate.Text = N2Str2Zero(rsJOBMAST!flatrate)
   txtPOCode.Text = Null2String(rsJOBMAST!pocode)
   txtValidate.Text = Null2String(rsJOBMAST!Validate)
Else
   MsgBox "No Such Record!", vbCritical, "Warning"
   cmdAdd.Value = True
End If
End Sub

Sub rsRefresh()
Set rsJOBMAST = New Recordset
    rsJOBMAST.Open "select * from JOBMAST order by jcode asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
End Sub
