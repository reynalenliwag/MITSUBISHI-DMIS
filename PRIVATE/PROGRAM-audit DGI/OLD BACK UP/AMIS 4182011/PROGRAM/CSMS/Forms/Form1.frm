VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "WIZPROGBAR.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin wizProgBar.Prg Prg1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   661
      Picture         =   "Form1.frx":0000
      ForeColor       =   0
      BarPicture      =   "Form1.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   510
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR As ADODB.Recordset
Dim rsOrd_Hd As ADODB.Recordset
Dim rsORD_HIST As ADODB.Recordset

Private Sub Command1_Click()
Dim vRONO As String
Dim i As Integer
Set rsREPOR = New ADODB.Recordset
Set rsREPOR = gconCSMIOS.Execute("Select * from Repor WHERE DEALER_TYPE = " & DEALER_TYPE & " Order by Rep_Or asc")
If Not rsREPOR.EOF And Not rsREPOR.BOF Then
   rsREPOR.MoveFirst: i = 0
   Do While Not rsREPOR.EOF
      vRONO = Left(Null2String(rsREPOR!rep_or), 1) & Right(Null2String(rsREPOR!rep_or), 6)
      If Null2String(rsREPOR!invoice) <> "" Then
         gconPMIOS.Execute ("Update Ord_HD Set IN_PROCESS = 'N' Where RONO = '" & vRONO & "'")
         gconPMIOS.Execute ("Update Ord_HIST Set IN_PROCESS = 'N' Where RONO = '" & vRONO & "'")
      Else
         gconPMIOS.Execute ("Update Ord_HD Set IN_PROCESS = 'Y' Where RONO = '" & vRONO & "'")
         gconPMIOS.Execute ("Update Ord_HIST Set IN_PROCESS = 'Y' Where RONO = '" & vRONO & "'")
      End If
      i = i + 1
      Prg1.Value = (i / rsREPOR.RecordCount) * 100
      Prg1.Text = Int(Prg1.Value) & "%"
      rsREPOR.MoveNext
      DoEvents
   Loop
End If
End Sub
