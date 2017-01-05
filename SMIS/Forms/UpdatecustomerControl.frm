VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizprogbar.ocx"
Begin VB.Form frmSMIS_Files_UpdateCustomerControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Customer Control "
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UpdatecustomerControl.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3375
   Begin wizProgBar.Prg Prg1 
      Height          =   390
      Left            =   225
      TabIndex        =   2
      Top             =   900
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   688
      Picture         =   "UpdatecustomerControl.frx":01CA
      ForeColor       =   0
      BarPicture      =   "UpdatecustomerControl.frx":01E6
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
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Update"
      Height          =   675
      Left            =   615
      MouseIcon       =   "UpdatecustomerControl.frx":0202
      MousePointer    =   99  'Custom
      Picture         =   "UpdatecustomerControl.frx":0354
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   675
      Left            =   1605
      MouseIcon       =   "UpdatecustomerControl.frx":05EF
      MousePointer    =   99  'Custom
      Picture         =   "UpdatecustomerControl.frx":0741
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmSMIS_Files_UpdateCustomerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCheck_Click()

'''CUSTOMER CONTROL
    Screen.MousePointer = 11
    gconDMIS.Execute "delete from ALL_CusCtl"
    Dim rsCUSTOMER                     As ADODB.Recordset
    Dim k                              As Integer
    Dim NewCtlCde                      As String
    Dim X                              As Integer
    Prg1.Max = 25

    For k = 65 To 90
        X = X + 1
        Set rsCUSTOMER = New ADODB.Recordset
        rsCUSTOMER.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS

        If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCUSTOMER!CODE, 2, 5)) + 1, "00000")
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
        Prg1.Value = X
    Next
    Screen.MousePointer = 0
    MessagePop InfoFriend, "Updated", "Customer Control Updated"
    '''END CONTROL
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


