VERSION 5.00
Begin VB.Form frmCRIS_Group 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4260
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboGroup 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   4065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3315
      TabIndex        =   1
      Top             =   540
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
End
Attribute VB_Name = "frmCRIS_Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CustomerID                            As Long
Public DataID                                As Long
Public CustomerType                          As String
Event AddEditID(MID As Long)


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If CustomerType = "PP" Or CustomerType = "PC" Then
        oConSQL.Execute ("Update CRIS_PROFILE Set ContactType=" & cboGroup.ItemData(cboGroup.ListIndex) & " Where ProfileID= " & CustomerID)
    Else
        oConSQL.Execute ("Update ALL_CUSTOMER Set ContactType=" & cboGroup.ItemData(cboGroup.ListIndex) & " Where ID= " & CustomerID)
    End If
    MessagePop RecSaveOk, "Record Updated", "Profile's Group Updated"
    RaiseEvent AddEditID(cboGroup.ItemData(cboGroup.ListIndex))
End Sub

Private Sub Form_Load()
    FillCombo "SELECT DataID, MasterData from CRIS_vw_Master_PullDown  Where MasterType='Contact Type'", 0, 1, cboGroup
    If DataID > 0 Then
        cboGroup.ListIndex = SelectCombo(cboGroup, CStr(DataID), True)
    Else
        cboGroup.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CustomerID = 0
    DataID = 0
    CustomerType = vbNullString
End Sub
