VERSION 5.00
Begin VB.UserControl AutoUpdate 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ScaleHeight     =   390
   ScaleWidth      =   3525
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   2670
      Top             =   30
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "popop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3465
   End
End
Attribute VB_Name = "AutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private CN                                             As ADODB.Recordset
Public Event RECORDUPDATED()
Private LAST_DATE                                      As String
Private Sub Timer1_Timer()
    Dim NEW_UPDATED_DATETIME                           As String
    NEW_UPDATED_DATETIME = CHECK_RECORD_UPDATES
    If DateDiff("M", NEW_UPDATED_DATETIME, LAST_DATE) > 0 Then
        LAST_DATE = NEW_UPDATED_DATETIME
        RaiseEvent RECORDUPDATED
    End If
End Sub

Public Property Get oCon() As ADODB.Recordset
    Set oCon = CN
End Property

Public Property Let oCon(ByVal vNewValue As ADODB.Recordset)
    Set CN = vNewValue
End Property
Private Function CHECK_RECORD_UPDATES() As String
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("")
    If Not RS.EOF Or Not RS.BOF Then

    End If
End Function
Private Sub UserControl_Initialize()
    Set CN = New ADODB.Recordset
End Sub

Public Property Get LAST_UPDATED_DATETIME() As String
    LAST_UPDATED_DATETIME = LAST_DATE
End Property

Public Property Let LAST_UPDATED_DATETIME(ByVal vNewValue As String)
    LAST_DATE = vNewValue
End Property
