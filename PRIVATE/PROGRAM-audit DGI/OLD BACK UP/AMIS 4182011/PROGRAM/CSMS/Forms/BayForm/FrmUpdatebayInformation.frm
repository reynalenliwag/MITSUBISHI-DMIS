VERSION 5.00
Begin VB.Form frmCSMSUpdatebayInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assigned Bay"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmUpdatebayInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblbaydesc 
      BackColor       =   &H8000000F&
      Height          =   345
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   570
      Width           =   4995
   End
   Begin VB.CommandButton cmdtech1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4830
      MouseIcon       =   "FrmUpdatebayInformation.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdatebayInformation.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Select Technician"
      Top             =   2340
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   6270
      Top             =   3840
   End
   Begin VB.TextBox txtemp 
      Height          =   315
      Left            =   7290
      TabIndex        =   20
      Top             =   3660
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdSelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6390
      MouseIcon       =   "FrmUpdatebayInformation.frx":05D8
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdatebayInformation.frx":072A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Select from Customer Source"
      Top             =   540
      Width           =   480
   End
   Begin VB.TextBox txtemp3 
      Height          =   315
      Left            =   4800
      TabIndex        =   16
      Top             =   3210
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtemp2 
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      Top             =   2790
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtemp1 
      Height          =   315
      Left            =   7920
      TabIndex        =   14
      Top             =   900
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2670
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.CommandButton cmdtech3 
      Caption         =   "Select"
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
      Left            =   5220
      TabIndex        =   11
      Top             =   3540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdtech2 
      Caption         =   "Select"
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
      Left            =   4020
      TabIndex        =   10
      Top             =   3180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txttech3 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3210
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txttech2 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txttech1 
      BackColor       =   &H8000000F&
      Height          =   345
      Left            =   690
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2430
      Width           =   4065
   End
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
      Height          =   840
      Left            =   6075
      MouseIcon       =   "FrmUpdatebayInformation.frx":0A66
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdatebayInformation.frx":0BB8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   1020
      Width           =   825
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Assigned"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5250
      MouseIcon       =   "FrmUpdatebayInformation.frx":0EF6
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdatebayInformation.frx":1048
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Assign Technician"
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label lblbaycode 
      Caption         =   "thebayCOde"
      Height          =   345
      Left            =   1680
      TabIndex        =   24
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Bay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -660
      TabIndex        =   22
      Top             =   630
      Width           =   1785
   End
   Begin VB.Label labItemNo 
      Caption         =   "labItemNo"
      Height          =   315
      Left            =   2850
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label labCust 
      Caption         =   "labCust"
      Height          =   315
      Left            =   3690
      TabIndex        =   18
      Top             =   4050
      Width           =   1275
   End
   Begin VB.Label labRO 
      Caption         =   "labRO"
      Height          =   315
      Left            =   390
      TabIndex        =   17
      Top             =   2190
      Width           =   1245
   End
   Begin VB.Label Label5 
      Caption         =   "Customer Source"
      Height          =   285
      Left            =   510
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label4 
      Caption         =   "Default Technician &3"
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   3270
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "Default Technician &2"
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   2850
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label2 
      Caption         =   "Select Technician"
      Height          =   285
      Left            =   -1080
      TabIndex        =   5
      Top             =   2460
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE :  The following Information need your attention."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   6435
   End
End
Attribute VB_Name = "frmCSMSUpdatebayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GetTechCode(XXX As String) As String
    Dim rsTechnician                                   As ADODB.Recordset
    Set rsTechnician = New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_vw_Technician where EmpNO = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        GetTechCode = Null2String(rsTechnician!Technician)
    End If
End Function

Function checkIfBayAssigned() As Boolean
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    If lblbaydesc.Text = "" Then
        MsgBox "Please Select a Bay", vbInformation, "CSMS"
        Exit Function
    End If
    SQL = "SELECT bay_code,bay_status from CSMS_baymonitoring where bay_code= '" & lblbaycode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then

        If Null2String(RS!Bay_status) = "Allocated" Then
            checkIfBayAssigned = True
            Exit Function
        End If

    End If

    checkIfBayAssigned = False
End Function

Sub SaveBayinformation()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "update CSMS_BayMonitoring set RO = '" & labRO.Caption & "',bay_status ='Allocated' where bay_code='" & lblbaycode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    MsgBox "Bay has been assigned.", vbInformation, "Information"

End Sub

Private Sub cmdAdd_Click()

    If checkIfBayAssigned = True Then
        MsgBox "Bay is allready allocated..please select other Bay", vbInformation, "Information"
        lblbaydesc.Text = ""
        Exit Sub
    End If

    If lblbaydesc.Text = "" Then
        MsgBox "Please select bay!", vbExclamation, "WARNING"
        Exit Sub
    End If

    SaveBayinformation
    cmdCancel.Value = True

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    frmshowbay.Show 1
End Sub

Private Sub cmdtech1_Click()
    frmCSMSShowTechnician.labselect.Caption = "1"
    frmCSMSShowTechnician.labRO.Caption = labRO.Caption
    frmCSMSShowTechnician.Show 1
End Sub

Private Sub cmdtech2_Click()
    frmCSMSShowTechnician.labselect.Caption = "2"
    frmCSMSShowTechnician.Show 1
End Sub

Private Sub cmdtech3_Click()
    frmCSMSShowTechnician.labselect.Caption = "3"
    frmCSMSShowTechnician.Show 1
End Sub

Private Sub Form_Load()
    labCust.Caption = "": labRO.Caption = "": cmdAdd.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Label1(1).ForeColor = vbRed Then
        Label1(1).ForeColor = vbBlack
    Else
        Label1(1).ForeColor = vbRed
    End If
End Sub

Private Sub txttech1_Change()
    If txttech1 = "" Then
        txtemp1 = ""
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub

Private Sub txttech2_Change()
    If txttech2 = "" Then
        txtemp2 = ""
    End If
End Sub

Private Sub txttech3_Change()
    If txttech3 = "" Then
        txtemp = ""
    End If
End Sub

