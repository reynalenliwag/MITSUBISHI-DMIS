VERSION 5.00
Begin VB.Form frmCSMSUpdateCustomerInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assigned Technician"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmUpdateCustomerInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6150
      MouseIcon       =   "FrmUpdateCustomerInfo.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdateCustomerInfo.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Select Technician"
      Top             =   540
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   2820
      Top             =   990
   End
   Begin VB.TextBox txtemp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   6030
      MouseIcon       =   "FrmUpdateCustomerInfo.frx":05D8
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdateCustomerInfo.frx":072A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Select from Customer Source"
      Top             =   2310
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtemp3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   16
      Top             =   3210
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtemp2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      Top             =   2790
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtemp1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   2310
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
      Left            =   3990
      TabIndex        =   11
      Top             =   3240
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
      Left            =   3990
      TabIndex        =   10
      Top             =   2790
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txttech3 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox txttech1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   4335
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
      Left            =   5805
      MouseIcon       =   "FrmUpdateCustomerInfo.frx":0A66
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdateCustomerInfo.frx":0BB8
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
      Left            =   4995
      MouseIcon       =   "FrmUpdateCustomerInfo.frx":0EF6
      MousePointer    =   99  'Custom
      Picture         =   "FrmUpdateCustomerInfo.frx":1048
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Assign Technician"
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label lblJOBCODE 
      BackColor       =   &H008080FF&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   1110
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label labItemNo 
      Caption         =   "labItemNo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      TabIndex        =   21
      Top             =   3300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label labCust 
      Caption         =   "labCust"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   18
      Top             =   3300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label labRO 
      BackColor       =   &H00C0C0FF&
      Caption         =   "labRO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   17
      Top             =   1710
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label5 
      Caption         =   "Customer Source"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label4 
      Caption         =   "Default Technician &3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   3270
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "Default Technician &2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   2850
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select Technician"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   660
      Width           =   1515
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
Attribute VB_Name = "frmCSMSUpdateCustomerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function GetTechCode(XXX As String) As String
    Dim rsTechnician                                   As New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_vw_Technician where EmpNO = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        GetTechCode = LTrim(RTrim(Null2String(rsTechnician!Technician)))
    End If
End Function

Private Sub cmdAdd_Click()
    Dim xEmpNo1                                             As String
    Dim xTech1                                              As String
    Dim xEmpNo2                                             As String
    Dim xTech2                                              As String
    Dim xEmpNo3                                             As String
    Dim xTech3                                              As String
    Dim XCustomerSourceLead                                 As String
    Dim mdone                                               As String
    
    xEmpNo1 = LTrim(RTrim(N2Str2Null(txtemp1)))
    xTech1 = LTrim(RTrim(N2Str2Null(txttech1)))
    xEmpNo2 = LTrim(RTrim(N2Str2Null(txtemp2)))
    xTech2 = LTrim(RTrim(N2Str2Null(txttech2)))
    xEmpNo3 = LTrim(RTrim(N2Str2Null(txtemp3)))
    xTech3 = LTrim(RTrim(N2Str2Null(txttech3)))
    XCustomerSourceLead = LTrim(RTrim(N2Str2Null(txtSource)))
    mdone = "N"

    If Not txttech1.Text = "" Then
        If cmdtech1.Enabled = True Then
            gconDMIS.Execute "update CSMS_RepairOrder set " & _
                           " jstatus = 'S'," & _
                           " EmpNo1 = " & xEmpNo1 & "," & _
                           " Tech1 = " & xTech1 & "" & _
                           " where RO_No = '" & labRO.Caption & "'"
                           
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & labRO & " - SERVICE COUNTER", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            
            If xTech1 <> "NULL" Then
                SQL_STATEMENT = "update CSMS_Ro_det set " & _
                                " TECHNICIAN = " & xTech1 & "," & _
                                " TECHCODE = " & N2Str2Null(GetTechCode(txtemp1.Text)) & _
                                " ,DONE = '" & mdone & _
                                "' where REP_OR = '" & labRO.Caption & "' AND DETCDE = " & N2Str2Null(LABITEMNO.Caption)
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & lblJobCode & " TECH CODE: " & txtemp1, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If
        End If

        If cmdtech2.Enabled = True Then
            SQL_STATEMENT = "update CSMS_RepairOrder set " & _
                          " jstatus = 'S'," & _
                          " EmpNo2 = " & xEmpNo2 & "," & _
                          " Tech2 = " & xTech2 & "" & _
                          " where RO_No = '" & labRO.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & labRO & " - SERVICE COUNTER", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            If xTech2 <> "NULL" Then
                SQL_STATEMENT = "update CSMS_Ro_det set " & _
                              " TECHNICIAN = " & xTech2 & "," & _
                              " TECHCODE = " & N2Str2Null(GetTechCode(txtemp2.Text)) & _
                              " where REP_OR = '" & labRO.Caption & "' AND DETCDE = " & N2Str2Null(LABITEMNO.Caption)
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & lblJobCode & " TECH CODE: " & txtemp1, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If
        End If
        If cmdtech3.Enabled = True Then
            SQL_STATEMENT = "update CSMS_RepairOrder set " & _
                          " jstatus = 'S'," & _
                          " EmpNo3 = " & xEmpNo3 & "," & _
                          " Tech3 = " & xTech3 & "" & _
                          " where RO_No = '" & labRO.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & labRO & " - SERVICE COUNTER", "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            If xTech3 <> "NULL" Then
                SQL_STATEMENT = "update CSMS_Ro_det set " & _
                              " TECHNICIAN = " & xTech3 & "," & _
                              " TECHCODE = " & N2Str2Null(GetTechCode(txtemp3.Text)) & _
                              " where REP_OR = '" & labRO.Caption & _
                              "' AND DETCDE = " & N2Str2Null(LABITEMNO.Caption)
                gconDMIS.Execute SQL_STATEMENT

                'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & lblJobCode & " TECH CODE: " & txtemp1, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            End If
        End If
        If cmdSelect.Enabled = True Then
            SQL_STATEMENT = "update CSMS_RepairOrder set " & _
                          " CustomerSourceLead = " & XCustomerSourceLead & "" & _
                          " where RO_No = '" & labRO.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & labRO & " - SERVICE COUNTER", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If

        'UPDATE BY: MJP 04-23-08 01:11 AM
        Dim rstmp                                      As New ADODB.Recordset
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE empno = '" & txtemp1.Text & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            If Null2String(rstmp!jstatus) = "W" Or Null2String(rstmp!jstatus) = "B" Or Null2String(rstmp!jstatus) = "G" Or Null2String(rstmp!jstatus) = "L" Then

            Else
                SQL_STATEMENT = "update HRMS_EmpInfo set" & _
                              " assignedro = '" & labRO.Caption & "'," & _
                              " jstatus = 'S'" & _
                              " where EmpNo = '" & txtemp1 & "'"
                gconDMIS.Execute SQL_STATEMENT
            End If
        Else
            Set rstmp = New ADODB.Recordset
            Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_EMPINFO WHERE empno = '" & txtemp1.Text & "'")
            If Null2String(rstmp!jstatus) = "W" Or Null2String(rstmp!jstatus) = "B" Or Null2String(rstmp!jstatus) = "G" Or Null2String(rstmp!jstatus) = "L" Then
            Else
                SQL_STATEMENT = "update CSMS_EmpInfo set" & _
                              " assignedro = '" & labRO.Caption & "'," & _
                              " jstatus = 'S'" & _
                              " where EmpNo = '" & txtemp1 & "'"
                gconDMIS.Execute SQL_STATEMENT
            End If
        End If
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("AS", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtemp1), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        'UPDATE BY: MJP 04-23-08 01:11 AM

        'gconDMIS.Execute "update HRMS_EmpInfo set" & _
         '               " assignedro = '" & labRO.Caption & "'," & _
         '               " jstatus = 'S'" & _
         '               " where EmpNo = '" & txtemp1 & "'"


        cmdCancel.Value = True

        MessagePop InfoFriend, "RO Information Updated", "Technician Succesfully Assigned to Job", 1000

        Call frmCSMS_ServiceCounter.Click_ScheduleGrid
    
        
    Else
        MsgBox "Please Select a Technician...", vbInformation, "Assign Technician"
        On Error Resume Next
        cmdtech1.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    frmCSMSSourceLead.Show 1
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
    labCust.Caption = "": labRO.Caption = "": cmdAdd.Enabled = False
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

