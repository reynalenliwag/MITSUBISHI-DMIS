VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmAISFILTER 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Applicant"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12840
   Begin Crystal.CrystalReport rptFILTER 
      Left            =   8550
      Top             =   7650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   525
      Left            =   9375
      TabIndex        =   22
      Top             =   7800
      Width           =   1605
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "Exit"
      Height          =   525
      Left            =   11025
      TabIndex        =   15
      Top             =   7800
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By"
      Height          =   3915
      Left            =   75
      TabIndex        =   8
      Top             =   3750
      Width           =   12585
      Begin VB.CheckBox chkGENDER 
         Height          =   435
         Left            =   4755
         TabIndex        =   21
         Top             =   330
         Width           =   255
      End
      Begin VB.CheckBox chkREGION 
         Height          =   435
         Left            =   8295
         TabIndex        =   20
         Top             =   2730
         Width           =   255
      End
      Begin VB.CheckBox chkFIELDS 
         Height          =   435
         Left            =   8295
         TabIndex        =   19
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox chkSTATUS 
         Height          =   435
         Left            =   5655
         TabIndex        =   18
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkAGE 
         Height          =   435
         Left            =   3810
         TabIndex        =   17
         Top             =   1350
         Width           =   255
      End
      Begin VB.CheckBox chkDEGREE 
         Height          =   435
         Left            =   8295
         TabIndex        =   16
         Top             =   1770
         Width           =   255
      End
      Begin VB.ComboBox cboREGION 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2760
         Width           =   6045
      End
      Begin VB.ComboBox cboCSTATUS 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   870
         Width           =   3405
      End
      Begin VB.TextBox txtAGE 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         TabIndex        =   3
         Text            =   "18"
         Top             =   1350
         Width           =   1575
      End
      Begin VB.ComboBox cboDEGREE 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   6045
      End
      Begin VB.ComboBox cboFIELDS 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   6045
      End
      Begin VB.CommandButton cndSEARCH 
         Caption         =   "Search"
         Height          =   525
         Left            =   6930
         TabIndex        =   7
         Top             =   3270
         Width           =   1605
      End
      Begin VB.ComboBox cboGENDER 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Civil Status"
         Height          =   240
         Index           =   5
         Left            =   840
         TabIndex        =   14
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Educational Degree"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   1920
         Width           =   1905
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Age From Up"
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1470
         Width           =   1260
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Study Fields"
         Height          =   240
         Index           =   2
         Left            =   780
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Region Address"
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   2850
         Width           =   1500
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         Caption         =   "Gender"
         Height          =   240
         Index           =   0
         Left            =   1290
         TabIndex        =   9
         Top             =   480
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView lsvFILTER 
      Height          =   3345
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Civil Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Age"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Educ. Attaintment"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Study Fields"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Address"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmAISFILTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAGE_Click()
    If chkAGE.Value = 1 Then
        txtAGE.Enabled = True
    Else
        txtAGE.Enabled = False
    End If
End Sub

Private Sub chkDEGREE_Click()
    If chkDEGREE.Value = 1 Then
        cboDEGREE.Enabled = True
    Else
        cboDEGREE.Enabled = False
    End If
End Sub

Private Sub chkFIELDS_Click()
    If chkFIELDS.Value = 1 Then
        cboFIELDS.Enabled = True
    Else
        cboFIELDS.Enabled = False
    End If
End Sub

Private Sub chkGENDER_Click()
    If chkGENDER.Value = 1 Then
        cboGENDER.Enabled = True
    Else
        cboGENDER.Enabled = False
    End If
End Sub

Private Sub chkREGION_Click()
    If chkREGION.Value = 1 Then
        cboREGION.Enabled = True
    Else
        cboREGION.Enabled = False
    End If
End Sub

Private Sub chkSTATUS_Click()
    If chkSTATUS.Value = 1 Then
        cboCSTATUS.Enabled = True
    Else
        cboCSTATUS.Enabled = False
    End If
End Sub

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cndSEARCH_Click()
    Dim rsPER As ADODB.Recordset, rsREG As ADODB.Recordset, rsEDU As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vcboDEGREE As String, vcboFIELDS As String, vcboREGION As String, Sql As String
    Dim vtxtAGE As Integer, TMP_AGE As Integer
    Dim vGENDER As String, vCSTATUS As String, vREGION As String
    Dim ITEM As ListItem
    
    If Not IsNumeric(txtAGE.Text) Then
        MsgBox "Invalid Age Entry", vbExclamation, "Filter"
        txtAGE.SetFocus
        Exit Sub
    End If
    
    vtxtAGE = CInt(txtAGE)
    vcboDEGREE = Null2String(cboDEGREE)
    vcboFIELDS = Null2String(cboFIELDS)
    vcboREGION = Null2String(cboREGION)
    
    frmAISFILTER.MousePointer = 11
    
'    TMP_AGE = DateDiff("Y", Birthdate, Date)
   
    If chkGENDER.Value = 1 Then                             'GENDER   YES
        If cboGENDER.ListIndex = 0 Then                 'GENDER   UNSPECIFIED
            If chkSTATUS.Value = 1 Then                     'STATUS   YES
                If chkAGE.Value = 1 Then                    'AGE      YES
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where CivilStatus = " & _
                            CInt(cboCSTATUS.ListIndex) & " And Age >= " & CInt(txtAGE) & _
                            " Order By LastName Asc")
                Else                                        'AGE      NO
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where CivilStatus = " & _
                            CInt(cboCSTATUS.ListIndex) & " Order By LastName Asc")
                End If
            Else
                If chkAGE.Value = 1 Then                    'STATUS   NO
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Age >= " & _
                        CInt(txtAGE) & " Order By LastName ASC")
                Else
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Order By LastName ASC")
                End If
            End If
        Else                                                'GENDER   SPECIFIED
            If chkSTATUS.Value = 1 Then                     'STATUS   YES
                If chkAGE.Value = 1 Then                    'AGE      YES
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Gender = " & CInt(cboGENDER.ListIndex) & _
                            " And CivilStatus = " & CInt(cboCSTATUS.ListIndex) & " And Age >= " & _
                            CInt(txtAGE) & " Order By LastName Asc")
                Else                                        'AGE      NO
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Gender = " & CInt(cboGENDER.ListIndex) & _
                            " And CivilStatus = " & CInt(cboCSTATUS.ListIndex) & " Order By LastName Asc")
                End If
            Else                                            'STATUS   NO
                If chkAGE.Value = 1 Then                    'AGE      YES
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Gender = " & CInt(cboGENDER.ListIndex) & _
                            " And Age >= " & CInt(txtAGE) & " Order By LastName Asc")
                Else                                        'AGE      NO
                    Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Gender = " & CInt(cboGENDER.ListIndex) & _
                            " Order By LastName Asc")
                End If
            End If
        End If
    Else                                                    'GENDER   NO
        If chkSTATUS.Value = 1 Then                         'STATUS   YES
            If chkAGE.Value = 1 Then                        'AGE      YES
                Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where CivilStatus = " & _
                        CInt(cboCSTATUS.ListIndex) & " And Age >= " & CInt(txtAGE) & " Order By LastName Asc")
            Else                                            'AGE      NO
                Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where CivilStatus = " & _
                        CInt(cboCSTATUS.ListIndex) & " Order By LastName Asc")
            End If
        Else
            If chkAGE.Value = 1 Then                        'STATUS   NO
                Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Where Age >= " & _
                    CInt(txtAGE) & " Order By LastName ASC")
            Else
                Set rsPER = GetRS("Select * From HRMS_APPLICANT_PERSONAL Order By LastName ASC")
            End If
        End If
    End If
        
    lsvFILTER.ListItems.Clear
    If Not (rsPER.BOF And rsPER.EOF) Then
        Do While Not rsPER.EOF
            If rsPER!GENDER = 1 Then vGENDER = "MALE"
            If rsPER!GENDER = 2 Then vGENDER = "FEMALE"
            If rsPER!GENDER = 0 Then vGENDER = "UNSPECIFIED"
            If rsPER!CIVILSTATUS = 0 Then vCSTATUS = "Single"
            If rsPER!CIVILSTATUS = 1 Then vCSTATUS = "Married"
            If rsPER!CIVILSTATUS = 2 Then vCSTATUS = "Separated"
            If rsPER!CIVILSTATUS = 3 Then vCSTATUS = "Divorced"
            If rsPER!CIVILSTATUS = 4 Then vCSTATUS = "Widowed"
            
            If chkDEGREE.Value = 1 Then 'DEGREE YES
                If chkFIELDS.Value = 1 Then 'FIELDS YES
                    If chkREGION.Value = 1 Then 'REGION YES
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And SchoolType = '" & Null2String(cboDEGREE) & _
                                "' And StudyFields = '" & vcboFIELDS & "'")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And Tmp_Region = " & CInt(cboREGION.ListIndex) & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If                  'REGION YES
                    Else                        'REGION NO
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And SchoolType = '" & Null2String(cboDEGREE) & _
                                "' And StudyFields = '" & vcboFIELDS & "'")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & _
                                rsPER!APPLICANT_ID & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If                  'REGION NO
                    End If                  'FIELDS YES
                Else                        'FIELDS NO
                    If chkREGION.Value = 1 Then 'REGION YES
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And SchoolType = '" & Null2String(cboDEGREE) & "'")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And Tmp_Region = " & CInt(cboREGION.ListIndex) & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If                  'REGION YES
                    Else                        'REGION NO
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And SchoolType = '" & Null2String(cboDEGREE) & "'")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & _
                                APPLICANT_ID & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If
                    End If                      'REGION NO
                End If                      'FIELDS NO
                                        'DEGREE YES
            Else                        'DEGREE NO
                If chkFIELDS.Value = 1 Then 'FIELDS YES
                    If chkREGION.Value = 1 Then 'REGION YES
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And StudyFields = '" & vcboFIELDS & "")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And Tmp_Region = " & CInt(cboREGION.ListIndex) - 1 & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If                  'REGION YES
                    Else                        'REGION NO
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And StudyFields = '" & vcboFIELDS & "'")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicent_ID = " & _
                                APPLICANT_ID & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If
                        End If                  'REGION NO
                    End If                  'FIELDS YES
                Else                        'FIELDS NO
                    If chkREGION.Value = 1 Then 'REGION YES
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & "")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & rsPER!APPLICANT_ID & _
                                " And Tmp_Region = " & CInt(cboREGION.ListIndex) & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If                  'REGION YES
                        End If
                    Else                        'REGION NO
                        Set rsEDU = GetRS("Select * From HRMS_APPLICANT_EDUC Where Applicant_ID = " & rsPER!APPLICANT_ID & "")
                        If Not (rsEDU.BOF And rsEDU.EOF) Then
                            Set rsREG = GetRS("Select * From HRMS_APPLICANT_ADDRESS Where Applicant_ID = " & rsPER!APPLICANT_ID & "")
                            If Not (rsREG.BOF And rsREG.EOF) Then
                                Set ITEM = lsvFILTER.ListItems.Add(, , rsPER!APPLICANT_ID)
                                ITEM.SubItems(1) = Null2String(rsPER!LastName) & "," & Null2String(rsPER!FirstName)
                                ITEM.SubItems(2) = vGENDER
                                ITEM.SubItems(3) = vCSTATUS
                                ITEM.SubItems(4) = rsPER!AGE
                                ITEM.SubItems(5) = rsEDU!SchoolType
                                ITEM.SubItems(6) = rsEDU!StudyFields
                                ITEM.SubItems(7) = ReturnRegion(rsREG!Per_Region)
                            End If                  'REGION YES
                        End If
                    End If                      'REGION NO
                End If                      'FIELDS NO
            End If                      'DEGREE NO

            rsPER.MoveNext
        Loop
    End If
    frmAISFILTER.MousePointer = 0
End Sub

Private Sub TestingSearch()
    Dim rsTmp As ADODB.Recordset
    
    
    Set rsTmp = GetRS("")
End Sub

Function ReturnRegion(INDEX As Integer) As String
    If INDEX = 0 Then ReturnRegion = "ARMM"
    If INDEX = 1 Then ReturnRegion = "Bicol Region"
    If INDEX = 2 Then ReturnRegion = "C.A.R."
    If INDEX = 3 Then ReturnRegion = "Cagayan Valley"
    If INDEX = 4 Then ReturnRegion = "Caraga"
    If INDEX = 5 Then ReturnRegion = "Central Luzon"
    If INDEX = 6 Then ReturnRegion = "Central Visayas"
    If INDEX = 7 Then ReturnRegion = "Eastern Visayas"
    If INDEX = 8 Then ReturnRegion = "Ilocos Region"
    If INDEX = 9 Then ReturnRegion = "National Capital Region"
    If INDEX = 10 Then ReturnRegion = "Nortern Mindanao"
    If INDEX = 11 Then ReturnRegion = "Southern Mindanao"
    If INDEX = 12 Then ReturnRegion = "Southern Tagalog"
    If INDEX = 13 Then ReturnRegion = "Western Mindanao"
    If INDEX = 14 Then ReturnRegion = "Western Visayas"
End Function

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call FillCboGENDER
    Call FillCboFIELDS
    Call FillcboCSTATUS
    Call FillCboDEGREE
    Call FillCboRegion
End Sub

Private Sub FillCboGENDER()
    cboGENDER.AddItem "Unspecified"
    cboGENDER.ItemData(cboGENDER.NewIndex) = 0
    cboGENDER.AddItem "Male"
    cboGENDER.ItemData(cboGENDER.NewIndex) = 1
    cboGENDER.AddItem "Female"
    cboGENDER.ItemData(cboGENDER.NewIndex) = 2
    cboGENDER.ListIndex = 0
End Sub

Private Sub FillCboFIELDS()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_FIELDS Order By Fields ASC")
    cboFIELDS.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboFIELDS.AddItem Null2String(rsTmp!FIELDS)
            rsTmp.MoveNext
        Loop
    End If
    cboFIELDS.ListIndex = 0
End Sub

Private Sub FillcboCSTATUS()
    cboCSTATUS.AddItem "Single"
    cboCSTATUS.AddItem "Married"
    cboCSTATUS.AddItem "Separated"
    cboCSTATUS.AddItem "Divorced"
    cboCSTATUS.AddItem "Widowed"
    cboCSTATUS.ListIndex = 0
End Sub

Private Sub FillCboDEGREE()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_DEGREE Order By Degree ASC")
    cboDEGREE.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboDEGREE.AddItem Null2String(rsTmp!DEGREE)
            rsTmp.MoveNext
        Loop
    End If
    cboDEGREE.ListIndex = 0
End Sub

Private Sub FillCboRegion()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = GetRS("Select * From HRMS_REGIONS Order By Region ASC")
    cboREGION.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            cboREGION.AddItem Null2String(rsTmp!Region)
            rsTmp.MoveNext
        Loop
    End If
    cboREGION.ListIndex = 0
End Sub

Private Sub lsvFILTER_DblClick()
    Dim rs As ADODB.Recordset
    Dim myString As String
    Dim INDEX As Long
    
    If Not lsvFILTER.ListItems.Count = 0 Then
        INDEX = lsvFILTER.SelectedItem.INDEX
        With lsvFILTER
            FROM_FORM_APPLY = "FILTER"
            APPLICANT_ID = CLng(.ListItems(INDEX).Text)
            'frmMain.tbMENU.Enabled = False
            frmAISFILTER.Enabled = False
            frmAISPOSITION_APPLY.Show
            
            Call DisplayApplicantInPosition
            Set rs = gconDMIS.Execute("Select * From HRMS_APPLICANT_IMAGE_LOCATION Where Applicant_ID = " & _
                APPLICANT_ID & "")
            If Not (rs.BOF And rs.EOF) Then
                If Null2String(rs!ImageLocation) <> "" Then
                    On Error Resume Next
                    LoadPic frmAISPOSITION_APPLY.imgAPP, Null2String(rs!ImageLocation)
                Else
                    LoadPic frmAISPOSITION_APPLY.imgAPP, ""
                End If
            Else
                LoadPic frmAISPOSITION_APPLY.imgAPP, ""
            End If
        End With
    End If
End Sub

Private Sub lsvFILTER_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
'    picProfile1.Cls
'        Set rs = GetRS("Select * from HRMS_Applicant_Educ where id = " & CLng(ITEM.Text))
'        If Not (rs.EOF And rs.BOF) Then
'            myString = rs.GetString(adClipString)
'            picProfile1.Print myString
'        End If
End Sub
