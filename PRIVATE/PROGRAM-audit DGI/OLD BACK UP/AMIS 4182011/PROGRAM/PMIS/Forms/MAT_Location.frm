VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISMAT_Location 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MATERIALS BY LOCATION"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_Location.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1350
   ScaleWidth      =   4440
   Begin wizButton.cmd cmdByLocation 
      Height          =   615
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "View the parts in the selected location"
      Top             =   660
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   1085
      TX              =   "View Materials By Location"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "MAT_Location.frx":0E42
   End
   Begin Crystal.CrystalReport rptLocation 
      Left            =   3810
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "BY LOCATION"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ComboBox cboLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select location from the list"
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "frmPMISMAT_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdByLocation_Click()

    Screen.MousePointer = 11
    If cboLocation.Text = "No Location" Then
        rptLocation.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLocation.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.TYPE} = 'M' AND isnull({partmas.location}) = true", DMIS_REPORT_Connection, 1
    ElseIf cboLocation.Text = "All Materials By Location" Then
        rptLocation.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLocation.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptLocation, PMIS_REPORT_PATH & "AllMaterialsbylocation.rpt", "{partmas.TYPE} = 'M'", DMIS_REPORT_Connection, 1
    Else
        rptLocation.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptLocation.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.location} =  '" & cboLocation.Text & "' AND {partmas.TYPE} = 'M'", DMIS_REPORT_Connection, 1
    End If
    Call NEW_LogAudit("V", "MATERIALS LOCATION", "", "", "", cboLocation, "", "")
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MATERIALS LOCATION)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MATERIALS LOCATION", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsLocation                                                    As ADODB.Recordset
    Set rsLocation = New ADODB.Recordset
    rsLocation.Open "select location from PMIS_vw_Location WHERE [TYPE] = 'M' order by location asc", gconDMIS
    If Not rsLocation.EOF And Not rsLocation.BOF Then
        rsLocation.MoveFirst
        cboLocation.Clear
        cboLocation.AddItem "All Materials By Location"
        Do While Not rsLocation.EOF
            If Null2String(rsLocation!Location) <> "" Then
                cboLocation.AddItem Null2String(rsLocation!Location)
            End If
            rsLocation.MoveNext
        Loop
        cboLocation.AddItem "No Location"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

