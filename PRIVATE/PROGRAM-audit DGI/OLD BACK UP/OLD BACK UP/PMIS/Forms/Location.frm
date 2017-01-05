VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_Location 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTS BY LOCATION"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Location.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1515
   ScaleWidth      =   4485
   Begin VB.CommandButton cmdByLocation 
      Caption         =   "View Parts By Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   4335
   End
   Begin Crystal.CrystalReport rptLocation 
      Left            =   4950
      Top             =   2550
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select location from the list"
      Top             =   420
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Location :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   1605
   End
End
Attribute VB_Name = "frmPMISReports_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOCAL_STOCKTYPE                                    As String
Dim LOCAL_ACCESS                                       As String

Sub SETSTOCK_TYPE(xxx As String)
    LOCAL_STOCKTYPE = xxx
    If xxx = "P" Then
        LOCAL_ACCESS = "PARTS LOCATION"
    ElseIf xxx = "A" Then
        LOCAL_ACCESS = "ACCESSORIES LOCATION"
    Else
        LOCAL_ACCESS = "MATERIALS LOCATION"
    End If
End Sub


Private Sub cmdByLocation_Click()

    Screen.MousePointer = 11

    rptLocation.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptLocation.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If LOCAL_STOCKTYPE = "P" Then
        If UCase(cboLocation.Text) = "NO LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.TYPE} = 'P' AND isnull({partmas.location}) = true", DMIS_REPORT_Connection, 1
        ElseIf UCase(cboLocation.Text) = "ALL PARTS BY LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "AllPartsbylocation.rpt", "{partmas.TYPE} = 'P' ", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.TYPE} = 'P' AND {partmas.location} =  '" & cboLocation.Text & "'", DMIS_REPORT_Connection, 1
        End If
    ElseIf LOCAL_STOCKTYPE = "A" Then
        If UCase(cboLocation.Text) = "NO LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.TYPE} = 'A' AND isnull({partmas.location}) = true", DMIS_REPORT_Connection, 1
        ElseIf UCase(cboLocation.Text) = "ALL ACCESSORIES BY LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "AllAccessoriesbylocation.rpt", "{partmas.type} = 'A'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.type} = 'A' and {partmas.location} =  '" & cboLocation.Text & "'", DMIS_REPORT_Connection, 1
        End If
    Else
        If UCase(cboLocation.Text) = "NO LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.TYPE} = 'M' AND isnull({partmas.location}) = true", DMIS_REPORT_Connection, 1
        ElseIf UCase(cboLocation.Text) = "ALL MATERIALS BY LOCATION" Then
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "AllMaterialsbylocation.rpt", "{partmas.TYPE} = 'M'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptLocation, PMIS_REPORT_PATH & "bylocation.rpt", "{partmas.location} =  '" & cboLocation.Text & "' AND {partmas.TYPE} = 'M'", DMIS_REPORT_Connection, 1
        End If
    End If
    Call NEW_LogAudit("V", LOCAL_ACCESS, "", "", "", cboLocation, "", "")
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = LOCAL_ACCESS
    
    
    Dim rsLocation                                     As ADODB.Recordset
    Set rsLocation = gconDMIS.Execute("select location from PMIS_vw_Location WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' order by location asc")

    cboLocation.Clear
    Do While Not rsLocation.EOF
        If LTrim(RTrim(Null2String(rsLocation!Location))) <> "" Then
            cboLocation.AddItem Null2String(rsLocation!Location)
        End If
        rsLocation.MoveNext
    Loop
    If cboLocation.ListCount > 0 Then
        If LOCAL_STOCKTYPE = "P" Then
            cboLocation.AddItem "All Parts By Location", 0
        ElseIf LOCAL_STOCKTYPE = "A" Then
            cboLocation.AddItem "All Accessories By Location", 0
        Else
            cboLocation.AddItem "All Materials By Location", 0
        End If
        cboLocation.AddItem "No Location", 1
    Else
    cboLocation.AddItem "No Location", 0
    End If
    
    Screen.MousePointer = 0
End Sub


