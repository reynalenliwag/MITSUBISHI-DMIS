VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_Log_SalesAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Sales Appointment"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SalesAppointment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   7590
   Begin VB.PictureBox picDataEntry 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   0
      ScaleHeight     =   4110
      ScaleWidth      =   7920
      TabIndex        =   0
      Top             =   1755
      Width           =   7920
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   390
         Width           =   3690
      End
      Begin VB.TextBox txtModelCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3150
         Width           =   1965
      End
      Begin VB.ComboBox cboVehicles 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2400
         Width           =   3570
      End
      Begin VB.TextBox txtModel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   3150
         Width           =   1515
      End
      Begin VB.ComboBox cboColors 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   3750
         Width           =   3570
      End
      Begin VB.ComboBox cboTerms 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3780
         TabIndex        =   18
         Top             =   2400
         Width           =   1740
      End
      Begin VB.ComboBox cboImportance 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   3570
      End
      Begin VB.ComboBox cboAttendingSE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   3570
      End
      Begin MSComCtl2.DTPicker txtExpectedPurchase 
         Height          =   360
         Left            =   3780
         TabIndex        =   20
         Top             =   3150
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   51576835
         CurrentDate     =   39171
      End
      Begin MSComCtl2.DTPicker txtStartTime 
         Height          =   360
         Left            =   1470
         TabIndex        =   9
         Top             =   1740
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   51576835
         UpDown          =   -1  'True
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker txtDate 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1740
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         Format          =   51576833
         CurrentDate     =   39171
      End
      Begin MSComCtl2.DTPicker txtEndTime 
         Height          =   360
         Left            =   2610
         TabIndex        =   10
         Top             =   1740
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   51576835
         UpDown          =   -1  'True
         CurrentDate     =   39084
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Expected Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   3780
         TabIndex        =   38
         Top             =   2880
         Width           =   1230
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   3810
         TabIndex        =   37
         Top             =   150
         Width           =   495
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code/Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2850
         Width           =   990
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Model Descript"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2130
         Width           =   1320
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "TIME FROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   1470
         TabIndex        =   7
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "TIME TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   2610
         TabIndex        =   8
         Top             =   1530
         Width           =   690
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   3510
         Width           =   450
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Expected Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   3780
         TabIndex        =   19
         Top             =   2190
         Width           =   1395
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Importance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "SAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture5 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   7590
      TabIndex        =   22
      Top             =   5865
      Width           =   7590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   795
         Left            =   6780
         MouseIcon       =   "SalesAppointment.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "SalesAppointment.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Cancel"
         Top             =   150
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   6090
         MouseIcon       =   "SalesAppointment.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "SalesAppointment.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Save this Record"
         Top             =   150
         Width           =   705
      End
      Begin VB.Label labid 
         Caption         =   "Label8"
         Height          =   510
         Left            =   270
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   7590
      TabIndex        =   24
      Top             =   0
      Width           =   7590
      Begin VB.TextBox txtEntityName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   210
         Width           =   4935
      End
      Begin VB.TextBox txtEntityContactperson 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtEntityAddress 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "SalesAppointment.frx":11FC
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtEntityPhone 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5070
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox txtEntityMobile 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   5070
         TabIndex        =   26
         Text            =   "09175041620"
         Top             =   720
         Width           =   2370
      End
      Begin VB.TextBox txtEntityEmail 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   5070
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1260
         Width           =   2370
      End
      Begin VB.Label labEntityName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CUSTOMER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   36
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label labEntityAddress 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   35
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CONTACT PERSON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   34
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   33
         Top             =   0
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   32
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5070
         TabIndex        =   31
         Top             =   510
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmCRIS_Log_SalesAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                          As Long
Dim AppointmentID                       As Long
Dim rs                                  As ADODB.Recordset
Friend Sub AddSalesAppointment(xProspectID As Long)
    ProspectID = xProspectID
End Sub
Friend Sub EditAppointment(ID As Long, xProspectID As Long)
    AppointmentID = ID
    ProspectID = xProspectID
End Sub
Private Sub cboVehicles_Click()
    If cboVehicles.ListIndex = -1 Then Exit Sub
    Dim TempRs                          As ADODB.Recordset
    Set TempRs = gconDMIS.Execute("SELECT CODE , MODEL FROM  ALL_MODEL  WHERE ID=" & cboVehicles.ItemData(cboVehicles.ListIndex))
    If Not (TempRs.EOF Or TempRs.BOF) Then
        txtModelCode = Null2String(TempRs!code)
        txtModel = Null2String(TempRs!Model)
    End If
End Sub
Private Sub cmdCancel_Click()
    AppointmentID = 0
    Unload Me
End Sub


'Upating Code       : AXP-0707200712:39
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:



    On Error Resume Next
    cboAttendingSE.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub
Private Sub cmdExit_Click()
    On Error GoTo Errorcode:
    Unload Me
    Exit Sub
Errorcode:
    ShowVBError
End Sub
Private Sub cmdSave_Click()
    Dim SAE                             As String
    Dim StartDateTime                   As String
    Dim EndDateTime                     As String
    Dim Model                           As String
    Dim Color                           As String
    Dim Terms                           As String
    Dim ExpectedPurchase                As String
    Dim ModelCode                       As String
    Dim ModelDescript                   As String
    Dim sql                             As String
    On Error GoTo Errorcode:
    SAE = N2Str2Null(cboAttendingSE.Text)
    StartDateTime = N2Str2Null(DateValue(txtDate.Value) & " " & TimeValue(txtStartTime))
    EndDateTime = N2Str2Null(DateValue(txtDate.Value) & " " & TimeValue(txtEndTime))
    Model = N2Str2Null(txtModel)
    Color = N2Str2Null(cboColors)
    Terms = N2Str2Null(cboTerms)
    ExpectedPurchase = N2Str2Null(txtExpectedPurchase.Value)
    ModelCode = N2Str2Null(txtModelCode)
    ModelDescript = N2Str2Null(cboVehicles)

    If AppointmentID <= 0 Then
        sql = "INSERT INTO CRIS_SalesAppointments("
        sql = sql & " ProspectID, SAE, StartDateTime, EndDateTime,  Model, "
        sql = sql & " Color, Terms, ExpectedPurchase, ModelCode,ModelDescript) " & vbCrLf
        sql = sql & " VALUES("
        sql = sql & ProspectID & " ,"
        sql = sql & SAE & ","
        sql = sql & StartDateTime & ","
        sql = sql & EndDateTime & ","
        sql = sql & Model & ","
        sql = sql & Color & ","
        sql = sql & Terms & ","
        sql = sql & ExpectedPurchase & ","
        sql = sql & ModelCode & ","
        sql = sql & ModelDescript & ")" & vbCrLf & "SELECT @@IDENTITY"
        LogAudit "A", "SALES APPOINTMENT", " SAE " & SAE & " " & " DATE " & StartDateTime & "-" & EndDateTime & "-" & txtEntityName
    Else

        sql = " Update CRIS_SalesAppointments SET "
        sql = sql & " ProspectID=" & ProspectID & ", "
        sql = sql & " SAE= " & SAE & " ,"
        sql = sql & " StartDateTime=" & StartDateTime & ", "
        sql = sql & " EndDateTime=" & EndDateTime & ", "
        sql = sql & " Model= " & Model & ", "
        sql = sql & " ModelCode = " & ModelCode & ", "
        sql = sql & " ModelDescript = " & ModelDescript & ", "
        sql = sql & " Color=" & Color & ", "
        sql = sql & " Terms=" & Terms & ", "
        sql = sql & " ExpectedPurchase=" & ExpectedPurchase
        sql = sql & " WHERE AppointmentID=" & AppointmentID
        LogAudit "E", "SALES APPOINTMENT", " SAE " & SAE & " " & " DATE " & StartDateTime & "-" & EndDateTime & "-" & txtEntityName
    End If


    Dim TempRs                          As ADODB.Recordset
    Set TempRs = gconDMIS.Execute(sql)
    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogAppointment=" & StartDateTime & " where prospectid=" & ProspectID)

    If AppointmentID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Schedule Sucessfully Added", 500, 2
    Else
        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 2
    End If

    Set TempRs = TempRs.NextRecordset
    If Not TempRs Is Nothing Then
        AppointmentID = TempRs.Collect(0)
    End If
    UpdateLog
    cmdCancel.Value = True
    Set TempRs = Nothing
    If FormExist("frmCRIS_Inquiry_SalesAppointment") Then
        frmCRIS_Inquiry_SalesAppointment.ShowMonthlyAppointments
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl Is Nothing Then Exit Sub
    If KeyCode = 13 And (Left(ActiveControl.Name, 3) = "txt" Or Left(ActiveControl.Name, 3) = "cbo") Then
        SendKeys ("{TAB}")
    End If
End Sub
Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitData
    If AppointmentID > 0 Then
        StoreMemvars
    Else
        initMemvars
    End If


    SetEntityDetails ProspectID, vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    AppointmentID = 0
End Sub

Sub InitData()

    Call FillCombo("SELECT DISTINCT Name from SMIS_vw_Srep  ORDER BY [name]", -1, 0, cboAttendingSE)
    Call FillCombo("Select DISTINCT 1, COLOR_DESC FROM ALL_COLOR ORDER BY COLOR_DESC", 0, 1, cboColors)
    Call FillCombo("select ID, DESCRIPT from ALL_MODEL", 0, 1, cboVehicles)
    With cboImportance
        .AddItem "Normal"
        .AddItem "High"
        .AddItem "Very High"
        .AddItem "Low"
        .ListIndex = 0
    End With
    With cboTerms
        .AddItem "Cash"
        .AddItem "Financing"
        .AddItem "Others"
        .ListIndex = 0
    End With





End Sub

Sub initMemvars()
    AppointmentID = 0
    cboAttendingSE.ListIndex = -1
    txtStartTime = TimeValue("8:00AM")

    txtStartTime.MinDate = TimeValue("8:00AM")
    txtStartTime.MaxDate = TimeValue("8:00PM")
    txtDate = DateValue(Now)
    txtEndTime = TimeValue("8:30AM")
    txtModel = ""
    txtModelCode = ""
    cboVehicles.ListIndex = -1
    cboColors.ListIndex = -1
    cboTerms.ListIndex = -1
    txtExpectedPurchase = DateValue(Now)
    txtnotes = ""

End Sub




Sub StoreMemvars()
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM CRIS_SALESAPPOINTMENTS WHERE APPOINTMENTID=" & AppointmentID & " ORDER BY STARTDATETIME DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly


    If Not rs.EOF And Not rs.BOF Then
        'AppointmentID, ProspectID, SAE, StartDateTime, EndDateTime, ClosedDate, Model, ModelDescript, ModelCode,Color, Terms, ExpectedPurchase, Notes
        AppointmentID = rs!AppointmentID
        ProspectID = rs!ProspectID
        cboAttendingSE.ListIndex = SelectCombo(cboAttendingSE, Null2String(rs!SAE))
        txtStartTime = TimeValue(rs!StartDateTime)
        txtEndTime = TimeValue(rs!EndDateTime)
        txtModel = Null2String(rs!Model)
        txtDate = Null2String(rs!StartDateTime)
        txtModelCode = Null2String(rs!ModelCode)
        cboVehicles.ListIndex = SelectCombo(cboVehicles, Null2String(rs!ModelDescript))
        cboColors.ListIndex = SelectCombo(cboColors, Null2String(rs!Color))
        cboTerms.ListIndex = SelectCombo(cboTerms, Null2String(rs!Terms))
        If IsNull(rs!ExpectedPurchase) = False Then
            txtExpectedPurchase = DateValue(rs!ExpectedPurchase)
        Else
            txtExpectedPurchase = Null
        End If
        txtnotes = Null2String(rs!Notes)


    Else

        ShowNoRecord

    End If
End Sub




Sub UpdateLog()

    Dim TSQL                            As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(StartDateTime) FROM CRIS_SalesAppointments  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGAPPOINTMENT=NULL  WHERE PROSPECTID=" & ProspectID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Private Sub txtStartTime_Change()
    If AppointmentID = 0 Then
        txtEndTime = DateAdd("n", 30, TimeValue(txtStartTime))
    End If

End Sub


Private Sub txtStartTime_LostFocus()
    If AppointmentID = 0 Then
        txtEndTime = DateAdd("n", 30, TimeValue(txtStartTime))
    End If

End Sub
Sub SetEntityDetails(xProspectID As Long, xCUSCODE As String)
    Dim TempRs                          As ADODB.Recordset
    txtEntityAddress = ""
    txtEntityContactperson = ""
    txtEntityEmail = ""
    txtEntityMobile = ""
    txtEntityName = ""
    txtEntityPhone = ""

    If xProspectID = 0 Then
        labEntityName = "CUSTOMER NAME"
        Set TempRs = gconDMIS.Execute("Select CUSTOMERNAME as [Name], CONTACTPERSON, PHONE, MOBILE, ADDRESS, EMAIL from CRIS_VW_ALLPROFILE WHERE CUSCDE=" & N2Str2Null(xCUSCODE))
    Else
        labEntityName = "PROSPECT NAME"
        Set TempRs = gconDMIS.Execute("Select ACCTNAME As [NAME], CONTACTPERSON, TELEPHONE as PHONE , MOBILE, ADDRESS , EMAIL  from CRIS_PROSPECTS WHERE PROSPECTID=" & N2Str2Null(xProspectID))
    End If

    If Not (TempRs.EOF Or TempRs.BOF) Then
        txtEntityAddress = Null2String(TempRs!Address)
        txtEntityContactperson = Null2String(TempRs!ContactPerson)
        txtEntityEmail = Null2String(TempRs!EMAIL)
        txtEntityMobile = Null2String(TempRs!Mobile)
        txtEntityName = Null2String(TempRs!Name)
        txtEntityPhone = Null2String(TempRs!Phone)
    End If
    Set TempRs = Nothing
End Sub
