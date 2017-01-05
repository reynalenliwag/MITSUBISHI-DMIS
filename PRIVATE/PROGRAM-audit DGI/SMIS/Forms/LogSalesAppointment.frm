VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Log_SalesAppointmentsdfsdf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Appointment"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewVStat 
      Caption         =   "Add From Vehicles Inventory"
      Height          =   675
      Left            =   9345
      TabIndex        =   28
      Top             =   1650
      Width           =   825
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8250
      MaskColor       =   &H0000FFFF&
      TabIndex        =   24
      Top             =   5610
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9330
      MaskColor       =   &H0000FFFF&
      TabIndex        =   23
      Top             =   5610
      Width           =   1020
   End
   Begin VB.PictureBox picModels 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   5385
      ScaleHeight     =   3165
      ScaleWidth      =   4905
      TabIndex        =   9
      Top             =   2325
      Width           =   4905
      Begin VB.TextBox txtYear 
         Height          =   345
         Left            =   1680
         TabIndex        =   27
         Top             =   1710
         Width           =   3090
      End
      Begin VB.ComboBox cboTerms 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Top             =   2160
         Width           =   3090
      End
      Begin VB.TextBox txtClass 
         Height          =   345
         Left            =   1680
         TabIndex        =   20
         Top             =   1320
         Width           =   3090
      End
      Begin VB.ComboBox cboColors 
         Height          =   330
         Left            =   1680
         TabIndex        =   19
         Top             =   900
         Width           =   3090
      End
      Begin VB.TextBox txtMake 
         Height          =   345
         Left            =   1680
         TabIndex        =   18
         Top             =   480
         Width           =   3090
      End
      Begin VB.TextBox txtModel 
         Height          =   345
         Left            =   1680
         TabIndex        =   17
         Top             =   45
         Width           =   3090
      End
      Begin MSComCtl2.DTPicker dtExpectedPurchase 
         Height          =   360
         Left            =   1680
         TabIndex        =   22
         Top             =   2580
         Width           =   3090
         _ExtentX        =   5450
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
         Format          =   51773443
         CurrentDate     =   39171
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "ExpectedPurchase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   16
         Top             =   2700
         Width           =   1560
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Expected Terms"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   7
         Left            =   255
         TabIndex        =   15
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   1230
         TabIndex        =   14
         Top             =   1860
         Width           =   390
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   1185
         TabIndex        =   13
         Top             =   1470
         Width           =   435
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   1185
         TabIndex        =   12
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   11
         Top             =   600
         Width           =   465
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   1110
         TabIndex        =   10
         Top             =   180
         Width           =   510
      End
   End
   Begin VB.PictureBox picEventDetails 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5745
      ScaleHeight     =   1275
      ScaleWidth      =   4815
      TabIndex        =   2
      Top             =   1050
      Width           =   4815
      Begin MSComCtl2.DTPicker dtStartTime 
         Height          =   360
         Left            =   1290
         TabIndex        =   3
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm:ss"
         Format          =   51773442
         CurrentDate     =   39084
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   360
         Left            =   1290
         TabIndex        =   4
         Top             =   30
         Width           =   3120
         _ExtentX        =   5503
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
         Format          =   51773440
         CurrentDate     =   39171
      End
      Begin MSComCtl2.DTPicker dtEndTime 
         Height          =   360
         Left            =   1290
         TabIndex        =   5
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51773442
         CurrentDate     =   39084
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Time To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   11
         Left            =   540
         TabIndex        =   26
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Time From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   10
         Left            =   315
         TabIndex        =   25
         Top             =   570
         Width           =   900
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   6
         Top             =   90
         Width           =   405
      End
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
      Left            =   5385
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   4740
   End
   Begin VB.ComboBox cboImportance 
      Height          =   330
      Left            =   7005
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   660
      Width           =   3120
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   1575
      ScaleHeight     =   5040
      ScaleWidth      =   8475
      TabIndex        =   29
      Top             =   6375
      Visible         =   0   'False
      Width           =   8505
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3795
         Left            =   60
         TabIndex        =   30
         Top             =   750
         Width           =   8355
         _Version        =   655364
         _ExtentX        =   14737
         _ExtentY        =   6694
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8130
         TabIndex        =   34
         Top             =   15
         Width           =   285
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "Select "
         Height          =   375
         Left            =   6750
         TabIndex        =   33
         Top             =   4560
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   7590
         TabIndex        =   32
         Top             =   4560
         Width           =   825
      End
      Begin VB.TextBox txtFitlerViewVehicles 
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Top             =   330
         Width           =   3915
      End
      Begin VB.Image ImgSearchProspect 
         Height          =   330
         Left            =   8040
         MousePointer    =   99  'Custom
         ToolTipText     =   "Enter Character(s) In Text Box And Press Enter To Search Record In Database"
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   3420
         TabIndex        =   36
         Top             =   390
         Width           =   2505
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   -15
         TabIndex        =   35
         Top             =   0
         Width           =   8535
         _Version        =   655364
         _ExtentX        =   15055
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Preview Vehicles On Stock :::"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
      End
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "Attending SAE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   8
      Left            =   5385
      TabIndex        =   8
      Top             =   -30
      Width           =   1200
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      Caption         =   "Importance Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   5415
      TabIndex        =   7
      Top             =   750
      Width           =   1500
   End
End
Attribute VB_Name = "frmSMIS_Log_SalesAppointmentsdfsdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                              As Long
Dim AppointmentID                           As Long

Private Sub Form_Load()
    InitVars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''
    ProspectID = 0

    AppointmentID = 0
End Sub
''''''CALLS
Friend Sub AddSalesAppointment(xProspectID As Long)
    ProspectID = xProspectID
    

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox picViewVehicles.hwnd, False, Me
End Sub

Private Sub cmdOk_Click()


    Dim SAE                                 As String
    Dim StartDateTime                       As String
    Dim EndDateTime                         As String
    Dim Model                               As String
    Dim Make                                As String
    Dim Color                               As String
    Dim Class                               As String
    Dim Year                                As String
    Dim Terms                               As String
    Dim ExpectedPurchase                    As String

    Dim SQL                                 As String




    SAE = N2Str2Null(cboAttendingSE.Text)
    StartDateTime = N2Str2Null(FormatDateTime(dtDate.Value, vbShortDate) & " " & FormatDateTime(dtStartTime.Value, vbLongTime))
    EndDateTime = N2Str2Null(FormatDateTime(dtDate.Value, vbShortDate) & " " & FormatDateTime(dtStartTime.Value, vbLongTime))

    Model = N2Str2Null(txtModel)
    Make = N2Str2Null(txtMake)
    Color = N2Str2Null(cboColors)
    Class = N2Str2Null(txtClass)
    Year = N2Str2Null(txtYear)
    Terms = N2Str2Null(cboTerms)
    ExpectedPurchase = N2Str2Null(dtExpectedPurchase.Value)


    If AppointmentID <= 0 Then
        SQL = "INSERT INTO CRIS_SalesAppointments(ProspectID, SAE, StartDateTime, EndDateTime,  Model, Make, Color, Class, Year, Terms, ExpectedPurchase) " _
            & " VALUES(@ProspectID, @SAE, @StartDateTime, @EndDateTime, @Model, @Make, @Color, @Class, @Year, @Terms, @ExpectedPurchase)" & vbCrLf & "SELECT @@IDENTITY"
    Else

        SQL = " Update CRIS_SalesAppointments " _
            & " SET ProspectID=@ProspectID, SAE=@SAE, StartDateTime=@StartDateTime, EndDateTime=@EndDateTime,  Model=@Model, Make=@Make, Color=@Color, Class=@Class, Year=@Year, Terms=@Terms, ExpectedPurchase=@ExpectedPurchase " _
            & " WHERE AppointmentID=@AppointmentID "
    End If
    SQL = Replace(SQL, "@AppointmentID", AppointmentID)
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@SAE", SAE)
    SQL = Replace(SQL, "@StartDateTime", StartDateTime)
    SQL = Replace(SQL, "@EndDateTime", EndDateTime)
    SQL = Replace(SQL, "@Model", Model)
    SQL = Replace(SQL, "@Make", Make)
    SQL = Replace(SQL, "@Color", Color)
    SQL = Replace(SQL, "@Class", Class)
    SQL = Replace(SQL, "@Year", Year)
    SQL = Replace(SQL, "@Terms", Terms)
    SQL = Replace(SQL, "@ExpectedPurchase", ExpectedPurchase)



    Dim temprs                              As ADODB.Recordset

    Set temprs = gconDMIS.Execute(SQL)
    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogAppointment=" & StartDateTime & " where prospectid=" & ProspectID)

    If AppointmentID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Schedule Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 1
    End If

    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        AppointmentID = temprs.Collect(0)
    End If


    Set temprs = Nothing
    MainForm.ProspectStatus.ProspectID = ProspectID
End Sub

'Friend Sub EditSalesAppointment(xProspectID As Long, xAcctName As String, xTestDriveScheduleID As Long)
'    ProspectID = xProspectID
'    AcctName = xAcctName
'    TestDriveScheduleID = xTestDriveScheduleID
'End Sub

Private Sub cmdSelectViewVehicles_Click()
    If lvViewVehicles.Records Is Nothing Then: Exit Sub

    With lvViewVehicles.SelectedRows.Row(0)
        txtMake = Null2String(.Record(3).Value)
        txtClass = Null2String(.Record(4).Value)
        txtModel = Null2String(.Record(5).Value)
        txtYear = Null2String(.Record(6).Value)
        cboColors = Null2String(.Record(7).Value)
    End With
    ShowHidePictureBox picViewVehicles.hwnd, False, Me
End Sub

Private Sub cmdViewVStat_Click()
    Dim temprs                              As ADODB.Recordset
    lvViewVehicles.FilterText = vbNullString
    Set temprs = gconDMIS.Execute("SELECT ID, CODE, DESCRIPT, MAKE, CLASS, MODEL, YEER,  color, ignkey, prodno, serialno, vino, engineno FROM SMIS_MRRINV ")
    flex_FillReportView temprs, lvViewVehicles
    ShowHidePictureBox picViewVehicles.hwnd, True, Me
End Sub

Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart      As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function

'''END CALLS




Private Sub InitVars()

    dtStartTime.Value = DateFromString(FormatDateTime(Now, vbShortDate), "8:00:00 AM")
    dtEndTime.Value = DateFromString(FormatDateTime(Now, vbShortDate), "9:00:00 AM")
    Call FillCombo("SELECT DISTINCT Name from SMIS_vw_Srep  ORDER BY [name]", -1, 0, cboAttendingSE)
    Call FillCombo("Select DISTINCT 1, COLOR_DESC FROM ALL_COLOR ORDER BY COLOR_DESC", 0, 1, cboColors)
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

    With lvViewVehicles
        .Columns.Add 0, "ID", 0, True
        .Columns.Add 1, "Code", 50, True
        .Columns.Add 2, "Description", 100, True
        .Columns.Add 3, "Make", 100, True
        .Columns.Add 4, "Class", 100, True
        .Columns.Add 5, "Model", 100, True
        .Columns.Add 6, "Year", 100, True
        .Columns.Add 7, "Color", 100, True
        .Columns(0).Visible = False
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GroupRowTextBold = True         ' = vbWhite
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.CaptionFont.Bold = True
    End With



End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdSelectViewVehicles_Click
End Sub

Private Sub txtFitlerViewVehicles_Change()
    lvViewVehicles.FilterText = txtFitlerViewVehicles.Text
    lvViewVehicles.Populate
End Sub

