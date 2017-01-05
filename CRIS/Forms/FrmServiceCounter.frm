VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "WIZFLEX.OCX"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "WIZFLEXCRACKER.OCX"
Begin VB.Form frmCRIS_Inquiry_ServiceAppointment 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Counter: Activity"
   ClientHeight    =   6630
   ClientLeft      =   -375
   ClientTop       =   435
   ClientWidth     =   11910
   FillColor       =   &H00808080&
   ForeColor       =   &H00F5F5F5&
   Icon            =   "FrmServiceCounter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11910
   Begin VB.ComboBox Combo1 
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
      Left            =   4140
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   420
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2700
      ScaleHeight     =   315
      ScaleWidth      =   7335
      TabIndex        =   9
      Top             =   120
      Width           =   7335
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Legend:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   750
         Top             =   30
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   2640
         Top             =   30
         Width           =   135
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4620
         Top             =   30
         Width           =   135
      End
      Begin VB.Label labPark 
         BackStyle       =   0  'Transparent
         Caption         =   "- Park"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   930
         MouseIcon       =   "FrmServiceCounter.frx":05CA
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   0
         Width           =   525
      End
      Begin VB.Label labWork 
         BackStyle       =   0  'Transparent
         Caption         =   "- Working"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2820
         MouseIcon       =   "FrmServiceCounter.frx":08D4
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labFinish 
         BackStyle       =   0  'Transparent
         Caption         =   "- Finish"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4800
         MouseIcon       =   "FrmServiceCounter.frx":0BDE
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   0
         Width           =   645
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5520
         Top             =   30
         Width           =   135
      End
      Begin VB.Label labBilled 
         BackStyle       =   0  'Transparent
         Caption         =   "- Billed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5700
         MouseIcon       =   "FrmServiceCounter.frx":0EE8
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
      Begin VB.Label labOver 
         BackStyle       =   0  'Transparent
         Caption         =   "- Over"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3960
         MouseIcon       =   "FrmServiceCounter.frx":11F2
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   0
         Width           =   555
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3780
         Top             =   30
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "- Release"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6540
         MouseIcon       =   "FrmServiceCounter.frx":14FC
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   0
         Width           =   825
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6360
         Top             =   30
         Width           =   135
      End
      Begin VB.Label labBackJob 
         BackStyle       =   0  'Transparent
         Caption         =   "- Back Job"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         MouseIcon       =   "FrmServiceCounter.frx":1806
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   0
         Width           =   945
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1500
         Top             =   30
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdDateBack 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdDateForward 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   60
      Width           =   1245
   End
   Begin FlexCell.Grid grdCounter 
      Height          =   5775
      Left            =   2700
      TabIndex        =   6
      Top             =   780
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   10186
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   30
   End
   Begin VB.Frame frmJobs 
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   2280
      TabIndex        =   3
      Top             =   7950
      Width           =   13095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1260
      TabIndex        =   0
      Top             =   3060
      Width           =   1305
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2460
      Left            =   60
      TabIndex        =   4
      Top             =   435
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4339
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmServiceCounter.frx":1B10
      MonthBackColor  =   16777215
      StartOfWeek     =   51052545
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   13932144
      CurrentDate     =   38458
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   3720
      TabIndex        =   5
      Top             =   2700
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2700
      TabIndex        =   19
      Top             =   480
      Width           =   1935
   End
   Begin MSForms.Label Label13 
      Height          =   570
      Left            =   60
      TabIndex        =   2
      Top             =   -180
      Width           =   615
      ForeColor       =   192
      VariousPropertyBits=   8388627
      PicturePosition =   262148
      Size            =   "1085;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Counter"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   4410
      TabIndex        =   1
      Top             =   90
      Width           =   2625
   End
End
Attribute VB_Name = "frmCRIS_Inquiry_ServiceAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thestatus                                          As String
Dim theRo                                              As String
Dim Thedate                                            As Date
Attribute Thedate.VB_VarUserMemId = 1073938435
Dim ChkStatus                                          As String
Attribute ChkStatus.VB_VarUserMemId = 1073938439
Dim zRONO                                              As String
Attribute zRONO.VB_VarUserMemId = 1073938440
Private Sub cmdDateBack_Click()
    MonthView1 = MonthView1 - 1
    cmdRefresh.Value = True
End Sub
Private Sub cmdDateForward_Click()
    MonthView1 = MonthView1 + 1
    cmdRefresh.Value = True
End Sub
Private Sub cmdRefresh_Click()
    cmdDateBack.Caption = " " & Format(MonthView1, "MMM") & " " & Format(MonthView1 - 1, "dd")
    cmdDateForward.Caption = " " & Format(MonthView1, "MMM") & " " & Format(MonthView1 + 1, "dd")
    ViewAppointment
End Sub

Sub ViewAppointment()
    Dim rsUpload                                       As ADODB.Recordset

    Set rsUpload = New ADODB.Recordset

    If ChkStatus <> "" Then
        If UCase(Combo1.Text) = "ALL" Then
            Set rsUpload = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate='" & CDate(MonthView1) & "' AND STATUS='" & ChkStatus & "' order by RO_No asc")
        Else
            Set rsUpload = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate='" & CDate(MonthView1) & "' AND STATUS='" & ChkStatus & "' AND CUSTOMER ='" & Combo1 & "'   order by RO_No asc")
        End If
    Else
        If UCase(Combo1.Text) = "ALL" Then
            Set rsUpload = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate='" & CDate(MonthView1) & "' order by RO_No asc")
        Else
            Set rsUpload = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate='" & CDate(MonthView1) & "' AND CUSTOMER ='" & Combo1 & "' order by RO_No asc")

        End If
    End If

    grdCounter.Rows = 1
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Dim pecnt                                      As Double
        Do While Not rsUpload.EOF
            If rsUpload![xHrsWork] > 0 And rsUpload![Hours] > 0 Then
                pecnt = Format(Round(((rsUpload![xHrsWork] / rsUpload![Hours]) * 100), 2), "0000.#0")
            Else
                pecnt = 0
            End If
            grdCounter.AddItem Format(rsUpload![AppointmentDate], "MM/dd/yyyy") & vbTab & _
                               rsUpload![Customer] & vbTab & _
                               rsUpload![Model] & vbTab & _
                               rsUpload![PLATE_NO] & vbTab & _
                               rsUpload![ro_no] & vbTab & _
                               rsUpload![Hours] & vbTab & _
                               rsUpload![xHrsWork] & vbTab & _
                               pecnt & vbTab & _
                               rsUpload![Status] & vbTab & _
                               rsUpload![promisedate] & vbTab & _
                               rsUpload![Today] & vbTab & _
                               rsUpload![writer] & vbTab & _
                               rsUpload![datefinish] & vbTab & _
                               rsUpload![tech2] & vbTab & _
                               rsUpload![tech3] & vbTab & _
                               rsUpload![ACCT_NO], False
            rsUpload.MoveNext
        Loop
    Else
        '.Rows = 1
    End If

    On Error Resume Next
    Dim xx                                             As Long
    If (grdCounter.Rows - 1) > 0 Then
        For xx = 1 To grdCounter.Rows - 1
            If Trim(grdCounter.Cell(xx, 9).Text) = "Park" Then    'Park
                grdCounter.Range(xx, 1, xx, 12).Selected
                grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                'grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &H0&
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Working" Then
                grdCounter.Range(xx, 1, xx, 12).Selected
                ' grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &HC0C000
                If (CDate(grdCounter.Cell(xx, 11).Text) > CDate(grdCounter.Cell(xx, 10).Text)) Then
                    grdCounter.Range(xx, 1, xx, 12).Selected
                    grdCounter.Range(xx, 1, xx, 12).ForeColor = vbRed
                End If
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Over" Then
                grdCounter.Range(xx, 1, xx, 12).Selected
                '  grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = vbRed
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Finish Job" Then
                grdCounter.Range(xx, 1, xx, 12).Selected
                '  grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = vbBlue
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Billed" Then
                grdCounter.Range(xx, 1, xx, 11).Selected
                ' grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &H800080
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Released" Then
                grdCounter.Range(xx, 1, xx, 12).Selected
                ' grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &H8000&
                'BTT - 06052007
            ElseIf Trim(grdCounter.Cell(xx, 9).Text) = "Back Job" Then
                grdCounter.Range(xx, 1, xx, 12).Selected
                ' grdCounter.Range(xx, 1, xx, 12).BackColor = &HFBFFFF
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &H80FF&
            Else
                grdCounter.Range(xx, 1, xx, 12).Selected    ' same with working
                grdCounter.Range(xx, 1, xx, 12).FontBold = True
                grdCounter.Range(xx, 1, xx, 12).ForeColor = &HC0C000
                If (CDate(grdCounter.Cell(xx, 10).Text) > CDate(grdCounter.Cell(xx, 9).Text)) Then
                    '   grdCounter.Range(xx, 1, xx, 12).Selected
                    '   grdCounter.Range(xx, 1, xx, 12).ForeColor = vbRed
                End If
            End If
            grdCounter.Cell(xx, 1).SetFocus
        Next
    End If
    ChkStatus = ""

End Sub


Private Sub Combo1_Change()
    cmdRefresh.Value = True
End Sub

Private Sub Combo1_Click()
    Combo1_Change
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    MonthView1.Value = Format(Now, "MM/dd/yyyy")
    InitGrid
    cmdRefresh.Value = True
    SetComboMaxLength Combo1, 600
    Combo_Loadval Combo1, gconDMIS.Execute("Select DISTINCT Customer From CSMS_vw_RepairOrder")
    cmdRefresh_Click
    Combo1.AddItem "ALL"
End Sub


Private Sub labBackJob_Click()
    ChkStatus = "Back-Job"
    ViewAppointment
End Sub

Private Sub labBackJob_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labBackJob.FontBold = True
    labBackJob.ForeColor = &HFF0000
    Shape7.BorderColor = &HFFFF&
End Sub



Private Sub grdCounter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        'BTT - 05232007 (DISABLE SHOW OF MENU IF ALL JOBS HAS BEEN FINISH)
        'theStatus = grdCounter.Cell(grdCounter.ActiveCell.Row, 9).Text
        Dim Test, test1                                As String
        Test = "Billed"
        If StrComp(Trim(thestatus), Test) = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then

        Else

        End If
        'CheckIfAllJobFinish
    End If
End Sub





Private Sub Label1_Click()
    ChkStatus = "Released"
    ViewAppointment
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'lst1.Visible = True
End Sub

Private Sub labPark_Click()
    ChkStatus = "Park"
    ViewAppointment
End Sub
Private Sub labWork_Click()
    ChkStatus = "Working"
    ViewAppointment
End Sub
Private Sub labOver_Click()
    ChkStatus = "Over"
    ViewAppointment
End Sub
Private Sub labFinish_Click()
    ChkStatus = "Finish"
    ViewAppointment
End Sub
Private Sub labBilled_Click()
    ChkStatus = "Billed"
    ViewAppointment
End Sub
Private Sub labPark_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labPark.FontBold = True
    labPark.ForeColor = &HFF0000
    Shape1.BorderColor = &HFFFF&
End Sub
Private Sub labWork_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labWork.FontBold = True
    labWork.ForeColor = &HFF0000
    Shape2.BorderColor = &HFFFF&
End Sub
Private Sub labOver_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labOver.FontBold = True
    labOver.ForeColor = &HFF0000
    Shape6.BorderColor = &HFFFF&
End Sub
Private Sub labFinish_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labFinish.FontBold = True
    labFinish.ForeColor = &HFF0000
    Shape4.BorderColor = &HFFFF&
End Sub
Private Sub labBilled_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labBilled.FontBold = True
    labBilled.ForeColor = &HFF0000
    Shape5.BorderColor = &HFFFF&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    labBilled.FontBold = False
    labBilled.ForeColor = &H0&
    labFinish.FontBold = False
    labFinish.ForeColor = &H0&
    labOver.FontBold = False
    labOver.ForeColor = &H0&
    labWork.FontBold = False
    labWork.ForeColor = &H0&
    labPark.FontBold = False
    labPark.ForeColor = &H0&
    Shape1.BorderColor = &H0&
    Shape2.BorderColor = &H0&
    Shape4.BorderColor = &H0&
    Shape5.BorderColor = &H0&
    Shape6.BorderColor = &H0&

    labBackJob.FontBold = False
    labBackJob.ForeColor = &H0&
    Shape7.BorderColor = &H0&

End Sub




Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Thedate = Format(Now, "MM/dd/yyyy")
    cmdRefresh.Value = True
End Sub

Sub InitGrid()
    With grdCounter
        .Cols = 17: .Rows = 2
        .DisplayFocusRect = False
        .AllowUserResizing = True
        .BackColorFixed = &HFFCFB5                    'RGB(90, 158, 214)
        .BackColorFixedSel = &H8000000F               ' RGB(110, 180, 230) 'YELLOW
        .BackColorBkg = &HF9EFE3                      'RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Date"
        .Cell(0, 2).Text = "Customer"
        .Cell(0, 3).Text = "Vehicle"
        .Cell(0, 4).Text = "Plate No."
        .Cell(0, 5).Text = "R/O"
        .Cell(0, 6).Text = "Std.Hrs"
        .Cell(0, 7).Text = "Hrs.Work"
        .Cell(0, 8).Text = "(%)"
        .Cell(0, 9).Text = "Status"
        .Cell(0, 10).Text = "Promise"
        .Cell(0, 11).Text = "TODAY"
        .Cell(0, 12).Text = "Service Adviser"
        .Cell(0, 13).Text = "Date Finish"
        '.Cell(0, 13).Text = "Tech-1"
        .Cell(0, 14).Text = "Tech-2"
        .Cell(0, 15).Text = "Tech-3"
        .Cell(0, 16).Text = "Account No"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:    '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:    '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox
        .Column(10).CellType = cellTextBox
        .Column(11).CellType = cellTextBox
        .Column(12).CellType = cellTextBox
        .Column(13).CellType = cellTextBox
        .Column(14).CellType = cellTextBox
        .Column(15).CellType = cellTextBox
        .Column(16).CellType = cellTextBox

        .Column(0).Width = 18
        .Column(1).Width = 60: .Column(1).Locked = True
        .Column(2).Width = 150: .Column(2).Locked = True
        .Column(3).Width = 120: .Column(3).Locked = True
        .Column(4).Width = 60: .Column(4).Locked = True
        .Column(5).Width = 65: .Column(5).Locked = True
        .Column(6).Width = 50: .Column(6).Locked = True
        .Column(7).Width = 60: .Column(7).Locked = True
        .Column(8).Width = 50: .Column(8).Locked = True
        .Column(9).Width = 60: .Column(9).Locked = True
        .Column(10).Width = 125: .Column(10).Locked = True
        .Column(11).Width = 125: .Column(11).Locked = True
        .Column(12).Width = 100: .Column(12).Locked = True
        .Column(13).Width = 100: .Column(13).Locked = True
        .Column(14).Width = 0: .Column(14).Locked = True
        .Column(15).Width = 0: .Column(15).Locked = True
        .Column(16).Width = 0: .Column(16).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 16, .Rows - 1, 16).ForeColor = RGB(0, 0, 128)
    End With
End Sub
