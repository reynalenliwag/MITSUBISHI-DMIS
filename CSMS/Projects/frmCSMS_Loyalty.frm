VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmCSMS_Loyalty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loyalty file Generation"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Loyalty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   12675
   Begin FlexCell.Grid Grid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10610
      BackColor2      =   12648384
      Cols            =   5
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.PictureBox picAdds 
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
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   14625
      TabIndex        =   14
      Top             =   6630
      Width           =   14625
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1260
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Decrypt"
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmCSMS_Loyalty.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Generate Date"
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComDlg.CommonDialog CDB1 
         Left            =   8130
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   795
         Left            =   11880
         MouseIcon       =   "frmCSMS_Loyalty.frx":1542
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":1694
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Save"
         Height          =   795
         Left            =   11100
         MouseIcon       =   "frmCSMS_Loyalty.frx":19FA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":1B4C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   795
         Left            =   10320
         MouseIcon       =   "frmCSMS_Loyalty.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Generate Date"
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.PictureBox Picgenerate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   4095
      ScaleHeight     =   2805
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   2318
      Width           =   4425
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3660
         MouseIcon       =   "frmCSMS_Loyalty.frx":308E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":31E0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cancel"
         Top             =   1950
         Width           =   705
      End
      Begin VB.CommandButton cmd_process 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2940
         MouseIcon       =   "frmCSMS_Loyalty.frx":351E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Loyalty.frx":3670
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Process Data"
         Top             =   1950
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "O.R Date Range"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   4275
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   345
            Left            =   2340
            TabIndex        =   6
            Top             =   570
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            Format          =   95944705
            CurrentDate     =   39562
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   345
            Left            =   270
            TabIndex        =   7
            Top             =   570
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            Format          =   95944705
            CurrentDate     =   39562
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date:"
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
            Left            =   2310
            TabIndex        =   9
            Top             =   330
            Width           =   765
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date:"
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
            Left            =   270
            TabIndex        =   8
            Top             =   330
            Width           =   1125
         End
      End
      Begin wizProgBar.Prg Prgbr 
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         Picture         =   "frmCSMS_Loyalty.frx":46F2
         ForeColor       =   255
         BorderStyle     =   2
         BarPicture      =   "frmCSMS_Loyalty.frx":470E
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         Height          =   795
         Left            =   3660
         Picture         =   "frmCSMS_Loyalty.frx":472A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " "
         Top             =   1950
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   5715
         _Version        =   655364
         _ExtentX        =   10081
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   " Generate Data"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   8388608
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   1065
      Index           =   1
      Left            =   -600
      TabIndex        =   1
      Top             =   6570
      Width           =   15165
      _Version        =   655364
      _ExtentX        =   26749
      _ExtentY        =   1879
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _Version        =   655364
      _ExtentX        =   25638
      _ExtentY        =   1085
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmCSMS_Loyalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i                                                   As Integer
Dim CTR                                                 As Integer
Dim xp As Long, xi As Long, xw As Long, xt As Long

Sub ApplyFlexCellSetting(XXX As Grid)
    With XXX
        .DisplayFocusRect = False
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .BackColorFixed = RGB(90, 158, 214)
        .BackColorFixedSel = RGB(110, 180, 230)
        .BackColorBkg = RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)
    End With

End Sub

Sub InitGrid()
   
   With Grid1
        'ApplyFlexCellSetting Grid1
        .Cols = 28
        
        .RowHeight(0) = 40
        .Column(0).Width = 40
        .Column(1).Width = 50
        .Column(2).Width = 100
        .Column(3).Width = 100
        .Column(4).Width = 170
        .Column(5).Width = 100
        .Column(6).Width = 120
        .Column(7).Width = 70
        .Column(8).Width = 80
        .Column(9).Width = 50
        .Column(10).Width = 70
        .Column(11).Width = 80
        .Column(12).Width = 120
        .Column(13).Width = 110
        .Column(14).Width = 70
        
        .Column(15).Width = 100
        .Column(16).Width = 60
        .Column(17).Width = 90
        .Column(18).Width = 50
        .Column(19).Width = 50
        .Column(20).Width = 60
        .Column(21).Width = 60
        .Column(22).Width = 150
        .Column(23).Width = 220
        .Column(24).Width = 50
        .Column(25).Width = 90
        .Column(26).Width = 80
        .Column(27).Width = 80
        
        .Cell(0, 0).Text = "L/N"
        .Cell(0, 1).Text = "DEALER"
        .Cell(0, 2).Text = "TRAN TYPE"
        .Cell(0, 3).Text = "CARD NO"
        .Cell(0, 4).Text = "NAME"
        .Cell(0, 5).Text = "MODEL"
        .Cell(0, 6).Text = "VIN NO"
        .Cell(0, 7).Text = "KM READING"
        .Cell(0, 8).Text = "RO NO/ TRAN NO"
        .Cell(0, 9).Text = "OR NO"
        .Cell(0, 10).Text = "INVOICE NO"
        .Cell(0, 11).Text = "INVOICE DATE"
        .Cell(0, 12).Text = "SERVICE ADVISOR"
        .Cell(0, 13).Text = "TECHNICIAN"
        .Cell(0, 14).Text = "DATE RECORDED"
        .Cell(0, 15).Text = "PROMISE TIME"
        .Cell(0, 16).Text = "PAYMENT TYPE"
        .Cell(0, 17).Text = "QUOTED AMOUNT"
        .Cell(0, 18).Text = "DETAIL TYPE"
        .Cell(0, 19).Text = "SALE SUB TYPE"
        .Cell(0, 20).Text = "WARR. CODE"
        .Cell(0, 21).Text = "PMS READING"
        .Cell(0, 22).Text = "DETAIL CODE"
        .Cell(0, 23).Text = "DETAIL DESCRIPTION"
        .Cell(0, 24).Text = "LTS/QTY"
        .Cell(0, 25).Text = "UNIT AMOUNT"
        .Cell(0, 26).Text = "TOTAL AMOUNT"
        .Cell(0, 27).Text = "OR DATE"
         
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
        .Column(6).Locked = True
        .Column(7).Locked = True
        .Column(8).Locked = True
        .Column(9).Locked = True
        .Column(10).Locked = True
        .Column(11).Locked = True
        .Column(12).Locked = True
        .Column(13).Locked = True
        .Column(14).Locked = True
        .Column(15).Locked = True
        .Column(16).Locked = True
        .Column(17).Locked = True
        .Column(18).Locked = True
        .Column(19).Locked = True
        .Column(20).Locked = True
        .Column(21).Locked = True
        .Column(22).Locked = True
        .Column(23).Locked = True
        .Column(24).Locked = True
        .Column(25).Locked = True
        .Column(26).Locked = True
        .Column(27).Locked = True
        
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Alignment = cellCenterCenter
        .Column(4).Alignment = cellLeftGeneral
        .Column(5).Alignment = cellLeftGeneral
        .Column(6).Alignment = cellLeftGeneral
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Alignment = cellCenterCenter
        .Column(9).Alignment = cellCenterCenter
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Alignment = cellCenterCenter
        .Column(12).Alignment = cellLeftGeneral
        .Column(13).Alignment = cellLeftGeneral
        .Column(14).Alignment = cellCenterCenter
        .Column(15).Alignment = cellCenterCenter
        .Column(16).Alignment = cellCenterCenter
        .Column(17).Alignment = cellRightGeneral
        .Column(18).Alignment = cellCenterCenter
        .Column(19).Alignment = cellCenterCenter
        .Column(20).Alignment = cellCenterCenter
        .Column(21).Alignment = cellCenterCenter
        .Column(22).Alignment = cellCenterCenter
        .Column(23).Alignment = cellLeftGeneral
        .Column(24).Alignment = cellCenterCenter
        .Column(25).Alignment = cellRightGeneral
        .Column(26).Alignment = cellRightGeneral
        .Column(27).Alignment = cellCenterCenter
        
        
        .Range(0, 0, 0, 27).WrapText = True
    End With

End Sub

Sub FillGrid()
    DoEvents
    Screen.MousePointer = 11
    
    Dim rsEMPINFO2                                                    As New ADODB.Recordset
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command

    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_CSMS_LOYALTY"
    cmd.ActiveConnection = gconDMIS
    cmd.CommandTimeout = 1000
    cmd.Parameters.Append cmd.CreateParameter("@FROMDATE", adDBDate, adParamInput, , dtpFromDate.Value)
    cmd.Parameters.Append cmd.CreateParameter("@TODATE", adDBDate, adParamInput, , dtpToDate.Value)
    Set rsEMPINFO2 = cmd.Execute
    
    If Not rsEMPINFO2.EOF And Not rsEMPINFO2.BOF Then
        Grid1.Rows = 1
        CTR = 1
        Prgbr.Max = rsEMPINFO2.RecordCount
        Prgbr.Value = 0
        rsEMPINFO2.MoveFirst

        While Not rsEMPINFO2.EOF
            DoEvents
            
            Grid1.AddItem Null2String(rsEMPINFO2!dealer) & Chr(9) & _
            Null2String(rsEMPINFO2!TRAN_TYPE) & Chr(9) & _
            Null2String(rsEMPINFO2!Loyalty_ID) & Chr(9) & _
            Null2String(Replace(rsEMPINFO2!NIYM, vbCrLf, "")) & Chr(9) & _
            Replace(Null2String(rsEMPINFO2!Model), vbCrLf, "") & Chr(9) & _
            Replace(Null2String(rsEMPINFO2!Vin), vbCrLf, "") & Chr(9) & _
            Null2String(rsEMPINFO2!km_rdg) & Chr(9) & _
            Null2String(rsEMPINFO2!REP_OR) & Chr(9) & _
            Null2String(rsEMPINFO2!OR_NUM) & Chr(9) & _
            Null2String(rsEMPINFO2!INVOICE) & Chr(9) & Null2String(rsEMPINFO2!dte_comp) & Chr(9) & _
            Replace(Null2String(rsEMPINFO2!NAYM), vbCrLf, "") & Chr(9) & Null2String(rsEMPINFO2!Technician) & Chr(9) & _
            Null2String(rsEMPINFO2!DTE_RECD) & Chr(9) & Null2String(rsEMPINFO2!DTE_PRO) & Chr(9) & _
            Null2String(rsEMPINFO2!TERM) & Chr(9) & Format(N2Str2Zero(rsEMPINFO2!amount), "#,###,##0.00") & Chr(9) & _
            Null2String(rsEMPINFO2!DET_TYPE) & Chr(9) & Null2String(rsEMPINFO2!S_TRAN_ST) & Chr(9) & _
            Null2String(rsEMPINFO2!wCode) & Chr(9) & NumericVal(N2Str2Zero(rsEMPINFO2!PMS_READING)) & Chr(9) & _
            Null2String(rsEMPINFO2!DETCDE) & Chr(9) & Null2String(Replace(Null2String(rsEMPINFO2!DETDSC), vbCrLf, "")) & Chr(9) & _
            Null2String(rsEMPINFO2!LTS_QTY) & Chr(9) & Format(N2Str2Zero(rsEMPINFO2!DetPrc), "#,###,##0.00") & Chr(9) & _
            Format(N2Str2Zero(rsEMPINFO2!DET_AMT), "#,###,##0.00") & Chr(9) & DateValue(Null2String(rsEMPINFO2!ORDATE)), False
            
            
            rsEMPINFO2.MoveNext
            
            Prgbr.Value = Prgbr.Value + 1
            'lblPercent.Caption = Round((Prgbr.Value / Prgbr.Max) * 100, 0) & "%"
            Prgbr.Text = "Generating Data " & " (" & Round((Prgbr.Value / Prgbr.Max) * 100, 0) & " %)"
            DoEvents
            
          Wend
        Grid1.Refresh
        MessagePop InfoFriend, "Info", "Generation Complete"
    Else
        Call ShowNoRecord
        Grid1.Rows = 1
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Set rsEMPINFO2 = Nothing
    Prgbr.Text = "Generating Data Complete (100%)"
    Screen.MousePointer = 0
End Sub

Private Sub cmd_cancel_Click()
    picAdds.Enabled = True
    Grid1.Enabled = True
    Picgenerate.Visible = False
    Picgenerate.ZOrder 1
End Sub

Private Sub cmd_process_Click()
    If MsgBox(" Do you want to proceed", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    Call FillGrid
    Call txtnonvisible
    cmdDone.Visible = True
    Call cmd_cancel_Click
End Sub

Private Sub cmdDone_Click()
    Picgenerate.Visible = False
    Picgenerate.ZOrder 1
    picAdds.Enabled = True
    Grid1.Enabled = True
End Sub

Private Sub cmdExit_Click()
    picAdds.Enabled = False
    Grid1.Enabled = False
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    Picgenerate.Visible = True
    Picgenerate.ZOrder 0
    dtpFromDate.Value = Date - 1
    dtpToDate.Value = Date - 1
    Call txtvisible
    Prgbr.Text = ""
End Sub

Sub txtvisible()
    cmd_process.Visible = True
    cmd_cancel.Visible = True
    picAdds.Enabled = False
    Grid1.Enabled = False
    Prgbr.Value = 0
End Sub

Sub txtnonvisible()
    cmd_process.Visible = False
    cmd_cancel.Visible = False
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo A:
    Dim Itmss                                           As String
    Dim i                                               As Integer
    Dim strFileToSave                                   As String
    Dim filenaym                                        As String
    If Grid1.Rows <> 1 Then
        filenaym = COMPANY_CODE & "_" & Replace(DateValue(Date), "/", "") & " " & Replace(TimeValue(Now), ":", "")
    
       With CDB1
            .CancelError = True
            .DialogTitle = "Save List As"
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
            .FileName = filenaym
            .Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
            .ShowSave
            strFileToSave = .FileName
        End With
        Close #1
        
        Open strFileToSave For Output As #1
            If strFileToSave = "" Then Exit Sub
            For i = 1 To Grid1.Rows - 1

                Itmss = xEncrypt(Grid1.Cell(i, 1).Text + "|" + Grid1.Cell(i, 2).Text _
                + "|" + Grid1.Cell(i, 3).Text + "|" + Grid1.Cell(i, 4).Text _
                + "|" + Grid1.Cell(i, 5).Text + "|" + Grid1.Cell(i, 6).Text _
                + "|" + Grid1.Cell(i, 7).Text + "|" + Grid1.Cell(i, 8).Text _
                + "|" + Grid1.Cell(i, 9).Text + "|" + Grid1.Cell(i, 10).Text _
                + "|" + Grid1.Cell(i, 11).Text + "|" + Grid1.Cell(i, 12).Text _
                + "|" + Grid1.Cell(i, 13).Text + "|" + Grid1.Cell(i, 14).Text _
                + "|" + Grid1.Cell(i, 15).Text + "|" + Grid1.Cell(i, 16).Text _
                + "|" + Grid1.Cell(i, 17).Text + "|" + Grid1.Cell(i, 18).Text _
                + "|" + Grid1.Cell(i, 19).Text + "|" + Grid1.Cell(i, 20).Text _
                + "|" + Grid1.Cell(i, 21).Text + "|" + Grid1.Cell(i, 22).Text _
                + "|" + Grid1.Cell(i, 23).Text + "|" + Grid1.Cell(i, 24).Text _
                + "|" + Grid1.Cell(i, 25).Text + "|" + Grid1.Cell(i, 26).Text _
                + "|" + Grid1.Cell(i, 27).Text)


'                Itmss = Grid1.Cell(i, 1).Text + "|" + Grid1.Cell(i, 2).Text _
'                + "|" + Grid1.Cell(i, 3).Text + "|" + Grid1.Cell(i, 4).Text _
'                + "|" + Grid1.Cell(i, 5).Text + "|" + Grid1.Cell(i, 6).Text _
'                + "|" + Grid1.Cell(i, 7).Text + "|" + Grid1.Cell(i, 8).Text _
'                + "|" + Grid1.Cell(i, 9).Text + "|" + Grid1.Cell(i, 10).Text _
'                + "|" + Grid1.Cell(i, 11).Text + "|" + Grid1.Cell(i, 12).Text _
'                + "|" + Grid1.Cell(i, 13).Text + "|" + Grid1.Cell(i, 14).Text _
'                + "|" + Grid1.Cell(i, 15).Text + "|" + Grid1.Cell(i, 16).Text _
'                + "|" + Grid1.Cell(i, 17).Text + "|" + Grid1.Cell(i, 18).Text _
'                + "|" + Grid1.Cell(i, 19).Text + "|" + Grid1.Cell(i, 20).Text _
'                + "|" + Grid1.Cell(i, 21).Text + "|" + Grid1.Cell(i, 22).Text _
'                + "|" + Grid1.Cell(i, 23).Text + "|" + Grid1.Cell(i, 24).Text _
'                + "|" + Grid1.Cell(i, 25).Text + "|" + Grid1.Cell(i, 26).Text _
'                + "|" + Grid1.Cell(i, 27).Text

                Print #1, Itmss
            Next i
            MsgBox "Text File Successfully created", vbInformation + vbOKOnly, "Info"
        Close #1
    Else
        MessagePop InfoFriend, "Info", "No record to save"
    End If

    Exit Sub
A:
    MsgBox Err.Description, vbInformation, "Info"
    Err.Clear
End Sub

Private Sub Command1_Click()
    'Call InitGrid
    Grid1.Rows = 1
    
    Dim objfso As Object
    Dim strdata
    Dim strtextfile
    Dim arrlines
    
'    With CommonDialog1
'        .CancelError = True
'
'        .DialogTitle = "Open Text File"
'        .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
'        .Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
'        .ShowOpen
'    End With
'
'    Set objfso = CreateObject("scripting.filesystemobject")
'    strtextfile = "C:\Documents and Settings\NETSPEED8\Desktop\HAI_10262010 100259 AM.txt"
'    strdata = objfso.OpenTextFile(strtextfile, ForReading).ReadAll
'
'
'    arrlines = Split(strdata, vbCrLf)
'
'    For Each strLine In arrlines
'        Debug.Print strLine
'    Next
'
'    Set objfso = Nothing
'    fileopen
' Get a free file number
'    Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
'    nFileNum = FreeFile
'
'    ' Open a text file for input. inputbox returns the path to read the file
'    Open "C:\Documents and Settings\NETSPEED8\Desktop\encryted.txt" For Input As nFileNum
'    lLineCount = 1
'    ' Read the contents of the file
'    Do While Not EOF(nFileNum)
'       Line Input #nFileNum, sNextLine
'       'do something with it
'       'add line numbers to it, in this case!
'       sNextLine = sNextLine & vbCrLf
'       Debug.Print sNextLine
'       sText = sText & sNextLine
'
'    Loop
'    'Text1.Text = sText
'
'    ' Close the file
'    Close nFileNum

    Dim fs As FileSystemObject
    Dim TS As TextStream
    Dim sample As String
   
    Set fs = New FileSystemObject
      'To write
    Set TS = fs.OpenTextFile("C:\Documents and Settings\NETSPEED8\Desktop\new encr.txt", ForReading, False)
    'ts.WriteLine "I Love"
    'ts.WriteLine "VB Forums"
    'ts.Close
     
'      'To Read
    If fs.FileExists("C:\Documents and Settings\NETSPEED8\Desktop\new encr.txt") Then
        Set TS = fs.OpenTextFile("C:\Documents and Settings\NETSPEED8\Desktop\new encr.txt")

        Do While Not TS.AtEndOfStream
            'Debug.Print TS.ReadLine
            sample = xDecrypt(TS.ReadLine)
            Debug.Print sample
            'Debug.Print Encrypt(TS.ReadLine)
            
        Loop
        TS.Close
    End If
       'clear memory used by FSO objects

    Set TS = Nothing
    Set fs = Nothing
'    Dim fso As New FileSystemObject
'    Dim ts As TextStream
'    Dim strOutput As String
'    Set ts = fso.OpenTextFile("C:\Documents and Settings\NETSPEED8\Desktop\eee.txt")
'    Do Until ts.AtEndOfStream
'        Debug.Print ts.ReadLine
'     strOutput = strOutput + ts.ReadLine
'    Loop
'
'    ts.Close
'    ReadTextFile = strOutput
    
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    InitGrid
    Grid1.Rows = 1
    CTR = 0
End Sub

Public Function Encrypt(STR As String) As String
    Dim i As Integer
    Dim intKeyChar As Integer
    Dim strTemp As String
    Dim strText As String
    Dim strKey As String
    Dim strChar1 As String * 1
    Dim strChar2 As String * 1

    strText = STR
    strKey = "LOYALTY"
    For i = 1 To Len(strText)
        strChar1 = Mid(strText, i, 1)
        intKeyChar = ((i - 1) Mod Len(strKey)) + 1
        strChar2 = Mid(strKey, intKeyChar, 1)
        strTemp = strTemp & Chr(Asc(strChar1) Xor Asc(strChar2))
    Next i
    Encrypt = strTemp
End Function

Public Function xEncrypt(ByVal icText As String) As String
    Dim icLen As Integer
    Dim icNewText As String
    Dim icChar  As String
    icChar = ""
    icLen = Len(icText)
    For i = 1 To icLen
        icChar = Mid(icText, i, 1)
        Select Case Asc(icChar)
            Case 65 To 90
                icChar = Chr(Asc(icChar) + 127)
            Case 97 To 122
                icChar = Chr(Asc(icChar) + 121)
            Case 48 To 57
                icChar = Chr(Asc(icChar) + 196)
            Case 32
                icChar = Chr(32)
        End Select
        icNewText = icNewText + icChar
    Next
    xEncrypt = icNewText
End Function

Public Function xDecrypt(ByVal icText As String) As String
    Dim icLen As Integer
    Dim icNewText As String
    Dim icChar As String
    icChar = ""
    icLen = Len(icText)
    For i = 1 To icLen
        icChar = Mid(icText, i, 1)
        Select Case Asc(icChar)
            Case 192 To 217
                icChar = Chr(Asc(icChar) - 127)
            Case 218 To 243
                icChar = Chr(Asc(icChar) - 121)
            Case 244 To 253
                icChar = Chr(Asc(icChar) - 196)
            Case 32
                icChar = Chr(32)
        End Select
        icNewText = icNewText + icChar
    Next
    xDecrypt = icNewText
End Function
