VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmOSMSProcessMonthEndProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Month-End Processing"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "MonthEndProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   5865
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   795
      Left            =   4980
      MouseIcon       =   "MonthEndProc.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "MonthEndProc.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      Height          =   795
      Left            =   4260
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "MonthEndProc.frx":08FA
      MousePointer    =   99  'Custom
      Picture         =   "MonthEndProc.frx":0A4C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Press F11 for Posting By Range"
      Top             =   660
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   1455
      Left            =   30
      ScaleHeight     =   1455
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "MonthEndProc.frx":0D71
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MonthEndProc.frx":0D8D
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   3
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "MonthEndProc.frx":0DA9
         End
      End
      Begin VB.Label labCPB 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmOSMSProcessMonthEndProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    If PROC_TYPE = "MONTH_END" Then
        If MsgQuestionBox("Close All Transactions, Are You Sure?", "Month End Processing") = True Then
            cmdPost.Enabled = False
            cmdExit.Enabled = False
            MonthEndUpdate
            cmdExit.Enabled = True
        End If
    End If
    If PROC_TYPE = "RANKING" Then
        If MsgQuestionBox("Generate Rank File, Are You Sure?", "Generate Rank File") = True Then
            cmdPost.Enabled = False
            cmdExit.Enabled = False
            GenRankFile
            cmdExit.Enabled = True
        End If
    End If
    If PROC_TYPE = "STKSTAT" Then
        If MsgQuestionBox("Generate Stock Status, Are You Sure?", "Generate Stock Status") = True Then
            cmdPost.Enabled = False
            cmdExit.Enabled = False
            CreateStockStatus
            cmdExit.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    If PROC_TYPE = "MONTH_END" Then Me.Caption = "Month-End Processing"
    If PROC_TYPE = "RANKING" Then Me.Caption = "Generate Rank File"
    If PROC_TYPE = "STKSTAT" Then Me.Caption = "Generate Stock Status"
    Screen.MousePointer = 0
End Sub

Sub MonthEndUpdate()
    On Error Resume Next
    Dim rsSupply, rsShipping As ADODB.Recordset

    Dim vSUPPLYID As Long
    Dim vSupplySupply_Code, vSupplySupply_Description As String
    Dim vSupplyOnHand As Long
    Dim vSupplyCost, vSupplyMad As Double
    Dim vSupplyOnOrder As Long
    Dim vSupplyInvClass As String
    Dim vSupplySStock As Long

    Dim i As Integer
    Screen.MousePointer = 11
    progCPB.Value = 0
    DoEvents
    MsgSpeech "Updating Supplies Master File"
    Me.Caption = "Updating Supplies Master File"
    labCPB.Caption = "Updating Supplies Master File... Please Wait..."
    DoEvents
    gconDMIS.Execute "update OSMS_SUPPLY set" & _
                   " lastm_oh = onhand," & _
                   " lastm_Cost = Cost," & _
                   " lastm_mad = Mad," & _
                   " lastm_oo = onorder," & _
                   " noship = noship + 1," & _
                   " mad = (Curr_Month + Prev_Month + Months_2 + Months_3 + Months_4 + Months_5) / 6 FROM OSMS_SHIPPING " & _
                   " where Curr_Month <= 0"
    progCPB.Value = 100
    DoEvents
    progCPB.Value = 0
    DoEvents
    gconDMIS.Execute "update OSMS_SUPPLY set" & _
                   " lastm_oh = onhand," & _
                   " lastm_Cost = Cost," & _
                   " lastm_mad = Mad," & _
                   " lastm_oo = onorder," & _
                   " noship = 0," & _
                   " mad = (Curr_Month + Prev_Month + Months_2 + Months_3 + Months_4 + Months_5) / 6 FROM OSMS_SHIPPING " & _
                   " where Curr_Month > 0"
    progCPB.Value = 100
    DoEvents
    Screen.MousePointer = 11
    progCPB.Value = 0
    DoEvents
    MsgSpeech "Updating Shipping File"
    Me.Caption = "Updating Shipping File"
    labCPB.Caption = "Updating Shipping File... Please Wait..."
    DoEvents
    gconDMIS.Execute "update OSMS_shipping set" & _
                   " months_60 = Months_59, months_59 = Months_58, months_58 = Months_57, months_57 = Months_56," & _
                   " months_56 = Months_55, months_55 = Months_54, months_54 = Months_53, months_53 = Months_52," & _
                   " months_52 = Months_51, months_51 = Months_50, months_50 = Months_49, months_49 = Months_48," & _
                   " months_48 = Months_47, months_47 = Months_46, months_46 = Months_45, months_45 = Months_44," & _
                   " months_44 = Months_43, months_43 = Months_42, months_42 = Months_41, months_41 = Months_40," & _
                   " months_40 = Months_39, months_39 = Months_38, months_38 = Months_37, months_37 = Months_36," & _
                   " months_36 = Months_35, months_35 = Months_34, months_34 = Months_33, months_33 = Months_32," & _
                   " months_32 = Months_31, months_31 = Months_30, months_30 = Months_29, months_29 = Months_28," & _
                   " months_28 = Months_27, months_27 = Months_26, months_26 = Months_25, months_25 = Months_24," & _
                   " months_24 = Months_23, months_23 = Months_22, months_22 = Months_21, months_21 = Months_20," & _
                   " months_20 = Months_19, months_19 = Months_18, months_18 = Months_17, months_17 = Months_16," & _
                   " months_16 = Months_15, months_15 = Months_14, months_14 = Months_13, months_13 = Months_12," & _
                   " months_12 = Months_11, months_11 = Months_10, months_10 = Months_9, months_9 = Months_8," & _
                   " months_8 = Months_7, months_7 = Months_6, months_6 = Months_5, months_5 = Months_4," & _
                   " months_4 = Months_3, months_3 = Months_2, months_2 = Prev_Month, prev_month = Curr_Month," & _
                   " curr_month = 0 "
    DoEvents
    progCPB.Value = 100
    Screen.MousePointer = 0
    Me.Caption = "Updating Complete!"
    labCPB.Caption = "Updating Complete!"
    'frmMain.mnuMonthEnd.Enabled = False

    MsgSpeechBox "Month End Processing Completed!"
End Sub

Sub GenRankFile()
    Dim rsSupply As ADODB.Recordset
    Dim rsShipping As ADODB.Recordset
    Dim i As Integer

    Dim SMonths_12 As Integer
    Dim SMonths_11 As Integer
    Dim SMonths_10 As Integer
    Dim SMonths_9 As Integer
    Dim SMonths_8 As Integer
    Dim SMonths_7 As Integer
    Dim SMonths_6 As Integer
    Dim SMonths_5 As Integer
    Dim SMonths_4 As Integer
    Dim SMonths_3 As Integer
    Dim SMonths_2 As Integer
    Dim SPrev_Month As Integer
    Dim vTotSales As Double
    Dim vMAD12 As Double
    Dim vRankType As String
    Dim vSubClass As String
    Dim vPrevClass As String
    Dim vPrevSClass As String
    Dim OldStock As Integer
    Dim S_year1 As Integer
    Dim S_year2 As Integer
    Dim S_year3 As Integer
    Dim S_year4 As Integer
    Dim S_year5 As Integer
    Dim P_Onhand As Integer
    Dim P_Cost As Double
    Dim P_lastrrdate As String
    Dim P_Supply_Description As String
    Set rsSupply = New ADODB.Recordset
    rsSupply.Open "select Supply_Code,Supply_Description,onhand,Cost,lastrrdate,invclass,subinvclas from OSMS_SUPPLY order by Supply_Code asc", gconDMIS
    If Not rsSupply.EOF And Not rsSupply.BOF Then
        rsSupply.MoveFirst
        MsgSpeech "Generating Rank File... This may take a while... Please wait..."
        Me.Caption = "Generating Rank File"
        DoEvents
        i = 0
        Do While Not rsSupply.EOF
            labProcessing.Caption = "Processing Supply Code: " & Null2String(rsSupply!Supply_Code)
            DoEvents
            SMonths_12 = 0: SMonths_11 = 0
            SMonths_10 = 0: SMonths_9 = 0
            SMonths_8 = 0: SMonths_7 = 0
            SMonths_6 = 0: SMonths_5 = 0
            SMonths_4 = 0: SMonths_3 = 0
            SMonths_2 = 0: SPrev_Month = 0
            vTotSales = 0: vMAD12 = 0
            S_year1 = 0: S_year2 = 0: S_year3 = 0: S_year4 = 0: S_year5 = 0
            OldStock = 0
            P_Onhand = N2Str2Zero(rsSupply!ONHAND)
            P_Cost = N2Str2Zero(rsSupply!Cost)
            P_lastrrdate = N2Date2Null(rsSupply!lastrrdate)
            P_Supply_Description = N2Str2Null(rsSupply!Supply_Description)
            vPrevClass = N2Str2Null(rsSupply!InvClass)
            vPrevSClass = N2Str2Null(rsSupply!SubInvClas)
            Set rsShipping = New ADODB.Recordset
            rsShipping.Open "Select * FROM OSMS_SHIPPING  where Supply_Code = " & N2Str2Null(rsSupply!Supply_Code), gconDMIS
            If Not rsShipping.EOF And Not rsShipping.BOF Then
                SMonths_12 = N2Str2Zero(rsShipping!Months_12)
                SMonths_11 = N2Str2Zero(rsShipping!Months_11)
                SMonths_10 = N2Str2Zero(rsShipping!Months_10)
                SMonths_9 = N2Str2Zero(rsShipping!Months_9)
                SMonths_8 = N2Str2Zero(rsShipping!Months_8)
                SMonths_7 = N2Str2Zero(rsShipping!Months_7)
                SMonths_6 = N2Str2Zero(rsShipping!Months_6)
                SMonths_5 = N2Str2Zero(rsShipping!Months_5)
                SMonths_4 = N2Str2Zero(rsShipping!Months_4)
                SMonths_3 = N2Str2Zero(rsShipping!Months_3)
                SMonths_2 = N2Str2Zero(rsShipping!Months_2)
                SPrev_Month = N2Str2Zero(rsShipping!Prev_Month)
                S_year1 = N2Str2Zero(rsShipping!Months_12) + N2Str2Zero(rsShipping!Months_11) + N2Str2Zero(rsShipping!Months_10) + N2Str2Zero(rsShipping!Months_9) + N2Str2Zero(rsShipping!Months_8) + N2Str2Zero(rsShipping!Months_7) + N2Str2Zero(rsShipping!Months_6) + N2Str2Zero(rsShipping!Months_5) + N2Str2Zero(rsShipping!Months_4) + N2Str2Zero(rsShipping!Months_3) + N2Str2Zero(rsShipping!Months_2) + N2Str2Zero(rsShipping!Prev_Month)
                S_year2 = N2Str2Zero(rsShipping!months_24) + N2Str2Zero(rsShipping!months_23) + N2Str2Zero(rsShipping!months_22) + N2Str2Zero(rsShipping!months_21) + N2Str2Zero(rsShipping!months_20) + N2Str2Zero(rsShipping!months_19) + N2Str2Zero(rsShipping!months_18) + N2Str2Zero(rsShipping!months_17) + N2Str2Zero(rsShipping!months_16) + N2Str2Zero(rsShipping!months_15) + N2Str2Zero(rsShipping!months_14) + N2Str2Zero(rsShipping!months_13)
                S_year3 = N2Str2Zero(rsShipping!months_36) + N2Str2Zero(rsShipping!months_35) + N2Str2Zero(rsShipping!months_34) + N2Str2Zero(rsShipping!months_33) + N2Str2Zero(rsShipping!months_32) + N2Str2Zero(rsShipping!months_31) + N2Str2Zero(rsShipping!months_30) + N2Str2Zero(rsShipping!months_29) + N2Str2Zero(rsShipping!months_28) + N2Str2Zero(rsShipping!months_27) + N2Str2Zero(rsShipping!months_26) + N2Str2Zero(rsShipping!months_25)
                S_year4 = N2Str2Zero(rsShipping!months_48) + N2Str2Zero(rsShipping!months_47) + N2Str2Zero(rsShipping!months_46) + N2Str2Zero(rsShipping!months_45) + N2Str2Zero(rsShipping!months_44) + N2Str2Zero(rsShipping!months_43) + N2Str2Zero(rsShipping!months_42) + N2Str2Zero(rsShipping!months_41) + N2Str2Zero(rsShipping!months_40) + N2Str2Zero(rsShipping!months_39) + N2Str2Zero(rsShipping!months_38) + N2Str2Zero(rsShipping!months_37)
                S_year5 = N2Str2Zero(rsShipping!months_60) + N2Str2Zero(rsShipping!months_59) + N2Str2Zero(rsShipping!months_58) + N2Str2Zero(rsShipping!months_57) + N2Str2Zero(rsShipping!Months_56) + N2Str2Zero(rsShipping!months_55) + N2Str2Zero(rsShipping!months_54) + N2Str2Zero(rsShipping!months_53) + N2Str2Zero(rsShipping!months_52) + N2Str2Zero(rsShipping!months_51) + N2Str2Zero(rsShipping!months_50) + N2Str2Zero(rsShipping!months_49)
                vTotSales = Format(S_year1, MAXIMUM_DIGIT)
                vMAD12 = Format(vTotSales / 12, MAXIMUM_DIGIT)
            End If

            If vTotSales < 99999 And vTotSales > 359 Then
                vRankType = "A": vSubClass = "1"
            ElseIf vTotSales < 360 And vTotSales > 239 Then
                vRankType = "A": vSubClass = "2 "
            ElseIf vTotSales < 240 And vTotSales > 119 Then
                vRankType = "A": vSubClass = "3"
            ElseIf vTotSales < 120 And vTotSales > 47 Then
                vRankType = "B": vSubClass = ""
            ElseIf vTotSales < 48 And vTotSales > 23 Then
                vRankType = "C": vSubClass = ""
            ElseIf vTotSales < 24 And vTotSales > 0 Then
                vRankType = "D": vSubClass = ""
            Else
                If IsNull(rsSupply!lastrrdate) = False Then
                    OldStock = Int((CDate(LOGDATE) - Null2Date(rsSupply!lastrrdate)) / 365)
                    If OldStock > 0 Then
                        vRankType = "E"
                        If OldStock >= 5 And S_year1 + S_year2 + S_year3 + S_year4 + S_year5 = 0 Then
                            vSubClass = "5"
                        ElseIf OldStock = 4 And S_year1 + S_year2 + S_year3 + S_year4 = 0 Then vSubClass = "4"
                        ElseIf OldStock = 3 And S_year1 + S_year2 + S_year3 = 0 Then vSubClass = "3"
                        ElseIf OldStock = 2 And S_year1 + S_year2 = 0 Then vSubClass = "2"
                        ElseIf OldStock = 1 Then vSubClass = "1"
                        Else
                            If S_year1 <> 0 Then
                                vSubClass = "1"
                            ElseIf S_year1 + S_year2 <> 0 Then vSubClass = "2"
                            ElseIf S_year1 + S_year2 + S_year3 <> 0 Then vSubClass = "3"
                            ElseIf S_year1 + S_year2 + S_year3 + S_year4 <> 0 Then vSubClass = "4"
                            ElseIf S_year1 + S_year2 + S_year3 + S_year4 + S_year5 <> 0 Then vSubClass = "5"
                            End If
                        End If
                    Else
                        vRankType = "F": vSubClass = ""
                    End If
                Else
                    vRankType = "E"
                    If S_year1 <> 0 Then
                        vSubClass = "1"
                    ElseIf S_year1 + S_year2 <> 0 Then vSubClass = "2"
                    ElseIf S_year1 + S_year2 + S_year3 <> 0 Then vSubClass = "3"
                    ElseIf S_year1 + S_year2 + S_year3 + S_year4 <> 0 Then vSubClass = "4"
                    ElseIf S_year1 + S_year2 + S_year3 + S_year4 + S_year5 <> 0 Then vSubClass = "5"
                    Else
                        If S_year1 + S_year2 + S_year3 + S_year4 + S_year5 = 0 Then vSubClass = "5"
                    End If
                End If
            End If
            gconDMIS.Execute "update OSMS_SUPPLY set " & _
                             "invclass = " & N2Str2Null(vRankType) & "," & _
                             "subinvclas = " & N2Str2Null(vSubClass) & "," & _
                             "mad = " & N2Str2Zero(vMAD12) & _
                           " where Supply_Code = " & N2Str2Null(rsSupply!Supply_Code)
            gconDMIS.Execute "insert into osms_rankfle " & _
                             "(Supply_Code,Supply_Description,invclass,subinvclas,onhand,mad12,sales12,lastrrdate,Cost,month_gen,prev_month,months_2,months_3,months_4,months_5,months_6,months_7,months_8,months_9,months_10,months_11,months_12,prevclass,prevsclas,date_gen)" & _
                           " values (" & N2Str2Null(rsSupply!Supply_Code) & ", " & P_Supply_Description & _
                             "," & N2Str2Null(vRankType) & ", " & N2Str2Null(vSubClass) & ", " & P_Onhand & _
                             "," & vMAD12 & ", " & vTotSales & ", " & P_lastrrdate & ", " & P_Cost & ", " & Month(LOGDATE) & ", " & SPrev_Month & _
                             "," & SMonths_2 & ", " & SMonths_3 & ", " & SMonths_4 & _
                             "," & SMonths_5 & ", " & SMonths_6 & ", " & SMonths_7 & _
                             "," & SMonths_8 & ", " & SMonths_9 & ", " & SMonths_10 & _
                             "," & SMonths_11 & ", " & SMonths_12 & ", " & vPrevClass & ", " & vPrevSClass & ", " & N2Date2Null(LOGDATE) & ")"
            DoEvents
            i = i + 1
            progCPB.Value = (i / rsSupply.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsSupply.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeechBox "Generate Rank File Complete!"
    Else
        MsgSpeechBox "Error opening Supplies Master File"
    End If
End Sub

Sub CreateStockStatus()
    Screen.MousePointer = 11
    progCPB.Value = 0
    Me.Caption = "Updating Supplies Master File"
    labCPB.Caption = "Updating Supplies Master File for Stock Status... Please Wait..."
    DoEvents
    progCPB.Value = 100
    gconDMIS.Execute "update OSMS_SUPPLY set" & _
                   " sstock = mad " & _
                   " where invclass='A'"
    DoEvents
    Screen.MousePointer = 11
    progCPB.Value = 0
    Me.Caption = "Creating Stock Status"
    labCPB.Caption = "Create Stock Status Master File... Please Wait..."
    DoEvents
    gconDMIS.Execute "delete from  OSMS_StkStat where date_gen = '" & LOGDATE & "'"
    progCPB.Value = 100
    gconDMIS.Execute "insert into OSMS_stkstat " & _
                     "(Supply_Code,Supply_Description,onhand,Cost,mad,sstock,onorder)" & _
                   " select Supply_Code,Supply_Description,OnHand,Cost,Mad,SStock,OnOrder from OSMS_SUPPLY order by Supply_Code asc"
    gconDMIS.Execute "update osms_stkstat set date_gen = " & N2Date2Null(LOGDATE) & " where date_gen IS NULL"
    'frmMain.mnuCreateStockStatus.Enabled = False

    MsgSpeechBox "Create Stock Status Complete!"
    Screen.MousePointer = 0
    DoEvents

End Sub
