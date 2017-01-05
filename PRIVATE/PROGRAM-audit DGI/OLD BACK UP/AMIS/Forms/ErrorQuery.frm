VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAMISErrorQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR FILES..."
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ErrorQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11700
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   180
      Top             =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grdDuplicat 
      Height          =   3435
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   8
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNo_Mstr 
      Height          =   3435
      Left            =   90
      TabIndex        =   2
      Top             =   4350
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   7
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNoHeader 
      Height          =   3435
      Left            =   5910
      TabIndex        =   1
      Top             =   420
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   7
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNoDetail 
      Height          =   3435
      Left            =   5910
      TabIndex        =   3
      Top             =   4350
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6059
      _Version        =   393216
      Cols            =   7
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin VB.Label labNoDetail 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "NO DETAIL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5910
      TabIndex        =   8
      Top             =   3960
      Width           =   5715
   End
   Begin VB.Label labNoHeader 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "NO HEADER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5910
      TabIndex        =   7
      Top             =   60
      Width           =   5715
   End
   Begin VB.Label labNo_Mstr 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "NO ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   3960
      Width           =   5715
   End
   Begin VB.Label labDUPFILES 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATE FILE(S)"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   5715
   End
   Begin VB.Label labAydi 
      Caption         =   "Label1"
      Height          =   195
      Left            =   11520
      TabIndex        =   4
      Top             =   5310
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "frmAMISErrorQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDuplicat                                    As ADODB.Recordset
Dim rsNO_Mstr                                     As ADODB.Recordset
Dim rsNOHeader                                    As ADODB.Recordset
Dim rsNODetail                                    As ADODB.Recordset
Attribute rsNODetail.VB_VarUserMemId = 1073938434

Dim kcnt                                          As Integer
Attribute kcnt.VB_VarUserMemId = 1073938437

Sub RefreshDisplay()
    cleargrd
    Set rsDuplicat = New ADODB.Recordset
    rsDuplicat.Open "Select * from AMIS_Duplicat", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNO_Mstr = New ADODB.Recordset
    rsNO_Mstr.Open "Select * from AMIS_No_Mstr", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNOHeader = New ADODB.Recordset
    rsNOHeader.Open "Select * from AMIS_NoHeader", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNODetail = New ADODB.Recordset
    rsNODetail.Open "Select * from AMIS_NoDetail", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshDup()
    cleargrd
    Set rsDuplicat = New ADODB.Recordset
    rsDuplicat.Open "Select * from AMIS_Duplicat", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNO_Mstr()
    cleargrd
    Set rsNO_Mstr = New ADODB.Recordset
    rsNO_Mstr.Open "Select * from AMIS_No_Mstr", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNoHead()
    cleargrd
    Set rsNOHeader = New ADODB.Recordset
    rsNOHeader.Open "Select * from AMIS_NoHeader", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNoDet()
    cleargrd
    Set rsNODetail = New ADODB.Recordset
    rsNODetail.Open "Select * from AMIS_NoDetail", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub cleargrd()
    cleargrid grdDuplicat
    cleargrid grdNo_Mstr
    cleargrid grdNoHeader
    cleargrid grdNoDetail
End Sub

Sub InitGrid()
    With grdDuplicat
        .Row = 0
        .FormatString = "TYPE | Tran Type | Tranno | File                          | " & _
                        "Rec No. 1 | Rec No. 2 | Status | ID"
        .ColWidth(6) = 1
    End With
    With grdNo_Mstr
        .Row = 0
        .FormatString = "TYPE | Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
    With grdNoHeader
        .Row = 0
        .FormatString = "TYPE | Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
    With grdNoDetail
        .Row = 0
        .FormatString = "TYPE | Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
End Sub

Sub FillGrid()
'On Error GoTo Errorcode
    kcnt = 0
    If Not rsDuplicat.EOF And Not rsDuplicat.BOF Then
        Screen.MousePointer = 11
        rsDuplicat.MoveFirst
        Do While Not rsDuplicat.EOF
            kcnt = kcnt + 1
            grdDuplicat.AddItem Null2String(rsDuplicat![Type]) & Chr(9) & _
                                Null2String(rsDuplicat!TRANTYPE) & Chr(9) & _
                                Null2String(rsDuplicat!TRANNO) & Chr(9) & _
                                Null2String(rsDuplicat!fileneym) & Chr(9) & _
                                Null2String(rsDuplicat!recno1) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!recno2) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!Status) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!ID)
            rsDuplicat.MoveNext
        Loop
        If kcnt <> 0 Then grdDuplicat.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNO_Mstr.EOF And Not rsNO_Mstr.BOF Then
        Screen.MousePointer = 11
        rsNO_Mstr.MoveFirst
        Do While Not rsNO_Mstr.EOF
            kcnt = kcnt + 1
            grdNo_Mstr.AddItem Null2String(rsNO_Mstr![Type]) & Chr(9) & _
                               Null2String(rsNO_Mstr!TRANTYPE) & Chr(9) & _
                               Null2String(rsNO_Mstr!TRANNO) & Chr(9) & _
                               Null2String(rsNO_Mstr!recno) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!stat_h) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!stat_d) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!ID)
            rsNO_Mstr.MoveNext
        Loop
        If kcnt <> 0 Then grdNo_Mstr.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNOHeader.EOF And Not rsNOHeader.BOF Then
        Screen.MousePointer = 11
        rsNOHeader.MoveFirst
        Do While Not rsNOHeader.EOF
            kcnt = kcnt + 1
            grdNoHeader.AddItem Null2String(rsNOHeader![Type]) & Chr(9) & _
                                Null2String(rsNOHeader!TRANTYPE) & Chr(9) & _
                                Null2String(rsNOHeader!TRANNO) & Chr(9) & _
                                Null2String(rsNOHeader!recno) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!stat_h) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!stat_d) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!ID)
            rsNOHeader.MoveNext
        Loop
        If kcnt <> 0 Then grdNoHeader.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNODetail.EOF And Not rsNODetail.BOF Then
        Screen.MousePointer = 11
        rsNODetail.MoveFirst
        Do While Not rsNODetail.EOF
            kcnt = kcnt + 1
            grdNoDetail.AddItem Null2String(rsNODetail!Type) & Chr(9) & _
                                Null2String(rsNODetail!TRANTYPE) & Chr(9) & _
                                Null2String(rsNODetail!TRANNO) & Chr(9) & _
                                Null2String(rsNODetail!recno) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!stat_h) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!stat_d) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!ID)
            rsNODetail.MoveNext
        Loop
        If kcnt <> 0 Then grdNoDetail.RemoveItem 1
        Screen.MousePointer = 0
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    RefreshDisplay
    LogAudit "G", "ERROR NO HEADER"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdDuplicat_DblClick()
    Dim fild, DupTranType, DupTranNo              As String
    Dim FieldName                                 As String
    Dim VTYPE                                     As String
    grdDuplicat.Row = grdDuplicat.Row
    grdDuplicat.Col = 0
    VTYPE = grdDuplicat.Text
    grdDuplicat.Col = 1
    DupTranType = grdDuplicat.Text
    grdDuplicat.Col = 2
    DupTranNo = grdDuplicat.Text
    grdDuplicat.Col = 3
    FieldName = grdDuplicat.Text
    grdDuplicat.Col = 6
    fild = grdDuplicat.Text
    If fild <> "" Then
        If VTYPE = "H" Then
            If MsgQuestionBox("View This Journal?", "Duplicate Journal") = True Then
                JOURNALTYPE = DupTranType
                frmAMISJournalEntry.Show
                frmAMISJournalEntry.SearchVoucherNo (DupTranNo)
                'Else
                '    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = vbYes Then
                '        gconDMIS.Execute "delete from " & FieldName & " where TYPE = '" & vTYPE & "' AND trantype = 'RIV' and tranno = " & N2Str2Null(DupTranNo)
                '        gconDMIS.Execute "delete from AMIS_Duplicat where id = " & fild
                '        rsRefreshDup
                '    End If
            End If
            Exit Sub
        End If
    End If
End Sub

Private Sub grdNoDetail_DblClick()
    grdNoDetail.Col = 1
    grdDuplicat.Row = 1
    If grdNoDetail.Text <> "No Entry" Then
        Dim Duplicat_ID, Journal_Det_ID           As Long
        grdNoDetail.Col = 6
        Duplicat_ID = grdNoDetail.Text
        grdNoDetail.Col = 3
        Journal_Det_ID = grdNoDetail.Text
        If MsgBox("Delete this Error Record?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
            gconDMIS.Execute ("Delete from AMIS_NoDetail where id = " & Duplicat_ID)
            gconDMIS.Execute ("Delete from AMIS_Journal_Det where id = " & Journal_Det_ID)
            RefreshDisplay
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Exit Sub
    grdDuplicat.Col = 1
    grdDuplicat.Row = 1
    If grdDuplicat.Text <> "No Entry" Then
        labDUPFILES.ForeColor = &HFF&
        labDUPFILES.Caption = "*** DUPLICATE FILE(S) ***"
        If labDUPFILES.Visible = False Then
            labDUPFILES.Visible = True
        Else
            labDUPFILES.Visible = False
        End If
    End If
    grdNoHeader.Col = 1
    grdNoHeader.Row = 1
    If grdNoHeader.Text <> "No Entry" Then
        labNoHeader.ForeColor = &HFF&
        labNoHeader.Caption = "*** NO HEADER ***"
        If labNoHeader.Visible = False Then
            labNoHeader.Visible = True
        Else
            labNoHeader.Visible = False
        End If
    End If
    grdNo_Mstr.Col = 1
    grdNo_Mstr.Row = 1
    If grdNo_Mstr.Text <> "No Entry" Then
        labNo_Mstr.ForeColor = &HFF&
        labNo_Mstr.Caption = "*** NO MASTER ***"
        If labNo_Mstr.Visible = False Then
            labNo_Mstr.Visible = True
        Else
            labNo_Mstr.Visible = False
        End If
    End If
    grdNoDetail.Col = 1
    grdNoDetail.Row = 1
    If grdNoDetail.Text <> "No Entry" Then
        labNoDetail.ForeColor = &HFF&
        labNoDetail.Caption = "*** NO DETAIL ***"
        If labNoDetail.Visible = False Then
            labNoDetail.Visible = True
        Else
            labNoDetail.Visible = False
        End If
    End If
End Sub

