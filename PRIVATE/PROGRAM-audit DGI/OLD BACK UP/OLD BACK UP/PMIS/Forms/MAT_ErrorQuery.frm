VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPMISMAT_ErrorQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR FILES..."
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_ErrorQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   11700
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   180
      Top             =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grdDuplicat 
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   7
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNo_Mstr 
      Height          =   2535
      Left            =   90
      TabIndex        =   2
      Top             =   3360
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   6
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNoHeader 
      Height          =   2535
      Left            =   5910
      TabIndex        =   1
      Top             =   420
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   6
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdNoDetail 
      Height          =   2535
      Left            =   5910
      TabIndex        =   3
      Top             =   3360
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   6
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
      Top             =   2970
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
      Caption         =   "NO MASTER"
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
      Top             =   2970
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
Attribute VB_Name = "frmPMISMAT_ErrorQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDuplicat, rsNO_Mstr              As ADODB.Recordset
Attribute rsNO_Mstr.VB_VarUserMemId = 1073938432
Dim rsNOHeader, rsNODetail             As ADODB.Recordset
Attribute rsNOHeader.VB_VarUserMemId = 1073938434
Attribute rsNODetail.VB_VarUserMemId = 1073938434

Dim AddorEdit                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938436
Dim kcnt                               As Integer
Attribute kcnt.VB_VarUserMemId = 1073938437

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    cleargrd
    Set rsDuplicat = New ADODB.Recordset
    rsDuplicat.Open "Select * from PMIS_Duplicat", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNO_Mstr = New ADODB.Recordset
    rsNO_Mstr.Open "Select * from PMIS_No_Mstr", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNOHeader = New ADODB.Recordset
    rsNOHeader.Open "Select * from PMIS_NoHeader", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Set rsNODetail = New ADODB.Recordset
    rsNODetail.Open "Select * from PMIS_NoDetail", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
    LogAudit "G", "ERROR NO HEADER"
    Screen.MousePointer = 0
End Sub

Sub rsRefreshDup()
    cleargrd
    Set rsDuplicat = New ADODB.Recordset
    rsDuplicat.Open "Select * from PMIS_Duplicat", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNO_Mstr()
    cleargrd
    Set rsNO_Mstr = New ADODB.Recordset
    rsNO_Mstr.Open "Select * from PMIS_No_Mstr", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNoHead()
    cleargrd
    Set rsNOHeader = New ADODB.Recordset
    rsNOHeader.Open "Select * from PMIS_NoHeader", gconDMIS, adOpenForwardOnly, adLockReadOnly
    InitGrid
    FillGrid
End Sub

Sub rsRefreshNoDet()
    cleargrd
    Set rsNODetail = New ADODB.Recordset
    rsNODetail.Open "Select * from PMIS_NoDetail", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    Dim kim                            As Integer
    With grdDuplicat
        .Row = 0
        .FormatString = "Tran Type | Tranno | File                          | " & _
                        "Rec No. 1 | Rec No. 2 | Status | ID"
        .ColWidth(6) = 1
    End With
    With grdNo_Mstr
        .Row = 0
        .FormatString = "Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
    With grdNoHeader
        .Row = 0
        .FormatString = "Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
    With grdNoDetail
        .Row = 0
        .FormatString = "Tran Type | Tranno | " & _
                        "Rec No. | Stat_H | Stat_D | ID"
        .ColWidth(5) = 1
    End With
End Sub

Sub FillGrid()
    On Error GoTo ErrorCode
    kcnt = 0
    If Not rsDuplicat.EOF And Not rsDuplicat.BOF Then
        Screen.MousePointer = 11
        rsDuplicat.MoveFirst
        Do While Not rsDuplicat.EOF
            kcnt = kcnt + 1
            grdDuplicat.AddItem Null2String(rsDuplicat!TRANTYPE) & Chr(9) & _
                                Null2String(rsDuplicat!Tranno) & Chr(9) & _
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
            grdNo_Mstr.AddItem Null2String(rsNO_Mstr!TRANTYPE) & Chr(9) & _
                               Null2String(rsNO_Mstr!Tranno) & Chr(9) & _
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
            grdNoHeader.AddItem Null2String(rsNOHeader!TRANTYPE) & Chr(9) & _
                                Null2String(rsNOHeader!Tranno) & Chr(9) & _
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
            grdNoDetail.AddItem Null2String(rsNODetail!TRANTYPE) & Chr(9) & _
                                Null2String(rsNODetail!Tranno) & Chr(9) & _
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

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISErrorQuery = Nothing
    UnloadForm Me
End Sub

Private Sub grdDuplicat_DblClick()
    Dim fild, DupTranType, DupTranNo   As String
    Dim FieldName                      As String
    grdDuplicat.Row = grdDuplicat.Row
    grdDuplicat.Col = 0
    DupTranType = grdDuplicat.Text
    grdDuplicat.Col = 1
    DupTranNo = grdDuplicat.Text
    grdDuplicat.Col = 2
    FieldName = grdDuplicat.Text
    grdDuplicat.Col = 6
    fild = grdDuplicat.Text
    If fild <> "" Then
        If DupTranType = "RIV" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    COUNTERTYPE = "RIV"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = vbYes Then
                        gconDMIS.Execute "delete from " & FieldName & " where trantype = 'RIV' and tranno = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction"
                Exit Sub
            End If
        End If

        If DupTranType = "ADB" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    COUNTERTYPE = "ADB"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = vbYes Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'ADB' and tranno = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction"
                Exit Sub
            End If
        End If

        If LOGLEVEL <> "RIV USER" Then
            If DupTranType = "CSH" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    COUNTERTYPE = "CSH"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CSH' and tranno = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
            End If
            If DupTranType = "CHG" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    COUNTERTYPE = "CHG"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CHG' and tranno = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
            End If
            If DupTranType = "RR" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    frmPMISReceiving2.Show
                    frmPMISReceiving2.FindDupRRno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_RR_Hd where rrno = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
            End If
            If DupTranType = "PO" Then
                If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                    'frmPMISPurchase.Show
                    'frmPMISPurchase.FindDupPOno (DupTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_PO_Hd where pono = " & N2Str2Null(DupTranNo)
                        gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                        rsRefreshDup
                    End If
                End If
            End If
        Else
            MsgSpeechBox "User level not allowed for this Transaction."
            Exit Sub
        End If
    End If
End Sub

Private Sub grdNo_Mstr_DblClick()
    Dim fild, No_MstrTranType, No_MstrTranNo As String

    grdNo_Mstr.Row = grdNo_Mstr.Row
    grdNo_Mstr.Col = 0
    No_MstrTranType = grdNo_Mstr.Text
    grdNo_Mstr.Col = 1
    No_MstrTranNo = grdNo_Mstr.Text
    grdNo_Mstr.Col = 5
    fild = grdNo_Mstr.Text
    If fild <> "" Then
        If No_MstrTranType = "RIV" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    COUNTERTYPE = "RIV"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'RIV' and tranno = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If

        If No_MstrTranType = "ADB" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    COUNTERTYPE = "ADB"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'ADB' and tranno = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If

        If LOGLEVEL <> "RIV USER" Then
            If No_MstrTranType = "CSH" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    COUNTERTYPE = "CSH"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CSH' and tranno = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
            End If
            If No_MstrTranType = "CHG" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    COUNTERTYPE = "CHG"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CHG' and tranno = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
            End If
            If No_MstrTranType = "RR" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    frmPMISReceiving2.Show
                    frmPMISReceiving2.FindDupRRno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_RR_Hd where rrno = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
            End If
            If No_MstrTranType = "PO" Then
                If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                    'frmPMISPurchase.Show
                    'frmPMISPurchase.FindDupPOno (No_MstrTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_PO_Hd where pono = " & N2Str2Null(No_MstrTranNo)
                        gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                        rsRefreshNO_Mstr
                    End If
                End If
            End If
        Else
            MsgSpeechBox "User level not allowed for this Transaction."
            Exit Sub
        End If
    End If
End Sub

Private Sub grdNoHeader_DblClick()
    Dim fild, NoHeadTranType, NoHeadTranNo As String

    grdNoHeader.Row = grdNoHeader.Row
    grdNoHeader.Col = 0
    NoHeadTranType = grdNoHeader.Text
    grdNoHeader.Col = 1
    NoHeadTranNo = grdNoHeader.Text
    grdNoHeader.Col = 5
    fild = grdNoHeader.Text
    If fild <> "" Then
        If NoHeadTranType = "RIV" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'RIV' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If

        If NoHeadTranType = "ADB" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'ADB' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If
        If LOGLEVEL <> "RIV USER" Then
            If NoHeadTranType = "CSH" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'CSH' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            End If
            If NoHeadTranType = "CHG" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'CHG' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            End If
            If NoHeadTranType = "RR" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            End If
            If NoHeadTranType = "PO" Then
                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                    gconDMIS.Execute "delete from PMIS_TdayTran where trantype = 'PO' and tranno = " & N2Str2Null(NoHeadTranNo)
                    gconDMIS.Execute "delete from PMIS_NoHeader where id = " & fild
                    rsRefreshNoHead
                End If
                Exit Sub
            End If
        Else
            MsgSpeechBox "User level not allowed for this Transaction."
            Exit Sub
        End If
    End If
End Sub

Private Sub grdNoDetail_DblClick()
    Dim fild, NODetTranType, NODetTranNo As String

    grdNoDetail.Row = grdNoDetail.Row
    grdNoDetail.Col = 0
    NODetTranType = grdNoDetail.Text
    grdNoDetail.Col = 1
    NODetTranNo = grdNoDetail.Text
    grdNoDetail.Col = 5
    fild = grdNoDetail.Text
    If fild <> "" Then
        If NODetTranType = "RIV" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    COUNTERTYPE = "RIV"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'RIV' and tranno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If

        If NODetTranType = "ADB" Then
            If LOGLEVEL = "RIV USER" Or LOGLEVEL = "ADM" Or LOGLEVEL = "AUTHOR" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    COUNTERTYPE = "ADB"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'ADB' and tranno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
                Exit Sub
            Else
                MsgSpeechBox "User level not allowed for this Transaction."
                Exit Sub
            End If
        End If

        If LOGLEVEL <> "RIV USER" Then
            If NODetTranType = "CSH" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    COUNTERTYPE = "CSH"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CSH' and tranno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
            End If
            If NODetTranType = "CHG" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    COUNTERTYPE = "CHG"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'CHG' and tranno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
            End If
            If NODetTranType = "DR" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    COUNTERTYPE = "DR"
                    frmPMISCustomerOrder.Show
                    frmPMISCustomerOrder.FindDupTranno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_Ord_Hd where trantype = 'DR' and tranno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
            End If
            If NODetTranType = "RR" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    frmPMISReceiving2.Show
                    frmPMISReceiving2.FindDupRRno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_RR_Hd where rrno = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
            End If
            If NODetTranType = "PO" Then
                If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                    'frmPMISPurchase.Show
                    'frmPMISPurchase.FindDupPOno (NODetTranNo)
                Else
                    If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                        gconDMIS.Execute "delete from PMIS_PO_Hd where pono = " & N2Str2Null(NODetTranNo)
                        gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                        rsRefreshNoDet
                    End If
                End If
            End If
        Else
            MsgSpeechBox "User level not allowed for this Transaction."
            Exit Sub
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    grdDuplicat.Col = 0
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
    grdNoHeader.Col = 0
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
    grdNo_Mstr.Col = 0
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
    grdNoDetail.Col = 0
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
