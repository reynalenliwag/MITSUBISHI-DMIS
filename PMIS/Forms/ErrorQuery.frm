VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPMISInquiry_ErrorQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR FILES..."
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11715
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ErrorQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11715
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
Attribute VB_Name = "frmPMISInquiry_ErrorQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDuplicat, rsNO_Mstr                              As ADODB.Recordset
Attribute rsNO_Mstr.VB_VarUserMemId = 1073938432
Dim rsNOHeader, rsNODetail                             As ADODB.Recordset
Attribute rsNOHeader.VB_VarUserMemId = 1073938434
Attribute rsNODetail.VB_VarUserMemId = 1073938434

Dim ADDOREDIT                                          As String
Attribute ADDOREDIT.VB_VarUserMemId = 1073938436
Dim KCNT                                               As Integer
Attribute KCNT.VB_VarUserMemId = 1073938437

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
    KCNT = 0
    If Not rsDuplicat.EOF And Not rsDuplicat.BOF Then
        Screen.MousePointer = 11
        rsDuplicat.MoveFirst
        Do While Not rsDuplicat.EOF
            KCNT = KCNT + 1
            grdDuplicat.AddItem Null2String(rsDuplicat![Type]) & Chr(9) & _
                                Null2String(rsDuplicat!TranType) & Chr(9) & _
                                Null2String(rsDuplicat!TRANNO) & Chr(9) & _
                                Null2String(rsDuplicat!fileneym) & Chr(9) & _
                                Null2String(rsDuplicat!recno1) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!recno2) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!STATUS) & Chr(9) & _
                                N2Str2IntZero(rsDuplicat!ID)
            rsDuplicat.MoveNext
        Loop
        If KCNT <> 0 Then grdDuplicat.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNO_Mstr.EOF And Not rsNO_Mstr.BOF Then
        Screen.MousePointer = 11
        rsNO_Mstr.MoveFirst
        Do While Not rsNO_Mstr.EOF
            KCNT = KCNT + 1
            grdNo_Mstr.AddItem Null2String(rsNO_Mstr![Type]) & Chr(9) & _
                               Null2String(rsNO_Mstr!TranType) & Chr(9) & _
                               Null2String(rsNO_Mstr!TRANNO) & Chr(9) & _
                               Null2String(rsNO_Mstr!recno) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!stat_h) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!stat_d) & Chr(9) & _
                               N2Str2IntZero(rsNO_Mstr!ID)
            rsNO_Mstr.MoveNext
        Loop
        If KCNT <> 0 Then grdNo_Mstr.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNOHeader.EOF And Not rsNOHeader.BOF Then
        Screen.MousePointer = 11
        rsNOHeader.MoveFirst
        Do While Not rsNOHeader.EOF
            KCNT = KCNT + 1
            grdNoHeader.AddItem Null2String(rsNOHeader![Type]) & Chr(9) & _
                                Null2String(rsNOHeader!TranType) & Chr(9) & _
                                Null2String(rsNOHeader!TRANNO) & Chr(9) & _
                                Null2String(rsNOHeader!recno) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!stat_h) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!stat_d) & Chr(9) & _
                                N2Str2IntZero(rsNOHeader!ID)
            rsNOHeader.MoveNext
        Loop
        If KCNT <> 0 Then grdNoHeader.RemoveItem 1
        Screen.MousePointer = 0
    End If

    If Not rsNODetail.EOF And Not rsNODetail.BOF Then
        Screen.MousePointer = 11
        rsNODetail.MoveFirst
        Do While Not rsNODetail.EOF
            KCNT = KCNT + 1
            grdNoDetail.AddItem Null2String(rsNODetail!Type) & Chr(9) & _
                                Null2String(rsNODetail!TranType) & Chr(9) & _
                                Null2String(rsNODetail!TRANNO) & Chr(9) & _
                                Null2String(rsNODetail!recno) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!stat_h) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!stat_d) & Chr(9) & _
                                N2Str2IntZero(rsNODetail!ID)
            rsNODetail.MoveNext
        Loop
        If KCNT <> 0 Then grdNoDetail.RemoveItem 1
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
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISInquiry_ErrorQuery = Nothing
    UnloadForm Me
End Sub

Private Sub grdDuplicat_DblClick()
    Dim FILD, DupTranType, DupTranNo                   As String
    Dim FieldName                                      As String
    Dim VTYPE                                          As String
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
    FILD = grdDuplicat.Text
    If FILD <> "" Then
        If DupTranType = "RIV" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                COUNTERTYPE = "RIV"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (DupTranNo)
                End If
                'JBFandEAP:021809 to prevent the deletion in transaction error
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = vbYes Then
                '                    gconDMIS.Execute "delete from " & FieldName & " where TYPE = '" & VTYPE & "' AND trantype = 'RIV' and tranno = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
            Exit Sub
        End If

        If DupTranType = "ADB" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                COUNTERTYPE = "ADB"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (DupTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = vbYes Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'ADB' and tranno = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
            Exit Sub
        End If

        If DupTranType = "CSH" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                COUNTERTYPE = "CSH"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (DupTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CSH' and tranno = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
        End If
        If DupTranType = "CHG" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                COUNTERTYPE = "CHG"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (DupTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (DupTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CHG' and tranno = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
        End If
        If DupTranType = "RR" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                If VTYPE = "A" Then
                    frmPMISTrans_Receiving2_AC.Show
                    frmPMISTrans_Receiving2_AC.FindDupRRno (DupTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_Receiving2_MAT.Show
                    frmPMISTrans_Receiving2_MAT.FindDupRRno (DupTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_Receiving2.Show
                    frmPMISTrans_Receiving2.FindDupRRno (DupTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_RR_Hd where TYPE = '" & VTYPE & "' AND rrno = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
        End If
        If DupTranType = "PO" Then
            If MsgQuestionBox("View This Transaction?", "Duplicate Transaction") = True Then
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "Duplicate Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_PO_Hd where TYPE = '" & VTYPE & "' AND pono = " & N2Str2Null(DupTranNo)
                '                    gconDMIS.Execute "delete from PMIS_Duplicat where id = " & fild
                '                    rsRefreshDup
                '                End If
            End If
        End If
    End If
End Sub

Private Sub grdNo_Mstr_DblClick()
    Dim FILD, No_MstrTranType, No_MstrTranNo           As String
    Dim VTYPE                                          As String
    grdNo_Mstr.Row = grdNo_Mstr.Row
    grdNo_Mstr.Col = 0
    VTYPE = grdNo_Mstr.Text
    grdNo_Mstr.Col = 1
    No_MstrTranType = grdNo_Mstr.Text
    grdNo_Mstr.Col = 2
    No_MstrTranNo = grdNo_Mstr.Text
    grdNo_Mstr.Col = 6
    FILD = grdNo_Mstr.Text
    If FILD <> "" Then
        If No_MstrTranType = "RIV" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                COUNTERTYPE = "RIV"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (No_MstrTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'RIV' and tranno = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
            Exit Sub
        End If

        If No_MstrTranType = "ADB" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                COUNTERTYPE = "ADB"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (No_MstrTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'ADB' and tranno = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
            Exit Sub
        End If

        If No_MstrTranType = "CSH" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                COUNTERTYPE = "CSH"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (No_MstrTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CSH' and tranno = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
        End If
        If No_MstrTranType = "CHG" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                COUNTERTYPE = "CHG"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (No_MstrTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (No_MstrTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CHG' and tranno = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
        End If
        If No_MstrTranType = "RR" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                If VTYPE = "A" Then
                    frmPMISTrans_Receiving2_AC.Show
                    frmPMISTrans_Receiving2_AC.FindDupRRno (No_MstrTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_Receiving2_MAT.Show
                    frmPMISTrans_Receiving2_MAT.FindDupRRno (No_MstrTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_Receiving2.Show
                    frmPMISTrans_Receiving2.FindDupRRno (No_MstrTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_RR_Hd where TYPE = '" & VTYPE & "' AND rrno = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
        End If
        If No_MstrTranType = "PO" Then
            If MsgQuestionBox("View This Transaction?", "No Master Transaction") = True Then
                frmPMISTrans_Purchase.Show
                frmPMISTrans_Purchase.FindDupPOno (No_MstrTranNo)
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Master Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_PO_Hd where TYPE = '" & VTYPE & "' AND pono = " & N2Str2Null(No_MstrTranNo)
                '                    gconDMIS.Execute "delete from PMIS_No_Mstr where id = " & fild
                '                    rsRefreshNO_Mstr
                '                End If
            End If
        End If
    End If
End Sub

Private Sub grdNoHeader_DblClick()
    Dim FILD, NoHeadTranType, NoHeadTranNo             As String
    Dim VTYPE                                          As String
    grdNoHeader.Row = grdNoHeader.Row
    grdNoHeader.Col = 0
    VTYPE = grdNoHeader.Text
    grdNoHeader.Col = 1
    NoHeadTranType = grdNoHeader.Text
    grdNoHeader.Col = 2
    NoHeadTranNo = grdNoHeader.Text
    grdNoHeader.Col = 6
    FILD = grdNoHeader.Text
    If FILD <> "" Then
        If NoHeadTranType = "RIV" Then
            If MsgQuestionBox("Delete this Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'RIV' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If

        If NoHeadTranType = "ADB" Then
            If MsgQuestionBox("Delete this Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'ADB' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If
        If NoHeadTranType = "CSH" Then
            If MsgQuestionBox("Delete This Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'CSH' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If
        If NoHeadTranType = "CHG" Then
            If MsgQuestionBox("Delete This Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'CHG' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If
        If NoHeadTranType = "RR" Then
            If MsgQuestionBox("Delete This Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'RR' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If
        If NoHeadTranType = "PO" Then
            If MsgQuestionBox("Delete This Transaction?", "No Header Transaction") = True Then
                SQL_STATEMENT = "delete from PMIS_TdayTran where TYPE = '" & VTYPE & "' AND trantype = 'PO' and tranno = '" & NoHeadTranNo & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, FindTransactionID(NoHeadTranNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null(VTYPE), "TYPE"), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                SQL_STATEMENT = "delete from PMIS_NoHeader where id = " & FILD
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "BROWSE ERROR FILES", SQL_STATEMENT, N2Str2Null(FILD), "", NoHeadTranNo, N2Str2Null(NoHeadTranType), ""

                rsRefreshNoHead
            End If
            Exit Sub
        End If
    End If
End Sub

Private Sub grdNoDetail_DblClick()
    Dim FILD, NODetTranType, NODetTranNo               As String
    Dim VTYPE                                          As String
    grdNoDetail.Row = grdNoDetail.Row
    grdNoDetail.Col = 0
    VTYPE = grdNoDetail.Text
    grdNoDetail.Col = 1
    NODetTranType = grdNoDetail.Text
    grdNoDetail.Col = 2
    NODetTranNo = grdNoDetail.Text
    grdNoDetail.Col = 6
    FILD = grdNoDetail.Text
    If FILD <> "" Then
        If NODetTranType = "RIV" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                COUNTERTYPE = "RIV"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (NODetTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where type = '" & VTYPE & "' and trantype = 'RIV' and tranno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
            Exit Sub
        End If

        If NODetTranType = "ADB" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                COUNTERTYPE = "ADB"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (NODetTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' and trantype = 'ADB' and tranno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
            Exit Sub
        End If

        If NODetTranType = "CSH" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                COUNTERTYPE = "CSH"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (NODetTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CSH' and tranno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
        End If
        If NODetTranType = "CHG" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                COUNTERTYPE = "CHG"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (NODetTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'CHG' and tranno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
        End If
        If NODetTranType = "DR" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                COUNTERTYPE = "DR"
                If VTYPE = "A" Then
                    frmPMISTrans_CustomerOrder_AC.Show
                    frmPMISTrans_CustomerOrder_AC.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_CustomerOrder_MAT.Show
                    frmPMISTrans_CustomerOrder_MAT.FindDupTranno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_CustomerOrder.Show
                    frmPMISTrans_CustomerOrder.FindDupTranno (NODetTranNo)
                End If
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_Ord_Hd where TYPE = '" & VTYPE & "' AND trantype = 'DR' and tranno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
        End If
        If NODetTranType = "RR" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                If VTYPE = "A" Then
                    frmPMISTrans_Receiving2_AC.Show
                    frmPMISTrans_Receiving2_AC.FindDupRRno (NODetTranNo)
                End If
                If VTYPE = "M" Then
                    frmPMISTrans_Receiving2_MAT.Show
                    frmPMISTrans_Receiving2_MAT.FindDupRRno (NODetTranNo)
                End If
                If VTYPE = "P" Then
                    frmPMISTrans_Receiving2.Show
                    frmPMISTrans_Receiving2.FindDupRRno (NODetTranNo)
                End If
                'JBFandEAP:021809 to prevent the deletion in transaction error
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_RR_Hd where TYPE = '" & VTYPE & "' AND rrno = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
        End If
        If NODetTranType = "PO" Then
            If MsgQuestionBox("View This Transaction?", "No Detail Transaction") = True Then
                '            Else
                '                If MsgQuestionBox("Delete This Transaction?", "No Detail Transaction") = True Then
                '                    gconDMIS.Execute "delete from PMIS_PO_Hd where TYPE = '" & VTYPE & "' AND pono = " & N2Str2Null(NODetTranNo)
                '                    gconDMIS.Execute "delete from PMIS_NoDetail where id = " & fild
                '                    rsRefreshNoDet
                '                End If
            End If
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

