VERSION 5.00
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmPMISTrans_POConfirmationProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO Confirmation"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   Icon            =   "POConfirmationProcess.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSEQ_NO 
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
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   480
      Width           =   585
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   11085
      TabIndex        =   16
      Top             =   3720
      Width           =   11115
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Back Order For Allocation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   4695
      End
      Begin VB.Label LABALLOCATED 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   8070
         TabIndex        =   22
         Top             =   30
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ALLOCATED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   6870
         TabIndex        =   21
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ORDERED:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   4770
         TabIndex        =   20
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BACK ORDER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   8970
         TabIndex        =   19
         Top             =   30
         Width           =   1185
      End
      Begin VB.Label LABTOTALORDERED 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   5970
         TabIndex        =   18
         Top             =   30
         Width           =   885
      End
      Begin VB.Label LABBACKORDER 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   10170
         TabIndex        =   17
         Top             =   30
         Width           =   885
      End
   End
   Begin VB.TextBox txtSEQ_NO 
      Height          =   375
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtTranno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8130
      TabIndex        =   12
      Text            =   "Text"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPOType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Text            =   "Text"
      Top             =   480
      Width           =   2205
   End
   Begin VB.TextBox txtConfirmedDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Text            =   "Text"
      Top             =   480
      Width           =   1635
   End
   Begin VB.TextBox txtPODate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2850
      TabIndex        =   6
      Text            =   "Text"
      Top             =   480
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      MouseIcon       =   "POConfirmationProcess.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "POConfirmationProcess.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   180
      Width           =   765
   End
   Begin FlexCell.Grid GridDetails 
      Height          =   2685
      Left            =   60
      TabIndex        =   3
      Top             =   4230
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   4736
      BackColor2      =   12648384
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.CommandButton cmdFindPO 
      Height          =   390
      Left            =   2370
      MouseIcon       =   "POConfirmationProcess.frx":0A1A
      MousePointer    =   99  'Custom
      Picture         =   "POConfirmationProcess.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Process Checking of Previous Cut-Off Balance"
      Top             =   480
      Width           =   450
   End
   Begin VB.TextBox txtPONum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9570
      MouseIcon       =   "POConfirmationProcess.frx":0FEB
      MousePointer    =   99  'Custom
      Picture         =   "POConfirmationProcess.frx":113D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save Changes"
      Top             =   180
      Width           =   765
   End
   Begin FlexCell.Grid GridHeader 
      Height          =   2595
      Left            =   60
      TabIndex        =   14
      Top             =   1080
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   4577
      BackColor2      =   12648384
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*Pls. Update the first row."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      TabIndex        =   24
      Top             =   6990
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tranno. No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8100
      TabIndex        =   13
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   11
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   210
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      TabIndex        =   7
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1425
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00F5D8BC&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00F5D8BC&
      Height          =   885
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   9465
   End
End
Attribute VB_Name = "frmPMISTrans_POConfirmationProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SEQNO_Status                        As Boolean
Dim aydi                                As Integer
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFindPO_Click()
    Screen.MousePointer = 11

    If Trim(txtPONum) = "" Then
        MsgSpeechBox "Please input a valid Purchase Order Number."
        txtPONum.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
'gconDMIS.Execute ("alter table PMIS_PO_DETAILS ADD SEQ_NO NVARCHAR(10)")
'gconDMIS.Execute ("alter table PMIS_PO_DETAILS ADD PODATE SMALLDATETIME")
'gconDMIS.Execute ("alter table PMIS_PO_DETAILS ADD STOCKNO NVARCHAR(30)")
'gconDMIS.Execute ("alter table PMIS_PO_DETAILS ADD CONFIRMEDDATE SMALLDATETIME")

    Dim VSeq_No As String
    Dim A, b, c As String
    Dim vBOAmount As Double
    Dim rsThan As ADODB.Recordset
    Set rsThan = New ADODB.Recordset
   
    Set rsThan = gconDMIS.Execute("SELECT SEQ_NO,PODATE,SONUM,PO_NO FROM PMIS_PO_DETAILS WHERE SONUM = " & N2Str2Null(Trim(txtPONum)) & " ORDER BY SEQ_NO ASC")
    If Not rsThan.EOF And Not rsThan.BOF Then
        cboSEQ_NO.Clear
        rsThan.MoveLast
        VSeq_No = Mid(Null2String(rsThan!SEQ_NO), 2, 1)
        A = NumericVal(VSeq_No) + 1
        b = "0" & CStr(A)
        'txtSEQ_NO.Text = "0" & CStr(a)
        c = 0
        While A > 0
            cboSEQ_NO.AddItem "0" & CStr(c)
            c = c + 1
            A = A - 1
        Wend
        cboSEQ_NO.AddItem b
        If VSeq_No = "" Then
            cboSEQ_NO.ListIndex = 0
        Else
            cboSEQ_NO.ListIndex = CInt(VSeq_No + 1)
        End If
        SEQNO_Status = True
        
    Else
        ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If

    txtPODate = Format(Null2Date(rsThan!PODATE), "SHORT DATE")
    txtConfirmedDate = Format(LOGDATE, "SHORT DATE")
    txtPOType = GETPOTYPE(rsThan!SONUM)
    txtTranno = Null2String(rsThan!PO_NO)

    GridHeader.Rows = 1
    GridHeader.Refresh
    GridHeader.AutoRedraw = False
    Dim rsPONUm As ADODB.Recordset
    Set rsPONUm = New ADODB.Recordset
    
    Set rsPONUm = gconDMIS.Execute("SELECT ID,STOCKNO,PODATE,SONUM,SEQ_NO,QTY_ORDERED, QTY_ALLOCATED, QTY_SERVED, QTY_BACKORDER, CONFIRMEDDATE FROM PMIS_PO_DETAILS WHERE SONUM= " & N2Str2Null(Trim(txtPONum)) & " ORDER BY ID ASC")
    If Not rsPONUm.EOF And Not rsPONUm.BOF Then
        Dim row_index, col_index, TOTAL_ALLOCATED, ORDERED_QTY As Integer
        row_index = 0: col_index = 4: TOTAL_ALLOCATED = 0
        rsPONUm.MoveFirst
        ORDERED_QTY = NumericVal(rsPONUm!QTY_ORDERED)
        Do While Not rsPONUm.EOF
            GridHeader.AddItem Null2Date(rsPONUm!PODATE) & Chr(9) & _
                               Null2String(rsPONUm!SONUM) & "- " & Null2String(rsPONUm!SEQ_NO) & Chr(9) & _
                               Null2String(rsPONUm!STOCKNO) & Chr(9) & _
                               Null2String(GETPOTYPE(rsThan!SONUM)) & Chr(9) & _
                               NumericVal(rsPONUm!QTY_ORDERED) & Chr(9) & _
                               NumericVal(rsPONUm!Qty_Allocated) & Chr(9) & _
                               NumericVal(rsPONUm!Qty_Served) & Chr(9) & _
                               NumericVal(rsPONUm!QTY_BACKORDER) & Chr(9) & _
                               Null2Date(rsPONUm!CONFIRMEDDATE)
            row_index = row_index + 1
            GridHeader.Cell(row_index, 0).Text = rsPONUm!ID
            TOTAL_ALLOCATED = TOTAL_ALLOCATED + NumericVal(rsPONUm!Qty_Allocated)
            rsPONUm.MoveNext
        Loop
    Else
        ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If
    LABALLOCATED = TOTAL_ALLOCATED
    LABTOTALORDERED = COMPUTE_TOTAL_ORDERED
    LABBACKORDER = COMPUTE_TOTAL_ORDERED - LABALLOCATED
    GridHeader.Refresh
    GridHeader.AutoRedraw = True
    Set rsPONUm = Nothing
    Set rsThan = Nothing
    Screen.MousePointer = 0
End Sub
Sub cboSEQ_No_Click()
    On Error Resume Next
    If SEQNO_Status = True Then Exit Sub
    Dim rsJ As ADODB.Recordset
    Set rsJ = New ADODB.Recordset
    Set rsJ = gconDMIS.Execute("SELECT ID,PODATE,SONUM,SEQ_NO,QTY_ORDERED, QTY_ALLOCATED, QTY_SERVED, QTY_BACKORDER, CONFIRMEDDATE FROM PMIS_PO_DETAILS WHERE SONUM = " & N2Str2Null(Trim(txtPONum)) & " AND SEQ_NO = " & N2Str2Null(cboSEQ_NO))
    If Not rsJ.EOF And Not rsJ.BOF Then
        GridHeader.Rows = 1
        GridHeader.Refresh
        GridHeader.AutoRedraw = True
        Dim row_index As Integer
        row_index = 0
        rsJ.MoveFirst
        Do While Not rsJ.EOF
            GridHeader.AddItem Null2Date(rsJ!PODATE) & Chr(9) & _
                               Null2String(rsJ!SONUM) & "- " & Null2String(rsJ!SEQ_NO) & Chr(9) & _
                               Null2String(GETPOTYPE(rsJ!SONUM)) & Chr(9) & _
                               NumericVal(rsJ!QTY_ORDERED) & Chr(9) & _
                               NumericVal(rsJ!Qty_Allocated) & Chr(9) & _
                               NumericVal(rsJ!Qty_Served) & Chr(9) & _
                               NumericVal(rsJ!QTY_BACKORDER) & Chr(9) & _
                               Null2Date(rsJ!CONFIRMEDDATE)
            row_index = row_index + 1
            GridHeader.Cell(row_index, 0).Text = rsJ!ID
            rsJ.MoveNext
        Loop
        GridHeader.Refresh
        GridHeader.AutoRedraw = False
    Else
        ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 0
End Sub


Function COMPUTE_TOTAL_ORDERED() As Integer
    Dim rsMyBday As ADODB.Recordset
    Set rsMyBday = New ADODB.Recordset
    Set rsMyBday = gconDMIS.Execute("SELECT SUM(QTY_ORDERED) AS TOTAL_ORDERED FROM PMIS_PO_DETAILS WHERE SEQ_NO = '00' AND SONUM = " & N2Str2Null(Trim(txtPONum)))
    If Not rsMyBday.EOF And Not rsMyBday.BOF Then
        COMPUTE_TOTAL_ORDERED = NumericVal(rsMyBday!TOTAL_ORDERED)
    End If
End Function

Function COMPUTE_BO_AMOUNT(BAK_ORDER As Integer) As Double
    Dim rsLorna As ADODB.Recordset
    Set rsLorna = New ADODB.Recordset
    Set rsLorna = gconDMIS.Execute("SELECT UNITPRICE FROM PMIS_PO_DETAILS WHERE SEQ_NO = '00' AND SONUM = " & N2Str2Null(Trim(txtPONum)))
    If Not rsLorna.EOF And Not rsLorna.BOF Then
        COMPUTE_BO_AMOUNT = BAK_ORDER * NumericVal(rsLorna!UnitPrice)
    End If
End Function
Function CALCULATE_CURR_BO(cleverj As Integer) As Double
    Dim rsBday As ADODB.Recordset
    Set rsBday = New ADODB.Recordset
    Set rsBday = gconDMIS.Execute("SELECT UNITPRICE FROM PMIS_PO_DETAILS WHERE SEQ_NO = '00' AND ID = " & aydi & " AND SONUM = " & N2Str2Null(Trim(txtPONum)))
    If Not rsBday.EOF And Not rsBday.BOF Then
        CALCULATE_CURR_BO = Round(cleverj * NumericVal(rsBday!UnitPrice), 2)
    End If
    Set rsBday = Nothing
End Function
Function CALCULATE_BO_Amount(thanzky As String) As Integer
    Dim rsBO As ADODB.Recordset
    Set rsBO = New ADODB.Recordset
    Set rsBO = gconDMIS.Execute("SELECT SUM(QTY_ORDERED) AS TOTAL_ORDERED, SUM(QTY_ALLOCATED) AS TOTAL_ALLOCATED FROM PMIS_PO_DETAILS WHERE SONUM = " & N2Str2Null(Trim(thanzky)))
    If Not rsBO.EOF And Not rsBO.BOF Then
        CALCULATE_BO_Amount = NumericVal(rsBO!TOTAL_ORDERED) - (NumericVal(rsBO!TOTAL_ALLOCATED) + NumericVal(GridDetails.Cell(1, 5).Text))
    End If
    Set rsBO = Nothing
End Function
Private Sub cmdSave_Click()
    Screen.MousePointer = 11
    If Trim(txtPONum) = "" Then
        MsgSpeechBox "PO Number must not be empty."
        Screen.MousePointer = 0
        Exit Sub
    End If
    If GridDetails.Cell(1, 2).Text = "" Then
        MsgSpeechBox "There's nothing to save."
        Screen.MousePointer = 0
        Exit Sub
    End If
    Dim rsSONum As ADODB.Recordset
    Set rsSONum = New ADODB.Recordset
    Set rsSONum = gconDMIS.Execute("SELECT SONUM,SEQ_NO,STOCKNO FROM PMIS_PO_DETAILS WHERE SONUM = '" & txtPONum & "'  AND SEQ_NO = '" & cboSEQ_NO & "' AND STOCKNO = '" & GridDetails.Cell(1, 2).Text & "'")
    If Not rsSONum.EOF And Not rsSONum.BOF Then
        MsgSpeechBox "Transaction cannot be saved." & vbCrLf & "Same Order Number and SEQ. No. already exist." & vbCrLf & " Pls. update the Sequence Number."
        Screen.MousePointer = 0
        Exit Sub
    End If
    Dim BACK_ORDER, TOTAL_ALLOCATED, ORDERED_QTY                As Integer
    Dim rsBackOrder                                             As ADODB.Recordset
    TOTAL_ALLOCATED = 0: BACK_ORDER = 0
    Set rsBackOrder = New ADODB.Recordset
    Set rsBackOrder = gconDMIS.Execute("SELECT QTY_ORDERED,QTY_ALLOCATED,QTY_BACKORDER FROM PMIS_PO_DETAILS WHERE SONUM = '" & txtPONum & "' AND STOCKNO = '" & GridDetails.Cell(1, 2).Text & "'")
    If Not rsBackOrder.EOF And Not rsBackOrder.BOF Then
       ORDERED_QTY = NumericVal(rsBackOrder!QTY_ORDERED)
       rsBackOrder.MoveFirst
       Do While Not rsBackOrder.EOF
            TOTAL_ALLOCATED = TOTAL_ALLOCATED + NumericVal(rsBackOrder!Qty_Allocated)
            rsBackOrder.MoveNext
       Loop
       TOTAL_ALLOCATED = TOTAL_ALLOCATED + NumericVal(GridDetails.Cell(1, 5).Text)
       If ORDERED_QTY < TOTAL_ALLOCATED Then
            MsgBox "Transaction cannot be proceed." & vbCrLf & "Qty. to be confirmed for Part Number '" & GridDetails.Cell(1, 2).Text & "' exceeded the remaining Back-Order Qty.", vbCritical
            Screen.MousePointer = 0
            Exit Sub
       Else
            BACK_ORDER = ORDERED_QTY - TOTAL_ALLOCATED
       End If
    End If
    
    Dim rsUpdatePO As ADODB.Recordset
    Set rsUpdatePO = New ADODB.Recordset
    Set rsUpdatePO = gconDMIS.Execute("INSERT INTO PMIS_PO_DETAILS " & _
                                    " (PO_NO,ITEMNO,SONUM,STOCKNO,SEQ_NO,PODATE,CONFIRMEDDATE,Qty_Ordered,Qty_Unserved,Qty_Allocated,Qty_Served,BackOrderAmt,Qty_BackOrder) " & _
                                    " VALUES (" & N2Str2Null(txtTranno.Text) & _
                                      "," & N2Str2Null(GridDetails.Cell(1, 1).Text) & _
                                      "," & N2Str2Null(txtPONum.Text) & _
                                      "," & N2Str2Null(GridDetails.Cell(1, 2).Text) & _
                                      "," & N2Str2Null(cboSEQ_NO.Text) & "," & N2Str2Null(txtPODate.Text) & _
                                      "," & N2Str2Null(txtConfirmedDate.Text) & _
                                      "," & NumericVal(GridDetails.Cell(1, 4).Text) & _
                                      "," & BACK_ORDER & _
                                      "," & NumericVal(GridDetails.Cell(1, 5).Text) & _
                                      "," & NumericVal(GridDetails.Cell(1, 6).Text) & _
                                      "," & CALCULATE_CURR_BO(CInt(BACK_ORDER)) & "," & BACK_ORDER & ")")
    ShowSuccessFullyUpdated
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initMemvars
    InitGridHeader
    InitGridDetails
End Sub
Sub initMemvars()
    txtPONum.Text = ""
    txtConfirmedDate.Text = ""
    txtPODate.Text = ""
    txtPONum.Text = ""
    txtPOType.Text = ""
    txtTranno.Text = ""
    cboSEQ_NO.Clear
End Sub
Sub InitGridHeader()
    With GridHeader
        .Cols = 10
        .Rows = 2

        .Cell(0, 0).Text = "ID"
        .Column(0).Width = 0

        .Cell(0, 1).Text = "PO Date"
        .Column(1).Width = 65
        .Column(1).Locked = True

        .Cell(0, 2).Text = "SO Number"
        .Column(2).Width = 85
        .Column(2).Locked = True
        
        .Cell(0, 3).Text = "Ordered Parts"
        .Column(3).Width = 90
        .Column(3).Locked = True

        .Cell(0, 4).Text = "Order Type"
        .Column(4).Width = 150
        .Column(4).Locked = True
        .Column(4).Alignment = cellCenterGeneral

        .Cell(0, 5).Text = "Ordered Qty"
        .Column(5).Alignment = cellCenterGeneral
        .Column(5).Locked = True
        .Column(5).Width = 70

        .Cell(0, 6).Text = "Confirmed Qty"
        .Column(6).Alignment = cellCenterGeneral
        .Column(6).Locked = True
        .Column(6).Width = 75

        .Cell(0, 7).Text = "Supplied Qty"
        .Column(7).Alignment = cellCenterGeneral
        .Column(7).Locked = True
        .Column(7).Width = 70

        .Cell(0, 8).Text = "B/O Qty."
        .Column(8).Alignment = cellCenterGeneral
        .Column(8).Locked = True
        .Column(8).Width = 50

        .Cell(0, 9).Text = "Last Confirmed"
        .Column(9).Alignment = cellRightGeneral
        .Column(9).Locked = True
        .Column(9).Width = 80
    End With
End Sub
Sub InitGridDetails()
    With GridDetails
        .Cols = 10
        .Rows = 2

        .Cell(0, 0).Text = "ID"
        .Column(0).Width = 0

        .Cell(0, 1).Text = "L/N"
        .Column(1).Width = 40

        .Cell(0, 2).Text = "Part Number"
        .Column(2).Alignment = cellLeftGeneral
        .Column(2).Locked = True
        .Column(2).Width = 110

        .Cell(0, 3).Text = "Part Description"
        .Column(3).Alignment = cellLeftGeneral
        .Column(3).Locked = True
        .Column(3).Width = 170

        .Cell(0, 4).Text = "Ordered"
        .Column(4).Width = 50
        .Column(4).Locked = True
        .Column(4).Alignment = cellCenterGeneral

        .Cell(0, 5).Text = "Confirmed"
        .Column(5).Alignment = cellCenterGeneral
        .Column(5).Width = 60

        .Cell(0, 6).Text = "Supplied"
        .Column(6).Alignment = cellCenterGeneral
        .Column(6).Width = 50

        .Cell(0, 7).Text = "Back Order"
        .Column(7).Alignment = cellCenterGeneral
        .Column(7).Locked = True
        .Column(7).Width = 65

        .Cell(0, 8).Text = "BO Amount"
        .Column(8).Alignment = cellRightGeneral
        .Column(8).Locked = True
        .Column(8).Width = 80

        .Cell(0, 9).Text = "Confirmed Date"
        .Column(9).Alignment = cellRightGeneral
        .Column(9).Locked = True
        .Column(9).Width = 90
    End With
End Sub

Function GETPOTYPE(thanzky As String) As String
    Dim cj As String
    cj = Mid(thanzky, 3, 1)
    Select Case cj
    Case "R"
        GETPOTYPE = "REGULAR ORDER"
    Case "E"
        GETPOTYPE = "EMERGENCY ORDER"
    Case "W"
        GETPOTYPE = "WARRANTY ORDER"
    Case "A"
        GETPOTYPE = "ADVANCE ORDER"
    Case "S"
        GETPOTYPE = "SPECIAL ORDER"
    Case "V"
        GETPOTYPE = "VEHICLE OFF-ROAD"
    End Select
End Function


Private Sub GridDetails_CellChange(ByVal Row As Long, ByVal Col As Long)
    Dim rsBOValue As ADODB.Recordset
    Set rsBOValue = New ADODB.Recordset

End Sub
Sub GridHeader_Click()
    'GridHeader.Range(GridHeader.ActiveCell.Row, 1, GridHeader.ActiveCell.Row, 8).BackColor = vbRed
    On Error GoTo ERROR_CODE
    If GridHeader.Rows > 1 Then
        VIEWDETAILS GridHeader.Cell(GridHeader.ActiveCell.Row, 0).Text
    aydi = GridHeader.Cell(GridHeader.ActiveCell.Row, 0).Text
    End If
Exit Sub
ERROR_CODE:
    err.Clear
        Exit Sub
End Sub
Sub VIEWDETAILS(di As Integer)
    Dim rsHeaderValue As ADODB.Recordset
    Set rsHeaderValue = New ADODB.Recordset
    Set rsHeaderValue = gconDMIS.Execute("SELECT PD.PO_NO, PD.ITEMNO, PD.EMERGENCY, PD.SOMonth, " & _
                                       " PD.SOYear, PD.SONum, PD.Qty_Ordered, PD.Qty_Allocated, " & _
                                       " PD.Qty_Served, PD.Qty_Unserved, PD.POFill, PD.POKill," & _
                                       " PD.Qty_BackOrder,PD.OrderAmount, PD.AllocAmount," & _
                                       " PD.BackOrderAmt," & _
                                       " PD.STATUS,PD.ID, PD.CONFIRMEDDATE, " & _
                                       " PD.SEQ_NO,PD.UNITPRICE, PD.USERCODE," & _
                                       " AllDayTran.TRANDATE," & _
                                       " AllDayTran.TranType, AllDayTran.Type, AllDayTran.STATUS, PartMas.STOCKDESC, PartMas.STOCKNO " & _
                                       " FROM dbo.PMIS_AllDayTran AllDayTran INNER JOIN " & _
                                       " dbo.PMIS_PartMas AS PartMas ON AllDayTran.TYPE = PartMas.TYPE AND AllDayTran.STOCK_ORD = PartMas.STOCKNO INNER JOIN " & _
                                       " dbo.PMIS_Po_Details AS PD ON AllDayTran.TRANNO = PD.PO_NO AND AllDayTran.ITEMNO = PD.ITEMNO " & _
                                       " WHERE AllDayTran.TYPE = 'P' AND AllDayTran.TRANTYPE = 'PO' AND " & _
                                       " AllDayTran.STATUS = 'P' AND PD.SONum = " & N2Str2Null(Trim(txtPONum)) & " AND PD.ID = " & di)
    If Not rsHeaderValue.EOF And Not rsHeaderValue.BOF Then
        GridDetails.Rows = 1
        GridDetails.Refresh
        GridDetails.AutoRedraw = True
        Do While Not rsHeaderValue.EOF
            GridDetails.AddItem Format(Null2String(rsHeaderValue!itemno), "0000") & Chr(9) & _
                                Null2String(rsHeaderValue!STOCKNO) & Chr(9) & _
                                Null2String(rsHeaderValue!STOCKDESC) & Chr(9) & _
                                NumericVal(rsHeaderValue!QTY_ORDERED) & Chr(9) & _
                                NumericVal(rsHeaderValue!Qty_Allocated) & Chr(9) & _
                                NumericVal(rsHeaderValue!Qty_Served) & Chr(9) & _
                                NumericVal(rsHeaderValue!QTY_BACKORDER) & Chr(9) & _
                                NumericVal(rsHeaderValue!BackOrderAmt) & Chr(9) & _
                                Null2Date(rsHeaderValue!CONFIRMEDDATE)
            rsHeaderValue.MoveNext
        Loop
        GridDetails.Refresh
        GridDetails.AutoRedraw = False
    End If
End Sub

Private Sub txtPONum_Change()
    If Trim(txtPONum) = "" Then
        cboSEQ_NO.Enabled = False
    Else
        cboSEQ_NO.Enabled = True
    End If
End Sub
