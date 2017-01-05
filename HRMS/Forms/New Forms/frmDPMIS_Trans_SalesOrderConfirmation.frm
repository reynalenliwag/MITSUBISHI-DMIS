VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmHRMS_LeaveEmployeeEffectivity 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Leave Settings"
   ClientHeight    =   6690
   ClientLeft      =   1110
   ClientTop       =   2625
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   12240
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picImport 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   30
      ScaleHeight     =   525
      ScaleWidth      =   9735
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import Employee Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   3105
      End
      Begin VB.Label Label2 
         Caption         =   "** New Employee Information Available Click Import to Update Information **"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3270
         TabIndex        =   19
         Top             =   120
         Width           =   6345
      End
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1620
      TabIndex        =   13
      Top             =   8340
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   12225
      TabIndex        =   0
      Top             =   480
      Width           =   12225
      Begin FlexCell.Grid Grid1 
         Height          =   5175
         Left            =   0
         TabIndex        =   1
         Top             =   30
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   9128
         AllowUserResizing=   0   'False
         Appearance      =   0
         BackColor2      =   12907725
         BackColorFixed  =   14737632
         BackColorBkg    =   8421504
         BackColorScrollBar=   14737632
         BackColorSel    =   8388608
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         ShowResizeTips  =   0   'False
         ReadOnlyFocusRect=   0
         Rows            =   30
         ScrollBarStyle  =   0
         SelectionMode   =   1
         EnterKeyMoveTo  =   1
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12195
      TabIndex        =   2
      Top             =   0
      Width           =   12225
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   7590
         TabIndex        =   14
         Top             =   30
         Width           =   4425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6630
         TabIndex        =   15
         Top             =   90
         Width           =   855
      End
   End
   Begin VB.PictureBox picPOC 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   3930
      ScaleHeight     =   1305
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   2250
      Visible         =   0   'False
      Width           =   5895
      Begin wizProgBar.Prg Prg1 
         Height          =   435
         Left            =   150
         TabIndex        =   4
         Top             =   450
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   767
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":058A
         BackColor       =   16777215
         ForeColor       =   0
         Appearance      =   0
         BarPicture      =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":05A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label labPOC 
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   180
         Width           =   5595
      End
   End
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   8220
      ScaleHeight     =   885
      ScaleWidth      =   3930
      TabIndex        =   10
      Top             =   5790
      Visible         =   0   'False
      Width           =   3930
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3150
         MouseIcon       =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":05C2
         MousePointer    =   99  'Custom
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2460
         MouseIcon       =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":0A52
         MousePointer    =   99  'Custom
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":0BA4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdd 
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
      Height          =   870
      Left            =   8340
      ScaleHeight     =   870
      ScaleWidth      =   3975
      TabIndex        =   7
      Top             =   5790
      Width           =   3975
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3030
         MouseIcon       =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":0EF4
         MousePointer    =   99  'Custom
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":1046
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2340
         MouseIcon       =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":13AC
         MousePointer    =   99  'Custom
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":14FE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1650
         MouseIcon       =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":1864
         MousePointer    =   99  'Custom
         Picture         =   "frmDPMIS_Trans_SalesOrderConfirmation.frx":19B6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Label LABID 
      Caption         =   "Label2"
      Height          =   315
      Left            =   12630
      TabIndex        =   6
      Top             =   60
      Width           =   705
   End
End
Attribute VB_Name = "frmHRMS_LeaveEmployeeEffectivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLeaves                                        As ADODB.Recordset
Attribute rsLeaves.VB_VarUserMemId = 1073938432
Dim rsNewEmployee                                   As ADODB.Recordset


Private Sub cmdCancel_Click()
    picSave.Visible = False
    picAdd.Visible = True

    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    Grid1.Column(9).Locked = True
    Grid1.Column(10).Locked = True
    
    FillGrid
End Sub

Private Sub cmdEdit_Click()
    picSave.Visible = True
    picAdd.Visible = False
    Grid1.Column(6).Locked = False
    Grid1.Column(7).Locked = False
    Grid1.Column(8).Locked = False
    Grid1.Column(9).Locked = False
    Grid1.Column(10).Locked = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdUpload_OK_Click()
    If IsDate(txtDate_FPL.Value) = False Then
        MsgBox "Cannot Confirm Transaction without Valid Date of Conirmation" & vbCrLf & "Please Edit Transaction.", vbCritical, "Error In Date"
        Exit Sub
    End If

    If LTrim(RTrim(txtFPLNo.Text)) = "" Then
        MsgBox "Cannot Confirm Transaction without Valid SOC Number" & vbCrLf & "Please Edit Transaction.", vbCritical, "Error In Date"
        Exit Sub
    End If

    Dim I                                           As Integer
    Dim RSPICKLISTDET                               As ADODB.Recordset
    Dim ID                                          As Integer
    Dim TOTAL_AMT                                   As Double
    Dim BK_QTY                                      As Integer
    Dim CONF_QTY                                    As Integer
    Dim ACT_RATE                                    As Double
    Dim DET_AMT                                     As Double
    Dim MAXVAL                                      As Long
    Dim STOCK_SUP                                   As String
    Dim DATE_FPICK                                  As String
    Dim PICKLIST_NO                                 As String
    Dim ETA

    MAXVAL = Grid1.Rows - 1
    Prg1.Max = MAXVAL
    PICKLIST_NO = txtFPLNo
    DATE_FPICK = txtDate_FPL.Value

    For I = 1 To Grid1.Rows - 1
        Prg1.Value = I
        labPOC = FormatNumber((I / MAXVAL) * 100)
        ID = Grid1.Cell(I, 10).SingleValue
        CONF_QTY = Grid1.Cell(I, 5).SingleValue
        BK_QTY = Grid1.Cell(I, 6).SingleValue
        ACT_RATE = Grid1.Cell(I, 7).DoubleValue
        DET_AMT = Grid1.Cell(I, 8).DoubleValue
        STOCK_SUP = Grid1.Cell(I, 1).Text
        ETA = N2Str2Null(Grid1.Cell(I, 11).Text)
        If IsDate(ETA) = True Then
            PRINTWITHETA = True
        End If
        gconDMIS.Execute ("UPDATE PMIS_SO_DET SET" _
                        & "  DATE_PPA=NULL, DATE_FPICKLIST= " & N2Str2Null(DATE_FPICK) _
                        & " , PICKLIST_NO= " & PICKLIST_NO _
                        & " , CONF_QTY = " & CONF_QTY _
                        & " , BK_QTY = " & BK_QTY _
                        & " , ACT_RATE= " & ACT_RATE _
                        & " , ETA= " & ETA _
                        & " , DET_AMT = " & DET_AMT & " WHERE ID=" & ID)

        gconDMIS.Execute ("UPDATE PMIS_SO_DET SET  DATE_PPA=NULL,DATE_PICKLIST=DATE_FPICKLIST   WHERE ID=" & ID & " AND DATE_PICKLIST IS NULL")
        TOTAL_AMT = TOTAL_AMT + DET_AMT
    Next

    gconDMIS.Execute ("UPDATE PMIS_SO_HDR SET  DATE_PPA=NULL,  DATE_FPICKLIST= " & N2Str2Null(DATE_FPICK) & " ,PICKLIST_NO= " & PICKLIST_NO & "  WHERE TRANNO='" & TRANNO & "'")
    gconDMIS.Execute ("UPDATE PMIS_SO_HDR SET  DATE_PPA=NULL, DATE_PICKLIST= DATE_FPICKLIST  WHERE TRANNO='" & TRANNO & "' AND DATE_PICKLIST IS NULL")


    ShowPictureBox picConfirmation, False, picMain


    RaiseEvent ConfirmedSO(rsLeaves!TRANNO)
End Sub

Sub FillGrid()
    Grid1.Rows = 1
    Dim I                                           As Integer
    Set rsLeaves = New ADODB.Recordset

    rsLeaves.Open "SELECT HRMS_EMPINFO.SEX, HRMS_EMPINFO.LASTNAME + ',' + HRMS_EMPINFO.FIRSTNAME + LEFT( ISNULL('.' + HRMS_EMPINFO.MIDDLENAME,'') ,1) AS EMPNAME ,HRMS_EMPINFO.DATEHIRED, HRMS_EMPINFO.POSITION,HRMS_EMPINFO.EMPLEVEL,HRMS_LEAVE_START.* FROM HRMS_LEAVE_START INNER JOIN HRMS_EMPINFO ON HRMS_EMPINFO.EMPNO=HRMS_LEAVE_START.EMPNO  ORDER BY 1 ASC", gconDMIS, adOpenKeyset, adLockReadOnly
    Grid1.AutoRedraw = False
    If Not rsLeaves.EOF Or Not rsLeaves.BOF Then
        While Not rsLeaves.EOF
            I = I + 1
            Grid1.AddItem _
                    rsLeaves!EMPNO & Chr(9) & _
                                   (rsLeaves!EMPNAME) & Chr(9) & _
                                   Format(rsLeaves!DateHired, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!Position) & Chr(9) & _
                                   Format(rsLeaves!EMPLEVEL) & Chr(9) & _
                                   Format(rsLeaves!SL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!VL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!EL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!ML_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!PL_START, "MM/DD/YYYY") & Chr(9) & _
                                   "DELETE" & Chr(9) & rsLeaves!ID, False
            If Null2String(rsLeaves!SEX) = "M" Then
                Grid1.Cell(I, 9).Locked = True
            End If
          
            
            Grid1.Range(I, 11, I, 11).FontBold = True
            Grid1.Range(I, 11, I, 11).ForeColor = vbBlue

            rsLeaves.MoveNext
        Wend
    End If
    Grid1.Refresh
    Grid1.AutoRedraw = True

    Set rsNewEmployee = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO NOT IN( SELECT EMPNO FROM HRMS_LEAVE_START) AND ACTIVEINACTIVE='A'")
    If Not (rsNewEmployee.EOF Or rsNewEmployee.BOF) Then
        picImport.Visible = True
    Else
        picImport.Visible = False
    End If

End Sub


Private Sub cmdImport_Click()
    If MsgBox("Are you Sure You want to Import Employee Information", vbInformation + vbYesNo) = vbNo Then Exit Sub
    rsNewEmployee.MoveFirst
    While Not rsNewEmployee.EOF
        gconDMIS.Execute ("insert into HRMS_LEAVE_START (EMPNO)VALUES(" & N2Str2Null(rsNewEmployee!EMPNO) & ")")
        rsNewEmployee.MoveNext
    Wend
    FillGrid
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Are you Sure You want to Update Information", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Dim LineId                                      As Integer
    Dim I                                           As Integer
    For I = 1 To Grid1.Rows - 1

        LineId = Grid1.Cell(I, 12).Text
        gconDMIS.Execute ("UPDATE HRMS_LEAVE_START SET " & _
                        " SL_START=" & N2Str2Null(Grid1.Cell(I, 6).Text) & _
                          ",VL_START=" & N2Str2Null(Grid1.Cell(I, 7).Text) & _
                          ",EL_START=" & N2Str2Null(Grid1.Cell(I, 8).Text) & _
                          ",ML_START=" & N2Str2Null(Grid1.Cell(I, 9).Text) & _
                          ",PL_START=" & N2Str2Null(Grid1.Cell(I, 10).Text) & _
                        " WHERE ID=" & LineId)

    Next
    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call initGrid
    FillGrid

End Sub


Sub initGrid()
    Dim rg                                          As FlexCell.Range
    With Grid1
        .Cols = 13
        .Rows = 1
        .FixedCols = 4

        Set rg = .Range(0, 0, 0, 12)
        rg.WrapText = True
        rg.Alignment = cellCenterCenter

        .DisplayFocusRect = False
        .AllowUserResizing = True


        .Cell(0, 0).Text = "L/N"
        .Column(0).WIDTH = 25

        .Cell(0, 1).Text = "EMPLOYEE NUMBER"
        .Column(1).WIDTH = 60
        .Column(1).Locked = True

        .Cell(0, 2).Text = "EMPLOYEE NAME"
        .Column(2).WIDTH = 145
        .Column(2).Locked = True
        .Column(2).Alignment = cellLeftGeneral

        .Cell(0, 3).Text = "DATE HIRED"
        .Column(3).WIDTH = 64
        .Column(3).Alignment = cellLeftGeneral
        .Column(3).Locked = True

        .Cell(0, 4).Text = "POSITION"
        .Column(4).WIDTH = 160
        .Column(4).Alignment = cellLeftGeneral
        .Column(4).Locked = True

        .Cell(0, 5).Text = "LEVEL"
        .Column(5).WIDTH = 40
        .Column(5).Alignment = cellLeftGeneral
        .Column(5).Locked = True



        .Cell(0, 6).Text = "S/L"
        .Column(6).WIDTH = 70
        .Column(6).Alignment = cellLeftGeneral
        .Column(6).Locked = True
        .Column(6).CellType = cellCalendar


        .Cell(0, 7).Text = "V/L"
        .Column(7).WIDTH = 70
        .Column(7).Alignment = cellLeftGeneral
        .Column(7).Locked = True
        .Column(7).CellType = cellCalendar



        .Cell(0, 8).Text = "E/L"
        .Column(8).WIDTH = 70
        .Column(8).Alignment = cellLeftGeneral
        .Column(8).Locked = True
        .Column(8).CellType = cellCalendar

        .Cell(0, 9).Text = "M/L"
        .Column(9).WIDTH = 70
        .Column(9).Locked = True
        .Column(9).Alignment = cellLeftGeneral
        .Column(9).CellType = cellCalendar





        .Cell(0, 10).Text = "P/L"
        .Column(10).WIDTH = 70
        .Column(10).Locked = True
        .Column(10).CellType = cellCalendar
        .Column(10).Alignment = cellLeftGeneral


        .Cell(0, 11).Text = "OPTION"
        .Column(11).WIDTH = 60
        .Column(11).CellType = cellTextBox
        .Column(11).Locked = True




        .Cell(0, 12).Text = "ID"
        .Column(12).WIDTH = 0






    End With
End Sub



Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
On Error Resume Next
    If IsDate(Grid1.ActiveCell.Text) = True And IsDate(Grid1.Cell(Grid1.ActiveCell.Row, 3).Text) = True Then
        If Not (DateDiff("D", Grid1.ActiveCell.Text, Grid1.Cell(Grid1.ActiveCell.Row, 3).Text) <= 0) Then
            Grid1.ActiveCell.Text = ""
        End If
    Else
        Grid1.ActiveCell.Text = ""
    End If

End Sub
 
Private Sub Grid1_EditRow(ByVal Row As Long)
On Error Resume Next
    If IsDate(Grid1.ActiveCell.Text) = True And IsDate(Grid1.Cell(Grid1.ActiveCell.Row, 3).Text) = True Then
        If Not (DateDiff("D", Grid1.ActiveCell.Text, Grid1.Cell(Grid1.ActiveCell.Row, 3).Text) <= 0) Then
            Grid1.ActiveCell.Text = ""
        End If
    Else
        Grid1.ActiveCell.Text = ""
    End If
End Sub

 

Sub FillSearchGrid(xxx As String)
    Grid1.Rows = 1
    Dim I                                           As Integer
    Set rsLeaves = New ADODB.Recordset

    rsLeaves.Open "SELECT HRMS_EMPINFO.SEX, HRMS_EMPINFO.LASTNAME + ',' + HRMS_EMPINFO.FIRSTNAME + LEFT( ISNULL('.' + HRMS_EMPINFO.MIDDLENAME,'') ,1) AS EMPNAME ,HRMS_EMPINFO.DATEHIRED, HRMS_EMPINFO.POSITION,HRMS_EMPINFO.EMPLEVEL,HRMS_LEAVE_START.* FROM HRMS_LEAVE_START INNER JOIN HRMS_EMPINFO ON HRMS_EMPINFO.EMPNO=HRMS_LEAVE_START.EMPNO  WHERE HRMS_EMPINFO.LASTNAME LIKE '" & Repleys(txtSearch) & "%' ORDER BY 1 ASC", gconDMIS, adOpenKeyset, adLockReadOnly
    Grid1.AutoRedraw = False
    If Not rsLeaves.EOF Or Not rsLeaves.BOF Then
        While Not rsLeaves.EOF
            I = I + 1
            Grid1.AddItem _
                    rsLeaves!EMPNO & Chr(9) & _
                                   (rsLeaves!EMPNAME) & Chr(9) & _
                                   Format(rsLeaves!DateHired, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!Position) & Chr(9) & _
                                   Format(rsLeaves!EMPLEVEL) & Chr(9) & _
                                   Format(rsLeaves!SL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!VL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!EL_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!ML_START, "MM/DD/YYYY") & Chr(9) & _
                                   Format(rsLeaves!PL_START, "MM/DD/YYYY") & Chr(9) & _
                                   "DELETE" & Chr(9) & rsLeaves!ID, False
            If Null2String(rsLeaves!SEX) = "M" Then
                Grid1.Cell(I, 9).Locked = True
            End If
            
            Grid1.Range(I, 11, I, 11).FontBold = True
            Grid1.Range(I, 11, I, 11).ForeColor = vbBlue

            rsLeaves.MoveNext
        Wend
    End If
    Grid1.Refresh
    Grid1.AutoRedraw = True
 
End Sub
 



Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FillSearchGrid txtSearch
    
    End If
End Sub
