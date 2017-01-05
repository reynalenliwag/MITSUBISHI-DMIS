VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTBUploadingTool 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer's TB Uploading Tool"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   Icon            =   "frmTBUploadingTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   8625
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   -10080
      Picture         =   "frmTBUploadingTool.frx":07AA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   -120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picChecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   -9720
      Picture         =   "frmTBUploadingTool.frx":0894
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   -120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmbClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear Uploading"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   2055
   End
   Begin VB.CommandButton cmdStartUploading 
      Appearance      =   0  'Flat
      Caption         =   "Start Uploading"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton cmdBrowse 
      Appearance      =   0  'Flat
      Caption         =   "Browse Excel File"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtPath 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin MSFlexGridLib.MSFlexGrid grdUpload 
      Height          =   6435
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   11351
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColor       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   16711680
      ForeColorSel    =   16777215
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   8160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Picture         =   "frmTBUploadingTool.frx":097E
      ForeColor       =   0
      BarPicture      =   "frmTBUploadingTool.frx":099A
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   120520707
      CurrentDate     =   42241
   End
   Begin MSForms.Label lblSumCredit 
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   7800
      Width           =   1695
      ForeColor       =   8388608
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2990;450"
      FontName        =   "Segoe UI"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblSumDebit 
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   7800
      Width           =   1695
      ForeColor       =   8388608
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2990;450"
      FontName        =   "Segoe UI"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cut Off:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblRows 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Rows:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   8520
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Sheet:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin MSForms.ComboBox cmbSheet 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   2535
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4471;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Segoe UI"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTBUploadingTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       ´$$$$`                             ,,,
'       ´$$$$$$$`                         ´$$$`
'        `$$$$$$$`      ,,        ,,      ´$$$$´
'         `$$$$$$$`    ´$$`     ´$$`    ´$$$$$´
'          `$$$$$$$`_´$$$$$`_´$$$$$`__´$$$$$$$´
'           `$$$$$$$_$$$$$$$_$$$$$$$_´$$$$$$$´
'            `$$$$$$_$$$$$$$_$$$$$$$`´$$$$$$´
'   ,,,,,    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
' ´$$$$$`    `$$$$$$_$$$$$$$_$$$$$$$_$$$$$$´
'´$$$$$$$$$`´$$$$$$$_SARAJOY_$$$$$$$_$$$$$´
'´$$$$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$$_$$$$$´
'   `$$$$$$$$$$$$$$$_$$$$$$$_$$$$$$_$$$$$$´
'      `$$$$$$$$$$$$$_$$$$$__$$_$$$$$$_$$´
'       `$$$$$$$$$$$$$__,$$$$,_____,$$$$$´
'         `$$$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'          `$$$$$$$$$$$$$$$$$$$$$$$$$$$´
'            `$$$$$$$$CREATED$ON$$$$$$´
'               `$$$$$$MAY172016$$$$$´
'
'              ROCK 'N ROLL TO THE WORLD

Option Explicit

Dim xlObject              As Excel.Application
Dim xlWB                  As Excel.Workbook

Dim CMD                   As New ADODB.Command
Dim cnExcel               As New ADODB.Connection

Dim rsExtractFromExcel    As New ADODB.Recordset
Dim rsCheckTable          As New ADODB.Recordset

Dim cnExcelStr            As String
Dim xSheetNeym            As String

Private Sub Form_Load()
    Screen.MousePointer = 0
    CenterMe frmMain, Me, 1
    initGrid
End Sub

Sub initGrid()
    With grdUpload
        .Rows = 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        
        .Row = 0
        .Col = 0
        .Col = 1: .Text = "ACCTCODE"
        .Col = 2: .Text = "ACCTNAME"
        .Col = 3: .Text = "DEBIT"
        .Col = 4: .Text = "CREDIT"
    End With
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.filter = ""
    CommonDialog1.filter = "Excel File (*.xlsx)|*.xlsx|Excel File (*.xls)|*.xls"
    
    CommonDialog1.ShowOpen
    txtPath.Text = CommonDialog1.FileName
    cmbSheet.SetFocus
End Sub

Private Sub cmbSheet_GotFocus()
On Error GoTo ErrHandler
    If txtPath.Text = "" Then Exit Sub
    
    Dim oSheet
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Open(txtPath.Text)

    cmbSheet.Clear
    
    For Each oSheet In xlWB.Sheets
        cmbSheet.AddItem (oSheet.Name)
    Next
    
    xlObject.Workbooks.Close
    Exit Sub
    
ErrHandler:
    MsgBox "There is a problem opening the workbook.", vbCritical, "Error"
End Sub

Private Sub cmbSheet_Change()
    If cmbSheet.Value = "" Then Exit Sub
    Call LoadGrid(txtPath.Text, cmbSheet.Value)
    DTPicker1.SetFocus
End Sub

Private Sub DTPicker1_Change()
    cmdStartUploading.SetFocus
End Sub

Sub LoadGrid(xPath, xSheetname As String)
    initGrid
    
    Select Case Right(xPath, 4)
        Case ".xls"
            cnExcelStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xPath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
        Case "xlsx"
            cnExcelStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & xPath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=2"""
    End Select
    
    Set cnExcel = New ADODB.Connection
    cnExcel.Open cnExcelStr
    
    
    Set rsExtractFromExcel = New ADODB.Recordset
    xSheetname = "[" & xSheetname & "$]"
    rsExtractFromExcel.Open "SELECT * FROM " & xSheetname & "", cnExcel, adOpenForwardOnly, adLockReadOnly
    
    If Not rsExtractFromExcel.EOF And Not rsExtractFromExcel.BOF Then
        Dim i As Integer
        i = 0
        
        rsExtractFromExcel.MoveFirst
        Do While Not rsExtractFromExcel.EOF
            grdUpload.AddItem Chr(9) & _
              rsExtractFromExcel!ACCT_CODE & Chr(9) & _
              rsExtractFromExcel!ACCOUNT_TITLE & Chr(9) & _
              ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!Debit)) & Chr(9) & _
              ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!Credit)) & Chr(9)
                                                             
            i = i + 1
            grdUpload.Row = i: grdUpload.Col = 0
            grdUpload.CellPictureAlignment = 4
            Set grdUpload.CellPicture = picUnchecked.Picture
            lblRows.Caption = i
            
            lblSumDebit.Caption = Val(lblSumDebit.Caption) + N2Str2Zero(rsExtractFromExcel!Debit)
            lblSumCredit.Caption = Val(lblSumCredit.Caption) + N2Str2Zero(rsExtractFromExcel!Credit)
            
        rsExtractFromExcel.MoveNext
        Loop
    End If
    
    lblSumDebit.Caption = ToDoubleNumber(N2Str2Zero(lblSumDebit.Caption))
    lblSumCredit.Caption = ToDoubleNumber(N2Str2Zero(lblSumCredit.Caption))
End Sub

Private Sub cmdStartUploading_Click()
    If txtPath.Text = "" Or cmbSheet.Value = "" Then MsgBox "All fields are required", vbExclamation, "Information"
    
    If lblSumCredit.Caption <> lblSumDebit.Caption Then MsgBox "Trial balance is not balance.", vbCritical, "Information": Exit Sub
    
    If MsgBox("Are you sure that cut off date is " & DTPicker1.Value & "?", vbYesNo, "Question") = vbNo Then Exit Sub
    
    xSheetNeym = "": xSheetNeym = cmbSheet.Value & "$"
    
'CHECK AND CREATE TABLE
    lblStatus.Caption = "Status: Checking table..."
    Set rsCheckTable = New ADODB.Recordset
    rsCheckTable.Open "IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='" & xSheetNeym & "') SELECT 1 AS RESULT ELSE SELECT 0 AS RESULT", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    If Not rsCheckTable.EOF And Not rsCheckTable.BOF Then
        If rsCheckTable!RESULT = "1" Then
            MsgBox "Table already exists. This sheet may have been already uploaded to database.", vbExclamation, "Information"
            lblStatus.Caption = "Status: "
            Exit Sub
        Else
            gconDMIS.Execute ("CREATE TABLE [DBO].[" & xSheetNeym & "] " & _
                "([ACCT_CODE] [NVARCHAR](11) NULL, " & _
                "[ACCOUNT_TITLE] [NVARCHAR](255) NULL, " & _
                "[DEBIT] [FLOAT] NULL, " & _
                "[CREDIT] [FLOAT] NULL) " & _
                "ON [PRIMARY]")
        End If
    End If
    lblStatus.Caption = "Status: Table created..."

'GET DATA OF EXCEL FROM RECORDSET AND INSERT TO NEWLY CREATED TEMP TABLE
    If rsExtractFromExcel.State = adStateOpen Then
        rsExtractFromExcel.Requery 1
                    
        If Not rsExtractFromExcel.EOF And Not rsExtractFromExcel.BOF Then
            Dim i As Integer
            i = 0
            
            Prg1.Value = 0
            Prg1.Max = lblRows.Caption + 5
        
            rsExtractFromExcel.MoveFirst
            Do While Not rsExtractFromExcel.EOF

                gconDMIS.Execute ("INSERT INTO [" & xSheetNeym & "] ([ACCT_CODE],[ACCOUNT_TITLE],[DEBIT],[CREDIT]) VALUES " & _
                                    "( " & _
                                    " '" & rsExtractFromExcel!ACCT_CODE & "'," & _
                                    " '" & rsExtractFromExcel!ACCOUNT_TITLE & "'," & _
                                    " '" & (N2Str2Zero(rsExtractFromExcel!Debit)) & "'," & _
                                    " '" & (N2Str2Zero(rsExtractFromExcel!Credit)) & "')")
                
                i = i + 1
                grdUpload.Row = i: grdUpload.Col = 0
                grdUpload.CellPictureAlignment = 4
                Set grdUpload.CellPicture = picChecked.Picture
                
                If Prg1.Value = Prg1.Max - 5 Then
                Else
                    Prg1.Value = Prg1.Value + 1
                End If
            
                lblStatus.Caption = "Status: Uploading account code: [" & rsExtractFromExcel!ACCT_CODE & "] to temporary table...Row [" & i & "]"
                rsExtractFromExcel.MoveNext
            Loop
            
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows were uploaded to temporary table."
            
            'INSERT TO AMIS_JOURNAL HD USING THE EXISTING QUERY
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows currently transferring to trial balance."
            Call TB_STORED(DTPicker1.Value, cmbSheet.Value)
            Call DELETE_DEBIT_CREDIT_ZERO
            
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows were uploaded to Trial balance."
            Prg1.Value = Prg1.Value + 5
            If MsgBox(cmbSheet.Value & " successfully uploaded to Trial Balance." & vbCrLf & "would you like to clear uploading", vbInformation + vbYesNo, "Question") = vbYes Then cmbClear.Value = 1
        End If
        
    End If
    
    cmbClear.SetFocus
    cnExcel.Close
End Sub

Private Sub cmbClear_Click()
    initGrid
    txtPath.Text = ""
    cmbSheet.Clear
    lblStatus.Caption = "Status: "
    lblRows.Caption = "0"
    lblSumDebit.Caption = "0"
    lblSumCredit.Caption = "0"
    Prg1.Value = 0
    cmdBrowse.SetFocus
End Sub

Sub TB_STORED(yBegdate, ySheet As String)
    Dim SQL_STORED As String
    
    SQL_STORED = ""
    SQL_STORED = SQL_STORED & "DECLARE @J_JNO           AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_ITEMNO        AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_VOUCHERNO     AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_DATE          AS SMALLDATETIME   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_TYPE          AS VARCHAR(5)      " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @DEBIT           AS DECIMAL(18,2)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @CREDIT          AS DECIMAL(18,2)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTCODE        AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTDESC        AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @REMARKS         AS VARCHAR(100)    " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_VOUCHERNO = '000001'                 " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_DATE =  '" & yBegdate & "'           " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_TYPE = 'OPB'                         " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @REMARKS = 'BEG BALANCE ' + (CONVERT(VARCHAR(50),@J_DATE,107))   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT ISNULL(MAX(JNO),0) + 1 JNO FROM AMIS_JOURNAL_HD)   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT (REPLICATE('0',6-LEN(@J_JNO)) + @J_JNO))           " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_HD(JDATE,VOUCHERNO,JTYPE,JNO,STATUS,REMARKS,LASTUPDATE)   " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,'N',@REMARKS,GETDATE())   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @TB AS CURSOR    " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @TB = CURSOR FOR   " & vbCrLf
    SQL_STORED = SQL_STORED & "SELECT  ACCT_CODE,ISNULL(DEBIT,0.00) AS DEBIT,ISNULL(CREDIT,0.00) AS CREDIT FROM " & ySheet & "$ WHERE ACCT_CODE <> 'NULL'" & vbCrLf
    SQL_STORED = SQL_STORED & "OPEN @TB   " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @TB INTO @ACCTCODE,@DEBIT,@CREDIT   " & vbCrLf
    SQL_STORED = SQL_STORED & "WHILE @@FETCH_STATUS = 0   " & vbCrLf
    SQL_STORED = SQL_STORED & "BEGIN   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_ITEMNO = (SELECT ISNULL(MAX(JITEMNO),0) + 1 JITEMNO FROM AMIS_JOURNAL_DET WHERE JTYPE = @J_TYPE AND VOUCHERNO = @J_VOUCHERNO  )   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_ITEMNO = (SELECT (REPLICATE('0',4-LEN(@J_ITEMNO)) + @J_ITEMNO))   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @ACCTDESC = (SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE ACCTCODE=@ACCTCODE)    " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_DET(JITEMNO,JDATE,VOUCHERNO,JTYPE,JNO,STATUS,ACCT_CODE,ACCT_NAME,LASTUPDATE,DEBIT,CREDIT)   " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_ITEMNO,@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,'N',@ACCTCODE,@ACCTDESC,GETDATE(),@DEBIT,@CREDIT)   " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @TB INTO @ACCTCODE,@DEBIT,@CREDIT   " & vbCrLf
    SQL_STORED = SQL_STORED & "END   " & vbCrLf
    SQL_STORED = SQL_STORED & "CLOSE @TB   " & vbCrLf
    SQL_STORED = SQL_STORED & "DEALLOCATE @TB   "
    
    gconDMIS.Execute (SQL_STORED)
End Sub

Sub DELETE_DEBIT_CREDIT_ZERO()
    gconDMIS.Execute ("DELETE FROM AMIS_JOURNAL_DET WHERE JTYPE = 'OPB' AND ISNULL(DEBIT,0) = 0 AND ISNULL(CREDIT,0) = 0")
End Sub

