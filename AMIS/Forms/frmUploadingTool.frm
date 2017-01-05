VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUploadingTool 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer's Uploading Tool"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14775
   Icon            =   "frmUploadingTool.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   14775
   Begin VB.CommandButton cmbClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear Uploading"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8520
      Width           =   2055
   End
   Begin VB.PictureBox picChecked 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      Picture         =   "frmUploadingTool.frx":07AA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Picture         =   "frmUploadingTool.frx":0894
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   8160
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   661
      Picture         =   "frmUploadingTool.frx":097E
      ForeColor       =   0
      BarPicture      =   "frmUploadingTool.frx":099A
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
   Begin VB.CommandButton cmdStartUploading 
      Appearance      =   0  'Flat
      Caption         =   "Start Uploading"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3615
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdUpload 
      Height          =   6705
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11827
      _Version        =   393216
      Cols            =   12
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   13200
      TabIndex        =   13
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   417071107
      CurrentDate     =   42241
   End
   Begin MSForms.Label lblAcctCode 
      Height          =   255
      Left            =   10920
      TabIndex        =   17
      Top             =   7800
      Width           =   1695
      ForeColor       =   8388608
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2990;450"
      FontName        =   "Segoe UI Semibold"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   600
   End
   Begin MSForms.Label lblSumAmount 
      Height          =   255
      Left            =   12840
      TabIndex        =   16
      Top             =   7800
      Width           =   1695
      ForeColor       =   8388608
      BackColor       =   16777215
      Caption         =   "0"
      Size            =   "2990;450"
      FontName        =   "Segoe UI Semibold"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cut Off:"
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
      Left            =   12360
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin MSForms.ComboBox cmbChoose 
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Width           =   2535
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4471;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Segoe UI"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   12360
      TabIndex        =   11
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status:"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   8640
      Width           =   9855
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
      Left            =   13800
      TabIndex        =   9
      Top             =   8640
      Width           =   735
   End
   Begin MSForms.ComboBox cmbSheet 
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   480
      Width           =   2415
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4260;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   6
      FontName        =   "Segoe UI"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Sheet:"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmUploadingTool"
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
'               `$$$$$$MAY122016$$$$$´
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
    cmbChoose.AddItem "ACCOUNTS RECEIVABLE"
    cmbChoose.AddItem "ACCOUNTS PAYABLE"
End Sub

Sub initGrid()
    With grdUpload
        .Rows = 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1100
        .ColWidth(2) = 600
        .ColWidth(3) = 850: .ColWidth(4) = 2800
        .ColWidth(5) = 1100: .ColWidth(6) = 1600
        .ColWidth(7) = 1100: .ColWidth(8) = 1100
        .ColWidth(9) = 1100: .ColWidth(10) = 1100
        .ColWidth(11) = 1100
        
        .Row = 0
        .Col = 0
        .Col = 1: .Text = "ACCTCODE"
        .Col = 2: .Text = "ENTITYCODE"
        
        If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
        .Col = 3: .Text = "CUSCODE"
        ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
        .Col = 3: .Text = "SUPCODE"
        End If
        
        .Col = 4: .Text = "REFERENCENAME"
        .Col = 5: .Text = "INVOICEDATE"
        .Col = 6: .Text = "INVOICENO"
        .Col = 7: .Text = "INVOICETYPE"
        .Col = 8: .Text = "DUEDATE"
        .Col = 9: .Text = "AMOUNT"
        .Col = 10: .Text = "PAYMENT"
        .Col = 11: .Text = "BALANCE"
        
    End With
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
            
            If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
                grdUpload.AddItem Chr(9) & _
                  rsExtractFromExcel!AcctCode & Chr(9) & _
                  rsExtractFromExcel!ENTITYCODE & Chr(9) & _
                  rsExtractFromExcel!CUSCODE & Chr(9) & _
                  rsExtractFromExcel!REFERENCENAME & Chr(9) & _
                  rsExtractFromExcel!invoicedate & Chr(9) & _
                  rsExtractFromExcel!INVOICENO & Chr(9) & _
                  rsExtractFromExcel!INVOICETYPE & Chr(9) & _
                  rsExtractFromExcel!DUEDATE & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!amount)) & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!payment)) & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!BALANCE))
            ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
                grdUpload.AddItem Chr(9) & _
                  rsExtractFromExcel!AcctCode & Chr(9) & _
                  rsExtractFromExcel!ENTITYCODE & Chr(9) & _
                  rsExtractFromExcel!supcode & Chr(9) & _
                  rsExtractFromExcel!REFERENCENAME & Chr(9) & _
                  rsExtractFromExcel!invoicedate & Chr(9) & _
                  rsExtractFromExcel!INVOICENO & Chr(9) & _
                  rsExtractFromExcel!INVOICETYPE & Chr(9) & _
                  rsExtractFromExcel!DUEDATE & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!amount)) & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!payment)) & Chr(9) & _
                  ToDoubleNumber(N2Str2Zero(rsExtractFromExcel!BALANCE))
            End If
                                                             
            i = i + 1
            grdUpload.Row = i: grdUpload.Col = 0
            grdUpload.CellPictureAlignment = 4
            Set grdUpload.CellPicture = picUnchecked.Picture
            lblRows.Caption = i
            
            lblAcctCode.Caption = rsExtractFromExcel!AcctCode
            lblSumAmount.Caption = Val(lblSumAmount.Caption) + N2Str2Zero(rsExtractFromExcel!amount)
            
        rsExtractFromExcel.MoveNext
        Loop
    End If
        
    lblSumAmount.Caption = ToDoubleNumber(N2Str2Zero(lblSumAmount.Caption))
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.filter = ""
    CommonDialog1.filter = "Excel File (*.xlsx)|*.xlsx|Excel File (*.xls)|*.xls"
    
    CommonDialog1.ShowOpen
    txtPath.Text = CommonDialog1.FileName
    cmbChoose.SetFocus
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

Private Sub cmbChoose_Change()
    cmbSheet.SetFocus
End Sub

Private Sub cmdStartUploading_Click()
    If txtPath.Text = "" Or cmbSheet.Value = "" Or cmbChoose.Value = "" Then MsgBox "All fields are required", vbExclamation, "Information"
    
    If CheckIfScheduleAcct(lblAcctCode.Caption) = False Then MsgBox lblAcctCode.Caption & " is not a scheduled account.": Exit Sub
    
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
            If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
                gconDMIS.Execute ("CREATE TABLE [DBO].[" & xSheetNeym & "] " & _
                    "([ACCTCODE] [NVARCHAR](11) NULL, " & _
                    "[ENTITYCODE] [NVARCHAR](1) NULL, " & _
                    "[CUSCODE] [NVARCHAR](6) NULL, " & _
                    "[REFERENCENAME] [NVARCHAR](255) NULL, " & _
                    "[INVOICEDATE] [DATETIME] NULL, " & _
                    "[INVOICENO] [NVARCHAR](60) NULL, " & _
                    "[INVOICETYPE] [NVARCHAR](2) NULL, " & _
                    "[DUEDATE] [DATETIME] NULL, " & _
                    "[AMOUNT] [FLOAT] NULL, " & _
                    "[PAYMENT] [FLOAT] NULL, " & _
                    "[BALANCE] [FLOAT] NULL) " & _
                    "ON [PRIMARY]")
            ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
                gconDMIS.Execute ("CREATE TABLE [DBO].[" & xSheetNeym & "] " & _
                    "([ACCTCODE] [NVARCHAR](11) NULL, " & _
                    "[ENTITYCODE] [NVARCHAR](1) NULL, " & _
                    "[SUPCODE] [NVARCHAR](6) NULL, " & _
                    "[REFERENCENAME] [NVARCHAR](255) NULL, " & _
                    "[INVOICEDATE] [DATETIME] NULL, " & _
                    "[INVOICENO] [NVARCHAR](60) NULL, " & _
                    "[INVOICETYPE] [NVARCHAR](2) NULL, " & _
                    "[DUEDATE] [DATETIME] NULL, " & _
                    "[AMOUNT] [FLOAT] NULL, " & _
                    "[PAYMENT] [FLOAT] NULL, " & _
                    "[BALANCE] [FLOAT] NULL) " & _
                    "ON [PRIMARY]")
            End If
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
                
                If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
                    gconDMIS.Execute ("INSERT INTO [" & xSheetNeym & "] ([ACCTCODE],[ENTITYCODE],[CUSCODE],[REFERENCENAME],[INVOICEDATE],[INVOICENO],[INVOICETYPE],[DUEDATE],[AMOUNT],[PAYMENT],[BALANCE]) VALUES " & _
                                        "( " & _
                                        " '" & rsExtractFromExcel!AcctCode & "'," & _
                                        " '" & rsExtractFromExcel!ENTITYCODE & "'," & _
                                        " '" & rsExtractFromExcel!CUSCODE & "'," & _
                                        " '" & rsExtractFromExcel!REFERENCENAME & "'," & _
                                        " '" & rsExtractFromExcel!invoicedate & "'," & _
                                        " '" & rsExtractFromExcel!INVOICENO & "'," & _
                                        " '" & rsExtractFromExcel!INVOICETYPE & "'," & _
                                        " '" & rsExtractFromExcel!DUEDATE & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!amount)) & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!payment)) & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!BALANCE)) & "')")
                                        
                ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
                    gconDMIS.Execute ("INSERT INTO [" & xSheetNeym & "] ([ACCTCODE],[ENTITYCODE],[SUPCODE],[REFERENCENAME],[INVOICEDATE],[INVOICENO],[INVOICETYPE],[DUEDATE],[AMOUNT],[PAYMENT],[BALANCE]) VALUES " & _
                                        "( " & _
                                        " '" & rsExtractFromExcel!AcctCode & "'," & _
                                        " '" & rsExtractFromExcel!ENTITYCODE & "'," & _
                                        " '" & rsExtractFromExcel!supcode & "'," & _
                                        " '" & rsExtractFromExcel!REFERENCENAME & "'," & _
                                        " '" & rsExtractFromExcel!invoicedate & "'," & _
                                        " '" & rsExtractFromExcel!INVOICENO & "'," & _
                                        " '" & rsExtractFromExcel!INVOICETYPE & "'," & _
                                        " '" & rsExtractFromExcel!DUEDATE & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!amount)) & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!payment)) & "'," & _
                                        " '" & (N2Str2Zero(rsExtractFromExcel!BALANCE)) & "')")
                                        
                End If
                
                i = i + 1
                grdUpload.Row = i: grdUpload.Col = 0
                grdUpload.CellPictureAlignment = 4
                Set grdUpload.CellPicture = picChecked.Picture
                
                If Prg1.Value = Prg1.Max - 5 Then
                Else
                    Prg1.Value = Prg1.Value + 1
                End If
            
                lblStatus.Caption = "Status: Uploading invoice number [" & rsExtractFromExcel!INVOICENO & "] to temporary table..."
                rsExtractFromExcel.MoveNext
            Loop
            
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows were uploaded to temporary table."
            
            'INSERT TO AMIS_JOURNAL HD USING THE EXISTING QUERY
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows currently transferring to opening balance."
            If cmbChoose.Text = "ACCOUNTS RECEIVABLE" Then
                Call AR_STORED(DTPicker1.Value, cmbSheet.Value)
            ElseIf cmbChoose.Text = "ACCOUNTS PAYABLE" Then
                Call AP_STORED(DTPicker1.Value, cmbSheet.Value)
            End If
        
            lblStatus.Caption = "Status: " & lblRows.Caption & " rows were uploaded to opening balance."
            Prg1.Value = Prg1.Value + 5
            If MsgBox(cmbSheet.Value & " successfully uploaded to Opening Balance." & vbCrLf & "would you like to clear uploading", vbInformation + vbYesNo, "Question") = vbYes Then cmbClear.Value = 1
        End If
        
    End If
    
    cmbClear.SetFocus
    cnExcel.Close
End Sub

Function CheckIfScheduleAcct(xAcct) As Boolean
    Dim RSSCHEDACCT As ADODB.Recordset
    Set RSSCHEDACCT = New ADODB.Recordset
    RSSCHEDACCT.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = '" & xAcct & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    If Not RSSCHEDACCT.EOF And Not RSSCHEDACCT.BOF Then
        If RSSCHEDACCT!Is_Schedule_Accnt = True Then
            CheckIfScheduleAcct = True
        Else
            CheckIfScheduleAcct = False
        End If
    End If
End Function

Private Sub cmbClear_Click()
    initGrid
    txtPath.Text = ""
    cmbSheet.Clear
    cmbChoose.Clear
    cmbChoose.AddItem "ACCOUNTS RECEIVABLE"
    cmbChoose.AddItem "ACCOUNTS PAYABLE"
    
    lblStatus.Caption = "Status: "
    lblRows.Caption = "0"
    Prg1.Value = 0
    
    lblAcctCode.Caption = "0"
    lblSumAmount.Caption = "0"
    
    cmdBrowse.SetFocus
End Sub

Sub AR_STORED(yBegdate, ySheet As String)
    Dim SQL_STORED As String
    SQL_STORED = ""
    SQL_STORED = SQL_STORED & "DECLARE @J_JNO          AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_VOUCHERNO    AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_DATE         AS SMALLDATETIME   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_TYPE         AS VARCHAR(5)      " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @CUSCDE         AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTNAME       AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICEDATE    AS SMALLDATETIME   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @AMOUNTTOPAY    AS DECIMAL(18,2)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTCODE       AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTDESC       AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICETYPE    AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICENO      AS NVARCHAR(100)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ENTITYCODE     AS NVARCHAR(1)     " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @REMARKS        AS VARCHAR(100)    " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @AR AS CURSOR                      " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @AR = CURSOR FOR                       " & vbCrLf
    SQL_STORED = SQL_STORED & "SELECT ACCTCODE,ENTITYCODE,CUSCODE,REFERENCENAME,INVOICEDATE,INVOICENO,INVOICETYPE,CAST(AMOUNT AS NUMERIC (10,2)) AS AMOUNT FROM " & ySheet & "$ WHERE ACCTCODE IS NOT NULL " & vbCrLf
    SQL_STORED = SQL_STORED & "OPEN @AR                                   " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @AR INTO @ACCTCODE,@ENTITYCODE,@CUSCDE,@ACCTNAME,@INVOICEDATE,@INVOICENO,@INVOICETYPE,@AMOUNTTOPAY " & vbCrLf
    SQL_STORED = SQL_STORED & "WHILE @@FETCH_STATUS = 0                   " & vbCrLf
    SQL_STORED = SQL_STORED & "BEGIN                                      " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_DATE = '" & yBegdate & "'           " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_TYPE = 'COB'                        " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @REMARKS = 'BEG BALANCE ' + (CONVERT(VARCHAR(50),@J_DATE,107)) " & vbCrLf
    SQL_STORED = SQL_STORED & "--SET @ENTITYCODE ='C'                       " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @ACCTNAME = (CASE WHEN @ENTITYCODE = 'C' THEN " & vbCrLf
    SQL_STORED = SQL_STORED & "(SELECT  AC.ACCTNAME FROM ALL_CUSTOMER_TABLE AC WHERE AC.CUSCDE=@CUSCDE) " & vbCrLf
    SQL_STORED = SQL_STORED & "ELSE                                       " & vbCrLf
    SQL_STORED = SQL_STORED & "(SELECT  AV.NAMEOFVENDOR FROM ALL_VENDOR_TABLE AV WHERE AV.CODE=@CUSCDE) " & vbCrLf
    SQL_STORED = SQL_STORED & "END)                                       " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @ACCTDESC = (SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE ACCTCODE=@ACCTCODE)                         " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT ISNULL(MAX(JNO),0) + 1 JNO FROM AMIS_JOURNAL_HD) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT (REPLICATE('0',6-LEN(@J_JNO)) + @J_JNO))                 " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_VOUCHERNO = (SELECT ISNULL(MAX(VOUCHERNO),0) + 1 FROM AMIS_JOURNAL_HD WHERE JTYPE = 'COB') " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_VOUCHERNO = (SELECT (REPLICATE('0',6-LEN(@J_VOUCHERNO)) + @J_VOUCHERNO))                         " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_HD(JDATE,VOUCHERNO,JTYPE,JNO,CUSTOMERCODE,INVOICEDATE,INVOICETYPE,INVOICENO,INVOICEAMT,STATUS,ENTITY_CLASS,REMARKS) " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,@CUSCDE,@INVOICEDATE,@INVOICETYPE,@INVOICENO,@AMOUNTTOPAY,'N',@ENTITYCODE,@REMARKS)                 " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_DET(JDATE,VOUCHERNO,JTYPE,JNO,ENTITY,STATUS,ACCT_CODE,ACCT_NAME) " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,@ENTITYCODE+@CUSCDE,'N',@ACCTCODE,@ACCTDESC)                 " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @AR INTO @ACCTCODE,@ENTITYCODE,@CUSCDE,@ACCTNAME,@INVOICEDATE,@INVOICENO,@INVOICETYPE,@AMOUNTTOPAY " & vbCrLf
    SQL_STORED = SQL_STORED & "END                                         " & vbCrLf
    SQL_STORED = SQL_STORED & "CLOSE @AR                                   " & vbCrLf
    SQL_STORED = SQL_STORED & "DEALLOCATE @AR"
                        
    gconDMIS.Execute (SQL_STORED)
End Sub

Sub AP_STORED(yBegdate, ySheet As String)
    Dim SQL_STORED As String
    SQL_STORED = ""
    
    SQL_STORED = SQL_STORED & "DECLARE @J_JNO           AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_VOUCHERNO     AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_DATE          AS SMALLDATETIME " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @J_TYPE          AS VARCHAR(5)    " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @CODE            AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @NAMEOFVENDOR    AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICEDATE     AS SMALLDATETIME " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @DUEDATE         AS SMALLDATETIME " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @AMOUNTPAID      AS DECIMAL(18,2) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @BALANCE         AS DECIMAL(18,2) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTCODE        AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ACCTDESC        AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICETYPE     AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @INVOICENO       AS NVARCHAR(100) " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @ENTITYCODE      AS NVARCHAR(1)   " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @REMARKS         AS VARCHAR(100)  " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @TERMS           AS INT           " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @PAYTYPE         AS VARCHAR(100)  " & vbCrLf
    SQL_STORED = SQL_STORED & "DECLARE @AP AS CURSOR                     " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @AP = CURSOR FOR                      " & vbCrLf
    SQL_STORED = SQL_STORED & "SELECT ACCTCODE,ENTITYCODE,SUPCODE,REFERENCENAME,INVOICEDATE,INVOICENO,INVOICEDATE,INVOICETYPE,AMOUNT,BALANCE FROM " & ySheet & "$ WHERE SUPCODE IS NOT NULL " & vbCrLf
    SQL_STORED = SQL_STORED & "OPEN @AP                                  " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @AP INTO @ACCTCODE,@ENTITYCODE,@CODE,@NAMEOFVENDOR,@INVOICEDATE,@INVOICENO,@DUEDATE,@INVOICETYPE,@AMOUNTPAID,@BALANCE " & vbCrLf
    SQL_STORED = SQL_STORED & "WHILE @@FETCH_STATUS = 0                  " & vbCrLf
    SQL_STORED = SQL_STORED & "BEGIN " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_DATE = '" & yBegdate & "'          " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_TYPE = 'VPJ'                       " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @REMARKS = 'BEG BALANCE ' + (CONVERT(VARCHAR(50),@J_DATE,107)) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @NAMEOFVENDOR =  (SELECT  NAMEOFVENDOR FROM ALL_VENDOR_TABLE  WHERE CODE=@CODE) " & vbCrLf
    SQL_STORED = SQL_STORED & "--SET @INVOICETYPE = 'OI'                   " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @ACCTDESC = (SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE ACCTCODE=@ACCTCODE) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT ISNULL(MAX(JNO),0) + 1 JNO FROM AMIS_JOURNAL_HD) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_JNO = (SELECT (REPLICATE('0',6-LEN(@J_JNO)) + @J_JNO)) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_VOUCHERNO = (SELECT ISNULL(MAX(VOUCHERNO),0) + 1 FROM AMIS_JOURNAL_HD WHERE JTYPE = 'VPJ') " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @J_VOUCHERNO = (SELECT (REPLICATE('0',6-LEN(@J_VOUCHERNO)) + @J_VOUCHERNO)) " & vbCrLf
    SQL_STORED = SQL_STORED & "SET @PAYTYPE =(SELECT PAY_CODE FROM ALL_PAYTERM WHERE NO_DAYS = @TERMS) " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_HD(JDATE,VOUCHERNO,JTYPE,JNO,VENDORCODE,INVOICEDATE,INVOICETYPE,INVOICENO,DUEDATE,PAYTYPE,AMOUNTTOPAY,AMOUNTPAID,BALANCE,STATUS,ENTITY_CLASS,REMARKS) " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,@CODE,@INVOICEDATE,@INVOICETYPE,@INVOICENO,@DUEDATE,@TERMS,@AMOUNTPAID,'0.00',@BALANCE,'N',@ENTITYCODE,@REMARKS) " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_JOURNAL_DET(JDATE,VOUCHERNO,JTYPE,JNO,ENTITY,STATUS,ACCT_CODE,ACCT_NAME,INVOICENO,INVOICETYPE) " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES(@J_DATE,@J_VOUCHERNO,@J_TYPE,@J_JNO,@ENTITYCODE+@CODE,'N',@ACCTCODE,@ACCTDESC,@INVOICENO,@INVOICETYPE) " & vbCrLf
    SQL_STORED = SQL_STORED & "INSERT INTO AMIS_PV_DETAIL (JDATE,JTYPE,VOUCHERNO,INV_NO,AMOUNT,ENTITYCODE) " & vbCrLf
    SQL_STORED = SQL_STORED & "VALUES (@J_DATE,@J_TYPE,@J_VOUCHERNO,@INVOICENO,@AMOUNTPAID,@ENTITYCODE) " & vbCrLf
    SQL_STORED = SQL_STORED & "FETCH NEXT FROM @AP INTO @ACCTCODE,@ENTITYCODE,@CODE,@NAMEOFVENDOR,@INVOICEDATE,@INVOICENO,@DUEDATE,@INVOICETYPE,@AMOUNTPAID,@BALANCE " & vbCrLf
    SQL_STORED = SQL_STORED & "END                                        " & vbCrLf
    SQL_STORED = SQL_STORED & "CLOSE @AP                                  " & vbCrLf
    SQL_STORED = SQL_STORED & "DEALLOCATE @AP"
    
    gconDMIS.Execute (SQL_STORED)
End Sub

