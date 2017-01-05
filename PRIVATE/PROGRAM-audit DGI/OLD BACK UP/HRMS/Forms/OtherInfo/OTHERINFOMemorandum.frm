VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOMemorandum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MEMORANDUM"
   ClientHeight    =   4500
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   6615
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
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
      Height          =   705
      Left            =   5820
      MouseIcon       =   "OTHERINFOMemorandum.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOMemorandum.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit Window"
      Top             =   3690
      Width           =   705
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5130
      MouseIcon       =   "OTHERINFOMemorandum.frx":04B8
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOMemorandum.frx":060A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Delete Selected Record"
      Top             =   3690
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
      Height          =   705
      Left            =   4440
      MouseIcon       =   "OTHERINFOMemorandum.frx":0935
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOMemorandum.frx":0A87
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Edit Selected Record"
      Top             =   3690
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3750
      MouseIcon       =   "OTHERINFOMemorandum.frx":0DE3
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOMemorandum.frx":0F35
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Add Record"
      Top             =   3690
      Width           =   705
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   150
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMemorandum 
      Height          =   3090
      Left            =   660
      ScaleHeight     =   3030
      ScaleWidth      =   5565
      TabIndex        =   8
      Top             =   180
      Width           =   5625
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   1860
         Width           =   3375
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "Open File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4530
         TabIndex        =   6
         ToolTipText     =   "Open File"
         Top             =   1860
         Width           =   945
      End
      Begin VB.TextBox txtNoOfOffense 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txtMemoDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   0
         Top             =   60
         Width           =   1455
      End
      Begin VB.TextBox txtDisciplinaryAction 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1110
         Width           =   5415
      End
      Begin VB.TextBox txtSubject 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   420
         Width           =   3930
      End
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
         Height          =   675
         Left            =   2760
         MouseIcon       =   "OTHERINFOMemorandum.frx":1248
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOMemorandum.frx":139A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel Entry"
         Top             =   2250
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
         Height          =   675
         Left            =   2070
         MouseIcon       =   "OTHERINFOMemorandum.frx":16D8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOMemorandum.frx":182A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   2250
         Width           =   705
      End
      Begin VB.CommandButton cmdAttachFile 
         Caption         =   "Attach File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3510
         TabIndex        =   5
         ToolTipText     =   "Attach File"
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1620
         TabIndex        =   14
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Offense"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   1530
         Width           =   1515
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Memo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Disciplinary Action"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   810
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   450
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdMemorandum 
      Height          =   3210
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   5662
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERINFOMemorandum.frx":1B7A
   End
   Begin MSComctlLib.ListView lstMemorandum 
      Height          =   3540
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   6244
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "OTHERINFOMemorandum.frx":1B96
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date Of Memo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Disciplinary Action"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No. of Offense"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
   End
End
Attribute VB_Name = "frmOTHERINFOMemorandum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsMemorandum                                                      As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim PAYLNAME                                                          As String
Dim EMPLIVIL                                                          As String

Public Function SaveFileToDB(ByVal Filename As String, _
                             RS As Object, FieldName As String) As Boolean
    '**************************************************************
    'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
    'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME

    'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACESS)

    'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

    'Dim sConn As String
    'Dim oConn As New ADODB.Connection
    'Dim oRs As New ADODB.Recordset
    '
    '
    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    '
    'oConn.Open sConn
    'oRs.Open "SELECT * FROM MYTABLE", oConn, adOpenKeyset, _
     adLockOptimistic
    'oRs.AddNew

    'SaveFileToDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
    'oRs.Update
    'oRs.Close
    '**************************************************************

    Dim iFileNum                                                      As Integer
    Dim lFileLength                                                   As Long

    Dim abBytes()                                                     As Byte
    Dim iCtr                                                          As Integer

    'On Error GoTo ErrorHandler
    If Dir(Filename) = "" Then Exit Function
    If Not TypeOf RS Is ADODB.Recordset Then Exit Function

    'read file contents to byte array
    iFileNum = FreeFile
    Open Filename For Binary Access Read As #iFileNum
    lFileLength = LOF(iFileNum)

    ReDim abBytes(lFileLength)
    Get #iFileNum, , abBytes()
    'put byte array contents into db field
    RS.FIELDS(FieldName).AppendChunk abBytes()
    'Stop
    Close #iFileNum

    SaveFileToDB = True
ErrorHandler:
    MsgBox Err.Description
End Function

Public Function LoadFileFromDB(Filename As String, _
                               RS As Object, FieldName As String) As Boolean
    '************************************************
    'PURPOSE: LOADS BINARY DATA IN RECORDSET RS,
    'FIELD FieldName TO a File Named by the FileName parameter

    'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

    'SAMPLE USAGE
    'Dim sConn As String
    'Dim oConn As New ADODB.Connection
    'Dim oRs As New ADODB.Recordset
    '
    '
    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    '
    'oConn.Open sConn
    'oRs.Open "SELECT * FROM MyTable", oConn, adOpenKeyset,
    ' adLockOptimistic
    'LoadFileFromDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
    'oRs.Close
    '************************************************
    Dim iFileNum                                                      As Integer
    Dim lFileLength                                                   As Long
    Dim abBytes()                                                     As Byte
    Dim iCtr                                                          As Integer

    On Error GoTo ErrorHandler
    If Not TypeOf RS Is ADODB.Recordset Then Exit Function

    iFileNum = FreeFile
    Open Filename For Binary As #iFileNum
    lFileLength = LenB(RS(FieldName))

    abBytes = RS(FieldName).GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum
    LoadFileFromDB = True

ErrorHandler:
    MsgBox Err.Description
End Function

Sub InitMemVars()
    txtMemoDate.Text = ""
    txtSubject.Text = ""
    txtDisciplinaryAction.Text = ""
    txtNoOfOffense.Text = ""
    txtFileName.Text = ""
End Sub

Sub SSubjectreEntry(XXX As Variant)
    Set rsMemorandum = New ADODB.Recordset
    Set rsMemorandum = gconDMIS.Execute("Select * from HRMS_Memorandum Where ID = " & XXX)
    If Not rsMemorandum.EOF And Not rsMemorandum.BOF Then
        labID.Caption = rsMemorandum!ID
        txtMemoDate.Text = Null2String(rsMemorandum!MemoDate)
        txtSubject.Text = Null2String(rsMemorandum!Subject)
        txtDisciplinaryAction.Text = Null2String(rsMemorandum!DisciplinaryAction)
        txtNoOfOffense.Text = Null2String(rsMemorandum!NoOfOffense)
    End If
End Sub

Sub FillGrid()
    lstMemorandum.Sorted = False: lstMemorandum.ListItems.Clear
    lstMemorandum.Enabled = False
    Set rsMemorandum = New ADODB.Recordset
    Set rsMemorandum = gconDMIS.Execute("select [MemoDate],[Subject],[DisciplinaryAction],NoOfOffense,ID from HRMS_Memorandum where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsMemorandum.EOF And rsMemorandum.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstMemorandum.ListItems, rsMemorandum
        lstMemorandum.Refresh
        lstMemorandum.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

'Upating Code       : AXP-0707200711:58
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_ADD", "DATA ENTRY") = False Then Exit Sub
    cmdMemorandum.ZOrder 0: picMemorandum.ZOrder 0
    AddorEdit = "ADD"
    InitMemVars
    On Error Resume Next
    txtMemoDate.SetFocus

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdAttachFile_Click()
    On Error Resume Next
    comDialog.FILTER = "All Document Files |*.doc |*.xls | *.ppt"
    comDialog.FilterIndex = 2
    comDialog.DefaultExt = "DOC"
    PAYLNAME = comDialog.Filename
    comDialog.Filename = comDialog.Filename
    If PAYLNAME = "" Then
        comDialog.Filename = "*.doc;*.xls;*.ppt"
    End If
    comDialog.Action = 2
    PAYLNAME = comDialog.Filename
    txtFileName.Text = PAYLNAME
    If Err = 32755 Then Exit Sub
    If Err = 32755 Then Exit Sub
End Sub

Private Sub cmdCancel_Click()
    cmdMemorandum.ZOrder 1: picMemorandum.ZOrder 1
End Sub

'Upating Code       : AXP-0707200711:57
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_DELETE", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstMemorandum.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_Memorandum Where ID = " & lstMemorandum.SelectedItem.SubItems(4))
                ShowDeletedMsg
                FillGrid
            End If
        End If
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200711:58
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "ACESS_EDIT", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstMemorandum.SelectedItem.SubItems(4) <> "" Then
            SSubjectreEntry lstMemorandum.SelectedItem.SubItems(4)
            cmdMemorandum.ZOrder 0: picMemorandum.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpenFile_Click()
    Dim sConn                                                         As String
    Dim oConn                                                         As New ADODB.Connection
    Dim oRs                                                           As New ADODB.Recordset

    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    sConn = DMIS_Connection
    oConn.Open sConn
    oRs.Open "SELECT * from HRMS_Memorandum Where ID = " & labID.Caption, oConn, adOpenKeyset, adLockOptimistic
    If Not oRs.EOF And Not oRs.BOF Then
        LoadFileFromDB App.Path & "\TempFile.DOC", oRs, "MEMO_DOC"
        oRs.Close
    End If

    'Referenziare "Microsoft Word 8.0 Object Library" o superiore
    'Reference "Microsoft Word 8.0 Object Library" or more
    Dim xWord                                                         As Word.Application    ' L'applicazione Word         (The Word Application
    Dim xRange                                                        As Range    ' Oggetto Range               (Object Range)
    Dim xSelection                                                    As Find    ' Oggetto Find                (Object Find)
    Dim xTabella                                                      As TABLE    ' Oggetto Tabella             (Object Table)
    Dim xCella                                                        As Cell    ' Oggetto Cella               (Object Cell)
    Set xWord = New Application
    xWord.Visible = False
    'Aggiungo un documento o un modello fatto precedentemente che si chiama prova.dot
    'Add a document or a model do precedent call "prova.dot"
    'xWord.Documents.Add App.Path & "\prova.dot"
    xWord.Documents.Add App.Path & "\TempFile.DOC"
    'Protetto da una password supponiamo sia "pass"
    'Protect by a password like "pass"
    'xWord.ActiveDocument.Unprotect "pass"
    'Aggiungo i valori dal record (io ho usato dati fissi per semplicità) al posto di
    '%%nome%%, %%cognome%%, %%data%% che devono comparire nel documento di word
    'come "%%<nome campo>%%"
    '
    'Add Record Value
    Set xRange = xWord.ActiveDocument.Range
    xRange.Find.Execute "%%nome%%", , , , , , , , , "Paperino", True:    '"Paperino" Can Be substitute by a TextBox
    Set xRange = xWord.ActiveDocument.Range
    xRange.Find.Execute "%%cognome%%", , , , , , , , , "Pippo", True    '"Pippo" Can Be substitute by a TextBox
    Set xRange = xWord.ActiveDocument.Range
    xRange.Find.Execute "%%data%%", , , , , , , , , "01/01/2000", True    '"01/01/2000" Can Be substitute by a TextBox
    'Ripristina la password                         (Put again the password)
    'xWord.ActiveDocument.Protect wdAllowOnlyFormFields, , "pass"
    'Per visualizze il documento                    (Show The Document)
    xWord.Visible = True
    xWord.WindowState = wdWindowStateMaximize
    xWord.Application.Activate
    'Per visualizzare l'anteprima di stampa         (Print Preview)
    'xWord.ActiveDocument.PrintPreview
    'Per inviare via email il documento             (Send By email)
    'xWord.ActiveDocument.SendMail
    'Per inviare via fax                            (Send A fax)
    'xWord.ActiveDocument.SendFax
    'Per stampare il documento                      (Print)
    'xWord.ActiveDocument.PrintOut
    'Per salvarlo in una directory                  (Save In A directory)
    'xWord.ActiveDocument.SaveAs App.Path & "\Docs\" & "MyDoc.doc"
End Sub

'Upating Code       : AXP-0707200711:57
Private Sub cmdSave_Click()
    Dim RecID                                                         As Long
    On Error GoTo Errorcode:

    cmdMemorandum.ZOrder 1: picMemorandum.ZOrder 1
    'SAMPLE USAGE
    'Dim sConn As String
    'Dim oConn As New ADODB.Connection
    'Dim oRs As New ADODB.Recordset
    '
    '
    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    'sConn = DMIS_Connection
    '
    'oConn.Open sConn
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Memorandum " & _
                         "(EMPLEVEL,EMPNO,[DisciplinaryAction],[MemoDate],[Subject],NoOfOffense,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtDisciplinaryAction.Text) & "," & N2Str2Null(txtMemoDate.Text) & "," & N2Str2Null(txtSubject.Text) & "," & N2Str2Null(txtNoOfOffense.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"
        'oRs.Open "SELECT * from HRMS_Memorandum order by id desc", oConn, adOpenKeyset, adLockOptimistic
        'If Not oRs.EOF And Not oRs.BOF Then
        '   RecID = oRs!ID
        'End If
        'Set oRs = New ADODB.Recordset
        'oRs.Open "SELECT * from HRMS_Memorandum WHERE ID = " & RecID, oConn, adOpenKeyset, adLockOptimistic
        'If Not oRs.EOF And Not oRs.BOF Then
        '   SaveFileToDB PAYLNAME, oRs, "MEMO_DOC"
        '   oRs.Update
        '   oRs.Close
        'End If
    Else
        gconDMIS.Execute "update HRMS_Memorandum set " & _
                       " [DisciplinaryAction] = " & N2Str2Null(txtDisciplinaryAction.Text) & "," & _
                       " [MemoDate] = " & N2Str2Null(txtMemoDate.Text) & "," & _
                       " [Subject] = " & N2Str2Null(txtSubject.Text) & "," & _
                       " NoOfOffense = " & N2Str2Null(txtNoOfOffense.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE OTHER INFORMATION", EMPLOYEE_NO)
        'Set oRs = New ADODB.Recordset
        'oRs.Open "SELECT * from HRMS_Memorandum WHERE ID = " & labID.Caption, oConn, adOpenForwardOnly, adLockOptimistic
        'If Not oRs.EOF And Not oRs.BOF Then
        '   SaveFileToDB PAYLNAME, oRs, "MEMO_DOC"
        '   oRs.Update
        '   oRs.Close
        'End If
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdMemorandum.ZOrder 1: picMemorandum.ZOrder 1
        Case vbKeyF3
            cmdMemorandum.ZOrder 0: picMemorandum.ZOrder 0
            AddorEdit = "ADD"
            InitMemVars
            On Error Resume Next
            txtMemoDate.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstMemorandum.SelectedItem.SubItems(5) <> "" Then
                    SSubjectreEntry lstMemorandum.SelectedItem.SubItems(5)
                    cmdMemorandum.ZOrder 0: picMemorandum.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstMemorandum.SelectedItem.SubItems(5) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_Memorandum Where ID = " & lstMemorandum.SelectedItem.SubItems(5))
                        ShowDeletedMsg
                        FillGrid
                    End If
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    cmdMemorandum.ZOrder 1: picMemorandum.ZOrder 1
    FillGrid
End Sub

Private Sub lstMemorandum_DblClick()
    If EmptyRecord = False Then
        If lstMemorandum.SelectedItem.SubItems(4) <> "" Then
            SSubjectreEntry lstMemorandum.SelectedItem.SubItems(4)
            cmdMemorandum.ZOrder 0: picMemorandum.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstMemorandum_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstMemorandum
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

