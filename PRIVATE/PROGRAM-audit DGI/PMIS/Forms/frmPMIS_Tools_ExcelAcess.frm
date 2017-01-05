VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmPMIS_Tools_ExcelAcess 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price List Conversion Tool"
   ClientHeight    =   6285
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11235
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
   Icon            =   "frmPMIS_Tools_ExcelAcess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11235
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   7905
      Left            =   30
      ScaleHeight     =   7905
      ScaleWidth      =   17205
      TabIndex        =   3
      Top             =   0
      Width           =   17205
      Begin FlexCell.Grid Grid1 
         Height          =   5415
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9551
         BackColor2      =   12648384
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   2
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   450
         Width           =   6675
      End
      Begin VB.CommandButton Command1 
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
         Left            =   10440
         MouseIcon       =   "frmPMIS_Tools_ExcelAcess.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Tools_ExcelAcess.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   675
      End
      Begin VB.CommandButton cmdOk 
         Enabled         =   0   'False
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
         Left            =   9780
         MouseIcon       =   "frmPMIS_Tools_ExcelAcess.frx":0A1A
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Tools_ExcelAcess.frx":0B6C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Update Master File"
         Top             =   60
         Width           =   675
      End
      Begin VB.CommandButton cmdCheck 
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
         Left            =   9120
         MouseIcon       =   "frmPMIS_Tools_ExcelAcess.frx":0E07
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Tools_ExcelAcess.frx":0F59
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Process Checking of Previous Cut-Off Balance"
         Top             =   60
         Width           =   675
      End
      Begin VB.TextBox txtExcel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   60
         Width           =   6675
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   3540
         Top             =   690
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Database Location of DMIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   10
         Top             =   480
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Browse for DNP File/ Price List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   2220
      End
   End
   Begin VB.PictureBox PICPROGRESS 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   60
      ScaleHeight     =   765
      ScaleWidth      =   11085
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   11115
      Begin VB.CommandButton cmd_mCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   9960
         TabIndex        =   4
         Top             =   210
         Width           =   1065
      End
      Begin wizProgBar.Prg Prg1 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   556
         Picture         =   "frmPMIS_Tools_ExcelAcess.frx":13D8
         BorderStyle     =   2
         BarPicture      =   "frmPMIS_Tools_ExcelAcess.frx":13F4
         BarPictureMode  =   0
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
      Begin VB.Label labexcel 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   60
         TabIndex        =   12
         Top             =   540
         Width           =   4260
      End
      Begin VB.Label LABPROGRESS 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   5010
      End
   End
End
Attribute VB_Name = "frmPMIS_Tools_ExcelAcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDNPP                                             As ADODB.Recordset
Dim m_cancel                                           As Boolean
Dim wsDNPP                                             As Workspace
Dim dbDNPP                                             As DATABASE



Sub ACESS_STRUCTURE()
    Dim i                                              As Integer
    Dim tdDNPP                                         As TableDef
    Dim FLDDNPP(7)                                     As Field
    Dim DNPPIDIndex                                    As Index
    Dim DNPPIDFLD                                      As Field

    Set tdDNPP = dbDNPP.CreateTableDef("ANNEX 1")
    Set FLDDNPP(0) = tdDNPP.CreateField("Field2", dbText, 60)
    Set FLDDNPP(1) = tdDNPP.CreateField("Field3", dbText, 255)
    Set FLDDNPP(2) = tdDNPP.CreateField("Field4", dbDouble)
    Set FLDDNPP(3) = tdDNPP.CreateField("Field5", dbDouble)
    Set FLDDNPP(4) = tdDNPP.CreateField("Field6", dbDouble)
    Set FLDDNPP(5) = tdDNPP.CreateField("Field7", dbDouble)
    Set FLDDNPP(6) = tdDNPP.CreateField("Field8", dbText, 255)

    For i = 0 To 6
        tdDNPP.Fields.Append FLDDNPP(i)
    Next i
    tdDNPP.Fields(1).AllowZeroLength = True
    tdDNPP.Fields(6).AllowZeroLength = True
    Set DNPPIDIndex = tdDNPP.CreateIndex("Field2")
    DNPPIDIndex.Primary = True
    DNPPIDIndex.Unique = True

    Set DNPPIDFLD = DNPPIDIndex.CreateField("Field2")
    DNPPIDIndex.Fields.Append DNPPIDFLD
    tdDNPP.Indexes.Append DNPPIDIndex
    dbDNPP.TableDefs.Append tdDNPP

    Set DNPPIDIndex = Nothing
    Set DNPPIDFLD = Nothing
    Set tdDNPP = Nothing
    For i = 0 To 7
        Set FLDDNPP(i) = Nothing
    Next i



    Set tdDNPP = dbDNPP.CreateTableDef("ANNEX 2")
    Set FLDDNPP(0) = tdDNPP.CreateField("Field2", dbText, 60)
    Set FLDDNPP(1) = tdDNPP.CreateField("Field3", dbText, 255)
    Set FLDDNPP(2) = tdDNPP.CreateField("Field4", dbDouble)
    Set FLDDNPP(3) = tdDNPP.CreateField("Field5", dbDouble)
    Set FLDDNPP(4) = tdDNPP.CreateField("Field6", dbDouble)
    Set FLDDNPP(5) = tdDNPP.CreateField("Field7", dbDouble)
    Set FLDDNPP(6) = tdDNPP.CreateField("Field8", dbText, 255)

    For i = 0 To 6
        tdDNPP.Fields.Append FLDDNPP(i)
    Next i
    tdDNPP.Fields(1).AllowZeroLength = True
    tdDNPP.Fields(6).AllowZeroLength = True
    Set DNPPIDIndex = tdDNPP.CreateIndex("Field2")
    DNPPIDIndex.Primary = True
    DNPPIDIndex.Unique = True


    Set DNPPIDFLD = DNPPIDIndex.CreateField("Field2")
    DNPPIDIndex.Fields.Append DNPPIDFLD
    tdDNPP.Indexes.Append DNPPIDIndex
    dbDNPP.TableDefs.Append tdDNPP

    Set DNPPIDIndex = Nothing
    Set DNPPIDFLD = Nothing
    Set tdDNPP = Nothing
    For i = 0 To 7
        Set FLDDNPP(i) = Nothing
    Next i



    Set tdDNPP = dbDNPP.CreateTableDef("ANNEX 3")
    Set FLDDNPP(0) = tdDNPP.CreateField("Field2", dbText, 60)
    Set FLDDNPP(1) = tdDNPP.CreateField("Field3", dbText, 255)
    Set FLDDNPP(2) = tdDNPP.CreateField("Field4", dbDouble)
    Set FLDDNPP(3) = tdDNPP.CreateField("Field5", dbDouble)
    Set FLDDNPP(4) = tdDNPP.CreateField("Field6", dbDouble)
    Set FLDDNPP(5) = tdDNPP.CreateField("Field7", dbDouble)
    Set FLDDNPP(6) = tdDNPP.CreateField("Field8", dbText, 255)

    For i = 0 To 6
        tdDNPP.Fields.Append FLDDNPP(i)
    Next i
    tdDNPP.Fields(1).AllowZeroLength = True
    tdDNPP.Fields(6).AllowZeroLength = True

    Set DNPPIDIndex = tdDNPP.CreateIndex("Field2")
    DNPPIDIndex.Primary = True
    DNPPIDIndex.Unique = True


    Set DNPPIDFLD = DNPPIDIndex.CreateField("Field2")
    DNPPIDIndex.Fields.Append DNPPIDFLD
    tdDNPP.Indexes.Append DNPPIDIndex
    dbDNPP.TableDefs.Append tdDNPP

    Set DNPPIDIndex = Nothing
    Set DNPPIDFLD = Nothing
    Set tdDNPP = Nothing
    For i = 0 To 7
        Set FLDDNPP(i) = Nothing
    Next i
End Sub


Private Sub cmd_mCancel_Click()
    Unload Me
End Sub


Private Sub cmdCheck_Click()
    dlg.Filter = ""
    dlg.FileName = ""






    dlg.Filter = "Hari DNP List(*.xls)|*.xls"
    dlg.ShowOpen


    If dlg.FileName = "" Then
        MessagePop InfoFriend, "File Selection", "No Files Being Selected", 1000, 2
        Exit Sub
    End If

    txtExcel = dlg.FileName

    dlg.FileName = ""
    dlg.Filter = "Access Files DMIS Format (*.MDB)|*.MDB"

    dlg.ShowSave

    Text1 = dlg.FileName

    If Text1 = "" Then Exit Sub

    Set wsDNPP = DBEngine.Workspaces(0)
    If Exists(Text1) Then
        If MsgQuestionBox("File already Exist! Do You want to Overwrite?", "File Exist") = False Then
            Exit Sub
        Else
            Kill Text1
        End If
    End If
    Set dbDNPP = wsDNPP.CreateDatabase(Text1, dbLangGeneral)
    Set dbDNPP = wsDNPP.OpenDatabase(Text1)

    frmSplash.Show
    frmSplash.labCon.Caption = "Creating HARI DNP Part Master File... Please wait..."
    DoEvents
    ACESS_STRUCTURE
    dbDNPP.Close
    Unload frmSplash
    MsgSpeechBox "DNP Structure Successfully Created! Please Click Generate to Write Database."
    cmdOk.Enabled = True


End Sub


Private Sub cmdOk_Click()

    On Error GoTo Errorcode:
    m_cancel = False

    Grid1.Rows = 1
    Grid1.AutoRedraw = False

    Dim j                                              As Long
    Dim i                                              As Long
    Dim LINEITEM
    Dim STOCKNO
    Dim STOCKNAME
    Dim DNP20                                          As Double
    Dim DNP32                                          As Double
    Dim DNP40                                          As Double
    Dim SRP                                            As Double
    Dim oConExcelToAcess                               As ADODB.Connection

    PICPROGRESS.Visible = True
    PICPROGRESS.ZOrder 0
    DoEvents
    Screen.MousePointer = 11
    labexcel = ""
    If LTrim(RTrim(txtExcel)) = "" Then: Exit Sub
    If Not rsDNPP Is Nothing Then
        Set rsDNPP = Nothing
    End If
    Set oConExcelToAcess = New ADODB.Connection
    oConExcelToAcess.CursorLocation = adUseClient
    Set rsDNPP = New ADODB.Recordset
    oConExcelToAcess.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text1 & ";Persist Security Info=False")



    Call rsDNPP.Open("select * from [annex 1]", oConExcelToAcess, adOpenDynamic, adLockOptimistic)
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(txtExcel)

    Dim int_wsheet                                     As Integer

    For int_wsheet = 1 To 3
        Set xlSheet = xlBook.Worksheets(int_wsheet)
        LINEITEM = xlSheet.Cells(7 + i, "A")
        If IsNumeric(LINEITEM) = False Or LINEITEM = "" Or IsEmpty(LINEITEM) = True Then
            err.Description = "DATA SHOULD FROM LINE NO 7" & " IN ANNEX [" & int_wsheet & "]"
            GoTo Errorcode
        End If
    Next


    For int_wsheet = 1 To 3
        labexcel = "IMPORTING DATA FROM ANNEX [" & int_wsheet & "]"
        i = 0
        Set xlSheet = xlBook.Worksheets(int_wsheet)
        LABPROGRESS = "Uploading of Price List. This might take a while.. Please Wait..."
        LABUPLOADSTATUS = ""

        LINEITEM = xlSheet.Cells(7 + i, "A")
        If i = 0 Then
            If IsNumeric(LINEITEM) = False Then
                err.Description = "Line Number Should Start at From Line No 7" & " IN ANNEX & " & int_wsheet
                GoTo Errorcode
            End If
        End If



        Prg1.Max = 100


        If IsNumeric(xlSheet.Cells(7 + i, "D")) = False Then

            err.Description = "Invalid File Format"
            GoTo Errorcode
        End If
        If IsNumeric(xlSheet.Cells(7 + i, "E")) = False Then

            err.Description = "Invalid File Format"
            GoTo Errorcode
        End If
        If IsNumeric(xlSheet.Cells(7 + i, "F")) = False Then

            err.Description = "Invalid File Format"
            GoTo Errorcode
        End If
        If IsNumeric(xlSheet.Cells(7 + i, "G")) = False Then


            err.Description = "Invalid File Format"
            GoTo Errorcode
        End If


        Do While LINEITEM <> ""
            If m_cancel = True Then Exit Do
            LINEITEM = xlSheet.Cells(7 + i, "A")
            STOCKNO = LTrim(RTrim(Replace(xlSheet.Cells(7 + i, "B"), "'", "")))
            STOCKNAME = Replace(xlSheet.Cells(7 + i, "C"), "'", "")
            DNP40 = xlSheet.Cells(7 + i, "D")
            DNP32 = xlSheet.Cells(7 + i, "E")
            DNP20 = xlSheet.Cells(7 + i, "F")
            SRP = xlSheet.Cells(7 + i, "G")
            Model = Replace(xlSheet.Cells(7 + i, "H"), "'", "")
            LABPROGRESS = STOCKNO & ":" & STOCKNAME

            If Prg1.Value >= 100 Then: Prg1.Value = 0: j = 0
            i = i + 1
            j = j + 1
            Prg1.Value = j
            DoEvents
            If Len(STOCKNO) > 0 Then
                Grid1.AddItem STOCKNO & Chr(9) & STOCKNAME & Chr(9) & DNP32 & Chr(9) & DNP20 & Chr(9) & DNP40 & Chr(9) & SRP & Chr(9) & Model, False
                If DNP32 > 0 Then
                    With rsDNPP
                        .AddNew
                        .Fields("FIELD2") = STOCKNO
                        .Fields("FIELD3") = STOCKNAME
                        .Fields("FIELD4") = DNP20
                        .Fields("FIELD5") = DNP32
                        .Fields("FIELD6") = DNP40
                        .Fields("FIELD7") = SRP
                        .Fields("FIELD8") = Model
                        .Update
                    End With
                End If
            End If
            Grid1.Refresh
            Grid1.TopRow = Grid1.Rows
        Loop
    Next

    Grid1.AutoRedraw = True
    Screen.MousePointer = 0
    Set rsDNPP = Nothing
    xlBook.Close
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Prg1.Value = 100
    Screen.MousePointer = 0
    'ShowPictureBox PICPROGRESS, False, picMain
    MessagePop RecSave, "DNP/SRP", "DNP/SRP DATA BASE CONVERTED", 400, 2
    cmdCheck.Enabled = False
    cmdOk.Enabled = False
    Exit Sub
Errorcode:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "ERROR"
    err.Clear

    If Not rsDNPP Is Nothing Then
        Set rsDNPP = Nothing
    End If
    If IsObject(xlBook) = True Then
        If Not xlBook Is Nothing Then
            xlBook.Close
            Set xlBook = Nothing
            Set xlSheet = Nothing
            Set xlApp = Nothing
        End If
    End If
    PICPROGRESS.Visible = False
    PICPROGRESS.ZOrder 1

    Screen.MousePointer = 0
End Sub




Private Sub Command1_Click()
    Unload Me
End Sub

Sub Form_Load()
    frmMain.Timer1.Enabled = False
    CenterMe frmMain, Me, 1
    m_cancel = False
    InitGrid
End Sub

Sub InitGrid()
    With Grid1
        .Cols = 8
        .Rows = 1
        .DisplayFocusRect = False
        .AllowUserResizing = True
        .FixedCols = 0
        .Cell(0, 0).Text = "L/N"
        .Column(0).Width = 75

        .Cell(0, 1).Text = "PART NO."
        .Column(1).Width = 100
        .Column(1).Locked = True
        .Column(1).Alignment = cellLeftGeneral

        .Cell(0, 2).Text = "DESCRIPTION"
        .Column(2).Width = 172
        .Column(2).Locked = True
        .Column(2).Alignment = cellLeftGeneral


        .Cell(0, 3).Text = "DNP32"
        .Column(3).Width = 60
        .Column(3).Alignment = cellRightGeneral
        .Column(3).CellType = cellTextBox
        .Column(3).Mask = cellValue
        .Column(3).FormatString = "0.00"
        .Column(3).DecimalLength = 2

        .Cell(0, 4).Text = "DNP20"
        .Column(4).Width = 60
        .Column(4).Alignment = cellRightGeneral
        .Column(4).CellType = cellTextBox
        .Column(4).Mask = cellValue
        .Column(4).FormatString = "0.00"
        .Column(4).Locked = True
        .Column(4).DecimalLength = 2

        .Cell(0, 5).Text = "DNP40"
        .Column(5).Width = 60
        .Column(5).Alignment = cellRightGeneral
        .Column(5).Locked = True
        .Column(5).CellType = cellTextBox
        .Column(5).Mask = cellValue
        .Column(5).FormatString = "0.00"
        .Column(5).DecimalLength = 2

        .Cell(0, 6).Text = "SRP"
        .Column(6).Width = 60
        .Column(6).Alignment = cellRightGeneral
        .Column(6).Locked = True
        .Column(6).CellType = cellTextBox
        .Column(6).Mask = cellValue
        .Column(6).FormatString = "0.00"
        .Column(6).DecimalLength = 2

        .Cell(0, 7).Text = "MODEL APPLICATION"
        .Column(7).Width = 150
        .Column(7).Locked = True
        .Column(7).Alignment = cellLeftGeneral
    End With

End Sub



