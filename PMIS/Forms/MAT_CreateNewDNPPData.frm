VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPMISMAT_CreateDNPPDATA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Master File"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   ForeColor       =   &H8000000F&
   Icon            =   "MAT_CreateNewDNPPData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MAT_CreateNewDNPPData.frx":1472
   ScaleHeight     =   720
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateData 
      Caption         =   "Create Distributor Database Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   750
      MouseIcon       =   "MAT_CreateNewDNPPData.frx":1AEC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Create Distributor Database Now"
      Top             =   75
      Width           =   3315
   End
   Begin MSComDlg.CommonDialog cmdDialogINV 
      Left            =   1050
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPMISMAT_CreateDNPPDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wsDNPP                                                            As Workspace
Dim dbDNPP                                                            As DATABASE
Dim FILNAME                                                           As String
Dim gconNewDNPP                                                       As ADODB.Connection

Function OpenOldDb() As Boolean
    On Error GoTo ConnErr
    Dim DNPSRP_Connection                                             As String
    With wizVar
        If .VerifyCryptoFile(App.Path & "\PMIS.crp") = True Then
            DNPSRP_Connection = .OpenCryptoFile("NEWDNPP", "CONNECT")
        End If
    End With
    On Error GoTo ConnErr
    Set gconNewDNPP = New ADODB.Connection
    gconNewDNPP.ConnectionString = DNPSRP_Connection
    gconNewDNPP.Open
    OpenOldDb = True
    Exit Function

ConnErr:
    ShowADOErrors gconNewDNPP
    Exit Function
End Function

Sub Create_Database()
    On Error GoTo ErrCode
    Set wsDNPP = DBEngine.Workspaces(0)
    If Exists(FILNAME) Then
        If MsgQuestionBox(FILNAME & " Already Exist! Overwrite?", "File Exist") = False Then
            Exit Sub
        Else
            Kill FILNAME
        End If
    End If
    Set dbDNPP = wsDNPP.CreateDatabase(FILNAME, dbLangGeneral)
    Set dbDNPP = wsDNPP.OpenDatabase(FILNAME)
    frmSplash.Show
    frmSplash.labCon.Caption = "Creating MMPC Part Master File... Please wait..."
    DoEvents
    Create_MMPC_Part_MasterFile_Table
    dbDNPP.Close
    Unload frmSplash
    MsgSpeechBox "MMPC DATABASE AND TABLES Successfully Created!!"
    Unload Me
    Exit Sub
ErrCode:
    If Err.Number = 3204 Then
        Resume Next
    Else
        ShowVBError
    End If
End Sub

Sub Create_MMPC_Part_MasterFile_Table()
    Dim I                                                             As Integer
    Dim tdDNPP                                                        As TableDef
    Dim FLDDNPP(7)                                                    As Field
    Dim DNPPIDIndex                                                   As Index
    Dim DNPPIDFLD                                                     As Field

    Set tdDNPP = dbDNPP.CreateTableDef("NEWDNPP")
    Set FLDDNPP(0) = tdDNPP.CreateField("ID", dbLong)
    FLDDNPP(0).Required = True
    Set FLDDNPP(1) = tdDNPP.CreateField("STOCKNUMBER", dbText, 12)
    FLDDNPP(1).Required = True
    FLDDNPP(1).AllowZeroLength = False
    Set FLDDNPP(2) = tdDNPP.CreateField("DESCRIPTIO", dbText, 16)
    Set FLDDNPP(3) = tdDNPP.CreateField("DNPP", dbDouble)
    Set FLDDNPP(4) = tdDNPP.CreateField("SRP", dbDouble)
    Set FLDDNPP(5) = tdDNPP.CreateField("MODEL", dbDouble)
    Set FLDDNPP(6) = tdDNPP.CreateField("ICC", dbDouble)
    For I = 0 To 6
        tdDNPP.Fields.Append FLDDNPP(I)
    Next I
    Set DNPPIDIndex = tdDNPP.CreateIndex("ID")
    DNPPIDIndex.Primary = True
    DNPPIDIndex.Unique = True
    Set DNPPIDFLD = DNPPIDIndex.CreateField("ID")
    DNPPIDIndex.Fields.Append DNPPIDFLD
    tdDNPP.Indexes.Append DNPPIDIndex
    dbDNPP.TableDefs.Append tdDNPP
    Set DNPPIDIndex = Nothing
    Set DNPPIDFLD = Nothing
    Set tdDNPP = Nothing
    For I = 0 To 7: Set FLDDNPP(I) = Nothing: Next I
End Sub

Private Sub cmdCreateData_Click()


    On Error Resume Next
    Dim MYPATH, PAYLNAME                                              As String
    MYPATH = App.Path
    cmdDialogINV.Filter = "Access Files (*.MDB)|*.MDB"
    cmdDialogINV.FilterIndex = 1
    cmdDialogINV.DefaultExt = "MDB"
    PAYLNAME = cmdDialogINV.FileName
    If MYPATH <> "\" Then
        cmdDialogINV.FileName = MYPATH & "\" & cmdDialogINV.FileName
    End If
    If PAYLNAME = "" Then
        cmdDialogINV.FileName = "*.MDB"
    End If
    cmdDialogINV.Action = 2
    If Err = 32755 Then Exit Sub
    FILNAME = cmdDialogINV.FileName
    Create_Database
    LogAudit "G", "CREATE DNPP DATABASE"
    If Err = 32755 Then Exit Sub
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe Me, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub

