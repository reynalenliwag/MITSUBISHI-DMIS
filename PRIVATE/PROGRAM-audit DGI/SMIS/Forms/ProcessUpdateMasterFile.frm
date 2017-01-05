VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_ProcessUpdateMasterFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Master File"
   ClientHeight    =   5685
   ClientLeft      =   270
   ClientTop       =   360
   ClientWidth     =   6390
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
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
   Icon            =   "ProcessUpdateMasterFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6390
   Begin VB.ComboBox cboYear 
      Height          =   345
      Left            =   2250
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboMonth 
      Height          =   345
      Left            =   2250
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
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
      Height          =   795
      Left            =   4950
      MouseIcon       =   "ProcessUpdateMasterFile.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUpdateMasterFile.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit Window"
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Update"
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
      Height          =   795
      Left            =   4230
      MouseIcon       =   "ProcessUpdateMasterFile.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "ProcessUpdateMasterFile.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Update Master File"
      Top             =   4800
      Width           =   735
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1095
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   30
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   750
         Width           =   3615
         Begin VB.Label labProcessing 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   2
            Top             =   -30
            Width           =   3525
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Update progress"
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "ProcessUpdateMasterFile.frx":0BAF
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ProcessUpdateMasterFile.frx":0BCB
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   3
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   4
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            TX              =   "cmd1"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "ProcessUpdateMasterFile.frx":0BE7
         End
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   5595
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   13
      Top             =   1140
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   1138
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DATE REL"
         Object.Width           =   2408
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "INV DATE"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "VI_NO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CODE "
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CS#"
         Object.Width           =   1455
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "STATUS"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "For The Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   3150
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmSMIS_ProcessUpdateMasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboMonth_Change()
    FillDetails
End Sub

Private Sub cbomonth_Click()
    FillDetails
End Sub

Private Sub cboYear_Change()
    FillDetails
End Sub
Private Sub cboYear_Click()
    FillDetails
End Sub

Sub FillDetails()
    Dim rsSO                            As ADODB.Recordset
    Dim fld                             As ADODB.Field
    Dim lst                             As ListItem
    Set rsSO = New ADODB.Recordset
'    If cboMonth <> "ALL" Then
'        Call rsSO.Open("Select convert(varchar,DATERELEASED,101) DATERELEASED,INVOICEDDATE,VI_NO, CODE,IGNKEY_NO,STATUS,ID from smis_salesorder where isnull(Status,'')<>'C' and isnull(soStatus,'')<>'C'   and month(deyt)=" & What_month(cboMonth) & " and Year(deyt)=" & cboYear, gconDMIS, adOpenKeyset, adLockReadOnly)
'    Else
'        Call rsSO.Open("Select convert(varchar,DATERELEASED,101) DATERELEASED,INVOICEDDATE,VI_NO, CODE,IGNKEY_NO,STATUS,ID   from smis_salesorder where isnull(Status,'')<>'C' and isnull(soStatus,'')<>'C'  and Year(deyt)=" & cboYear, gconDMIS, adOpenKeyset, adLockReadOnly)
'    End If
    Call rsSO.Open(" SELECT CONVERT(VARCHAR,DATERELEASED,101) DATERELEASED,INVOICEDDATE,VI_NO, CODE,IGNKEY,STATUS,ISTATUS,CUSTOMERCODE,ID   FROM SMIS_MRRINV_TABLE", gconDMIS, adOpenKeyset, adLockReadOnly)


    ListView1.ListItems.Clear
    If Not rsSO.EOF Or Not rsSO.BOF Then
        cmdProcess.Enabled = True
    Else
        cmdProcess.Enabled = False
    End If

    While Not rsSO.EOF
        j = j + 1
        Set lst = ListView1.ListItems.Add(, , j)
        lst.Checked = False
        For Each fld In rsSO.Fields
            If IsNull(fld.Value) Then
                lst.ListSubItems.Add , , vbNullString
            Else
                lst.ListSubItems.Add , , fld.Value
            End If
        Next

'        If IsDate(rsSO!DATERELEASED) = True Then
'            If CDate(rsSO!DATERELEASED) > CDate(LOGDATE) Then
'                SetColorX vbRed, lst
'            End If
'        Else
'            SetColorX &H4000&, lst
'        End If



        rsSO.MoveNext
    Wend

End Sub


Private Sub cmdProcess_Click()
    If MsgBox("Are You Sure You Want To Update Master File ", vbInformation + vbYesNo) = vbNo Then Exit Sub
    initMemvars
    
Dim rs As ADODB.Recordset
Set rs = gconDMIS.Execute("SELECT * FROM SMIS_SALESORDER WHERE STATUS='P'")
While Not rs.EOF
    gconDMIS.Execute ("update smis_MRRINV_TABLE SET ISTATUS='R', RELEASED=1 , DateReleased=" & N2Str2Null(rs!DateReleased) & "',InvoicedDate=" & N2Str2Null(rs!InvoicedDate) & " ,vi_no=" & N2Str2Null(rs!vi_no) & " WHERE IGNKEY='" & rs!IGNKEY_NO & "'")
    



 rs.MoveNext
Wend
End Sub

Private Sub cmdExit_Click()

    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    initMemvars
    cboYear = Year(LOGDATE)
    FillcboYear cboYear
    cmdProcess.Enabled = False
    fillcbomonth cboMonth
    cboMonth.AddItem "ALL", 0
    cboMonth.ListIndex = Month(LOGDATE)
    FillDetails
End Sub
Sub initMemvars()
    labCPB = "0%": labProcessing = ""
End Sub

Sub UpdateMasterFile()





Exit Sub





















    Dim rsSO                            As ADODB.Recordset
    Dim GMax                            As Integer
    Dim GVal                            As Integer
    Dim LGN                             As Long
    Set rsSO = New ADODB.Recordset

    'If cboMonth <> "ALL" Then
    'Call rsSO.Open("SELECT * FROM SMIS_SALESORDER WHERE ISNULL(SOSTATUS,'')<>'C' AND MONTH(DEYT)=" & What_month(cboMonth) & " AND YEAR(DEYT)=" & cboYear, gconDMIS, adOpenKeyset, adLockReadOnly)
    'Else
    '   Call rsSO.Open("SELECT * FROM SMIS_SALESORDER WHERE ISNULL(SOSTATUS,'')<>'C' AND YEAR(DEYT)=" & cboYear, gconDMIS, adOpenKeyset, adLockReadOnly)
    'End If

'    gconDMIS.Execute ("Update SMIS_MRRINV_TABLE" _
'                    & " SET [DateReleased] = NULL ," _
'                    & " [InvoicedDate] = null , " _
'                    & " [Released] =  0, " _
'                    & " [LastInvDate] = " & N2Str2Null(LOGDATE) & " , " _
'                    & " [WithProsBuyers] = 'n' , " _
'                    & " [ProspectID] = null, " _
'                    & " [CustomerCode] =NULL , " _
'                    & " [VI_NO] =NULL, " _
'                    & " [IStatus] ='O'  ")




    Call rsSO.Open("SELECT * FROM SMIS_SALESORDER WHERE ISDATE(DATERELEASED)=1 AND   STATUS='P'", gconDMIS, adOpenKeyset, adLockReadOnly)

    If rsSO.EOF Or rsSO.BOF Then
        MsgBox "No Record In Selection", vbInformation
        Exit Sub
    End If
    GMax = rsSO.RecordCount
    progCPB.Max = GMax
    progCPB.Value = 0
    rsSO.MoveFirst
    gconDMIS.Execute ("UPDATE SMIS_MRRINV_TABLE SET STATUS='P' WHERE STATUS<>'C' ")

    While Not rsSO.EOF
        DoEvents
        i = i + 1
        progCPB.Value = i
        labCPB = FormatPercent(i / GMax)
        labProcessing = rsSO!SO_NO

        If IsDate(rsSO!DateReleased) = True Then
            Me.Caption = NumericVal(Me.Caption) + 1
            gconDMIS.Execute ("Update SMIS_MRRINV_TABLE" _
                            & " SET [DateReleased] = " & N2Str2Null(rsSO!DateReleased) & " ," _
                            & " [InvoicedDate] = " & N2Str2Null(rsSO!InvoicedDate) & " , " _
                            & " [Released] =  1, " _
                            & " [LastInvDate] = " & N2Str2Null(LOGDATE) & " , " _
                            & " [WithProsBuyers] = 'Y' , " _
                            & " [ProspectID] = " & rsSO!PROSPECTID & ", " _
                            & " [CustomerCode] =" & N2Str2Null(rsSO!CODE) & " , " _
                            & " [VI_NO] =" & N2Str2Null(rsSO!vi_no) & " , " _
                            & " [STATUS] ='P'" & " , " _
                            & " [IStatus] ='R'  " _
                            & " WHERE IGNKEY='" & rsSO!IGNKEY_NO & "'")

            

'        ElseIf IsDate(rsSO!DATERELEASED) = False And Null2String(rsSO!SOSTATUS) = "P" Then
'            gconDMIS.Execute ("Update SMIS_MRRINV_TABLE" _
'                            & " SET [DateReleased] = NULL ," _
'                            & " [InvoicedDate] = " & N2Str2Null(rsSO!InvoicedDate) & " , " _
'                            & " [Released] =  0, " _
'                            & " [LastInvDate] = " & N2Str2Null(LOGDATE) & " , " _
'                            & " [WithProsBuyers] = 'Y' , " _
'                            & " [ProspectID] = " & rsSO!PROSPECTID & ", " _
'                            & " [CustomerCode] =" & N2Str2Null(rsSO!CODE) & " , " _
'                            & " [VI_NO] =" & N2Str2Null(rsSO!VI_NO) & " , " _
'                            & " [STATUS] ='P'" & " , " _
'                            & " [IStatus] ='S'  " _
'                            & " WHERE IGNKEY='" & rsSO!ignkey_no & "'")
            
        End If

        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).ListSubItems(7) = rsSO!ID Then
                ListView1.ListItems(i).Checked = True
                SetColorX vbBlue, ListView1.ListItems(i)
                Exit For
            End If
        Next

        rsSO.MoveNext
    Wend

End Sub
Sub SetColorX(colorx As OLE_COLOR, lstitem As ListItem)
    Dim i
    lstitem.ForeColor = colorx
    For i = 1 To lstitem.ListSubItems.Count - 1
        lstitem.ListSubItems(i).ForeColor = colorx
    Next

End Sub

Private Sub LISTVIEW1_DblClick()
    Load frmSMIS_Trans_VehicleInvoice
    'Debug.Print ListView1.SelectedItem.ListSubItems(7).Text

    frmSMIS_Trans_VehicleInvoice.SearchInvoice ListView1.SelectedItem.ListSubItems(7).Text
    frmSMIS_Trans_VehicleInvoice.Show
End Sub
