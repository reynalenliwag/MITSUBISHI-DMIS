VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISTrans_ADB_IssuancesSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearchADB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   6120
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   375
      Left            =   3780
      TabIndex        =   5
      Top             =   6120
      Width           =   1395
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Search Advance Bill Number"
      Height          =   315
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   30
      Width           =   2925
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Search RO Number"
      Height          =   315
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      Value           =   -1  'True
      Width           =   2925
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   3525
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5145
      Left            =   0
      TabIndex        =   1
      Top             =   930
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   9075
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ADB#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RO#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ADB Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Keyword"
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   2715
   End
End
Attribute VB_Name = "frmPMISTrans_ADB_IssuancesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOCAL_STOCKTYPE                                    As String
Event SETCUSTOMERINFO(XCUSTOMERCODE As String, XCUSTOMERNAME As String, XRONUMBER As String, XREMARK As String)
Sub SETSTOCKTYPE(XXX As String)
    LOCAL_STOCKTYPE = XXX

End Sub
Private Sub Command1_Click()
    ListView1_DblClick
End Sub

Private Sub Form_Load()
    Text1_Change
End Sub

Private Sub ListView1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Dim TRAN_NO                                        As Long
    Dim RSORD                                          As ADODB.Recordset
    TRAN_NO = ListView1.SelectedItem.ListSubItems(5).Text
    If ListView1.SelectedItem.ListSubItems(4) = "CURRT" Then
        Set RSORD = gconDMIS.Execute("SELECT rono,CUSTCODE,CUSTNAME,REMARKS FROM PMIS_ORD_HD WHERE ID=" & TRAN_NO)
    Else
        Set RSORD = gconDMIS.Execute("SELECT rono,CUSTCODE,CUSTNAME,REMARKS FROM PMIS_ORD_HIST WHERE ID=" & TRAN_NO)
    End If


    If Not (RSORD.EOF Or RSORD.BOF) Then
        RaiseEvent SETCUSTOMERINFO(Null2String(RSORD!CUSTCODE), Null2String(RSORD!CUSTNAME), Null2String(RSORD!RONO), Null2String(RSORD!REMARKS))

    End If


End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ListView1_DblClick
    End If
End Sub

Private Sub Text1_Change()
    Dim RSLIST                                         As ADODB.Recordset
    Dim SQLX                                           As String
    Dim str_RONO                                       As String
    Dim SEARCHSTRING                                   As String

    If LTrim(RTrim(Text1)) <> "" Then

        str_RONO = Text1.Text
        If Option1.Value = True Then

            If Left(str_RONO, 2) = "R-" Then
                str_RONO = "R-" & Format(NumericVal(Right(str_RONO, Len(str_RONO) - 2)), "00000000")
            Else
                str_RONO = "R-" & Format(NumericVal(Right(str_RONO, Len(str_RONO))), "00000000")
            End If
            SEARCHSTRING = " AND  rono like " & N2Str2Null(str_RONO & "%")
        Else
            SEARCHSTRING = " AND  TRANNO like " & N2Str2Null(Format(str_RONO, "000000") & "%")
        End If
    End If

    SQLX = "SELECT TRANDATE,TRANNO ,RONO ,SALES_ORIGIN, 'CURRT' AS DSTATUS,ID FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND ISNULL(STATUS3,'')  <>'F' AND ISNULL(STATUS2,'')  <>'R' AND STATUS='P' " & SEARCHSTRING & vbCrLf
    SQLX = SQLX & " UNION " & vbCrLf
    SQLX = SQLX & "SELECT TRANDATE,TRANNO ,RONO ,SALES_ORIGIN,'HIST' AS DSTATUS ,ID FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & LOCAL_STOCKTYPE & "' AND ISNULL(STATUS3,'')  <>'F' AND  ISNULL(STATUS2,'')  <>'R' AND STATUS='P' " & SEARCHSTRING & " ORDER BY RONO"
    Set RSLIST = gconDMIS.Execute(SQLX)
    Listview_Loadval ListView1.ListItems, RSLIST
End Sub
