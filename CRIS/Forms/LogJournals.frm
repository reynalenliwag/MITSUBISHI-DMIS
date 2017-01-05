VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCRIS_LogCall 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogJournals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6720
      TabIndex        =   25
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Height          =   435
      Left            =   5730
      TabIndex        =   24
      Top             =   4800
      Width           =   945
   End
   Begin VB.TextBox txtDuration 
      Height          =   375
      Left            =   6210
      TabIndex        =   23
      Text            =   "0"
      Top             =   900
      Width           =   1305
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6630
      Top             =   90
   End
   Begin VB.CheckBox chkReminders 
      Caption         =   "Need Followups ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3300
      TabIndex        =   18
      Top             =   1980
      Width           =   2040
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   3330
      TabIndex        =   17
      Top             =   1470
      Width           =   4275
   End
   Begin VB.ComboBox cboCallType 
      Height          =   330
      Left            =   3270
      TabIndex        =   13
      Top             =   300
      Width           =   1845
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   3300
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2640
      Width           =   4425
   End
   Begin VB.PictureBox picProfileCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   30
      ScaleHeight     =   4755
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Label lblEmail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   15
         TabIndex        =   11
         Top             =   3795
         Width           =   3090
      End
      Begin VB.Label lblContactNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15
         TabIndex        =   10
         Top             =   3045
         Width           =   3090
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   15
         TabIndex        =   9
         Top             =   2055
         Width           =   3090
      End
      Begin VB.Label lblAccountName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15
         TabIndex        =   8
         Top             =   1350
         Width           =   3090
      End
      Begin VB.Label lblCustomerName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   15
         TabIndex        =   7
         Top             =   615
         Width           =   3090
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   15
         TabIndex        =   6
         Top             =   1815
         Width           =   3090
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   3120
         _Version        =   655364
         _ExtentX        =   5503
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Profile"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   64
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   315
         Width           =   3090
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   15
         TabIndex        =   3
         Top             =   1080
         Width           =   3090
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Contact"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Index           =   3
         Left            =   15
         TabIndex        =   2
         Top             =   2760
         Width           =   3090
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   " Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Index           =   4
         Left            =   15
         TabIndex        =   1
         Top             =   3510
         Width           =   3090
      End
   End
   Begin MSComCtl2.DTPicker dtDateCall 
      Height          =   345
      Left            =   3300
      TabIndex        =   14
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55181313
      CurrentDate     =   39139
   End
   Begin MSComCtl2.DTPicker dtTimeCall 
      Height          =   345
      Left            =   4740
      TabIndex        =   15
      Top             =   900
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55181314
      CurrentDate     =   39139
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3300
      TabIndex        =   22
      Top             =   1230
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Time Called"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3300
      TabIndex        =   21
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Call Bound"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3270
      TabIndex        =   20
      Top             =   30
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3300
      TabIndex        =   19
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   6210
      TabIndex        =   16
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frmCRIS_LogCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProfileId                                As Long
Dim ProfileType                              As String
Dim ProspectID                               As Long
Dim AcctName                                 As String
Dim LOGID                                    As Long




Private Sub cmdSave_Click()
    If Runvalidation("@R") = False Then: Exit Sub
    Dim temprs                               As ADODB.Recordset
    If LOGID > 0 Then
        SQL = " Update CRIS_PROSPECT_CALLS " _
            & " SET DateTimeCall=@DateTimeCall, Duration=@Duration, Subject=@Subject, Comments=@Comments " _
            & " WHERE LogID=@LogID"
    Else
        SQL = "INSERT INTO CRIS_PROSPECT_CALLS(ProfileID, CustomerType, DateTimeCall, Duration, Subject, Comments) " _
            & " VALUES(@ProfileID, @CustomerType, @DateTimeCall, @Duration, @Subject, @Comments)" & vbCrLf & " SELECT @@IDENTITY"
    End If
    SQL = Replace(SQL, "@ProfileID", ProfileId)
    SQL = Replace(SQL, "@CustomerType", N2Str2Null(CustomerType))
    SQL = Replace(SQL, "@DateTimeCall", "'" & FormatDateTime(dtDateCall.Value, vbShortDate) & " " & FormatDateTime(dtTimeCall, vbLongTime) & "'")
    SQL = Replace(SQL, "@Duration", txtDuration.Value)
    SQL = Replace(SQL, "@Subject", N2Str2Null(cboSubject.Text))
    SQL = Replace(SQL, "@Comments", N2Str2Null(txtComments.Text))
    SQL = Replace(SQL, "@LogID", N2Str2Null(LOGID))
    'Set temprs = gconDMIS.Execute(SQL)
    SQL = Replace(SQL, "@LogID", N2Str2Null(LOGID))
    Set temprs = gconDMIS.Execute(SQL)
    If LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "Profile Sucessfully Added"
    Else
        MessagePop RecSave, "RecordSaved", "Profile Sucessfully Updated"
    End If

    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        LOGID = temprs.Collect(0)
        cmdSave.caption = "UPDATE"
        cmdCancel.caption = "Close"
    End If
    Set temprs = Nothing

    dtFilter_Change
    PicAdd.Visible = True
    PicSave.Visible = False
    picDetail.Enabled = False
End Sub



Private Sub dtFilter_Change()
    Dim SQL                                  As String
    If optDated(0).Value = True Then
        SQL = "SELECT LogID, " _
            & " Convert(varchar, DateTimeCall,101) as DateCall, " _
            & " Convert(varchar, DateTimeCall,108) as TimeCall, " _
            & " Duration, " _
            & " Subject FROM CRIS_PROSPECT_Calls WHERE DateTimeCall >='" & FormatDateTime(dtFilter.Value, vbShortDate) & "' and ProfileID=" & ProfileId

    Else

        SQL = "SELECT LogID, " _
            & " Convert(varchar, DateTimeCall,101) as DateCall, " _
            & " Convert(varchar, DateTimeCall,108) as TimeCall, " _
            & " Duration, " _
            & " Subject FROM CRIS_PROSPECT_Calls WHERE DateTimeCall <='" & FormatDateTime(dtFilter.Value, vbShortDate) & "' and ProfileID=" & ProfileId
    End If
    'SQL = "SELECT LogID, Convert(varchar, DateTimeCall,101) as DateCall, Convert(varchar, DateTimeCall,108) as TimeCall, Duration, Subject FROM CRIS_PROSPECT_CALLS WHERE DateTimeCall >'" & FormatDateTime(dtFilter.Value, vbShortDate) & "' and ProfileID=" & ProfileID
    flex_FillReportView gconDMIS.Execute(SQL), lvGrid, False
End Sub

Private Sub FillView()

End Sub

Private Sub cmdOk_Click()
    Dim t1 As String, t2                     As String
    Dim temprs                               As ADODB.Recordset
    Dim SQL  As String
    If LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospect_Calls " _
            & " (ProspectID,  DateTimeCall, Duration, Subject, Comments,Bound) " _
            & " VALUES(@ProspectID, @DateTimeCall, @Duration, @Subject, @Comments,@Bound)" & vbCrLf & "SELECT @@IDENTITY"


        
    Else
        SQL = "Update CRIS_Prospect_Calls SET " _
            & " ProspectID=@ProspectID , " _
            & " DateTimeCall=@DateTimeCall , " _
            & " Duration=@Duration , " _
            & " Subject=@Subject , " _
            & " Comments=@Comments , " _
            & " Bound=@Bound " _
            & " WHERE LogID=@LogID "
    End If



    SQL = Replace(SQL, "@LogID", LOGID)
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@DateTimeCall", N2Str2Null(FormatDateTime(dtDateCall.Value, vbShortDate)))
    SQL = Replace(SQL, "@Duration", CInt(txtDuration.Text))
    SQL = Replace(SQL, "@Subject", N2Str2Null(txtSubject))
    SQL = Replace(SQL, "@Comments", N2Str2Null(txtComments))
    SQL = Replace(SQL, "@Bound", N2Str2Null(cboCallType.Text))


   
    Set temprs = gconDMIS.Execute(SQL)
    gconDMIS.Execute ("update CRIS_PROSPECTs SET LogCall=" & N2Str2Null(t1) & " where prospectid=" & ProspectID)

    If LOGID <= 0 Then
        MessagePop RecSave, "Record Added ", "New Schedule Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Schedule Sucessfully Updated", 500, 1
    End If

    Set temprs = temprs.NextRecordset
    If Not temprs Is Nothing Then
        LOGID = temprs.Collect(0)
    End If


    Set temprs = Nothing

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    InitData
End Sub



Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then Exit Sub
    LOGID = Row.Record(0).Value
    InitData

End Sub

Sub InitData()
With cboCallType
    .AddItem ("In Bound")
    .AddItem ("Out Bound")
    .ListIndex = 0
End With
dtDateCall.Value = Now
dtTimeCall.Value = Now
End Sub
Private Sub optDated_Click(Index As Integer)
    txtFilter.Text = vbNullString
End Sub

Private Sub Timer1_Timer()
    Dim cntrl                                As Control
    For Each cntrl In Me.Controls
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If
        End If
    Next
    Timer1.Enabled = False
End Sub

Private Sub txtFilter_Change()
    lvGrid.FilterText = txtFilter.Text
    lvGrid.Populate
End Sub

Friend Sub AddCall(xProfileID As Long, xProfileType As String, xAcctName As String, xProspectID As Long)
    LOGID = 0
    ProfileType = xProfileType
    ProfileId = xProfileID
    AcctName = xAcctName
    ProspectID = xProspectID
    LabelIt
End Sub
Sub LabelIt()
    Dim temprs                               As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select * from   CRIS_vW_AllProfile where Profileid=" & ProfileId & " and ProfileTYpe =" & N2Str2Null(ProfileType))

    If Not (temprs.EOF Or temprs.BOF) Then

        lblCustomerName.caption = Null2String(temprs("ProfileName").Value)
        AcctName = Null2String(temprs("AcctName").Value)
        lblAccountName.caption = Null2String(temprs("AcctName").Value)
        lblAddress.caption = Null2String(temprs("Address").Value)
        lblContactNo.caption = Null2String(temprs("Phone").Value)
        lblEmail.caption = Null2String(temprs("Email").Value)
    End If
    Set temprs = Nothing

End Sub



Private Function Runvalidation(strcase As String) As Boolean
    Runvalidation = False
    Dim txt                                  As Control
    For Each txt In Me.Controls
        If (TypeOf txt Is TextBox Or TypeOf txt Is ComboBox) And txt.Tag = strcase Then
            If Trim(txt.Text) = vbNullString Then
                MessagePop RecSaveError, "Required Filed Missing", txt.ToolTipText & " is Required Field", 1000
                Call ColorIt(txt, Timer1)
                txt.SetFocus
                Exit Function
            End If
        End If
    Next
    Runvalidation = True
End Function





