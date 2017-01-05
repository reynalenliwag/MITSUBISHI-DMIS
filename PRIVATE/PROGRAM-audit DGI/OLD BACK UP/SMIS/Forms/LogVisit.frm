VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSMIS_Log_Visit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Visit"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LogVisit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtResults 
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
      Left            =   150
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2010
      Width           =   4425
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
      Height          =   825
      Left            =   150
      TabIndex        =   4
      Top             =   840
      Width           =   4275
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Height          =   435
      Left            =   2640
      TabIndex        =   7
      Top             =   4200
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   435
      Left            =   3630
      TabIndex        =   8
      Top             =   4200
      Width           =   945
   End
   Begin MSComCtl2.DTPicker txtDateVisit 
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   1545
      _ExtentX        =   2725
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
      Format          =   52232193
      CurrentDate     =   39139
   End
   Begin MSComCtl2.DTPicker txtTimeVisit 
      Height          =   345
      Left            =   1770
      TabIndex        =   2
      Top             =   270
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
      Format          =   52232194
      CurrentDate     =   39139
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Results"
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
      Left            =   150
      TabIndex        =   5
      Top             =   1710
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Time Visited"
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
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   1485
   End
   Begin VB.Label Label4 
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
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "frmSMIS_Log_Visit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProspectID                         As Long
Dim LOGID                              As Long

Private Sub cmdOk_Click()
    Dim TEMPRS                         As ADODB.Recordset
    Dim SQL                            As String
    If LOGID <= 0 Then
        SQL = "INSERT INTO CRIS_Prospect_Visits " _
            & " (ProspectID,  DateTimeVisit, Comments,Results) " _
            & " VALUES(@ProspectID,  @DateTimeVisit,  @Comments, @Results)" & vbCrLf & "SELECT @@IDENTITY"
    Else
        SQL = "Update CRIS_Prospect_Visits SET  " _
            & " ProspectID=@ProspectID, " _
            & " DateTimeVisit=@DateTimeVisit, " _
            & " Comments=@Comments, " _
            & " Results=@Results " _
            & " WHERE LogID=@LogID "
    End If
    SQL = Replace(SQL, "@LogID", LOGID)
    SQL = Replace(SQL, "@ProspectID", ProspectID)
    SQL = Replace(SQL, "@DateTimeVisit", N2Str2Null(txtDateVisit.Value))
    SQL = Replace(SQL, "@Comments", N2Str2Null(txtComments))
    SQL = Replace(SQL, "@Results", N2Str2Null(txtResults))

    Set TEMPRS = gconDMIS.Execute(SQL)
    gconDMIS.Execute ("update CRIS_PROSPECTS SET LogVisit=" & N2Str2Null(FormatDateTime(txtDateVisit.Value)) & " where prospectid=" & ProspectID)

    If LOGID <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Visit Sucessfully Added", 500, 1
    Else
        MessagePop RecSaveOk, "RecordSaved", "Visit Sucessfully Updated", 500, 1
    End If

    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        LOGID = TEMPRS.Collect(0)
    End If
    
       If FormExist("MainForm") Then
        MainForm.ProspectStatus.ShowStatus ProspectID
        End If
 
    Set TEMPRS = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitData
End Sub

Sub InitData()
    txtDateVisit.Value = Now
    txtTimeVisit.Value = Now
    txtComments.Text = vbNullString
End Sub


Friend Sub AddVisit(xProsID As Long)
    LOGID = 0
    ProspectID = xProsID
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProspectID = 0
    LOGID = 0
End Sub
