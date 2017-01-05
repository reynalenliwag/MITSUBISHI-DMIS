VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "Biokey.ocx"
Begin VB.Form frmHRMSEmpEnrollBio 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Finger Print Enroll"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12930
   StartUpPosition =   1  'CenterOwner
   Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX 
      Left            =   8130
      Top             =   480
      EnrollCount     =   3
      SensorIndex     =   0
      Threshold       =   10
      VerTplFileName  =   ""
      RegTplFileName  =   ""
      OneToOneThreshold=   10
      Active          =   0   'False
      IsRegister      =   0   'False
      EnrollIndex     =   0
      SensorSN        =   ""
      FPEngineVersion =   "Biokey 4.0"
      ImageWidth      =   0
      ImageHeight     =   0
      SensorCount     =   0
      TemplateLen     =   1152
      EngineValid     =   0   'False
      ForceSecondMatch=   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   4860
      ScaleHeight     =   2355
      ScaleWidth      =   2535
      TabIndex        =   17
      Top             =   990
      Width           =   2595
      Begin VB.Image Image1 
         Height          =   2355
         Left            =   0
         Picture         =   "EmpEnrollBio.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.PictureBox picTemplate1 
      Height          =   1905
      Left            =   240
      ScaleHeight     =   1845
      ScaleWidth      =   2175
      TabIndex        =   15
      Top             =   990
      Width           =   2235
      Begin VB.Image imgTemplate1 
         Height          =   1815
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label labFP1 
         Alignment       =   2  'Center
         Caption         =   "Register you Finger Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   480
         TabIndex        =   16
         Top             =   510
         Width           =   1215
      End
   End
   Begin VB.PictureBox picTemplate2 
      Height          =   1905
      Left            =   2550
      ScaleHeight     =   1845
      ScaleWidth      =   2175
      TabIndex        =   13
      Top             =   990
      Width           =   2235
      Begin VB.Image imgTemplate2 
         Height          =   1815
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label labFP2 
         Alignment       =   2  'Center
         Caption         =   "Register you Finger Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   480
         TabIndex        =   14
         Top             =   510
         Width           =   1155
      End
   End
   Begin VB.OptionButton optFingerPrint2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finger Print 2"
      Height          =   345
      Left            =   2550
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.OptionButton optFingerPrint1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finger Print 1"
      Height          =   345
      Left            =   270
      TabIndex        =   9
      Top             =   3000
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6990
      Top             =   60
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6375
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12118
            MinWidth        =   12118
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4740
      TabIndex        =   5
      Top             =   5850
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enroll Finger Print"
      Height          =   465
      Left            =   210
      TabIndex        =   3
      Top             =   3510
      Width           =   2265
   End
   Begin VB.TextBox txtEmpName 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1770
      MaxLength       =   250
      TabIndex        =   2
      Text            =   "LESLIE ARANZA"
      Top             =   540
      Width           =   5685
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6180
      TabIndex        =   6
      Top             =   5850
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   5400
      Left            =   7830
      Picture         =   "EmpEnrollBio.frx":42A2
      Stretch         =   -1  'True
      Top             =   270
      Width           =   4890
   End
   Begin VB.Label labEmpNo 
      BackStyle       =   0  'Transparent
      Caption         =   "0001"
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   210
      Width           =   2415
   End
   Begin VB.Label labFingerStatus 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   7830
      TabIndex        =   11
      Top             =   5850
      Width           =   4905
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5505
      Left            =   7770
      Top             =   240
      Width           =   4995
   End
   Begin VB.Label StatusBar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1455
      Left            =   180
      TabIndex        =   8
      Top             =   4260
      Width           =   7275
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   7470
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Label labStatus1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2580
      TabIndex        =   4
      Top             =   3510
      Width           =   4875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp. Name. :"
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   570
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp. No.      :"
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1605
   End
End
Attribute VB_Name = "frmHRMSEmpEnrollBio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FTempLen                                                          As Integer
Dim FRegTemplate                                                      As Variant
Dim FingerCount                                                       As Long
Dim fpcHandle                                                         As Long
Dim FFingerNames()                                                    As String
Dim FMatchType                                                        As Integer
Dim rsEmpInfo                                                         As ADODB.Recordset

Sub rsrefresh()
    On Error GoTo ERROR_REFRESH
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("Select FPIMAGE1,FPIMAGE2,EMPNO,LASTNAME,FIRSTNAME,MIDDLENAME from HRMS_EmpInfo WHERE EMPNO = '" & EmpInfoEmpno.Caption & "'")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        EmpInfoEmpno.Caption = Null2String(rsEmpInfo!EMPNO)
        labEmpNo.Caption = Null2String(rsEmpInfo!EMPNO)
        txtEmpName.Text = Null2String(rsEmpInfo!lastname) & "," & Null2String(rsEmpInfo!FIRSTNAME) & " " & Null2String(rsEmpInfo!MIDDLENAME) & "."
        Dim st                                                        As New ADODB.Stream
        Dim strTemp                                                   As String
        st.Type = adTypeBinary
        If IsNull(rsEmpInfo!FPIMAGE1) = False Then
            st.Open: labFP1.Visible = False
            st.Write rsEmpInfo.FIELDS("FPIMAGE1").Value
            st.SaveToFile Environ("TEMP") & "\FPIMAGE1", adSaveCreateOverWrite
            imgTemplate1.Picture = LoadPicture(Environ("TEMP") & "\FPIMAGE1")
            Kill (Environ("TEMP") & "\FPIMAGE1")
            st.Close
        Else
            LoadPic imgTemplate1, "": labFP1.Visible = True
        End If
        If IsNull(rsEmpInfo!FPIMAGE2) = False Then
            st.Open: labFP2.Visible = False
            st.Write rsEmpInfo.FIELDS("FPIMAGE2").Value
            st.SaveToFile Environ("TEMP") & "\FPIMAGE2", adSaveCreateOverWrite
            imgTemplate2.Picture = LoadPicture(Environ("TEMP") & "\FPIMAGE2")
            Kill (Environ("TEMP") & "\FPIMAGE2")
            st.Close
        Else
            LoadPic imgTemplate2, "": labFP2.Visible = True
        End If
        Set st = Nothing
        Image2.ZOrder 0
    End If
    
    Exit Sub
ERROR_REFRESH:
    MessagePop InfoStop, "Unknown Error", "" & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub

Sub InitializeScanner()
    On Error GoTo INITIALZE_ERROR
    ZKFPEngX.SensorIndex = 0
    If ZKFPEngX.InitEngine = 0 Then
        StatusBar1.Panels(1).Text = "Sensor Connected"
        StatusBar1.Panels(2).Text = "S/N: " & ZKFPEngX.SensorSN
        FMatchType = 0
    Else
        MsgBox "Finger Print Scanner Not Found!"
    End If
    
    Exit Sub
INITIALZE_ERROR:
    MessagePop InfoStop, "Unknown Error", "" & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub

Private Sub Command1_Click()
    If labEmpNo.Caption = "" Then
        MsgBox "Please Check Employee No."
        Exit Sub
    End If
    
    'MsgBox "Test", vbInformation, "Button Click"
    Command4.Enabled = False
    ZKFPEngX.BeginEnroll
    labFingerStatus.Caption = ""
    labStatus1.Caption = "Begin Register"
    StatusBar.Caption = ""
End Sub

Private Sub Command4_Click()
    Dim fi As Long, I                                                 As Long
    Dim Score As Long, ProcessNum                                     As Long
    Dim RegChanged                                                    As Boolean
    Dim sTemp                                                         As String
    sTemp = ZKFPEngX.GetTemplateAsString()
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("Select * from HRMS_EMPINFO Where EMPNO <> '" & labEmpNo.Caption & "' Order by EmpNo asc")
    If Not rsEmpInfo.EOF Then
        Do While Not rsEmpInfo.EOF
            If ZKFPEngX.VerFingerFromStr(Null2String(rsEmpInfo!fptemplate1), sTemp, False, RegChanged) = True Then
                'If MsgBox("Finger Print Already registered to " & Null2String(RSEMPINFO!lastname) & "," & Null2String(RSEMPINFO!FIRSTNAME) & " would you like to continue?", vbQuestion + vbYesNo, "Existing Found!") = vbNo Then
                MsgBox "Finger Print Already registered to " & Null2String(rsEmpInfo!lastname) & "," & Null2String(rsEmpInfo!FIRSTNAME), vbCritical, "Existing Found!"
                Exit Sub
                'End If
                Exit Do
            End If
            If ZKFPEngX.VerFingerFromStr(Null2String(rsEmpInfo!fptemplate2), sTemp, False, RegChanged) = True Then
                'If MsgBox("Finger Print Already registered to " & Null2String(RSEMPINFO!lastname) & "," & Null2String(RSEMPINFO!FIRSTNAME) & " would you like to continue?", vbQuestion + vbYesNo, "Existing Found!") = vbNo Then
                MsgBox "Finger Print Already registered to " & Null2String(rsEmpInfo!lastname) & "," & Null2String(rsEmpInfo!FIRSTNAME), vbCritical, "Existing Found!"
                Exit Sub
                'End If
                Exit Do
            End If
            rsEmpInfo.MoveNext
        Loop
    End If

    Dim mstream                                                       As ADODB.Stream
    Set mstream = New ADODB.Stream
    mstream.Type = adTypeBinary
    mstream.Open
    mstream.LoadFromFile Environ("TEMP") & "\FingerPrint.Jpg"
    Dim rsEmpInfoUpdate                                               As ADODB.Recordset
    Set rsEmpInfoUpdate = New ADODB.Recordset
    rsEmpInfoUpdate.Open "Select * from HRMS_EMPINFO WHERE EMPNO = '" & labEmpNo.Caption & "'", gconDMIS, adOpenKeyset, adLockOptimistic
    If optFingerPrint1.Value = True Then
        rsEmpInfoUpdate.FIELDS("FPTEMPLATE1").Value = ZKFPEngX.GetTemplateAsString()
        rsEmpInfoUpdate.FIELDS("FPIMAGE1").Value = mstream.Read
    Else
        rsEmpInfoUpdate.FIELDS("FPTEMPLATE2").Value = ZKFPEngX.GetTemplateAsString()
        rsEmpInfoUpdate.FIELDS("FPIMAGE2").Value = mstream.Read
    End If
    rsEmpInfoUpdate.Update
    'labStatus1.Caption = "Click Enroll Finger"
    Call rsrefresh: DoEvents
    MsgBox "Finger Print Successfully Saved in Database.", vbInformation, "Saved..."
    'Command1.Value = True
    Command4.Enabled = False
    Set rsEmpInfoUpdate = Nothing
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()

'orig jan 6
'    Screen.MousePointer = 11
'    StatusBar1.Panels(1).Text = "Initializing Sensor... Please Wait..."
'    fpcHandle = ZKFPEngX.CreateFPCacheDB
'    Call InitializeScanner
'    Call rsrefresh
'    Screen.MousePointer = 0

    Screen.MousePointer = 11
    StatusBar1.Panels(1).Text = "Initializing Sensor... Please Wait..."
    fpcHandle = ZKFPEngX.CreateFPCacheDB: InitializeScanner: rsrefresh
    Screen.MousePointer = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ZKFPEngX.FreeFPCacheDB (fpcHandle)
End Sub

Private Sub optFingerPrint1_Click()
    On Error GoTo ERROR_OPTION1
    ZKFPEngX.FlushFPImages
    labFingerStatus.Caption = ""
    labStatus1.Caption = "Click Enroll Finger"
    StatusBar.Caption = ""
    
    Exit Sub
ERROR_OPTION1:
    MessagePop InfoStop, "Unknown Error", "" & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub

Private Sub optFingerPrint2_Click()
    ZKFPEngX.FlushFPImages
    labFingerStatus.Caption = ""
    labStatus1.Caption = "Click Enroll Finger"
    StatusBar.Caption = ""
End Sub

Private Sub ZKFPEngX_OnCapture(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)


'orig jan 6
'    On Error GoTo ERROR_1
'    Dim sTemp                                                         As String
'    sTemp = ZKFPEngX.GetTemplateAsString()
'
'    'MsgBox sTemp, vbInformation, "Info"
'
'    Exit Sub
'ERROR_1:
'    MessagePop InfoStop, "Unknown Error", " " & Err.NUMBER & " : " & Err.Description
'    Err.Clear

Dim sTemp                                                         As String
sTemp = ZKFPEngX.GetTemplateAsString()

End Sub

Private Sub ZKFPEngX_OnCaptureToFile(ByVal ActionResult As Boolean)

End Sub

Private Sub ZKFPEngX_OnEnroll(ByVal ActionResult As Boolean, ByVal ATemplate As Variant)
    On Error GoTo ERROR_2
    
    If Not ActionResult Then
        Command4.Enabled = False
        MsgBox "Finger Print Register Failed!", vbCritical, "Not Matched"
        Command1.Value = True
    Else
        Command4.Enabled = True
        MsgBox "Finger Print Register Succeeded! Press the save button to register your finger print in Database.", vbInformation, "Success"
        FRegTemplate = ATemplate
        ZKFPEngX.SaveJPG Environ("TEMP") & "\FingerPrint.Jpg"
        ZKFPEngX.AddRegTemplateStrToFPCacheDB fpcHandle, FingerCount, ZKFPEngX.GetTemplateAsString()
        ReDim Preserve FFingerNames(FingerCount + 1)
        
        'MsgBox "test"
        FFingerNames(FingerCount) = labEmpNo.Caption
        FingerCount = FingerCount + 1
    End If
    
    Exit Sub
ERROR_2:
    MessagePop InfoStop, "Unknown Error", " " & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub

Private Sub ZKFPEngX_OnFeatureInfo(ByVal AQuality As Long)
    On Error GoTo ERROR_3
    
    StatusBar = ""
    If ZKFPEngX.IsRegister Then
        StatusBar = "Register Rtatus: Remaining Scan for Finger Print " & vbCrLf & "[" & ZKFPEngX.EnrollIndex - 1 & " time(s)]"
        'MsgBox "test", vbInformation, "1"
    End If
    If AQuality <> 0 Then
        labFingerStatus = "[Not Good]"
        
        'MsgBox "test", vbInformation, "not good"
    Else
        labFingerStatus = "[Good]"
        'MsgBox "test", vbInformation, "good"
    End If
    
    Exit Sub
ERROR_3:
    MessagePop InfoStop, "Unknown Error", " " & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub

Private Sub ZKFPEngX_OnImageReceived(AImageValid As Boolean)
    On Error GoTo ERROR_4
    
    'MsgBox "test", vbInformation, "test"
    ZKFPEngX.PrintImageAt hdc, 520, 20, ZKFPEngX.ImageWidth, ZKFPEngX.ImageHeight
        
        
    Exit Sub
ERROR_4:
    MessagePop InfoStop, "Unknown Error", " " & Err.NUMBER & " : " & Err.Description
    Err.Clear
End Sub


