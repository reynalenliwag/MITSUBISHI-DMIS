VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMIS SMS"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowMessage 
      Height          =   705
      Left            =   12120
      Picture         =   "frmSMS.frx":20D2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Show Inbox"
      Top             =   4380
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   705
      Left            =   2550
      Picture         =   "frmSMS.frx":3154
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   4380
      Width           =   555
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4005
      Left            =   3120
      TabIndex        =   3
      Top             =   330
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tranno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Message"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Height          =   705
      Left            =   2010
      Picture         =   "frmSMS.frx":5226
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Send"
      Top             =   4380
      Width           =   555
   End
   Begin VB.Timer TimerCheckGSMSignal 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8220
      Top             =   4620
   End
   Begin VB.TextBox txtBody 
      Height          =   3225
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1110
      Width           =   3075
   End
   Begin VB.TextBox txtDestination 
      Height          =   405
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   3105
   End
   Begin VB.CommandButton cmdNew 
      Height          =   705
      Left            =   0
      Picture         =   "frmSMS.frx":72F8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "New Message"
      Top             =   4380
      Width           =   555
   End
   Begin MSComctlLib.ListView lsvPhoneBook 
      Height          =   3465
      Left            =   0
      TabIndex        =   12
      Top             =   5130
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   6112
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mobile No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   345
      Index           =   1
      Left            =   3150
      TabIndex        =   10
      Top             =   0
      Width           =   9525
      _Version        =   655364
      _ExtentX        =   16801
      _ExtentY        =   609
      _StockProps     =   14
      Caption         =   "Inbox"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   8388608
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   3105
      _Version        =   655364
      _ExtentX        =   5477
      _ExtentY        =   609
      _StockProps     =   14
      Caption         =   "Mobile No."
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   8388608
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   750
      Width           =   3075
      _Version        =   655364
      _ExtentX        =   5424
      _ExtentY        =   609
      _StockProps     =   14
      Caption         =   "Message"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   8388608
   End
   Begin VB.Label labTranno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   3270
      TabIndex        =   6
      Top             =   4800
      Width           =   1275
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   765
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   4350
      Width           =   12675
      _Version        =   655364
      _ExtentX        =   22357
      _ExtentY        =   1349
      _StockProps     =   14
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Public WithEvents Modem As SMSModem
Attribute Modem.VB_VarHelpID = -1
Dim FIRST_LOD       As Integer

Private mLoading As Boolean
Private Settings As RegistrySettings
Private Type RegistrySettings
    ModemType As GSMModemTypeConstants
    Port As Integer
    CheckGSMSignal As Integer
    UnregisteredMode As Integer
    AutoReceive As Integer
    AutoDelete As Integer
    PhoneNumber As String
    LongMsg As Integer
    StatusReport As Integer
    HideHelp As Boolean
End Type




Private Sub ConnectToSIM()
    On Error GoTo hErr
    Screen.MousePointer = vbHourglass
    
    ' open/reset communications, initialize modem & wait for modem registration
    Modem.CloseComm
    Modem.LogTrace = True
    Modem.OpenComm _
        4, _
        4, _
        "", _
        smsNotifyAll, _
        60, _
        False
    
    
    Call chkGSMSignal_Click
    Modem.AutoDelete = False
        
    Screen.MousePointer = vbDefault

    ' check SIM busy locations
    If Modem.GetSimBusyLocationsCount = Modem.GetSimLocationsCount Then
        MsgBox "SIM card is full." & vbCrLf & "Please read and delete SIM messages.", vbExclamation
    End If

Exit Sub
hErr:
    Screen.MousePointer = vbDefault
    Report "Error connecting to modem " & Modem.GetManufacturer & " " & Modem.GetModel
    Call ShowError
End Sub

Private Sub chkGSMSignal_Click()
    Settings.CheckGSMSignal = False
    TimerCheckGSMSignal.Enabled = False
    If TimerCheckGSMSignal.Enabled Then Call TimerCheckGSMSignal_Timer
End Sub



Private Sub cmdNew_Click()
    txtDestination.Text = ""
    txtBody.Text = ""
    labTranno.Caption = ""
    txtDestination.Enabled = True
End Sub

Private Sub cmdSend_Click()
    On Error GoTo hErr
    
    If Modem.Status <> smsModemReady Then _
        MsgBox "Please open modem communication before!", vbExclamation: Exit Sub


    Dim msg                         As New SMSSubmit
    Dim msgRef                      As Integer
    Dim xTranno                     As String
    DATA_SMS.SMS_CONN.Open
    If labTranno.Caption = "" Then
        xTranno = GenerateTranno
    Else
        xTranno = labTranno
    End If
    
    With msg
        Set .Destination = GetGSMAddress(txtDestination)
        
        .Alphabet = gsmAlphabetText
        .Body = xTranno & " " & txtBody
        .UserDataHeader = Hex2Str("")
                
        .StatusReportRequest = False
    End With
    
    '// Send message
    Screen.MousePointer = vbHourglass
    msgRef = Modem.SendMessage(msg)
    
    DATA_SMS.SMS_CONN.Execute ("INSERT INTO ALL_SMS (TRANNO, TRANDATE, TRANTIME, SENDER, RECIEVER, MESSAGE) " & _
        " VALUES('" & xTranno & _
        "', '" & DateValue(Now) & _
        "', '" & TimeValue(Now) & _
        "', '+649233370215', '" & txtDestination & _
        "', '" & txtBody & "')")
    DATA_SMS.SMS_CONN.Close
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IMPORTANT NOTE:
    ' Text messages can also be sent as follows:
    ' msgRef = Modem.SendTextMessage(txtDestination, txtBody)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    txtBody.SetFocus
    Screen.MousePointer = vbDefault
    MsgBox "Message Send", vbInformation, "Send"
    
    

Exit Sub
hErr:
    Screen.MousePointer = vbDefault
    Call ShowError
End Sub

Function GenerateTranno() As String
    Dim rstmp    As New ADODB.Recordset
    Set rstmp = DATA_SMS.SMS_CONN.Execute("SELECT CAST(TRANNO AS INT) FROM ALL_SMS ORDER BY TRANNO DESC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GenerateTranno = Format(rstmp.Fields(0) + 1, "000000")
    Else
        GenerateTranno = "000001"
    End If
    Set rstmp = Nothing
End Function
Public Function Hex2Str(ByVal hexStr As String) As String
    Dim ret$, hex As String * 2, c%, k%, slen As Integer
    slen = Len(hexStr)
    If slen Mod 2 = 1 Then _
        Err.Raise 13, , "Invalid hex string (string is composed by an odd number of digits)" ' Type mismatch
    
    For k = 1 To slen Step 2
        hex = Mid$(hexStr, k, 2)
        If Not IsNumeric("&h" & hex) Then Err.Raise 13, , "Invalid hex string" ' Type mismatch
        c = Val("&h" & hex)
        ret = ret & Chr$(c)
    Next k
    Hex2Str = ret
End Function

Private Sub cmdShowMessage_Click()
    'On Error GoTo hErr
    Screen.MousePointer = vbHourglass
    Dim rstmp   As New ADODB.Recordset
    Dim Item    As ListItem
    DATA_SMS.SMS_CONN.Open
    ListView1.ListItems.Clear
    Set rstmp = DATA_SMS.SMS_CONN.Execute("SELECT * FROM ALL_SMS WHERE  RECIEVER IS NULL ORDER BY TRANDATE, TRANTIME ASC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = ListView1.ListItems.Add(, , rstmp!TRANNO)
            Item.SubItems(1) = rstmp!SENDER
            Item.SubItems(2) = DateValue(rstmp!TRANDATE)
            Item.SubItems(3) = TimeValue(rstmp!TRANTIME)
            Item.SubItems(4) = rstmp!message
            Item.SubItems(5) = rstmp!TRAN_ID
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
    Screen.MousePointer = 0
    DATA_SMS.SMS_CONN.Close
'    ListView1.ListItems.Clear
'    '// Query for received messages
'    Dim messages                            As Collection
'    Dim msg                                 As SMSDeliver
'
'    Set messages = Modem.ReadReceivedMessages(False, False)
'
'
'
'    '// Print to report
'    If messages.Count = 0 Then
'        MsgBox "No messages received", vbInformation, "Info"
'    Else
'        For Each msg In messages
'            ShowMessageReceived msg
'        Next msg
'    End If
'    Screen.MousePointer = vbDefault
'
'Exit Sub
'hErr:
'    Screen.MousePointer = vbDefault
'    Call ShowError
End Sub

Private Sub cmdCancel_Click()
    txtDestination.Text = ""
    txtBody.Text = ""
    txtDestination.Enabled = True
End Sub

Private Sub Form_Load()
    mLoading = True
    
    ' create modem instance
    Set Modem = New SMSModem
    FIRST_LOD = 0
    ' init form controls & show help
    
    Call LoadSettings
    Call ConnectToSIM
    
    Call FillPhoneBook
    'Call Settings2Form
    'App.HelpFile = App.Path & "\SMSLibX.CHM"
    'If Not Settings.HideHelp Then frmHelp.Show vbModeless
    'Me.Width = 3240
    ListView1.ListItems.Clear
    FIRST_LOD = 1
End Sub

Sub FillPhoneBook()
    On Error GoTo hErr
    Screen.MousePointer = vbHourglass
    lsvPhoneBook.ListItems.Clear
    '// Query phonebook items
    Dim pbItems         As Collection
    Dim pbi             As PhonebookItem
    Dim Item            As ListItem
    ' a) full read method (it can take some seconds)
    Set pbItems = Modem.ReadPhonebook()
    ' b) quick read method (only for phonebooks containing subsequent locations)
    'Set pbItems = Modem.ReadPhonebook(1, Modem.GetPhonebookBusy())
    
    '// Print to report
    If pbItems.Count = 0 Then
        MsgBox "No phonebook items", vbInformation, "Info"
    Else
        For Each pbi In pbItems
            Set Item = lsvPhoneBook.ListItems.Add(, , pbi.Number)
            Item.SubItems(1) = pbi.Name
            'ShowPhonebookItem pbi
        Next pbi
    End If
    Screen.MousePointer = vbDefault

Exit Sub
hErr:
    Screen.MousePointer = vbDefault
    Call ShowError
End Sub

'Private Sub ShowPhonebookItem(pbi As PhonebookItem, Optional comment As String)
'    Report "Phonebook (" & Format$(pbi.Index) & "): " _
'            & pbi.Name & ", " & pbi.Number _
'            & " " & comment
'End Sub
'Private Sub Settings2Form()
'    Dim i As Integer
'    For i = 0 To cmbModemType.ListCount - 1
'        If cmbModemType.ItemData(i) = Settings.ModemType Then cmbModemType.ListIndex = i
'    Next i
'    Call cmbModemType_Click
'    cmbPort.ListIndex = Settings.Port - 1
'    chkGSMSignal.Value = IIf(Settings.CheckGSMSignal, vbChecked, vbUnchecked)
'    chkSkipNetworkRegistration.Value = IIf(Settings.UnregisteredMode, vbChecked, vbUnchecked)
'    chkAutoReceive.Value = IIf(Settings.AutoReceive, vbChecked, vbUnchecked)
'    chkAutoDelete.Value = IIf(Settings.AutoDelete, vbChecked, vbUnchecked)
'    txtDestination = Settings.PhoneNumber
'    chkLongMsg.Value = IIf(Settings.LongMsg, vbChecked, vbUnchecked)
'    chkStatusReport.Value = IIf(Settings.StatusReport, vbChecked, vbUnchecked)
'    Call chkAutoReceive_Click
'    cmbAlphabet.ListIndex = 0
'    cmbClass.ListIndex = 0
'End Sub

Private Sub Report(ByVal v As Variant)
    Dim row$
    Dim Item As ListItem
    
    row = Format$(Time, "hh:nn:ss  ") & Format$(v)
    'Debug.Print row
    row = Replace(row, vbCrLf, " ")
    row = Replace(row, vbLf, " ")
    
    Set Item = ListView1.ListItems.Add(, , row)
    '
    'lstReport .AddItem row
    'lstReport.ListIndex = lstReport.ListCount - 1
End Sub

Private Sub LoadSettings()
    Const AppName = "SMSLibX"
    Dim Section As String: Section = App.EXEName
    On Error Resume Next
    Settings.ModemType = CInt(GetSetting(AppName, Section, "ModemType", gsmModemDummy))
    Settings.Port = CInt(GetSetting(AppName, Section, "Port", 1))
    Settings.CheckGSMSignal = CInt(GetSetting(AppName, Section, "CheckGSMSignal", 0))
    Settings.UnregisteredMode = CInt(GetSetting(AppName, Section, "UnregisteredMode", 0))
    Settings.AutoReceive = CInt(GetSetting(AppName, Section, "AutoReceive", 1))
    Settings.AutoDelete = 0 'CInt(GetSetting(AppName, Section, "AutoDelete", 0))
    Settings.PhoneNumber = GetSetting(AppName, Section, "PhoneNumber", "+")
    Settings.LongMsg = CInt(GetSetting(AppName, Section, "LongMsg", 0))
    Settings.StatusReport = CInt(GetSetting(AppName, Section, "StatusReport", 0))
    Settings.HideHelp = CBool(GetSetting(AppName, Section, "HideHelp", 0))
End Sub

Private Function IsObjectErr(ByVal raisedErrNumber As Long) As Boolean
    IsObjectErr = (raisedErrNumber >= vbObjectError And raisedErrNumber <= vbObjectError + 65535)
End Function

Private Sub ShowError()
    Dim errNum As Long
    errNum = IIf(IsObjectErr(Err.Number), Err.Number - vbObjectError, Err.Number)
    Report "Error #" & Format$(errNum) & " from " & Err.Source _
            & ": " & Err.Description
    MsgBox Err.Description, vbExclamation, _
           "Error #" & Format$(errNum) & " from " & Err.Source
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtDestination.Enabled = False
    labTranno.Caption = Item.Text
    txtDestination.Text = Item.SubItems(1)
End Sub

Private Sub Modem_MessageReceived(message As SMSLibX.SMSDeliver)
    'Command4_Click
    'Command4_Click
    'ShowMessageReceived message
    If FIRST_LOD <> 0 Then SaveMessageToDatabase message
End Sub

Sub SaveMessageToDatabase(msg As SMSDeliver)
    DATA_SMS.SMS_CONN.Open
    
    Dim xTranno As String
    Dim xMSG    As String
    Dim xSMS    As String
    
    xSMS = IIf(msg.Alphabet = gsmAlphabetData, _
                    HexDump(msg.UserData), _
                    HexDump(msg.UserDataHeader) & msg.Body)
    
    xTranno = Left(xSMS, 6)
    xMSG = Mid(xSMS, 8, Len(xSMS))
    
    DATA_SMS.SMS_CONN.Execute ("INSERT INTO ALL_SMS (TRANNO, TRANDATE, TRANTIME, SENDER, MESSAGE) " & _
        " VALUES('" & xTranno & _
        "', '" & DateValue(msg.TimeStamp) & _
        "', '" & TimeValue(msg.TimeStamp) & _
        "', '" & msg.Originator & _
        "', '" & xMSG & "')")
    DATA_SMS.SMS_CONN.Close
End Sub

Private Sub ShowMessageReceived(msg As SMSDeliver)
    Dim simLoc As String
    Dim strNew As String
    Dim Item       As ListItem
    
    
    If msg.SimIndex >= 0 Then simLoc = "[SIM loc." & msg.SimIndex & "] "
    If msg.Unread Then strNew = "NEW "
    
    Set Item = ListView1.ListItems.Add(, , msg.Originator)
    Item.SubItems(1) = Format$(msg.TimeStamp, "dd/mm/yyyy hh:nn:ss")
    Item.SubItems(2) = IIf(msg.Alphabet = gsmAlphabetData, _
                    HexDump(msg.UserData), _
                    HexDump(msg.UserDataHeader) & msg.Body)
    Item.SubItems(3) = strNew
    

End Sub

Public Function HexDump(ByVal s As String) As String
    Dim ret$
    ret = DumpStr(s)
    If Len(ret) Then ret = "[" & ret & "]"
    HexDump = ret
End Function


Public Function DumpStr(ByVal s As String) As String
    Dim ret$, c%, pos%, slen%
    slen = Len(s)
    For pos = 1 To slen
        c = Asc(Mid$(s, pos, 1))
        If c <= &HF Then ret = ret & "0"
        ret = ret & hex$(c)
    Next pos
    DumpStr = ret
End Function

Private Sub TimerCheckGSMSignal_Timer()
    Dim ret As Boolean, rssi As Integer, ber As Integer
    On Error Resume Next
    If Modem.Status = smsModemReady Then
        ret = Modem.CheckSignalQuality(rssi, ber)
        If ret Then
            Report "GSM signal OK  (RSSI=" & rssi & ", BER=" & ber & ")"
        Else
            Report "GSM signal too low  (RSSI=" & rssi & ", BER=" & ber & "; required value for RSSI is [11..31])"
        End If
    End If
End Sub
