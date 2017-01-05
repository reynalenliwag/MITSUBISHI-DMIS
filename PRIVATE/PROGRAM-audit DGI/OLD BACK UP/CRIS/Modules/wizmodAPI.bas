Attribute VB_Name = "wizmodAPI"
Option Explicit
Public Const MAX_COMPUTERNAME_LENGTH         As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oVoice                                As SpeechLib.SpVoice

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public gconDMIS                              As ADODB.Connection
Public gconACCESS                            As ADODB.Connection
Public gconBIR_RELIEF                        As ADODB.Connection

Public Const ConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS;Data Source=SERVER"
Public Const DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS;Data Source=SERVER"
Public Const DMIS_REPORT_Connection = "DSN=DMIS;DSQ=DMIS"

Public COA_AR_TRADE_UNITS                    As String
Public COA_AR_TRADE_SERVICE                  As String
Public COA_AR_TRADE_PARTS                    As String

Public HYUNDAI_COA_AR_TRADE_UNITS            As String
Public HYUNDAI_COA_AR_TRADE_SERVICE          As String
Public HYUNDAI_COA_AR_TRADE_PARTS            As String

Public COA_OUTPUT_TAX                        As String

'SALES - CASH
Public COA_SALES_SERVICE_CASH_TINSPAINT      As String
Public COA_SALES_SERVICE_CASH_SUBLET         As String
Public COA_SALES_SERVICE_CASH_AIRCON         As String
Public COA_SALES_SERVICE_CASH_LABOR          As String
Public COA_SALES_SERVICE_CASH_GOL            As String
Public COA_SALES_SERVICE_CASH_PARTS          As String

Public COA_SALES_GOL_CASH                    As String
Public COA_SALES_PARTS_CASH                  As String
Public COA_SALES_VEHICLES_CASH               As String

'SALES - CHARGE
Public COA_SALES_SERVICE_CHARGE_TINSPAINT    As String
Public COA_SALES_SERVICE_CHARGE_SUBLET       As String
Public COA_SALES_SERVICE_CHARGE_AIRCON       As String
Public COA_SALES_SERVICE_CHARGE_LABOR        As String
Public COA_SALES_SERVICE_CHARGE_GOL          As String
Public COA_SALES_SERVICE_CHARGE_PARTS        As String

Public COA_SALES_GOL_CHARGE                  As String
Public COA_SALES_PARTS_CHARGE                As String

'SALES - DISCOUNT - CASH
Public COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_LABOR As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_GOL   As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_PARTS As String

Public COA_SALES_DISCOUNT_GOL_CASH           As String
Public COA_SALES_DISCOUNT_PARTS_CASH         As String
Public COA_SALES_DISCOUNT_VEHICLES_CASH      As String

'SALES - DISCOUNT - CHARGE
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS As String

Public COA_SALES_DISCOUNT_GOL_CHARGE         As String
Public COA_SALES_DISCOUNT_PARTS_CHARGE       As String

'CHARGE TO WARRANTY
Public COA_DIRECT_EXPENSE_LABOR              As String
Public COA_DIRECT_EXPENSE_SPAREPARTS         As String
Public COA_DIRECT_EXPENSE_GOL                As String

Public COA_WARRANTY_SALES                    As String
Public COA_WARRANTY_SERVICE                  As String
Public COA_WARRANTY_PARTS                    As String

'CHARGE TO COMPANY
Public COA_COMPANY_CAR_SALES                 As String
Public COA_COMPANY_CAR_SERVICE               As String

'CHARGE TO SALES
Public COA_GFSI_SALES                        As String
Public COA_GFSI_SERVICE                      As String
Public COA_GFSI_PARTS                        As String

'CASH RECEIPTS
Public COA_CASH_ON_HAND                      As String
Public COA_BRANCH_LEGASPI                    As String
Public COA_CUSTOMER_DEPOSIT                  As String

Public COA_INSURANCE_PREMIUM_PAYABLE         As String
Public COA_CHATTEL_MORTGAGE_FEE_PAYABLE      As String
Public COA_NEW_VEHICLE_REGISTRATION          As String
Public COA_WARRANTY_CLAIMS_RECEIVABLE        As String

Public COA_ACCOUNTS_RECEIVABLE_NONTRADE_EMPLOYEES As String
Public COA_OTHER_PAYABLES                    As String
Public COA_INCIDENTAL_CHARGES_UNITS          As String
Public COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD   As String

Public COA_CORPORATE_TAX_WHELD               As String
Public COA_CORPORATE_VAT_WHELD               As String

Public HYUNDAI_COA_SALES_SERVICE_CASH_TINSPAINT As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_SUBLET As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_AIRCON As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_LABOR  As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_GOL    As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_PARTS  As String

Public HYUNDAI_COA_SALES_GOL_CASH            As String
Public HYUNDAI_COA_SALES_PARTS_CASH          As String

'SALES - CHARGE
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_TINSPAINT As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_SUBLET As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_AIRCON As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_LABOR As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_GOL  As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_PARTS As String

Public HYUNDAI_COA_SALES_GOL_CHARGE          As String
Public HYUNDAI_COA_SALES_PARTS_CHARGE        As String

'SALES - DISCOUNT - CASH
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_LABOR As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_GOL As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_PARTS As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CASH   As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CASH As String

'SALES - DISCOUNT - CHARGE
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CHARGE As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CHARGE As String

Public HYUNDAI_COA_WARRANTY_SALES            As String
Public HYUNDAI_COA_WARRANTY_SERVICE          As String
Public HYUNDAI_COA_WARRANTY_PARTS            As String

Public HYUNDAI_COA_COMPANY_CAR_SALES         As String
Public HYUNDAI_COA_COMPANY_CAR_SERVICE       As String

Public HYUNDAI_COA_GFSI_SALES                As String
Public HYUNDAI_COA_GFSI_SERVICE              As String
Public HYUNDAI_COA_GFSI_PARTS                As String

Public HYUNDAI_COA_NEW_VEHICLE_REGISTRATION  As String
Public HYUNDAI_COA_WARRANTY_CLAIMS_RECEIVABLE As String

Public HYUNDAI_COA_INCIDENTAL_CHARGES_UNITS  As String
Public HYUNDAI_COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD As String

Public COA_INVENTORIES_PARTS                 As String
Public COA_INPUT_TAX                         As String
Public COA_INCOME_TAX_WITHHELD               As String
Public COA_ACCOUNTS_PAYABLE                  As String

Public OPEN_AR_SHOW                          As Boolean
Public SJ_SHOW                               As Boolean
Public PMIS_ORDER_SHOW                       As Boolean

Public Const SYSTEM_OWNER_NAME = "ABC COMPANY"
Public Const SYSTEM_OWNER_ADDRESS = "METRO MANILA"
Public Const SYSTEM_OWNER_CONTACT = "(054) 811-61-19"
Public Const SYSTEM_OWNER_TIN = "VAT TIN 000-000-000-000 VAT"
Public Const SYSTEM_POLE_TRANSITION = "CONGRATULATIONS! ..."
Public Const SYSTEM_POLE_GRACE = "Thank You Come Again"
Public Const MAXIMUM_DIGIT = "###,###,##0.00"
Public Const DIGIT_FORMAT = "###,###,##0"
Public Const ZERO = "0.00"
Public Const MOVEDOWN = "{TAB}^{HOME}+{END}"
Public Const MOVEUP = "+{TAB}^{HOME}+{END}"
Public Const ControlA = "^{A}"
Public Const ControlB = "^{B}"
Public Const ControlC = "^{C}"
Public Const ControlD = "^{D}"
Public Const ControlE = "^{E}"
Public Const ControlF = "^{F}"
Public Const ControlG = "^{G}"
Public Const ControlH = "^{H}"
Public Const ControlI = "^{I}"
Public Const ControlJ = "^{J}"
Public Const ControlK = "^{K}"
Public Const ControlL = "^{L}"
Public Const ControlM = "^{M}"
Public Const ControlN = "^{N}"
Public Const ControlO = "^{O}"
Public Const ControlP = "^{P}"
Public Const ControlQ = "^{Q}"
Public Const ControlR = "^R"
Public Const ControlS = "^S"
Public Const ControlT = "^T"
Public Const ControlU = "^U"
Public Const ControlV = "^V"
Public Const ControlW = "^W"
Public Const ControlX = "^X"
Public Const ControlY = "^Y"
Public Const ControlZ = "^Z"
Public Const POLE_LENGTH = 20
Public Const TOTAL_POLE_LENGTH = 40
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
'Public Const ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=E:\SQLDATA\AMIS_NAGA\DATA\AMISDat.DAT"
Public LOGPASS                               As String
Public ApplySecurityValidation               As Boolean
Public LOGID                                 As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000
Private m_bInIDE                             As Boolean
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName                             As String * CCDEVICENAME
    dmSpecVersion                            As Integer
    dmDriverVersion                          As Integer
    dmSize                                   As Integer
    dmDriverExtra                            As Integer
    dmFields                                 As Long
    dmOrientation                            As Integer
    dmPaperSize                              As Integer
    dmPaperLength                            As Integer
    dmPaperWidth                             As Integer
    dmScale                                  As Integer
    dmCopies                                 As Integer
    dmDefaultSource                          As Integer
    dmPrintQuality                           As Integer
    dmColor                                  As Integer
    dmDuplex                                 As Integer
    dmYResolution                            As Integer
    dmTTOption                               As Integer
    dmCollate                                As Integer
    dmFormName                               As String * CCFORMNAME
    dmUnusedPadding                          As Integer
    dmBitsPerPel                             As Integer
    dmPelsWidth                              As Long
    dmPelsHeight                             As Long
    dmDisplayFlags                           As Long
    dmDisplayFrequency                       As Long
End Type
Dim DevM                                     As DEVMODE

Global ResolutionWidth As Single
Global ResolutionHeight As Single
Global ScreenResolution As String
Global CurrentWidth As Single
Global CurrentHeight As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'Option Explicit
'Type Critter
'    FrIndex As Integer
'    Xp As Integer
'    Yp As Integer
'    Width As Integer
'    Height As Integer
'    Xmove As Integer
'    Ymove As Integer
'    Frames As Integer
'    Show As Boolean
'    ImageSrcX(3) As Integer
'    ImageSrcY(3) As Integer
'    ImageMaskX(3) As Integer
'    ImageMaskY(3) As Integer
'End Type
'Private Sprites(75) As Critter
'Private SpCnt, SpriteCount As Integer
Public RightX, BottomY, ASpeed               As Integer
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008
Private Ind, Xo, Yo, Xs, Ys, XSrc            As Long
Private YSrc, DDC, SDC, res                  As Long
Dim z2                                       As Long

Public Function OpenAccessDb() As Boolean
'    frmSplash.Show
'    frmSplash.labCon.caption = "Connecting to DMIS Database... Please wait..."
'    DoEvents
'    Dim ACCESS_Connection                    As String
'    AccessCNT = 0
'    On Error GoTo ConnErr
'    With wizVar
'        If .VerifyCryptoFile(App.Path & "\Access.crp") = True Then
'            ACCESS_Connection = .OpenCryptoFile("ACCESS", "CONNECT")
'        End If
'        On Error GoTo ConnErr
'        ACCESS_Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=\\SERVER\HMI\DATA\ACCESS\access.dat"
'        Set gconACCESS = New ADODB.Connection
'        gconACCESS.ConnectionString = ACCESS_Connection
'        gconACCESS.Properties(.DecryptAccess("4A6574~§³‘¨·{…d¨¥|–¸{y~\y‰f„dŽšŽ")) = .DecryptAccess("6B696Dm±²r°ª¤¤")
'        gconACCESS.Open
'        OpenAccessDb = True
'        Unload frmSplash
'        Exit Function
'    End With
'
'ConnErr:
'    MsgBox Err.Description
'    MsgBox "I can't open a connection!!! Cryptofile or Datafile for Access is missing or Invalid " & vbCrLf & _
'           "Contact your friendly neighborhood SysAdministrator.", _
'           vbOKOnly + vbCritical, "ERROR"
End Function

Public Sub FlattenCombo(Ctl As ComboBox, bCut As Boolean)
    Dim hRgn                                 As Long
    On Error Resume Next
    If bCut = True Then
        hRgn = CreateRectRgn(1, 1, ((Ctl.Width / Screen.TwipsPerPixelX) - 3), _
                             ((Ctl.Height / Screen.TwipsPerPixelY) - 3))
    Else
        hRgn = CreateRectRgn(0, 0, (Ctl.Width / Screen.TwipsPerPixelX), _
                             (Ctl.Height / Screen.TwipsPerPixelY))
    End If
    SetWindowRgn Ctl.hwnd, hRgn, True
End Sub

Public Function GetMachineName() As String
    Dim plngSize                             As Long
    Dim pstrBuffer                           As String
    pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    plngSize = Len(pstrBuffer)
    If GetComputerName(pstrBuffer, plngSize) Then
        GetMachineName = Left$(pstrBuffer, plngSize)
    End If
End Function

Public Sub MoveKeyPress(KeyCode As Integer)

    Dim First3Letters                        As String
    If Screen.ActiveForm.ActiveControl Is Nothing Then
        Exit Sub
    End If
    First3Letters = Mid(Screen.ActiveForm.ActiveControl.Name, 1, 3)
    '''''BUGLIST: CHECK NOTHING FOR FORM
    Select Case KeyCode
        Case 13
            If First3Letters = "cbo" Then
                If Screen.ActiveForm.ActiveControl.Text = "" Then Call VBComBoBoxDroppedDown(Screen.ActiveForm.ActiveControl) Else SendKeys MOVEDOWN
            Else
                If First3Letters = "txt" Or First3Letters = "opt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
            End If
        Case 40
            If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEDOWN
        Case 38
            If First3Letters = "txt" Or First3Letters = "chk" Then SendKeys MOVEUP
    End Select
End Sub

Public Sub ShowADOErrors(gcon As ADODB.Connection)
    Screen.MousePointer = 0
    Dim errLoop                              As ADODB.Error
    Dim strHelp                              As String
    For Each errLoop In gcon.Errors
        If errLoop.HelpFile = "" Then strHelp = " No Helpfile available" Else strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
        MsgBoxXP "ADO Error '" & errLoop.Number & vbCrLf & "Source: " & errLoop.source _
               & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
               & vbCrLf & vbCrLf & "Description: " & errLoop.Description & vbCrLf & vbCrLf & strHelp, "ADO Error", XP_OKOnly, msg_Critical
    Next
End Sub

Public Sub ShowVBError()
    Screen.MousePointer = 0
    If CBool(Err) Then
        'MsgBoxXP "VB Error '" & Err.Number & vbCrLf & "Source: " & Err.source & vbCrLf & vbCrLf & "Description: " & Err.Description, "VB Runtime Error", XP_OKOnly, msg_Critical
        'MsgBox "VB Error " & Err.Number & vbCrLf & "Source: " & Err.source & vbCrLf & vbCrLf & "Description: " & Err.Description, "VB Runtime Error", vbOK + vbCritical
        MsgBox "VB Error " & Err.Number & vbCrLf & "Source: " & Err.source & vbCrLf & vbCrLf & "Description: " & Err.Description, vbOK + vbCritical, "VB Runtime Error"
        'Err.Clear
    End If
End Sub

Public Sub ShowNoRecord()
    On Error Resume Next
    MessagePop RecNotFound, "Empty", "No Such Record", 1000
    'oVoice.Speak "No Such Record!", SVSFlagsAsync
    'MsgBoxXP "No Such Record!", "No Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowCantFind(str2find As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecNotFound, "Not Found", "Can't find" & str2find, 1000
    'oVoice.Speak "Can't find " & str2find, SVSFlagsAsync
    'MsgBoxXP "Can't find " & str2find, "Not Found", XP_OKOnly, msg_Information
End Sub

Public Function ShowConfirmDelete() As Boolean
    On Error Resume Next
    oVoice.Speak "Delete selected record? are you sure?...", SVSFlagsAsync
    If MsgBox("Delete selected record? Are you sure...", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        ShowConfirmDelete = True
    Else
        ShowConfirmDelete = False
    End If
End Function

Public Sub ShowDeletedMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Record Successfully Deleted...", SVSFlagsAsync
    'MsgBoxXP "Record Successfully Deleted...", "Confirmed", XP_OKOnly, msg_Information
    MsgBox "Record Successfully Deleted...", vbOK + vbInformation, "Confirmed"
End Sub

Public Sub ShowNothingToDeleteMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak "Nothing to Delete...", SVSFlagsAsync
    MsgBoxXP "Nothing to Delete...", "Empty Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowFirstRecordMsg()
    On Error Resume Next
    MessagePop NaviBegin, "Beginning of Record", "First Record"
    'oVoice.Speak "Beginning of Record...", SVSFlagsAsync
    'MsgBoxXP "Beginning of Record...", "First Record", XP_OKOnly, msg_Information
End Sub

Public Sub ShowLastRecordMsg()
    On Error Resume Next
    MessagePop NaviEnd, "End of Record", "Last Record"
    '   oVoice.Speak "End of Record...", SVSFlagsAsync
    '    MsgBoxXP "End of Record...", "Last Record", XP_OKOnly, msg_Information
End Sub

Public Sub MsgSpeechBox(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
    'oVoice.Speak Msg, SVSFlagsAsync
    'MsgBoxXP Msg, "Info", XP_OKOnly, msg_Information
    MsgBox Msg, vbInformation, "Info"
End Sub

Public Sub MsgSpeech(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
End Sub

Public Function MsgQuestionBox(Msg As String, BoxTitle As String) As Boolean
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
    'MsgQuestionBox = MsgBoxXP(Msg, BoxTitle, XP_YesNo, msg_Question)
    'MsgQuestionBox = MsgBox(Msg, vbQuestion + vbYesNo, BoxTitle)
    If MsgBox(Msg, vbQuestion + vbYesNo, BoxTitle) = vbYes Then
        MsgQuestionBox = True
    Else
        MsgQuestionBox = False
    End If
End Function

Public Function InputSpeechBox(ByRef Msg As String, Optional ByRef DefaultValue As String) As Variant
    Screen.MousePointer = 0
    On Error Resume Next
    oVoice.Speak Msg, SVSFlagsAsync
    'InputSpeechBox = InputBoxXP(Msg, "Find", DefaultValue)
    InputSpeechBox = InputBox(Msg, "Find", DefaultValue)
End Function

Public Function isTransparent(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg                                  As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
    If Err Then isTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg                                  As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If Err Then MakeTransparent = 2
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
    Dim Msg                                  As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If Err Then MakeOpaque = 2
End Function

Public Sub ChangeRes(ByVal iWidth As Single, ByVal iHeight As Single)
    Dim a                                    As Boolean
    Dim i&
    i = 0
    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)
    Dim b&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
End Sub

Public Sub GetRes()
    CurrentWidth = Screen.Width / 15
    CurrentHeight = Screen.Height / 15
    ResolutionWidth = Screen.Width / 15
    ResolutionHeight = Screen.Height / 15
    ScreenResolution = str(ResolutionWidth) + ", " + str(ResolutionHeight)
End Sub

Public Sub UnloadApp()
    SetErrorMode SEM_NOGPFAULTERRORBOX
    If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
        Call ChangeRes(CurrentWidth, CurrentHeight)
    End If
    End
End Sub

Public Sub UnloadForm(frm As Object)
    Dim ShowCount                            As Integer
    SetErrorMode SEM_NOGPFAULTERRORBOX
    'If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
    '   Call ChangeRes(CurrentWidth, CurrentHeight)
    'End If
    For ShowCount = 200 To 1 Step -50
        MakeTransparent frm.hwnd, ShowCount
        DoEvents
    Next
    Unload frm
    Set frm = Nothing
End Sub

Public Property Get InIDE() As Boolean
Debug.Assert (IsInIDE())
    InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
    m_bInIDE = True
    IsInIDE = m_bInIDE
End Function

Public Function AutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
    Dim strFindThis As String, bContinueSearch As Boolean
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle

    If KeyAscii < 32 Then
        bContinueSearch = False
        cbBox.SelLength = 0
        If KeyAscii = Asc(vbBack) Then
            If lLength = 0 Then
                If Len(cbBox) > 0 Then
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                End If
            Else
                cbBox.Text = Left(cbBox.Text, lStart)
            End If
            cbBox.SelStart = Len(cbBox)
        ElseIf KeyAscii = vbKeyReturn Then
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii
        End If
    Else
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii)
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If

    If bContinueSearch Then
        Call VBComBoBoxDroppedDown(cbBox)
        lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        If lResult = CB_ERR Then
            cbBox.Text = strFindThis
            cbBox.SelLength = 0
            cbBox.SelStart = Len(cbBox)
        Else
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
        End If
    End If
    On Error GoTo 0
    Exit Function

ErrHandle:
Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

Public Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub

Public Function rsFindDuplicate(rs2Find As ADODB.Recordset, ByVal rsField2find, ByVal str2find) As Boolean
    Screen.MousePointer = 0
    On Error GoTo BFoundErr
    Dim rsToFind                             As ADODB.Recordset
    If Len(str2find) > 1 And Len(rsField2find) > 1 Then
        Set rsToFind = New ADODB.Recordset
        Set rsToFind = rs2Find.Clone
        rsToFind.Find rsField2find & " = '" & str2find & "'"
        If Not rsToFind.EOF Then rsFindDuplicate = True Else rsFindDuplicate = False
    End If
    Exit Function
BFoundErr:
    MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
    rsFindDuplicate = False
End Function

Public Sub ShowAlreadyExistMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecSaveError, "Duplicate Record", Ricord & " Already Exist!..."
    'oVoice.Speak Ricord & " Already Exist!...", SVSFlagsAsync
    'MsgBoxXP Ricord & " Already Exist!...", "Duplicate Record Found", XP_OKOnly, msg_Exclamation
End Sub

Public Sub ShowIsRequiredMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    'oVoice.Speak Ricord & " is Required!...", SVSFlagsAsync
    'MsgBoxXP Ricord & " is Required!...", "Field must have a Value", XP_OKOnly, msg_Exclamation
    MessagePop RecSaveError, "Missing Filelds", "Field must have a Value!..."
End Sub

Public Sub ShowSuccessFullyAdded()
    Screen.MousePointer = 0
    On Error Resume Next
    'oVoice.Speak "Data Successfully Added!...", SVSFlagsAsync
    'MsgBoxXP "Data Successfully Added!...", "wizweirdo's Message", XP_OKOnly, msg_Information
    MessagePop RecSaveOk, "Record Added", "Data Successfully Added!..."
End Sub

Public Sub ShowSuccessFullyUpdated()
    Screen.MousePointer = 0

    On Error Resume Next
    MessagePop RecSaveOk, "Record Updated", "Data Successfully Updated!..."
    'oVoice.Speak "Data Successfully Updated!...", SVSFlagsAsync
    'MsgBoxXP "Data Successfully Updated!...", "wizweirdo's Message", XP_OKOnly, msg_Information
End Sub

Public Function UpperAscii(Askey As Integer)
    'BGLIST: Invalid Procedures or call
    UpperAscii = Asc(UCase(Chr(Askey)))
End Function

Public Function OnlyNumeric(KeyCode As Integer) As Integer
    If KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 27 And KeyCode <> 46 Then
        If (KeyCode < 48 Or KeyCode > 57) And KeyCode <> 110 Then
            OnlyNumeric = 0
        Else
            OnlyNumeric = KeyCode
        End If
    Else
        OnlyNumeric = KeyCode
    End If
End Function

Public Function ToDoubleNumber(ByRef NumericText As Variant) As String
    Dim Counter                              As Integer
    Dim TempNumber                           As String
    Dim FoundPeriod                          As Boolean
    FoundPeriod = False: TempNumber = ""
    For Counter = 1 To Len(NumericText)
        If Mid(NumericText, Counter, 1) = "." Then
            If FoundPeriod = False Then
                TempNumber = TempNumber & Mid(NumericText, Counter, 1)
                FoundPeriod = True
            End If
        Else
            TempNumber = TempNumber & Mid(NumericText, Counter, 1)
        End If
    Next
    ToDoubleNumber = Format(TempNumber, MAXIMUM_DIGIT)
End Function

Public Function NumericVal(NumericText As Variant) As Double
    Dim Counter                              As Integer
    Dim NumericValue                         As String
    NumericValue = ""
    If Trim(NumericText) <> "" Then
        If IsNumeric(NumericText) = True Then
            If Val(Abs(NumericText)) > 0 Then
                For Counter = 1 To Len(NumericText)
                    If Mid(NumericText, Counter, 1) <> "," Then
                        NumericValue = NumericValue & Mid(NumericText, Counter, 1)
                    End If
                Next
                NumericVal = NumericValue
            Else
                NumericVal = 0
            End If
        Else
            NumericVal = 0
        End If
    Else
        NumericVal = 0
    End If
End Function

Public Sub Listview_Loadval(TisoyView As ListItems, RecSet As ADODB.Recordset)
    Dim Indx                                 As Long
    Dim i                                    As Long
    TisoyView.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            Indx = TisoyView.Count + 1
            TisoyView.Add Indx, , IIf(IsNull(RecSet(0)), "", Trim(RecSet(0)))
            For i = 1 To RecSet.Fields.Count - 1
                TisoyView(Indx).ListSubItems.Add i, , IIf(IsNull(RecSet(i)), "", Trim(RecSet(i)))
            Next i
            RecSet.MoveNext
        Wend
    End If
    Set RecSet = Nothing
End Sub

Public Sub Combo_Loadval(WeirdoCombo As ComboBox, RecSet As ADODB.Recordset)
    WeirdoCombo.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            WeirdoCombo.AddItem Null2String(RecSet(0))
            RecSet.MoveNext
        Wend
    End If
    Set RecSet = Nothing
End Sub

Public Function ValidPassword(login As String, Passwd As String, Module As String, Optional AdminPass As Boolean = False) As Boolean
    '===============================
    'Used to validate password.
    'first, it checks for the existence of the username,
    'password, module requested
    '===============================
    If ApplySecurityValidation = False Then
        ValidPassword = True
        Exit Function
    End If

    Dim Conn                                 As ADODB.Connection
    Dim rs                                   As ADODB.Recordset
    Dim PasswdStr                            As String
    Set Conn = New ADODB.Connection
    Conn.Open ConnStr
    Set rs = Conn.Execute("Select * from AMIS_vw_USERACCESS where username = '" & login & "' and description = '" & Module & "'")

    If Not (rs.BOF And rs.EOF) Then
        If rs!lock = True Then
            MsgBox "This user account has expired, Please contact your Administrator", vbCritical, "Access Denied!"
            ValidPassword = False
            Exit Function
        End If

        If Passwd = wizVar.DecryptAccess(Trim(rs!Password)) Then
            If AdminPass And rs!usergroup = "ADM" Then        'OK - Admin
                ValidPassword = True
                LOGNAME = Trim(rs!Username)
                LOGID = rs!UserID

                Exit Function
            ElseIf AdminPass And rs!usergroup <> "ADM" Then   'OK - but not admin
                MsgBox "Please contact your Sys Administrator.", vbCritical, "Access denied!"
                ValidPassword = False
                Exit Function
            ElseIf Not AdminPass Then                         'OK - admin not required
                ValidPassword = True
                LOGNAME = Trim(rs!Username)
                LOGID = rs!UserID
                Exit Function
            End If
        Else
            MsgBox "Please contact your System Administrator.", vbCritical, "Access denied!"
            ValidPassword = False
            Exit Function
        End If
    Else
        MsgBox "User " & login & " not found!", vbCritical, "Unknown User"
        ValidPassword = False
        Exit Function
    End If
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
End Function



Public Function populateCbo(QueryStr As String, ByVal cboConn As ADODB.Connection, ByRef myCbo, Optional Add_NA As Boolean = False) As Boolean
    '-- function that loads recordset into a referenced combobox
    '-- returns false if error occured
    '-- add "N/A" in combo if add_NA is true
    Dim fieldCount                           As Integer
    Dim rowCtr                               As Long
    Dim colCtr                               As Integer
    Dim cboRS                                As ADODB.Recordset

    On Error GoTo loadCboErr


    Set cboRS = cboConn.Execute(QueryStr)
    rowCtr = 0
    myCbo.Clear
    If Add_NA Then
        myCbo.AddItem "N/A"

        myCbo.List(rowCtr, 1) = "N/A"

        rowCtr = 1
    End If
    If Not cboRS.State = 0 Then
        If Not (cboRS.EOF And cboRS.BOF) Then
            Do While Not cboRS.EOF
                If cboRS.Fields(0) <> "N/A" Then
                    myCbo.AddItem Trim(cboRS.Fields(0)) & ""
                    For colCtr = 1 To cboRS.Fields.Count - 1
                        myCbo.List(rowCtr, colCtr) = Trim(cboRS.Fields(colCtr)) & ""
                    Next colCtr
                    rowCtr = rowCtr + 1
                End If
                cboRS.MoveNext
            Loop
            myCbo.ListIndex = -1
            populateCbo = True
        End If
    End If
    Set cboRS = Nothing
    Set cboConn = Nothing
    Exit Function
loadCboErr:
    Set cboRS = Nothing
    Set cboConn = Nothing
    MsgBox Err.Description
    populateCbo = False
End Function

Public Function ConvertToBIRDecimalFormat(XXX As Variant) As Double
    ConvertToBIRDecimalFormat = 1# + (XXX / 100)
End Function

Public Function VatPercentRate(XXX As Variant) As Double
    VatPercentRate = (XXX / 100)
End Function


Sub CScrKua(canvas As Object)
    Dim screendc&
    canvas.AutoRedraw = True
    screendc = CreateDC("DISPLAY", "", "", 0&)
    StretchBlt canvas.hdc, 0, 0, canvas.Width, canvas.Height, screendc, 0, 0, Screen.Width, Screen.Height, SRCCOPY
    DeleteDC screendc
    canvas.AutoRedraw = False
End Sub

Sub CaptureScreen(PityurBox As PictureBox)
    Dim DestDC, XPixels, YPixels, destX      As Long
    Dim destY, srcDC, SrcX, SrcY, RasterOp   As Long
    BottomY = PityurBox.ScaleHeight
    RightX = PityurBox.ScaleWidth
    CScrKua PityurBox
    PityurBox.Refresh
    DoEvents
    destX = 0: destY = 0
    XPixels = PityurBox.ScaleWidth
    YPixels = PityurBox.ScaleHeight
    srcDC = PityurBox.hdc
    SrcX = 0: SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, destX, destY, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    XPixels = PityurBox.ScaleWidth
    YPixels = PityurBox.ScaleHeight
    srcDC = PityurBox.hdc
    SrcX = 0: SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, 0, 0, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    DestDC = 0: XPixels = 0: YPixels = 0:
    srcDC = 0: SrcX = 0: SrcY = 0: RasterOp = 0
    Dim FileNaeym                            As String
    If LOGCODE <> "" And LOGDATE <> "" Then
        FileNaeym = "C:\" & App.EXEName & "_" & LOGCODE & "_" & Trim(str(Month(LOGDATE))) & Trim(str(Day(LOGDATE))) & Trim(str(Year(LOGDATE))) & "_" & Left(str(Time), 2) & Mid(str(Time), 4, 2) & Mid(str(Time), 7, 2) & Right(str(Time), 2) & ".jpg"
        'SavePicture PityurBox.Image, FileNaeym
        'Dim Conn As ADODB.Connection
        'Set Conn = New ADODB.Connection
        'Set Conn = DataEnvironment1.Connection1
        'Conn.Open

        'Conn.Execute ("Insert into LOGS (PictureFile) values (" & PityurBox.Image & ")")
    End If
    Set PityurBox = Nothing
End Sub

Function Date2Month(Value As String)
    Dim MO                                   As String
    MO = "January  February March    April    May      June     July     August   SeptemberOctober  November December "
    Date2Month = Mid$(MO, (Month(Value) - 1) * 9 + 1, 9)
End Function


