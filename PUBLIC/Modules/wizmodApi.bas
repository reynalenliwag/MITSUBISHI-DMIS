Attribute VB_Name = "wizmodAPI"

Option Explicit

Public CANCEL_ANS                                                   As String
Public REPRINT_CAPTION                                              As String
Public Const MAX_COMPUTERNAME_LENGTH                                As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oVoice                                                       As SpeechLib.SpVoice

Public sarahbaby As String

'FOR HELP SYSTEM
Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
'END FOR HELP SYSTEM

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public gconDMIS                                                     As ADODB.Connection
Public gconACCESS                                                   As ADODB.Connection

Public ConnStr
Public DMIS_Connection
Public DMIS_Audit_Connection
Public DMIS_REPORT_Connection

Public DEALER_CODE                                                  As String
Public COMPANY_CODE                                                 As String
Public COMPANY_VERSION                                              As String
Public COMPANY_NAME                                                 As String
Public COMPANY_ADDRESS                                              As String
Public COMPANY_TIN                                                  As String
Public TransactionID                                                As String
Public PMIS_REPORT_PATH                                             As String
Public CSMS_REPORT_PATH                                             As String
Public SMIS_REPORT_PATH                                             As String
Public CMIS_REPORT_PATH                                             As String
Public HRMS_REPORT_PATH                                             As String
Public AMIS_REPORT_PATH                                             As String
Public CRIS_REPORT_PATH                                             As String
Public OSMS_REPORT_PATH                                             As String
Public HRMS_PICTURES_PATH                                           As String

Public PMIS_REPORT_CONNECTION                                       As String
Public CSMS_REPORT_CONNECTION                                       As String
Public SMIS_REPORT_CONNECTION                                       As String
Public CMIS_REPORT_CONNECTION                                       As String
Public HRMS_REPORT_Connection                                       As String
Public AMIS_REPORT_CONNECTION                                       As String
Public CRIS_REPORT_CONNECTION                                       As String
Public OSMS_REPORT_CONNECTION                                       As String

Public SERVERNAME                                                   As String
Public SQLSERVERNAME                                                As String
Public DATABASE                                                     As String
''''FOR USER ACESS MANAGEMENT

Public LOGPASS                                                      As String
Attribute LOGPASS.VB_VarUserMemId = 1073741856
Public LOGID                                                        As Long
Attribute LOGID.VB_VarUserMemId = 1073741857
'''REMINDERS WINDOW
Public TIMER_REMIND                                                 As String
''''END USER ACESS MANAGEMENT

Public COA_AR_TRADE_UNITS                                           As String
Public COA_AR_TRADE_SERVICE                                         As String
Public COA_AR_TRADE_PARTS                                           As String

Public HYUNDAI_COA_AR_TRADE_UNITS                                   As String
Public HYUNDAI_COA_AR_TRADE_SERVICE                                 As String
Public HYUNDAI_COA_AR_TRADE_PARTS                                   As String

Public COA_OUTPUT_TAX                                               As String

'SALES - CASH
Public COA_SALES_SERVICE_CASH_TINSPAINT                             As String
Public COA_SALES_SERVICE_CASH_SUBLET                                As String
Public COA_SALES_SERVICE_CASH_AIRCON                                As String
Public COA_SALES_SERVICE_CASH_LABOR                                 As String
Public COA_SALES_SERVICE_CASH_GOL                                   As String
Public COA_SALES_SERVICE_CASH_PARTS                                 As String

Public COA_SALES_GOL_CASH                                           As String
Public COA_SALES_PARTS_CASH                                         As String
Public COA_SALES_VEHICLES_CASH                                      As String

'SALES - CHARGE
Public COA_SALES_SERVICE_CHARGE_TINSPAINT                           As String
Public COA_SALES_SERVICE_CHARGE_SUBLET                              As String
Public COA_SALES_SERVICE_CHARGE_AIRCON                              As String
Public COA_SALES_SERVICE_CHARGE_LABOR                               As String
Public COA_SALES_SERVICE_CHARGE_GOL                                 As String
Public COA_SALES_SERVICE_CHARGE_PARTS                               As String

Public COA_SALES_GOL_CHARGE                                         As String
Public COA_SALES_PARTS_CHARGE                                       As String

'SALES - DISCOUNT - CASH
Public COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT                    As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET                       As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON                       As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_LABOR                        As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_GOL                          As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_PARTS                        As String

Public COA_SALES_DISCOUNT_GOL_CASH                                  As String
Public COA_SALES_DISCOUNT_PARTS_CASH                                As String
Public COA_SALES_DISCOUNT_VEHICLES_CASH                             As String

'SALES - DISCOUNT - CHARGE
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT                  As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET                     As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON                     As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR                      As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL                        As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS                      As String

Public COA_SALES_DISCOUNT_GOL_CHARGE                                As String
Public COA_SALES_DISCOUNT_PARTS_CHARGE                              As String

'CHARGE TO WARRANTY
Public COA_DIRECT_EXPENSE_LABOR                                     As String
Public COA_DIRECT_EXPENSE_SPAREPARTS                                As String
Public COA_DIRECT_EXPENSE_GOL                                       As String

Public COA_WARRANTY_SALES                                           As String
Public COA_WARRANTY_SERVICE                                         As String
Public COA_WARRANTY_PARTS                                           As String

'CHARGE TO COMPANY
Public COA_COMPANY_CAR_SALES                                        As String
Public COA_COMPANY_CAR_SERVICE                                      As String

'CHARGE TO SALES
Public COA_GFSI_SALES                                               As String
Public COA_GFSI_SERVICE                                             As String
Public COA_GFSI_PARTS                                               As String

'CASH RECEIPTS
Public COA_CASH_ON_HAND                                             As String
Public COA_BRANCH_LEGASPI                                           As String
Public COA_CUSTOMER_DEPOSIT                                         As String

Public COA_INSURANCE_PREMIUM_PAYABLE                                As String
Public COA_INSURANCE_PREMIUM_RENEWAL                                As String
Public COA_LTO_PAYMENT                                              As String
Public COA_CHATTEL_MORTGAGE_FEE_PAYABLE                             As String
Public COA_NEW_VEHICLE_REGISTRATION                                 As String
Public COA_WARRANTY_CLAIMS_RECEIVABLE                               As String

Public COA_ACCOUNTS_RECEIVABLE_NONTRADE_EMPLOYEES                   As String
Public COA_OTHER_PAYABLES                                           As String
Public COA_INCIDENTAL_CHARGES_UNITS                                 As String
Public COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD                          As String
Public COA_PRE_DELIVERY                                             As String

Public COA_CORPORATE_TAX_WHELD                                      As String
Public COA_CORPORATE_VAT_WHELD                                      As String

Public HYUNDAI_COA_SALES_SERVICE_CASH_TINSPAINT                     As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_SUBLET                        As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_AIRCON                        As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_LABOR                         As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_GOL                           As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_PARTS                         As String

Public HYUNDAI_COA_SALES_GOL_CASH                                   As String
Public HYUNDAI_COA_SALES_PARTS_CASH                                 As String

'SALES - CHARGE
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_TINSPAINT                   As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_SUBLET                      As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_AIRCON                      As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_LABOR                       As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_GOL                         As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_PARTS                       As String

Public HYUNDAI_COA_SALES_GOL_CHARGE                                 As String
Public HYUNDAI_COA_SALES_PARTS_CHARGE                               As String

'SALES - DISCOUNT - CASH
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT            As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET               As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON               As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_LABOR                As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_GOL                  As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_PARTS                As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CASH                          As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CASH                        As String

'SALES - DISCOUNT - CHARGE
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT          As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET             As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON             As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR              As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL                As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS              As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CHARGE                        As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CHARGE                      As String

Public HYUNDAI_COA_WARRANTY_SALES                                   As String
Public HYUNDAI_COA_WARRANTY_SERVICE                                 As String
Public HYUNDAI_COA_WARRANTY_PARTS                                   As String

Public HYUNDAI_COA_COMPANY_CAR_SALES                                As String
Public HYUNDAI_COA_COMPANY_CAR_SERVICE                              As String

Public HYUNDAI_COA_GFSI_SALES                                       As String
Public HYUNDAI_COA_GFSI_SERVICE                                     As String
Public HYUNDAI_COA_GFSI_PARTS                                       As String

Public HYUNDAI_COA_NEW_VEHICLE_REGISTRATION                         As String
Public HYUNDAI_COA_WARRANTY_CLAIMS_RECEIVABLE                       As String

Public HYUNDAI_COA_INCIDENTAL_CHARGES_UNITS                         As String
Public HYUNDAI_COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD                  As String

Public COA_INVENTORIES_PARTS                                        As String
Public COA_INVENTORIES_GOL                                          As String
Public COA_INVENTORIES_VEHICLES                                     As String

Public COA_COST_OF_SALES_PARTS                                      As String
Public COA_COST_OF_SALES_GOL                                        As String
Public COA_COST_OF_SALES_VEHICLES                                   As String


Public COA_INPUT_TAX                                                As String
Public COA_INCOME_TAX_WITHHELD                                      As String
Public COA_ACCOUNTS_PAYABLE                                         As String

Public OPEN_AR_SHOW                                                 As Boolean
Public SJ_SHOW                                                      As Boolean
Public PMIS_ORDER_SHOW                                              As Boolean

Public Const SYSTEM_OWNER_CODE = "01"
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
Public AMIS_Invoicetype                                             As String
Public AMIS_Invoiceno                                               As String
'Public Const ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=E:\SQLDATA\AMIS_NAGA\DATA\AMISDat.DAT"

Public ApplySecurityValidation                                      As Boolean

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000
Private m_bInIDE                                                    As Boolean
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName                                                    As String * CCDEVICENAME
    dmSpecVersion                                                   As Integer
    dmDriverVersion                                                 As Integer
    dmSize                                                          As Integer
    dmDriverExtra                                                   As Integer
    dmFields                                                        As Long
    dmOrientation                                                   As Integer
    dmPaperSize                                                     As Integer
    dmPaperLength                                                   As Integer
    dmPaperWidth                                                    As Integer
    dmScale                                                         As Integer
    dmCopies                                                        As Integer
    dmDefaultSource                                                 As Integer
    dmPrintQuality                                                  As Integer
    dmColor                                                         As Integer
    dmDuplex                                                        As Integer
    dmYResolution                                                   As Integer
    dmTTOption                                                      As Integer
    dmCollate                                                       As Integer
    dmFormName                                                      As String * CCFORMNAME
    dmUnusedPadding                                                 As Integer
    dmBitsPerPel                                                    As Integer
    dmPelsWidth                                                     As Long
    dmPelsHeight                                                    As Long
    dmDisplayFlags                                                  As Long
    dmDisplayFrequency                                              As Long
End Type
Dim DevM                                                            As DEVMODE

Global ResolutionWidth As Single
Attribute ResolutionWidth.VB_VarUserMemId = 1073741986
Global ResolutionHeight As Single
Attribute ResolutionHeight.VB_VarUserMemId = 1073741987
Global ScreenResolution As String
Attribute ScreenResolution.VB_VarUserMemId = 1073741988
Global CurrentWidth As Single
Attribute CurrentWidth.VB_VarUserMemId = 1073741989
Global CurrentHeight As Single
Attribute CurrentHeight.VB_VarUserMemId = 1073741990

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
Public RightX, BottomY, ASpeed                                      As Integer
Attribute RightX.VB_VarUserMemId = 1073741991
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)

Private Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008
Private Ind, Xo, Yo, Xs, Ys, xSrc                                   As Long
Attribute Ind.VB_VarUserMemId = 1073741992
Private ySrc, DDC, SDC, res                                         As Long
Attribute ySrc.VB_VarUserMemId = 1073741963
Dim z2                                                              As Long
Attribute z2.VB_VarUserMemId = 1073741967

Public Const TTLDYSIN1YR = 365

Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_RIGHT = &H27
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Dim retrieveFilePath                                                As String
Dim sPic                                                            As IPictureDisp

Private Const FLAG_ICC_FORCE_CONNECTION = &H1
Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Public Function OpenAccessDb() As Boolean
    frmSplash.Show
    frmSplash.labCon.Caption = "Connecting to DMIS User Database... Please wait..."
    DoEvents
    AccessCNT = 0
    On Error GoTo ConnErr
    Set gconACCESS = New ADODB.Connection
    gconACCESS.ConnectionString = DMIS_Connection
    gconACCESS.Open
    OpenAccessDb = True
    Unload frmSplash
    Exit Function
ConnErr:
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "register again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "Database Connection Failed!"
    End
End Function

Public Sub FlattenCombo(CTL As ComboBox, bCut As Boolean)
    Dim hRgn                                                As Long
    On Error Resume Next
    If bCut = True Then
        hRgn = CreateRectRgn(1, 1, ((CTL.Width / Screen.TwipsPerPixelX) - 3), _
                             ((CTL.Height / Screen.TwipsPerPixelY) - 3))
    Else
        hRgn = CreateRectRgn(0, 0, (CTL.Width / Screen.TwipsPerPixelX), _
                             (CTL.Height / Screen.TwipsPerPixelY))
    End If
    SetWindowRgn CTL.hwnd, hRgn, True
End Sub

Public Function GetMachineName() As String
    Dim plngSize                                                    As Long
    Dim pstrBuffer                                                  As String
    pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    plngSize = Len(pstrBuffer)
    If GetComputerName(pstrBuffer, plngSize) Then
        GetMachineName = Left$(pstrBuffer, plngSize)
    End If
End Function

Public Sub MoveKeyPress(KeyCode As Integer)
    Dim First3Letters                                               As String
    On Error Resume Next
    
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
    Dim errLoop                                             As ADODB.error
    Dim strHelp                                             As String
    For Each errLoop In gcon.Errors
        If errLoop.HelpFile = "" Then strHelp = " No Helpfile available" Else strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
        MsgBoxXP "ADO Error '" & errLoop.Number & vbCrLf & "Source: " & errLoop.Source _
                 & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
                 & vbCrLf & vbCrLf & "Description: " & errLoop.DESCRIPTION & vbCrLf & vbCrLf & strHelp, "ADO Error", XP_OKOnly, msg_Critical
    Next
End Sub

Public Sub ShowVBError()
    On Error Resume Next
    Screen.MousePointer = 0
    If CBool(err) Then
        MessagePop RecLocekd, "System Info", "Cannot Process Your Request... " & vbCrLf & _
                                             "Please Try Again Or Rather Contact System Administrator..." & vbCrLf & _
                                             vbCrLf & _
                                             "  Module Name : " & MODULENAME & vbCrLf & _
                                             "  Ref No:" & err.Number & vbCrLf & _
                                             "  Description: " & err.DESCRIPTION, 4500, 1, 180
        Dim FORMNAME, EventName
        FORMNAME = Screen.ActiveForm.Name

        If IsEmpty(Screen.ActiveControl.Name) = False Then
            EventName = Screen.ActiveControl.Name
        End If

        gconAudit.Execute ("Insert into DMIS_XOT (DTVDP1, WLDIOP,MWTCTT,WTCTTT, WXOPOP,XOTNON ,TOCXXX) VALUES ('" & _
                           Now & "' , '" & _
                           LOGNAME & "' , '" & _
                           App.TITLE & "' , '" & _
                           FORMNAME & "' , '" & _
                           EventName & "' , " & _
                           N2Str2Null(err.Number) & " , " & _
                           N2Str2Null(err.DESCRIPTION) & ")")
        Screen.MousePointer = 0
        err.Clear
    End If

End Sub

Function EncrypStr(XXX, YYY As Boolean)
    Dim EncStr                                              As String    '
    Dim nard
    Dim MARK

    EncStr = ""
    If YYY = True Then
        For nard = 1 To Len(XXX)
            MARK = Chr(Asc(Mid(XXX, nard, 1)) + 5)
            EncStr = EncStr & MARK
        Next
    Else
        For nard = 1 To Len(XXX)
            MARK = Chr(Asc(Mid(XXX, nard, 1)) - 5)
            EncStr = EncStr & MARK
        Next
    End If
    EncrypStr = EncStr
End Function


Public Sub ShowNoRecord()
    On Error Resume Next
    MessagePop RecNotFound, "Empty", "No Such Record", 1000
End Sub

Public Sub ShowCantFind(str2find As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecNotFound, "Not Found", "Can't find" & str2find, 1000
End Sub

Public Function ShowConfirmDelete() As Boolean
    On Error Resume Next
    oVoice.Speak "Delete selected record, Are you Sure?", SVSFlagsAsync
    If MsgBox("Delete selected record, Are you Sure?", vbYesNo + vbQuestion, "Confirm Delete") = vbYes Then
        ShowConfirmDelete = True
    Else
        ShowConfirmDelete = False
    End If
End Function
Public Sub ShowVoidedMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop Delete, "Confirmed", "Record Successfully Voided..."
End Sub

Public Sub ShowDeletedMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop Delete, "Confirmed", "Record Successfully Deleted..."
End Sub

Public Sub ShowNothingToDeleteMsg()
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecNotFound, "Empty Record", "Nothing to Delete..."
End Sub

Public Sub ShowFirstRecordMsg()
    On Error Resume Next
    MessagePop NaviBegin, "Beginning of Record", "First Record"
End Sub

Public Sub ShowLastRecordMsg()
    On Error Resume Next
    MessagePop NaviEnd, "End of Record", "Last Record", 1500
End Sub

Public Sub MsgSpeechBox(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
    MsgBox Msg, vbInformation, "Info"
End Sub

Public Sub MsgSpeech(Msg As String)
    Screen.MousePointer = 0
    On Error Resume Next
End Sub

Public Function MsgQuestionBox(Msg As String, BoxTitle As String) As Boolean
    Screen.MousePointer = 0
    On Error Resume Next

    If MsgBox(Msg, vbQuestion + vbYesNo, BoxTitle) = vbYes Then
        MsgQuestionBox = True
    Else
        MsgQuestionBox = False
    End If
End Function

Public Function InputSpeechBox(ByRef Msg As String, Optional ByRef DefaultValue As String) As Variant
    Screen.MousePointer = 0
    On Error Resume Next
    InputSpeechBox = InputBox(Msg, "Find", DefaultValue)
End Function

Public Function isTransparent(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg                                                 As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
    If err Then isTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg                                                 As Long
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
    If err Then MakeTransparent = 2
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
    Dim Msg                                                 As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If err Then MakeOpaque = 2
End Function

Public Sub ChangeRes(ByVal iWidth As Single, ByVal iHeight As Single)
    Dim A                                                   As Boolean
    Dim i&
    i = 0
    Do
        A = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (A = False)
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
    ScreenResolution = STR(ResolutionWidth) + ", " + STR(ResolutionHeight)
End Sub

Public Sub UnloadApp()
    SetErrorMode SEM_NOGPFAULTERRORBOX
End Sub

Public Sub UnloadForm(frm As Object)
    Dim ShowCount                                           As Integer
    SetErrorMode SEM_NOGPFAULTERRORBOX
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
    Dim strFindThis As String, bContinueSearch              As Boolean
    Dim lResult As Long, lStart As Long, lLength            As Long
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
    Debug.Print "Failed: AutoCompleteComboBox due to : " & err.DESCRIPTION
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
    Dim rsToFind                                            As ADODB.Recordset
    If Len(str2find) > 1 And Len(rsField2find) > 1 Then
        Set rsToFind = New ADODB.Recordset
        Set rsToFind = rs2Find.Clone
        rsToFind.Find rsField2find & " = '" & str2find & "'"
        If Not rsToFind.EOF Then rsFindDuplicate = True Else rsFindDuplicate = False
    End If
    Exit Function
BFoundErr:
    MsgBox "Error:" & err & " " & error, vbOKOnly, "Error"
    rsFindDuplicate = False
End Function

Public Sub ShowAlreadyExistMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecSaveError, "Duplicate Record", Ricord & " Already Exist!..."
End Sub

Public Sub ShowIsRequiredMsg(Ricord As Variant)
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecSaveError, "Missing Filelds", "Field must have a Value!..." & Ricord, 1500
End Sub

Public Sub ShowSuccessFullyAdded()
    Screen.MousePointer = 0
    On Error Resume Next
    MessagePop RecSaveOk, "Record Added", "Data Successfully Added!..."
End Sub

Public Sub ShowSuccessFullyUpdated()
    Screen.MousePointer = 0

    On Error Resume Next
    MessagePop RecSaveOk, "Record Updated", "Data Successfully Updated!..."
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
    Dim COUNTER                                             As Integer
    Dim TempNumber                                          As String
    Dim FoundPeriod                                         As Boolean
    
    FoundPeriod = False: TempNumber = ""
    If Val(NumericText) < 0 Then
        ToDoubleNumber = NumericText
    Else
        For COUNTER = 1 To Len(NumericText)
            If Mid(NumericText, COUNTER, 1) = "." Then
                If FoundPeriod = False Then
                    TempNumber = TempNumber & Mid(NumericText, COUNTER, 1)
                    FoundPeriod = True
                End If
            Else
                TempNumber = TempNumber & Mid(NumericText, COUNTER, 1)
            End If
        Next
        ToDoubleNumber = Format(TempNumber, MAXIMUM_DIGIT)
    End If
End Function

Public Function NumericVal(NumericText As Variant) As Double
    Dim COUNTER                                             As Integer
    Dim NumericValue                                        As String
    
    NumericValue = ""
    
    If Trim(NumericText) <> "" Then
        
        If IsNumeric(NumericText) = True Then
            If Val(NumericText) >= 0 Then
                For COUNTER = 1 To Len(NumericText)
                    If Mid(NumericText, COUNTER, 1) <> "," Then
                         NumericValue = NumericValue & Mid(NumericText, COUNTER, 1)
                    
                    End If
                Next
                NumericVal = NumericValue
            Else
                NumericVal = Val(NumericText)
            End If
        Else
            NumericVal = 0
        End If
    Else
        NumericVal = 0
    End If
End Function

Public Sub Listview_Loadval(TisoyView As ListItems, RecSet As ADODB.Recordset)
    Dim Indx                                                As Long
    Dim i                                                   As Long
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
    If ApplySecurityValidation = False Then
        ValidPassword = True
        Exit Function
    End If

    Dim Conn                                                As ADODB.Connection
    Dim RS                                                  As ADODB.Recordset
    Dim PasswdStr                                           As String
    Set Conn = New ADODB.Connection
    Conn.Open ConnStr
    Set RS = Conn.Execute("Select * from ALL_RAMS_USERS where username = '" & login & "'")

    If Not (RS.BOF And RS.EOF) Then
        If RS!lock = True Then
            MsgBox "This user account has expired, Please contact your Administrator", vbCritical, "Access Denied!"
            ValidPassword = False
            Exit Function
        End If

        If Passwd = wizVar.DecryptAccess(Trim(RS!Password)) Then
            If AdminPass And RS!USERGROUP = "ADM" Then    'OK - Admin
                ValidPassword = True
                LOGNAME = Trim(RS!UserName)
                LOGID = RS!USERID
                Exit Function
            ElseIf AdminPass And RS!USERGROUP <> "ADM" Then    'OK - but not admin
                MsgBox "Please contact your System Administrator.", vbCritical, "Access denied!"
                ValidPassword = False
                Exit Function
            ElseIf Not AdminPass Then                 'OK - admin not required
                ValidPassword = True
                LOGNAME = Trim(RS!UserName)
                LOGID = RS!USERID

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
    Set RS = Nothing
    Conn.Close
    Set Conn = Nothing
End Function

Public Function populateCbo(QueryStr As String, ByVal cboConn As ADODB.Connection, ByRef myCbo, Optional Add_NA As Boolean = False) As Boolean
'-- function that loads recordset into a referenced combobox
'-- returns false if error occured
'-- add "N/A" in combo if add_NA is true
    Dim fieldCount                                          As Integer
    Dim rowCtr                                              As Long
    Dim colCtr                                              As Integer
    Dim cboRS                                               As ADODB.Recordset

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
    MsgBox err.DESCRIPTION
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
    StretchBlt canvas.hDC, 0, 0, canvas.Width, canvas.Height, screendc, 0, 0, Screen.Width, Screen.Height, SRCCOPY
    DeleteDC screendc
    canvas.AutoRedraw = False
End Sub

Sub sendtohelpdesk(errr As String)
    Dim retrieveFilePath                                               As String
    On Error GoTo errordaa:
    Dim sPic As IPictureDisp
    Dim MESS As String
    Dim retVal As String
    Dim ictr As Integer

    sarahbaby = "realsystems"
    retrieveFilePath = App.path + "\screenshot.JPG"
    Clipboard.Clear
    keybd_event VK_MENU, 0, 0, 0
    DoEvents
    keybd_event VK_SNAPSHOT, 1, 0, 0
    DoEvents
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    DoEvents
    Set sPic = Clipboard.GetData(0)
    SavePicture sPic, retrieveFilePath
    Clipboard.Clear
    Set sPic = Nothing

    frmMain.Enabled = False

    MESS = "From : (" & COMPANY_CODE & ") " & COMPANY_NAME & vbCrLf & " " & _
            "Module name: " & App.FileDescription & vbCrLf & " " & _
            "Module version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & " " & _
            "Computer name: " & GetMachineName & vbCrLf & " " & _
            "User name: " & LOGNAME & vbCrLf & " " & _
            "Date error occured: " & DateValue(Now) & vbCrLf & " " & _
            "Time error occured: " & TimeValue(Now) & vbCrLf & " " & _
            "Error Message: " & errr
    'send email
    DoEvents
    If Ping("http://www.gmail.com/") = False Then Exit Sub
    Call SendMail(Trim$("sarahjoyreal2@gmail.com"), _
            Trim$("DMIS CONCERN"), _
            Trim$(COMPANY_NAME & "-" & LOGNAME) & "<" & Trim$("sarahjoyreal2@gmail.com") & ">", _
            Trim$(MESS), _
            Trim$("smtp.gmail.com"), _
            CInt(Trim$("465")), _
            Trim$("bsuintranet@gmail.com"), _
            Trim$(sarahbaby), _
            Trim$(retrieveFilePath), _
            CBool(1))
    Exit Sub
errordaa:
    Exit Sub
End Sub

Public Function SendMail(sTo As String, sSubject As String, sFrom As String, sBody As String, sSmtpServer As String, iSmtpPort As Integer, sSmtpUser As String, sSmtpPword As String, sFilePath As String, bSmtpSSL As Boolean)
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.message
    Set lobj_cdomsg = New CDO.message

    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 10
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    
    DoEvents
    lobj_cdomsg.Send
    MsgBox "Bug sent", vbInformation, "Sent"
    frmMain.Enabled = True
    Set lobj_cdomsg = Nothing
    Kill sFilePath
    Exit Function
SendMail_Error:
Exit Function
End Function

Public Function Ping(prmIPaddr As String) As Boolean
    Ping = InternetCheckConnection(prmIPaddr, FLAG_ICC_FORCE_CONNECTION, 0&)
End Function
   
Sub CaptureScreen(PityurBox As PictureBox)
    Dim DestDC, XPixels, YPixels, destX                     As Long
    Dim destY, srcDC, SrcX, SrcY, RasterOp                  As Long
    BottomY = PityurBox.ScaleHeight
    RightX = PityurBox.ScaleWidth
    CScrKua PityurBox
    PityurBox.Refresh
    DoEvents
    destX = 0: destY = 0
    XPixels = PityurBox.ScaleWidth
    YPixels = PityurBox.ScaleHeight
    srcDC = PityurBox.hDC
    SrcX = 0: SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, destX, destY, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    XPixels = PityurBox.ScaleWidth
    YPixels = PityurBox.ScaleHeight
    srcDC = PityurBox.hDC
    SrcX = 0: SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, 0, 0, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    DestDC = 0: XPixels = 0: YPixels = 0:
    srcDC = 0: SrcX = 0: SrcY = 0: RasterOp = 0
    Dim FileNaeym                                           As String
    If LOGCODE <> "" And LOGDATE <> "" Then
        FileNaeym = "C:\" & App.EXEName & "_" & LOGCODE & "_" & Trim(STR(Month(LOGDATE))) & Trim(STR(Day(LOGDATE))) & Trim(STR(Year(LOGDATE))) & "_" & Left(STR(Time), 2) & Mid(STR(Time), 4, 2) & Mid(STR(Time), 7, 2) & Right(STR(Time), 2) & ".jpg"
        SavePicture PityurBox.Image, FileNaeym
    End If
    Set PityurBox = Nothing
End Sub

Function Date2Month(Value As String)
    Dim MO                                                  As String
    MO = "January  February March    April    May      June     July     August   SeptemberOctober  November December "
    Date2Month = Mid$(MO, (Month(Value) - 1) * 9 + 1, 9)
End Function

Function MonthToInt(Xmonth As String) As Integer
    'If Xmonth = "January" Then getmonthcode = "A"
    'If Xmonth = "February" Then getmonthcode = "B"
    'If Xmonth = "March" Then getmonthcode = "C"
    'If Xmonth = "April" Then getmonthcode = "D"
    'If Xmonth = "May" Then getmonthcode = "E"
    'If Xmonth = "June" Then getmonthcode = "F"
    'If Xmonth = "July" Then getmonthcode = "G"
'    If Xmonth = "August" Then getmonthcode = "H"
'    If Xmonth = 9 Then getmonthcode = "I"
'    If Xmonth = 10 Then getmonthcode = "J"
'    If Xmonth = 11 Then getmonthcode = "K"
'    If Xmonth = 12 Then getmonthcode = "L"
End Function

Public Function Module_Access(USERID As Long, SelectedModule As String, ModuleType) As Boolean
    Dim RS                                                  As ADODB.Recordset
    Dim SQL                                                 As String
    Dim checkRs                                             As ADODB.Recordset
    Set checkRs = gconDMIS.Execute("select count(*) from ALL_RAMS_MODULES where descriptions='" & SelectedModule & "' AND MAINMODULENAME='" & App.TITLE & "' AND MODULE_TYPE='" & ModuleType & "'")
    If checkRs.Fields(0).Value = 0 Then
        MessagePop InfoVoid, "Module Not Config !", "Modules Reconfigure! Please contact your System Administrator for Assigning of Module."
        gconDMIS.Execute ("INSERT INTO ALL_Rams_Modules (MAINMODULENAME, DESCRIPTIONS, MODULE_TYPE) values ('" & App.TITLE & "', '" & SelectedModule & "', '" & ModuleType & "')")
        Exit Function
    Else
        SQL = " SELECT"
        SQL = SQL & " COUNT(*) from"
        SQL = SQL & " ALL_vW_RAMS_USERACESS"
        SQL = SQL & " Where"
        SQL = SQL & " USERID=" & USERID & " and MAINMODULENAME='" & App.TITLE & " ' and "
        SQL = SQL & " descriptions='" & SelectedModule & "'"
        SQL = SQL & " AND MODULE_TYPE='" & ModuleType & "'"
        Set RS = gconDMIS.Execute(SQL)

        If RS.Fields(0).Value <> 0 Then
            Module_Access = True
        Else
            Module_Access = False
            MessagePop InfoVoid, "Access denied!", "Access denied! Contact Sys Ad!" & vbCrLf & "::" & SelectedModule & vbCrLf & "::" & ModuleType
        End If
    End If
    Set RS = Nothing
End Function

Public Function AllowReprint(ModuleDescription As String)
    Dim RS                                                  As ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT  COUNT(*) FROM ALL_Rams_UsersAcess  INNER JOIN ALL_Rams_Modules ON ALL_Rams_UsersAcess.MODULEID = ALL_Rams_Modules.MODULEID WHERE descriptions='" & ModuleDescription & "' AND ACESS_REPRINT=1 and mainmodulename='" & App.TITLE & "' AND USERID=" & LOGID)
    If RS.Fields(0).Value <> 0 Then
        AllowReprint = True
    Else
        AllowReprint = False
        MessagePop InfoVoid, "Re-Print Disabled!", ":: Reprinting of " & ModuleDescription & vbCrLf & ":: Please Contact Your Sys-Ad!"
    End If

End Function

Public Function Function_Access(USERID As Long, SelectedFeature As String, MODULENAME As String) As Boolean
    Dim SQL                                                 As String
    Dim RS                                                  As ADODB.Recordset
    Dim checkRs                                             As ADODB.Recordset
    Dim xSelectedFeature                                    As String
    Dim xModuleType                                         As String
    Dim rsModuleType                                        As ADODB.Recordset

    SQL = "SELECT COUNT(MODULEID) FROM ALL_vW_USERACESS"
    SQL = SQL & " WHERE USERID=" & USERID
    SQL = SQL & " AND " & SelectedFeature & " = 1 "
    SQL = SQL & " AND MAINMODULENAME= '" & App.TITLE & "'"
    SQL = SQL & "  AND ltrim(rtrim(DESCRIPTIONS))='" & LTrim(RTrim(MODULENAME)) & "'"
    Set RS = gconDMIS.Execute(SQL)
    
    If USERID = 1 Then GoTo Netspeed
    
    If MODULENAME <> "" Then
        Set rsModuleType = gconDMIS.Execute("select module_type from ALL_RAMS_MODULES WHERE MAINMODULENAME='" & App.TITLE & "' AND DESCRIPTIONS='" & MODULENAME & "'")
        If Not rsModuleType.EOF Or Not rsModuleType.BOF Then
            xModuleType = Null2String(rsModuleType!MODULE_TYPE)
        End If
    End If
    
    If RS.Fields(0).Value = 0 Then
        Function_Access = False
        Select Case UCase(SelectedFeature)
        Case "ACESS_ADD"
            xSelectedFeature = "ADD NEW ENTRY"
        Case "ACESS_EDIT"
            xSelectedFeature = "EDIT ENTRY"
        Case "ACESS_DELETE"
            xSelectedFeature = "DELETE ENTRY"
        Case "ACESS_VIEW"
            xSelectedFeature = "VIEW TRANSACTION DETAIL"
        Case "ACESS_PRINT"
            xSelectedFeature = "PRINT ENTRY OR TRANSACTION"
        Case "ACESS_PROCESS"
            xSelectedFeature = "PROCESS TRANSACTION(S)"
        Case "ACESS_SYSTEM"
            xSelectedFeature = "ACCESS SYSTEM"
        Case "ACESS_POST"
            xSelectedFeature = "POST TRANSACTION OR ENTRY"
        Case "ACESS_UNPOST"
            xSelectedFeature = "UN POST TRANSACTION OR ENTRY"
        Case "ACESS_CANCELENTRY"
            xSelectedFeature = "CANCEL TRANSACTION OR ENTRY"
        End Select

        MessagePop InfoVoid, "Invalid Access Level", "Please Contact your SYS AD." & vbCrLf & "Module Name: " & MODULENAME & vbCrLf & "Function:" & xSelectedFeature & vbCrLf & "Module Type:" & xModuleType, 3500, 2
    Else
Netspeed:
        Function_Access = True
    End If
    Set RS = Nothing

End Function

Public Function INVERSE_FUNCTION(XXX As Double) As Double
    Const A1 = -39.6968302866538, a2 = 220.946098424521, a3 = -275.928510446969
    Const a4 = 138.357751867269, a5 = -30.6647980661472, a6 = 2.50662827745924
    Const b1 = -54.4760987982241, b2 = 161.585836858041, b3 = -155.698979859887
    Const b4 = 66.8013118877197, b5 = -13.2806815528857, C1 = -7.78489400243029E-03
    Const c2 = -0.322396458041136, c3 = -2.40075827716184, c4 = -2.54973253934373
    Const c5 = 4.37466414146497, c6 = 2.93816398269878, d1 = 7.78469570904146E-03
    Const d2 = 0.32246712907004, d3 = 2.445134137143, d4 = 3.75440866190742
    Const p_low = 0.02425, p_high = 1 - p_low
    Dim q As Double, r                                      As Double

    If XXX < 0 Or XXX > 1 Then
        err.Raise vbObjectError, , "Inverse Function: Argument out of range."
    ElseIf XXX < p_low Then
        If XXX = 1 Then XXX = 0.99
        'q = Sqr(-2 * Log(XXX))
        If XXX = 0 Then q = 0 Else q = Sqr(-2 * Log(XXX))
        INVERSE_FUNCTION = (((((C1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                           ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    ElseIf XXX <= p_high Then
        q = XXX - 0.5: r = q * q
        INVERSE_FUNCTION = (((((A1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
                           (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
    Else
        If XXX = 1 Then XXX = 0.99
        q = Sqr(-2 * Log(1 - XXX))
        INVERSE_FUNCTION = -(((((C1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                           ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
    End If
End Function

Public Sub ReminderModule(xxTime)
    On Error GoTo ADDER:
    If gconDMIS Is Nothing Then Exit Sub
    Dim temprs                                              As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select MIN(NextTime) from CRIS_Reminders where EntityType='E' and SNOOZED=0 and  MONTH(nexttime)=MONTH(getdate()) and YEAR(nexttime)=YEAR(getdate()) and USERID=" & LOGID & "  and nexttime < = getdate()")

    If IsNull(temprs.Fields(0).Value) = False Then
        TIMER_REMIND = temprs.Fields(0).Value & ""
    Else
        TIMER_REMIND = xxTime
    End If
    Exit Sub

ADDER:
    err.Clear
    Exit Sub
End Sub

Sub SetUserPathSettings()
    Dim CURRENT_REPORTS_PATH                                As String
    Dim plngSize                                            As Long
    Dim pstrBuffer                                          As String

    'CURRENT_REPORTS_PATH = "LOCAL"
    CURRENT_REPORTS_PATH = "NSI_SERVER"
    'CURRENT_REPORTS_PATH = "DEALER"
    AMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "AMIS") & "\"
    CMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CMIS") & "\"
    CRIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CRIS") & "\"
    CSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "CSMS") & "\"
    HRMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "HRMS") & "\"
    OSMS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "OSMS") & "\"
    SMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "SMIS") & "\"
    PMIS_REPORT_PATH = GetSetting("DMIS 2.0", "REPORTS", "PMIS") & "\"
    HRMS_PICTURES_PATH = GetSetting("DMIS 2.0", "REPORTS", "HRMS") & "\images\"
End Sub

Sub SetCompanyProfile()
    MODULENAME = App.EXEName
    Dim rsProfile                                           As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = '" & MODULENAME & "'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        DEALER_CODE = Null2String(rsProfile!DEALERCODE)
        COMPANY_CODE = Null2String(rsProfile!COMPANYCODE)
        COMPANY_NAME = Null2String(rsProfile!CompanyName)
        COMPANY_ADDRESS = Null2String(rsProfile!Companyaddress)
        COMPANY_TIN = Null2String(rsProfile!companytinno)
        PREPARED_BY = Null2String(rsProfile!PreparedBy)
        CHECKED_BY = Null2String(rsProfile!CheckedBy)
        APPROVED_BY = Null2String(rsProfile!ApprovedBy)
        ACCOUNT_NO = Null2String(rsProfile!ACCOUNTNO)
        BANK_MANAGER = Null2String(rsProfile!bankmanager)
        SECRETARY = Null2String(rsProfile!SECRETARY)
        NOTED_BY = Null2String(rsProfile!notedby1)
        GENERAL_MANAGER = Null2String(rsProfile!GeneralManager)
    End If
    Set rsProfile = Nothing
End Sub

Public Function OnlyInteger(KeyCode As Integer) As Integer
    If KeyCode <> vbKeyHome And KeyCode <> vbKeyEnd And KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 27 Then
        If (KeyCode < 48 Or KeyCode > 57) And KeyCode <> 110 Then
            OnlyInteger = 0
        Else
            OnlyInteger = KeyCode
        End If
    Else
        OnlyInteger = KeyCode
    End If
End Function

Sub SaveReprintInformation(XApplication_type As Variant, XModule_name As Variant, Xtransaction_no As Variant, XReason As String, XDate_Reprint As Variant, Who_Reprint As String, IsReprint As Byte)
'Update By : BTT - 07212008
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset
    Dim RSReprint                                           As New ADODB.Recordset
    Dim nard                                                As String

    On Error GoTo RYAN:

    nard = "SELECT * from ALL_reprint_transaction where Application_type = '" & XApplication_type & "' and Transaction_no = '" & Xtransaction_no & "' AND REPRINT = " & 0 & ""

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(nard)
    If RS.EOF And RS.BOF Then
        SQL = "INSERT INTO ALL_Reprint_transaction (Application_type,Module_Name,Transaction_no,Reason,date_reprint,Who_Reprint,reprint)Values('" & XApplication_type & _
              "','" & XModule_name & "','" & Xtransaction_no & "','" & XReason & "'," & N2Date2Null(XDate_Reprint) & ",'" & Who_Reprint & "'," & IsReprint & ")"
        gconDMIS.Execute (SQL)
        REPRINT_CAPTION = "NO"
    Else
        REPRINT_CAPTION = "YES"
        With FrmReprintTransaction 'Look into AMIS folder AMIS form
            .LblTransactionNo = Xtransaction_no
            .lblTransaction_type = XApplication_type
            FrmReprintTransaction.Show 1
        End With
    End If
    Set RS = Nothing

    Exit Sub
RYAN:
    MsgBox err.DESCRIPTION, "Error", "Please Contact NSI Administrator"
End Sub

Sub ReturnInvoiceNo(xVOUCHERNO As String, xJType As String)
' update by BTT
    Dim RSHD                                                As New ADODB.Recordset
    Set RSHD = gconDMIS.Execute("Select voucherno,invoicetype,invoiceno,jtype from AMIS_journal_hd where voucherno ='" & xVOUCHERNO & _
                                "' and jtype='" & xJType & "'")
    If Not (RSHD.EOF And RSHD.BOF) Then
        AMIS_Invoiceno = Null2String(RSHD!INVOICENO)
        AMIS_Invoicetype = Null2String(RSHD!INVOICETYPE)
    Else
        AMIS_Invoiceno = N2Str2Null("")
        AMIS_Invoicetype = N2Str2Null("")
    End If
    Set RSHD = Nothing
End Sub

Function UpdateBalanceSJ(CRJVoucherno As String, is_posted As Boolean) As Double
    Dim RSCRJ                                               As New ADODB.Recordset
    Dim totalpayment                                        As Double
    Dim RSSJ                                                As New ADODB.Recordset
    Set RSCRJ = gconDMIS.Execute("SELECT * FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO ='" & CRJVoucherno & "'")
    If Not (RSCRJ.EOF And RSCRJ.BOF) Then
        Do While Not RSCRJ.EOF
            Set RSSJ = gconDMIS.Execute("Select voucherno,invoiceamt,invoiceno,invoicetype,balance from AMIS_journal_hd where jtype='SJ' and invoiceno='" & Null2String(RSCRJ!INVOICENO) & _
                                        "' and invoicetype='" & Null2String(RSCRJ!INVOICETYPE) & "'")
            If Not (RSSJ.EOF And RSSJ.BOF) Then
                If is_posted = True Then
                    UpdateBalanceSJ = NumericVal(RSSJ!BALANCE) - NumericVal(RSCRJ!invoiceamount)
                    gconDMIS.Execute ("Update amis_journal_hd set balance='" & NumericVal(UpdateBalanceSJ) & _
                                      "' where voucherno='" & RSSJ!VOUCHERNO & "' and jtype = 'SJ'")
                Else
                    gconDMIS.Execute ("Update amis_journal_hd set balance='" & NumericVal(RSSJ!InvoiceAmt) & _
                                      "' where voucherno='" & RSSJ!VOUCHERNO & "' and jtype = 'SJ'")
                End If

            End If
            RSCRJ.MoveNext
        Loop
    End If
    Set RSSJ = Nothing
    Set RSCRJ = Nothing
End Function

Public Function LimitChar(ByVal ALPHA As String, ByVal k As Integer)
    If InStr(ALPHA, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function


Function EXTRACT_FILES(CUSTOMID As Long, FileName As String) As Boolean
    On Error GoTo ErrorCode:
'---------------------------------------
'    FOR PARTS
'---------------------------------------
'    CHANGES IN MAC.xlt  = 102
'    DPI.xlt             = 103
'    MACMAC.xlt          = 104
'    PartsRundown.xlt    = 105
'    PO.xlt              = 106
'    PQIR.xlt            = 107
'    ADB.XLT             = 108
'---------------------------------------
'FOR SERVICE
'AfterSalesReportsSERVICE.xlt   = 111
'LABOR COST REPORTS             = 112
'---------------------------------------
    Dim b()                                                 As Byte
    Dim s                                                   As String
    Dim i                                                   As Long

    Dim temp                                                As String
    Dim StartPosition                                       As Long
    Dim mHandle                                             As Integer

    s = ""
    b = LoadResData(CUSTOMID, "CUSTOM")
    For i = 0 To UBound(b())
        s = s & Chr(b(i))
    Next i
    Erase b

    mHandle = FreeFile
    If FileName = "AccountGeneralLedgerAllAccount.xlt" Then
        Open AMIS_REPORT_PATH & "Ledgers\" & FileName For Binary As #mHandle
    Else
        Open App.path & "\" & FileName For Binary As #mHandle
    End If
    StartPosition = LOF(mHandle)
    temp = s
    Put #mHandle, , temp
    Put #mHandle, , StartPosition
    Close #mHandle
    EXTRACT_FILES = True
    Exit Function
ErrorCode:
    err.Clear
    EXTRACT_FILES = False
End Function

Function CheckIfRoIsAlreadyInvoice(XXX As String) As Boolean
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Not Null2String(RSTMP!INVOICE) = "" Then
            CheckIfRoIsAlreadyInvoice = True
        End If
    End If
    Set RSTMP = Nothing
End Function

Function CheckIfROStillExist(XXX As String) As Boolean
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfROStillExist = True
    End If
    Set RSTMP = Nothing
End Function

Function GetUserAction(xUSERACTION) As String
    Select Case xUSERACTION
    Case "A"
        GetUserAction = "Added Record"
    Case "E"
        GetUserAction = "Edit Record"
    Case "P"
        GetUserAction = "Post Record"
    Case "U"
        GetUserAction = "Unpost Record"
    Case "C"
        GetUserAction = "Cancel Record"
    Case "X"
        GetUserAction = "Delete Record"
    Case "V"
        GetUserAction = "View Record"
    Case "I"
        GetUserAction = "Inquire Record"
    Case "R"
        GetUserAction = "Process"
    Case "G"
        GetUserAction = "Generate"
    Case "O"
        GetUserAction = "Batch Posting"
    Case "B"
        GetUserAction = "Billed"
    Case "M"
        GetUserAction = "Import"
    Case "AA"
        GetUserAction = "Added Details"
    Case "EE"
        GetUserAction = "Edit Details"
    Case "UU"
        GetUserAction = "Unpost Details"
    Case "CC"
        GetUserAction = "Cancel Details"
    Case "XX"
        GetUserAction = "Cancel Details"
    Case "PP"
        GetUserAction = "Post Details"
    Case "AP"
        GetUserAction = "Approved"
    Case "DS"
        GetUserAction = "Disapproved"
    Case "AT"
        GetUserAction = "Attached"
    Case "DT"
        GetUserAction = "Dettached"
    Case "RD"
        GetUserAction = "Released"
    Case "UR"
        GetUserAction = "Unreleased"
    Case "JI"
        GetUserAction = "Clock In"
    Case "JO"
        GetUserAction = "Clock Out"
    Case "AS"
        GetUserAction = "Assigned"
    Case "RE"
        GetUserAction = "Removed"
    Case "UP"
        GetUserAction = "Upload"
    Case "RC"
        GetUserAction = "Recover"
    Case "JP"
        GetUserAction = "Job Passed"
    Case "JF"
        GetUserAction = "Job Failed"
    Case "CF"
        GetUserAction = "Confirmed"
    Case "MM"
        GetUserAction = "Import Details"
    Case "MP"
        GetUserAction = "Parts Memo"
    Case "CL"
        GetUserAction = "Search Button Click"
    Case "CS"

    Case "D"

    Case "PX"

    Case "UD"
        GetUserAction = "Upload Details"

    Case Else
        GetUserAction = xUSERACTION
    End Select
End Function

Function CheckIfTheJobIsFinish(XRONO As String, xJOBCODE As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DONE FROM CSMS_RODET " & _
                                 " WHERE LIVIL = 1 " & _
                                 " AND DETCODE = " & N2Str2Null(xJOBCODE) & _
                                 " AND REP_OR = " & N2Str2Null(XRONO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!DONE) = "Y" Then
            CheckIfTheJobIsFinish = "Finish"
        Else
            CheckIfTheJobIsFinish = "Not Finish"
        End If
    End If
    Set RSTMP = Nothing
End Function

Function CheckIfAppointmentTimeIsAvailable(xDate As String, xtime As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CUSNAM FROM CSMS_APPOINTMENT WHERE TRANDATE = " & N2Str2Null(xDate) & _
                                 " AND APPTTIME = " & N2Str2Null(xtime) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfAppointmentTimeIsAvailable = Null2String(RSTMP!CusNam)
    End If
    Set RSTMP = Nothing
End Function

Function CheckAppointmentStatus(xDate As String, xtime As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_APPOINTMENT WHERE TRANDATE = " & N2Str2Null(xDate) & _
                                 " AND APPTTIME = " & N2Str2Null(xtime) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckAppointmentStatus = Null2String(RSTMP!Status)
    End If
    Set RSTMP = Nothing
End Function

Function CheckEstimateStatus(xESTNO As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT UPLOAD_STATUS FROM CSMS_ESTHD WHERE ESTIMATENO = " & N2Str2Null(xESTNO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!UPLOAD_Status) = "Y" Then
            CheckEstimateStatus = "UPLOADED"
        Else
            CheckEstimateStatus = "NOT UPLOADED"
        End If
    Else
        CheckEstimateStatus = "NOT FOUND"
    End If
    Set RSTMP = Nothing
End Function

Function GetFreshServiceCounterStatus(XRONO As String) As String
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_REPAIRORDER WHERE RO_NO = " & N2Str2Null(XRONO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetFreshServiceCounterStatus = LTrim(RTrim(Null2String(RSTMP!Status)))
    End If
    Set RSTMP = Nothing
End Function

'UPDATE BY   : MJP 11092009 0134PM
'DESCRIPTION : CRF 108
Function CheckIfUserIsAnServiceAdviser() As Boolean
    Dim RSDSA                                               As New ADODB.Recordset
    Dim rsHRMS                                              As New ADODB.Recordset
    Set RSDSA = gconDMIS.Execute("SELECT EMPNO FROM ALL_RAMS_USERS WHERE LTRIM(RTRIM(USERCODE)) = " & N2Str2Null(LTrim(RTrim(LOGCODE))) & "")
    If Not (RSDSA.BOF And RSDSA.EOF) Then
        Set rsHRMS = gconDMIS.Execute("SELECT IS_SERVICE_ADVISER FROM HRMS_EMPINFO WHERE EMPNO = '" & Null2String(RSDSA!empno) & "'")
        If Not (rsHRMS.EOF And rsHRMS.BOF) Then
            If Null2Bool(rsHRMS!IS_SERVICE_ADVISER) = "1" Then
                CheckIfUserIsAnServiceAdviser = True
            Else
                CheckIfUserIsAnServiceAdviser = False
            End If
        Else
            CheckIfUserIsAnServiceAdviser = False
        End If
    End If
    Set RSDSA = Nothing
End Function
'UPDATE BY   : MJP 11092009 0134PM

'UPDATE BY   : MJP 11082009 0254 PM
'DESCRIPTION : CRF 121
Function CheckJobStatusIfIdle(xREPOR As String, xDETCODE As String)
    Dim RSTMP                                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_RO_dET WHERE " & _
                                 " REP_OR = " & N2Str2Null(xREPOR) & _
                                 " AND LIVIL = 1 " & _
                                 " AND DETCDE = " & N2Str2Null(xDETCODE) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckJobStatusIfIdle = Null2String(RSTMP!Status)
    End If
    Set RSTMP = Nothing
End Function
'DESCRIPTION : CRF 121

'UPDATE BY   : MJP 11092009 0133PM
'DESCRIPTION : CRF 116
Function GenerateNextInvoiceno()
    Dim RSTMP                                               As New ADODB.Recordset


    Set RSTMP = Nothing
End Function
'UPDATE BY   : MJP 11082009 0254 PM

'UPDATE BY   : MJP 04212010 0244 PM
'DESCRIPTION : THIS UPDATE IS TO AVOID THE NEGATIVE VALUE IN THE BILLING SYSTEM,
'              THIS IS HAPPEN DUE TO THE INSURANCE VALUE IS ALREADY SET AND AMOUNT IS
'              CHANGE IN DETAILS VALUE
Function CheckIfInsuranceIsAlreadySet(xRO As String, xDETID As Integer, xAMOUNT As Currency, xDISCOUNT As Currency) As Boolean
    Dim rsINS                                               As New ADODB.Recordset
    Dim xINS                                                As Currency
    Dim xDET_AMT                                            As Currency
    Dim xDET_DIS                                            As Currency

    Set rsINS = gconDMIS.Execute("SELECT INSAMT FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(xRO) & "")
    If Not (rsINS.BOF And rsINS.EOF) Then
        xINS = NumericVal(rsINS.Fields(0))
    End If
    Set rsINS = Nothing

    Set rsINS = gconDMIS.Execute("SELECT DET_AMT, DISCOUNT_2 FROM CSMS_RO_DET WHERE ID = " & xDETID & "")
    If Not (rsINS.BOF And rsINS.EOF) Then
        xDET_AMT = NumericVal(rsINS.Fields(0))
        xDET_DIS = NumericVal(rsINS.Fields(1))
    End If
    Set rsINS = Nothing

    If Not xINS = 0 Then
        If xAMOUNT <> xDET_AMT Then
            CheckIfInsuranceIsAlreadySet = True
        Else
            If xDISCOUNT <> xDET_DIS Then
                CheckIfInsuranceIsAlreadySet = True
            Else
                CheckIfInsuranceIsAlreadySet = False
            End If
        End If
    End If
End Function
'UPDATE BY   : MJP 04212010 0244 PM

Function CheckifDetailsIsSublet(xID As Long) As Boolean
    Dim RSTMP                                               As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT ISNULL(ROTYPE,'') FROM CSMS_RO_DET WHERE ID = " & xID & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP.Fields(0)) = "SR" Then
            CheckifDetailsIsSublet = True
        Else
            CheckifDetailsIsSublet = False
        End If
    End If
    Set RSTMP = Nothing
End Function

Function CheckServerDate() As Boolean
    On Error GoTo ErrorCode
    Dim RS                                                  As ADODB.Recordset

    Set RS = gconDMIS.Execute("Select getdate() as DateNow, host_name() as PCName")
    If RS!PCNAME <> "SERVER" Or RS!PCNAME <> "DMISSERVER" Or RS!PCNAME <> "MASTER" Then
        If Date <> DateValue(RS!DateNow) Then
            CheckServerDate = False
            MessagePop InfoStop, "Warning", "Computer Date is not Equal with the Server Date, System not allow user to do a Backdate transaction. System will correct your Computer Date to Proceed in Saving"
            Date = RS!DateNow
            Time = RS!DateNow
        Else
            CheckServerDate = True
        End If
    End If
    Set RS = Nothing

    Exit Function

ErrorCode:
    MessagePop InfoStop, "Error", "" & err.Number & " " & err.DESCRIPTION
    err.Clear
End Function

Function checkdup(XTYPE As String, xTRANTYPE As String, xtrano As String) As Boolean
    Dim rsfindDup                                           As ADODB.Recordset
    Dim sqlcommand                                          As String

    Set rsfindDup = New ADODB.Recordset
    sqlcommand = "select trantype,tranno from PMIS_vw_PRS where trantype = '" & xTRANTYPE & "' and tranno = '" & xtrano & "' "
    sqlcommand = sqlcommand + "UNION ALL "
    sqlcommand = sqlcommand + " select trantype,tranno from PMIS_ORD_HIST where trantype = '" & xTRANTYPE & "' and tranno = '" & xtrano & "' and [type] = '" & XTYPE & "' "
    rsfindDup.Open (sqlcommand), gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsfindDup.EOF And Not rsfindDup.BOF Then
        checkdup = True
    Else
        checkdup = False
    End If
    Set rsfindDup = Nothing
    Exit Function
End Function

Function checkdup_PO(XTYPE As String, xpono As String) As Boolean
    Dim rsfindDup                                           As ADODB.Recordset
    Dim sqlcommand                                          As String

    Set rsfindDup = New ADODB.Recordset
    sqlcommand = "select pono from PMIS_PO_HD where pono = '" & xpono & "' and [type] = '" & XTYPE & "' "
    sqlcommand = sqlcommand + "UNION ALL "
    sqlcommand = sqlcommand + "select pono from PMIS_PO_HIST where pono = '" & xpono & "' and [type] = '" & XTYPE & "' "
    rsfindDup.Open (sqlcommand), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsfindDup.EOF And Not rsfindDup.BOF Then
        checkdup_PO = True
    Else
        checkdup_PO = False
    End If
    Set rsfindDup = Nothing
    Exit Function

End Function

Function checkdup_rr(XTYPE As String, xtrano As String) As Boolean
    Dim rsfindDup                                           As ADODB.Recordset
    Dim sqlcommand                                          As String

    sqlcommand = ""
    sqlcommand = "select rrno from PMIS_RR_Hd where [TYPE] = '" & XTYPE & "' AND rrno = '" & xtrano & "' "
    sqlcommand = sqlcommand + "UNION ALL "
    sqlcommand = sqlcommand + "select rrno from PMIS_Rec_hist where [TYPE] = '" & XTYPE & "' AND rrno = '" & xtrano & "' "
    
    Set rsfindDup = New ADODB.Recordset
    rsfindDup.Open (sqlcommand), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsfindDup.EOF And Not rsfindDup.BOF Then
        checkdup_rr = True
    Else
        checkdup_rr = False
    End If
    Set rsfindDup = Nothing
    Exit Function
    
End Function

Function checkdup_ISS(XTYPE As String, xTRANTYPE As String, xtrano As String) As Boolean
    Dim rsfindDup                                           As ADODB.Recordset
    Dim sqlcommand                                          As String

    sqlcommand = ""
    sqlcommand = "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = '" & XTYPE & "' AND trantype = '" & xTRANTYPE & "' and tranno = '" & xtrano & "'"
    sqlcommand = sqlcommand + "UNION ALL "
    sqlcommand = sqlcommand + "select trantype,tranno from PMIS_Ord_Hist where [TYPE] =  '" & XTYPE & "' AND trantype = '" & xTRANTYPE & "' and tranno = '" & xtrano & "'"
    Set rsfindDup = New ADODB.Recordset

    rsfindDup.Open (sqlcommand), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsfindDup.EOF And Not rsfindDup.BOF Then
        checkdup_ISS = True
    Else
        checkdup_ISS = False
    End If
    Set rsfindDup = Nothing
    Exit Function
End Function

'Updated By: IEBV
'description: to get all numeric in a string
Public Function TONUMERIC(NumericText As Variant) As Double
    Dim COUNTER                                        As Integer
    Dim NumericValue                                   As String
    NumericValue = ""
    If Trim(NumericText) <> "" Then
        If Val(NumericText) >= 0 Then
            For COUNTER = 1 To Len(NumericText)
                If Mid(NumericText, COUNTER, 1) <> "," And IsNumeric(Mid(NumericText, COUNTER, 1)) = True Then
                    NumericValue = NumericValue & Mid(NumericText, COUNTER, 1)
                End If
            Next
            TONUMERIC = NumericValue
        Else
            TONUMERIC = Val(NumericText)
        End If
    Else
        TONUMERIC = 0
    End If
End Function

Function checkdup_INVO(XTYPE As String, xinvo As String, xsupcode As String) As Boolean
    Dim rsINVNO                                             As ADODB.Recordset
    Dim XINVNO                                              As String

    Set rsINVNO = New ADODB.Recordset
    Dim strsqlcommand                                       As String
    strsqlcommand = "Select invno from pmis_rr_hd where invno = '" & Null2String(xinvo) & "' and [TYPE] = '" & XTYPE & "' and isnull(status,'N') <> 'C' and recvd_code = '" & xsupcode & "'"
    strsqlcommand = strsqlcommand + " UNION ALL "
    strsqlcommand = strsqlcommand + "Select invno from pmis_rec_hist where invno = '" & Null2String(xinvo) & "' and [TYPE] = '" & XTYPE & "' and isnull(status,'N') <> 'C' and recvd_code = '" & xsupcode & "'"
    rsINVNO.Open (strsqlcommand), gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not (rsINVNO.EOF And rsINVNO.BOF) Then
        checkdup_INVO = True
    Else
        checkdup_INVO = False
    End If
    Set rsINVNO = Nothing
    Exit Function
End Function

Function Post_ValidQuantity(XTYPE As String, xpono As String, XRRNO As String) As Boolean
    Dim rr_hd                                               As ADODB.Recordset
    Dim RR_DT                                               As ADODB.Recordset
    Dim PO_DT                                               As ADODB.Recordset
    Dim tmP_RR                                              As ADODB.Recordset

    Set RR_DT = New ADODB.Recordset
    Set rr_hd = New ADODB.Recordset
    Set RR_DT = New ADODB.Recordset
    Set tmP_RR = New ADODB.Recordset

    Dim i                                                   As Integer
    Dim iok                                                 As Integer

    Set tmP_RR = gconDMIS.Execute("Select * from pmis_tdaytran where [type] = '" & XTYPE & "' and trantype = 'RR' and tranno = '" & XRRNO & "' ")
    If Not (tmP_RR.EOF And tmP_RR.BOF) Then
        tmP_RR.MoveFirst
        Do While Not tmP_RR.EOF
            Set PO_DT = gconDMIS.Execute("Select * from pmis_alldaytran where tranno = '" & xpono & "' and [TYPE] = '" & XTYPE & "' and status = 'P' and trantype = 'PO' and stock_ord = '" & tmP_RR!stock_ord & "'")
            If Not (PO_DT.EOF And PO_DT.BOF) Then
                i = 0
                Set rr_hd = gconDMIS.Execute("Select * from PMIS_vw_RR_Trans where PONO = '" & xpono & "' and type = '" & XTYPE & "' and isnull(status,'N') in ('P','N')")
                If Not (rr_hd.EOF And rr_hd.BOF) Then
                    rr_hd.MoveFirst
                    i = 0
                    Do While Not rr_hd.EOF
                        Set RR_DT = gconDMIS.Execute("Select * from PMIS_vw_RR_Trans_Details where tranno = '" & rr_hd!RRNO & "' and [type] = '" & rr_hd!Type & "' and stock_ord = '" & tmP_RR!stock_ord & "' and ISNULL(status,'N') IN ('P','N') and tremarks is not null")
                        If Not (RR_DT.EOF And RR_DT.BOF) Then
                            i = i + RR_DT!tranqty
                        End If
                        rr_hd.MoveNext
                    Loop
                    If i > PO_DT!tranqty Then
                        iok = iok + 1
                    End If
                End If
            End If
            tmP_RR.MoveNext
        Loop
    End If

    If iok > 0 Then
        Post_ValidQuantity = False
    Else
        Post_ValidQuantity = True
    End If

    Exit Function
End Function

Function checkfcleartodelete(XTYPE As String, XRRNO As String, XPART As String) As Boolean
    Dim RSRR                                                As ADODB.Recordset
    Dim RSISS                                               As ADODB.Recordset

    Set RSRR = New ADODB.Recordset
    Set RSISS = New ADODB.Recordset
    Set RSRR = gconDMIS.Execute("Select * from pmis_tdaytran where [type] = '" & XTYPE & "' and tranno = '" & XRRNO & "' and stock_ord = '" & XPART & "'")
    If Not (RSRR.EOF And RSRR.BOF) Then
        Set RSISS = gconDMIS.Execute("SELECT * from pmis_tdaytran where [type] = '" & XTYPE & "' and trantype in('RIV','CSH','CHG','DR') and stock_ord = '" & XPART & "' and trandate >= '" & RSRR!trandate & "' and ID > '" & RSRR!ID & "' and isnull(status,'N') in ('N','P') and isnull(status,'N') not in ('C')")
        If Not (RSISS.EOF And RSISS.BOF) Then
            checkfcleartodelete = False
            Exit Function
        Else
            checkfcleartodelete = True
            Exit Function
        End If
    End If
    Set RSRR = Nothing
    Set RSISS = Nothing
End Function

Sub void()
    Dim sqlcommand                                          As String
    sqlcommand = ""
    sqlcommand = "IF NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='CSMS_REPOR' AND COLUMN_NAME='CANCEL_REASON')" & vbCrLf
    sqlcommand = sqlcommand & "ALTER TABLE CSMS_REPOR" & vbCrLf
    sqlcommand = sqlcommand & "ADD  CANCEL_REASON NVARCHAR(200)" & vbCrLf
    gconDMIS.Execute (sqlcommand)

    sqlcommand = ""
    sqlcommand = "IF  NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='CSMS_REPOR' AND COLUMN_NAME='CANCEL_DATE')" & vbCrLf
    sqlcommand = sqlcommand & "ALTER TABLE CSMS_REPOR" & vbCrLf
    sqlcommand = sqlcommand & "ADD  CANCEL_DATE smalldatetime" & vbCrLf
    gconDMIS.Execute (sqlcommand)

End Sub

Sub adddealercode()
    Dim sqlcommand                                          As String
    
    sqlcommand = ""
    sqlcommand = "IF  NOT EXISTS(SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='ALL_Vendor_Table' AND COLUMN_NAME='Dcode')" & vbCrLf
    sqlcommand = sqlcommand & "ALTER TABLE ALL_Vendor_Table" & vbCrLf
    sqlcommand = sqlcommand & "ADD  Dcode NVARCHAR(5)" & vbCrLf
    gconDMIS.Execute (sqlcommand)
End Sub

