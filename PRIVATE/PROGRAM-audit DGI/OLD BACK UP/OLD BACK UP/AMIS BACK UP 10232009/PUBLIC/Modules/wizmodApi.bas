Attribute VB_Name = "wizmodAPI"
Option Explicit


Public CANCEL_ANS                                      As String
Public REPRINT_CAPTION                                 As String
Public Const MAX_COMPUTERNAME_LENGTH                   As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public oVoice                                          As SpeechLib.SpVoice

'FOR HELP SYSTEM
Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
'END FOR HELP SYSTEM

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public gconDMIS                                        As ADODB.Connection
Public gconACCESS                                      As ADODB.Connection

Public ConnStr
Public DMIS_Connection
Public DMIS_Audit_Connection
Public DMIS_REPORT_Connection

Public DEALER_CODE                                     As String
Public COMPANY_CODE                                    As String
Public COMPANY_NAME                                    As String
Public COMPANY_ADDRESS                                 As String
Public COMPANY_TIN                                     As String
Public TransactionID                                   As String
Public PMIS_REPORT_PATH                                As String
Public CSMS_REPORT_PATH                                As String
Public SMIS_REPORT_PATH                                As String
Public CMIS_REPORT_PATH                                As String
Public HRMS_REPORT_PATH                                As String
Public AMIS_REPORT_PATH                                As String
Public CRIS_REPORT_PATH                                As String
Public OSMS_REPORT_PATH                                As String
Public HRMS_PICTURES_PATH                              As String

Public PMIS_REPORT_CONNECTION                          As String
Public CSMS_REPORT_CONNECTION                          As String
Public SMIS_REPORT_CONNECTION                          As String
Public CMIS_REPORT_CONNECTION                          As String
Public HRMS_REPORT_Connection                          As String
Public AMIS_REPORT_CONNECTION                          As String
Public CRIS_REPORT_CONNECTION                          As String
Public OSMS_REPORT_CONNECTION                          As String
Public SERVERNAME                                      As String
Public SQLSERVERNAME                                   As String
Public DATABASE                                        As String
''''FOR USER ACESS MANAGEMENT

Public LOGPASS                                         As String
Attribute LOGPASS.VB_VarUserMemId = 1073741856
Public LOGID                                           As Long
Attribute LOGID.VB_VarUserMemId = 1073741857
'''REMINDERS WINDOW
Public TIMER_REMIND                                    As String
''''END USER ACESS MANAGEMENT

Public COA_AR_TRADE_UNITS                              As String
Public COA_AR_TRADE_SERVICE                            As String
Public COA_AR_TRADE_PARTS                              As String

Public HYUNDAI_COA_AR_TRADE_UNITS                      As String
Public HYUNDAI_COA_AR_TRADE_SERVICE                    As String
Public HYUNDAI_COA_AR_TRADE_PARTS                      As String

Public COA_OUTPUT_TAX                                  As String

'SALES - CASH
Public COA_SALES_SERVICE_CASH_TINSPAINT                As String
Public COA_SALES_SERVICE_CASH_SUBLET                   As String
Public COA_SALES_SERVICE_CASH_AIRCON                   As String
Public COA_SALES_SERVICE_CASH_LABOR                    As String
Public COA_SALES_SERVICE_CASH_GOL                      As String
Public COA_SALES_SERVICE_CASH_PARTS                    As String

Public COA_SALES_GOL_CASH                              As String
Public COA_SALES_PARTS_CASH                            As String
Public COA_SALES_VEHICLES_CASH                         As String

'SALES - CHARGE
Public COA_SALES_SERVICE_CHARGE_TINSPAINT              As String
Public COA_SALES_SERVICE_CHARGE_SUBLET                 As String
Public COA_SALES_SERVICE_CHARGE_AIRCON                 As String
Public COA_SALES_SERVICE_CHARGE_LABOR                  As String
Public COA_SALES_SERVICE_CHARGE_GOL                    As String
Public COA_SALES_SERVICE_CHARGE_PARTS                  As String

Public COA_SALES_GOL_CHARGE                            As String
Public COA_SALES_PARTS_CHARGE                          As String

'SALES - DISCOUNT - CASH
Public COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT       As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET          As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON          As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_LABOR           As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_GOL             As String
Public COA_SALES_DISCOUNT_SERVICE_CASH_PARTS           As String

Public COA_SALES_DISCOUNT_GOL_CASH                     As String
Public COA_SALES_DISCOUNT_PARTS_CASH                   As String
Public COA_SALES_DISCOUNT_VEHICLES_CASH                As String

'SALES - DISCOUNT - CHARGE
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT     As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET        As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON        As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR         As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL           As String
Public COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS         As String

Public COA_SALES_DISCOUNT_GOL_CHARGE                   As String
Public COA_SALES_DISCOUNT_PARTS_CHARGE                 As String

'CHARGE TO WARRANTY
Public COA_DIRECT_EXPENSE_LABOR                        As String
Public COA_DIRECT_EXPENSE_SPAREPARTS                   As String
Public COA_DIRECT_EXPENSE_GOL                          As String

Public COA_WARRANTY_SALES                              As String
Public COA_WARRANTY_SERVICE                            As String
Public COA_WARRANTY_PARTS                              As String

'CHARGE TO COMPANY
Public COA_COMPANY_CAR_SALES                           As String
Public COA_COMPANY_CAR_SERVICE                         As String

'CHARGE TO SALES
Public COA_GFSI_SALES                                  As String
Public COA_GFSI_SERVICE                                As String
Public COA_GFSI_PARTS                                  As String

'CASH RECEIPTS
Public COA_CASH_ON_HAND                                As String
Public COA_BRANCH_LEGASPI                              As String
Public COA_CUSTOMER_DEPOSIT                            As String

Public COA_INSURANCE_PREMIUM_PAYABLE                   As String
Public COA_INSURANCE_PREMIUM_RENEWAL                   As String
Public COA_LTO_PAYMENT                                 As String
Public COA_CHATTEL_MORTGAGE_FEE_PAYABLE                As String
Public COA_NEW_VEHICLE_REGISTRATION                    As String
Public COA_WARRANTY_CLAIMS_RECEIVABLE                  As String

Public COA_ACCOUNTS_RECEIVABLE_NONTRADE_EMPLOYEES      As String
Public COA_OTHER_PAYABLES                              As String
Public COA_INCIDENTAL_CHARGES_UNITS                    As String
Public COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD             As String
Public COA_PRE_DELIVERY                                As String

Public COA_CORPORATE_TAX_WHELD                         As String
Public COA_CORPORATE_VAT_WHELD                         As String

Public HYUNDAI_COA_SALES_SERVICE_CASH_TINSPAINT        As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_SUBLET           As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_AIRCON           As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_LABOR            As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_GOL              As String
Public HYUNDAI_COA_SALES_SERVICE_CASH_PARTS            As String

Public HYUNDAI_COA_SALES_GOL_CASH                      As String
Public HYUNDAI_COA_SALES_PARTS_CASH                    As String

'SALES - CHARGE
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_TINSPAINT      As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_SUBLET         As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_AIRCON         As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_LABOR          As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_GOL            As String
Public HYUNDAI_COA_SALES_SERVICE_CHARGE_PARTS          As String

Public HYUNDAI_COA_SALES_GOL_CHARGE                    As String
Public HYUNDAI_COA_SALES_PARTS_CHARGE                  As String

'SALES - DISCOUNT - CASH
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_TINSPAINT As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_SUBLET  As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_AIRCON  As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_LABOR   As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_GOL     As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CASH_PARTS   As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CASH             As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CASH           As String

'SALES - DISCOUNT - CHARGE
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_TINSPAINT As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_SUBLET As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_AIRCON As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_LABOR As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_GOL   As String
Public HYUNDAI_COA_SALES_DISCOUNT_SERVICE_CHARGE_PARTS As String

Public HYUNDAI_COA_SALES_DISCOUNT_GOL_CHARGE           As String
Public HYUNDAI_COA_SALES_DISCOUNT_PARTS_CHARGE         As String

Public HYUNDAI_COA_WARRANTY_SALES                      As String
Public HYUNDAI_COA_WARRANTY_SERVICE                    As String
Public HYUNDAI_COA_WARRANTY_PARTS                      As String

Public HYUNDAI_COA_COMPANY_CAR_SALES                   As String
Public HYUNDAI_COA_COMPANY_CAR_SERVICE                 As String

Public HYUNDAI_COA_GFSI_SALES                          As String
Public HYUNDAI_COA_GFSI_SERVICE                        As String
Public HYUNDAI_COA_GFSI_PARTS                          As String

Public HYUNDAI_COA_NEW_VEHICLE_REGISTRATION            As String
Public HYUNDAI_COA_WARRANTY_CLAIMS_RECEIVABLE          As String

Public HYUNDAI_COA_INCIDENTAL_CHARGES_UNITS            As String
Public HYUNDAI_COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD     As String

Public COA_INVENTORIES_PARTS                           As String
Public COA_INVENTORIES_GOL                             As String
Public COA_INVENTORIES_VEHICLES                        As String

Public COA_COST_OF_SALES_PARTS                         As String
Public COA_COST_OF_SALES_GOL                           As String
Public COA_COST_OF_SALES_VEHICLES                      As String


Public COA_INPUT_TAX                                   As String
Public COA_INCOME_TAX_WITHHELD                         As String
Public COA_ACCOUNTS_PAYABLE                            As String

Public OPEN_AR_SHOW                                    As Boolean
Public SJ_SHOW                                         As Boolean
Public PMIS_ORDER_SHOW                                 As Boolean

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
Public AMIS_Invoicetype                                As String
Public AMIS_Invoiceno                                  As String
'Public Const ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=E:\SQLDATA\AMIS_NAGA\DATA\AMISDat.DAT"

Public ApplySecurityValidation                         As Boolean



Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000
Private m_bInIDE                                       As Boolean
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName                                       As String * CCDEVICENAME
    dmSpecVersion                                      As Integer
    dmDriverVersion                                    As Integer
    dmSize                                             As Integer
    dmDriverExtra                                      As Integer
    dmFields                                           As Long
    dmOrientation                                      As Integer
    dmPaperSize                                        As Integer
    dmPaperLength                                      As Integer
    dmPaperWidth                                       As Integer
    dmScale                                            As Integer
    dmCopies                                           As Integer
    dmDefaultSource                                    As Integer
    dmPrintQuality                                     As Integer
    dmColor                                            As Integer
    dmDuplex                                           As Integer
    dmYResolution                                      As Integer
    dmTTOption                                         As Integer
    dmCollate                                          As Integer
    dmFormName                                         As String * CCFORMNAME
    dmUnusedPadding                                    As Integer
    dmBitsPerPel                                       As Integer
    dmPelsWidth                                        As Long
    dmPelsHeight                                       As Long
    dmDisplayFlags                                     As Long
    dmDisplayFrequency                                 As Long
End Type
Dim DevM                                               As DEVMODE

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
Public RightX, BottomY, ASpeed                         As Integer
Attribute RightX.VB_VarUserMemId = 1073741991
Attribute BottomY.VB_VarUserMemId = 1073741991
Attribute ASpeed.VB_VarUserMemId = 1073741991
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)

Private Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008
Private Ind, Xo, Yo, Xs, Ys, XSrc                      As Long
Attribute Ind.VB_VarUserMemId = 1073741992
Attribute Xo.VB_VarUserMemId = 1073741992
Attribute Yo.VB_VarUserMemId = 1073741992
Attribute Xs.VB_VarUserMemId = 1073741992
Attribute Ys.VB_VarUserMemId = 1073741992
Attribute XSrc.VB_VarUserMemId = 1073741992
Private YSrc, DDC, SDC, res                            As Long
Attribute YSrc.VB_VarUserMemId = 1073741963
Attribute DDC.VB_VarUserMemId = 1073741963
Attribute SDC.VB_VarUserMemId = 1073741963
Attribute res.VB_VarUserMemId = 1073741963
Dim z2                                                 As Long
Attribute z2.VB_VarUserMemId = 1073741967

Public Const TTLDYSIN1YR = 365

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
    Dim hRgn                                           As Long
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
    Dim plngSize                                       As Long
    Dim pstrBuffer                                     As String
    pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
    plngSize = Len(pstrBuffer)
    If GetComputerName(pstrBuffer, plngSize) Then
        GetMachineName = Left$(pstrBuffer, plngSize)
    End If
End Function

Public Sub MoveKeyPress(KeyCode As Integer)

    Dim First3Letters                                  As String
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
    Dim errLoop                                        As ADODB.Error
    Dim strHelp                                        As String
    For Each errLoop In gcon.Errors
        If errLoop.HelpFile = "" Then strHelp = " No Helpfile available" Else strHelp = " Helpfile: " & errLoop.HelpFile & "; HelpContext: " & errLoop.HelpContext
        MsgBoxXP "ADO Error '" & errLoop.Number & vbCrLf & "Source: " & errLoop.Source _
               & vbCrLf & "SQL State: " & errLoop.SQLState & "; Native Error: " & errLoop.NativeError _
               & vbCrLf & vbCrLf & "Description: " & errLoop.Description & vbCrLf & vbCrLf & strHelp, "ADO Error", XP_OKOnly, msg_Critical
    Next
End Sub

Public Sub ShowVBError()
    On Error Resume Next
    Screen.MousePointer = 0
    If CBool(Err) Then
        'MsgBoxXP "VB Error '" & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & vbCrLf & "Description: " & Err.Description, "VB Runtime Error", XP_OKOnly, msg_Critical
        MessagePop RecLocekd, "System Info", "Cannot Process Your Request... " & vbCrLf & _
                                             "Please Try Again Or Rather Contact System Administrator..." & vbCrLf & _
                                             vbCrLf & _
                                           "  Module Name : " & MODULENAME & vbCrLf & _
                                           "  Ref No:" & Err.Number & vbCrLf & _
                                           "  Description: " & Err.Description, 4500, 1, 180
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
                           N2Str2Null(Err.Number) & " , " & _
                           N2Str2Null(Err.Description) & ")")
        Screen.MousePointer = 0
        'Call Shell("C:\DER.exe " & COMPANY_CODE & "~" & LOGCODE & "~" & App.TITLE & "~" & Formname & "~" & EventName & "~" & Err.Number & "~" & Err.Description & "~", vbNormalFocus)

        'DTVDP1=date time
        'WLDIOP=username
        'MWTCTT=modulename
        'WTCTTT=formname
        'WXOPOP=EventName
        'XOTNON=error number
        'TOCXXX=error description
        'Open "C:\error.txt" For Output As #1
        'Open "\\SERVER\DMIS 2.0\REPORTS\NSI_Submission.nsi" For Append As #1
        'Print #1, EncrypStr(Date & "," & App.TITLE & "," & Formname & "," & EventName & "," & Err.Description & "," & Err.Number, True)
        'Close #1
        Err.Clear
    End If

End Sub

Function EncrypStr(XXX, yyy As Boolean)
    Dim EncStr                                         As String    '
    Dim nard
    Dim MARK

    EncStr = ""
    If yyy = True Then
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
    'oVoice.Speak Msg, SVSFlagsAsync
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
    Dim Msg                                            As Long
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
    If Err Then isTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg                                            As Long
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
    Dim Msg                                            As Long
    On Error Resume Next
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If Err Then MakeOpaque = 2
End Function

Public Sub ChangeRes(ByVal iWidth As Single, ByVal iHeight As Single)
    Dim A                                              As Boolean
    Dim I&
    I = 0
    Do
        A = EnumDisplaySettings(0&, I&, DevM)
        I = I + 1
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
    '    If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
    '       Call ChangeRes(CurrentWidth, CurrentHeight)
    '  End If
    '  End
End Sub

Public Sub UnloadForm(frm As Object)
    Dim ShowCount                                      As Integer
    SetErrorMode SEM_NOGPFAULTERRORBOX
    'If ResolutionWidth <> CurrentWidth And ResolutionHeight <> CurrentHeight Then
    '   Call ChangeRes(CurrentWidth, CurrentHeight)
    'End If
    '    For ShowCount = 200 To 1 Step -50
    '      MakeTransparent frm.hwnd, ShowCount
    '     DoEvents
    '   Next
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
    Dim strFindThis As String, bContinueSearch         As Boolean
    Dim lResult As Long, lStart As Long, lLength       As Long
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
    Dim rsToFind                                       As ADODB.Recordset
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
    MessagePop RecSaveError, "Missing Filelds", "Field must have a Value!..." & Ricord, 1500
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
    Dim COUNTER                                        As Integer
    Dim TempNumber                                     As String
    Dim FoundPeriod                                    As Boolean
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
    Dim COUNTER                                        As Integer
    Dim NumericValue                                   As String
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
    Dim Indx                                           As Long
    Dim I                                              As Long
    TisoyView.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            Indx = TisoyView.Count + 1
            TisoyView.Add Indx, , IIf(IsNull(RecSet(0)), "", Trim(RecSet(0)))
            For I = 1 To RecSet.Fields.Count - 1
                TisoyView(Indx).ListSubItems.Add I, , IIf(IsNull(RecSet(I)), "", Trim(RecSet(I)))
            Next I
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

    'Used to validate password.
    'first, it checks for the existence of the username,
    'password, module requested

    If ApplySecurityValidation = False Then
        ValidPassword = True
        Exit Function
    End If

    Dim Conn                                           As ADODB.Connection
    Dim RS                                             As ADODB.Recordset
    Dim PasswdStr                                      As String
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
    Dim fieldCount                                     As Integer
    Dim rowCtr                                         As Long
    Dim colCtr                                         As Integer
    Dim cboRS                                          As ADODB.Recordset

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
    Dim DestDC, XPixels, YPixels, destX                As Long
    Dim destY, srcDC, SrcX, SrcY, RasterOp             As Long
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
    Dim FileNaeym                                      As String
    If LOGCODE <> "" And LOGDATE <> "" Then
        FileNaeym = "C:\" & App.EXEName & "_" & LOGCODE & "_" & Trim(STR(Month(LOGDATE))) & Trim(STR(Day(LOGDATE))) & Trim(STR(Year(LOGDATE))) & "_" & Left(STR(Time), 2) & Mid(STR(Time), 4, 2) & Mid(STR(Time), 7, 2) & Right(STR(Time), 2) & ".jpg"
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
    Dim MO                                             As String
    MO = "January  February March    April    May      June     July     August   SeptemberOctober  November December "
    Date2Month = Mid$(MO, (Month(Value) - 1) * 9 + 1, 9)
End Function
Public Function Module_Access(USERID As Long, SelectedModule As String, ModuleType) As Boolean
    Dim RS                                             As ADODB.Recordset
    Dim SQL                                            As String
    Dim checkRs                                        As ADODB.Recordset
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
            '            gconAudit.Execute ("insert into AUDIT_NOTINLIST(MOD_NAME,FUNC_NAME,USER_ID,USER_NAME) VALUES('" & SelectedModule & "','" & ModuleType & "'," & USERID & ", '" & LOGNAME & "')")

        End If
    End If
    Set RS = Nothing
End Function

Public Function AllowReprint(ModuleDescription As String)
    Dim RS                                             As ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT  COUNT(*)   FROM  ALL_Rams_UsersAcess  INNER JOIN ALL_Rams_Modules ON ALL_Rams_UsersAcess.MODULEID = ALL_Rams_Modules.MODULEID WHERE descriptions='" & ModuleDescription & "' AND ACESS_REPRINT=1 and mainmodulename='" & App.TITLE & "' AND USERID=" & LOGID)
    If RS.Fields(0).Value <> 0 Then
        AllowReprint = True
    Else
        AllowReprint = False
        MessagePop InfoVoid, "Re-Print Disabled!", ":: Reprinting of " & ModuleDescription & vbCrLf & ":: Please Contact Your Sys-Ad!"
    End If

End Function

Public Function Function_Access(USERID As Long, SelectedFeature As String, MODULENAME As String) As Boolean
    Dim SQL                                            As String
    Dim RS                                             As ADODB.Recordset
    Dim checkRs                                        As ADODB.Recordset
    Dim xSelectedFeature                               As String
    Dim xModuleType                                    As String
    Dim rsModuleType                                   As ADODB.Recordset

    SQL = "SELECT COUNT(MODULEID) FROM ALL_vW_USERACESS"
    SQL = SQL & " WHERE USERID=" & USERID
    SQL = SQL & " AND " & SelectedFeature & " = 1 "
    SQL = SQL & " AND MAINMODULENAME= '" & App.TITLE & "'"
    SQL = SQL & "  AND ltrim(rtrim(DESCRIPTIONS))='" & LTrim(RTrim(MODULENAME)) & "'"
    Set RS = gconDMIS.Execute(SQL)
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
        '    gconAudit.Execute ("insert into AUDIT_NOTINLIST(MOD_NAME,FUNC_NAME,USER_ID,USER_NAME) VALUES('" & MODULENAME & "','" & SelectedFeature & "'," & USERID & ", '" & LOGNAME & "')")
    Else
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
    Dim q As Double, r                                 As Double

    If XXX < 0 Or XXX > 1 Then
        Err.Raise vbObjectError, , "Inverse Function: Argument out of range."
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
    Dim temprs                                         As ADODB.Recordset
    Set temprs = gconDMIS.Execute("select MIN(NextTime) from CRIS_Reminders where EntityType='E' and SNOOZED=0 and  MONTH(nexttime)=MONTH(getdate()) and YEAR(nexttime)=YEAR(getdate()) and USERID=" & LOGID & "  and nexttime < = getdate()")

    If IsNull(temprs.Fields(0).Value) = False Then
        TIMER_REMIND = temprs.Fields(0).Value & ""
    Else
        TIMER_REMIND = xxTime
    End If
    Exit Sub

ADDER:
    Err.Clear
    Exit Sub
End Sub

Sub SetUserPathSettings()
    Dim CURRENT_REPORTS_PATH                           As String
    Dim plngSize                                       As Long
    Dim pstrBuffer                                     As String

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
    '    If CURRENT_REPORTS_PATH = "DEALER" Then
    '        'DEALERS SERVER PATH
    '        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HMH" Then
    '            AMIS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\AMIS\"
    '            CMIS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\CMIS\"
    '            CRIS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\CRIS\"
    '            CSMS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\CSMS\"
    '            HRMS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\HRMS\"
    '            OSMS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\OSMS\"
    '            SMIS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\SMIS\"
    '            PMIS_REPORT_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\PMIS\"
    '            HRMS_PICTURES_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\HRMS\images\"
    '        ElseIf COMPANY_CODE = "HAI" Then
    '            AMIS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\AMIS\"
    '            CMIS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\CMIS\"
    '            CRIS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\CRIS\"
    '            CSMS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\CSMS\"
    '            HRMS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\HRMS\"
    '            OSMS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\OSMS\"
    '            SMIS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\SMIS\"
    '            PMIS_REPORT_PATH = "\\SERVER\DMIS 2.0\REPORTS\PMIS\"
    '            HRMS_PICTURES_PATH = "\\SERVER\DMIS 2.0\REPORTS\HRMS\images\"
    '        ElseIf COMPANY_CODE = "HBK" Then
    '            AMIS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\AMIS\"
    '            CMIS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\CMIS\"
    '            CRIS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\CRIS\"
    '            CSMS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\CSMS\"
    '            HRMS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\HRMS\"
    '            OSMS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\OSMS\"
    '            SMIS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\SMIS\"
    '            PMIS_REPORT_PATH = "\\DMIS\DMIS 2.0\REPORTS\PMIS\"
    '            HRMS_PICTURES_PATH = "\\DMIS\DMIS 2.0\REPORTS\HRMS\images\"
    '        End If
    '    End If
    '    If CURRENT_REPORTS_PATH = "LOCAL" Then
    '        'PROGRAMMERS LOCAL PATH
    '        AMIS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\AMIS\"
    '        CMIS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\CMIS\"
    '        CRIS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\CRIS\"
    '        CSMS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\CSMS\"
    '        HRMS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\HRMS\"
    '        OSMS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\OSMS\"
    '        SMIS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\SMIS\"
    '        PMIS_REPORT_PATH = "D:\" & COMPANY_CODE & "\REPORTS\PMIS\"
    '        HRMS_PICTURES_PATH = "D:\" & COMPANY_CODE & "\REPORTS\HRMS\images\"
    '    End If
    '
    '    If CURRENT_REPORTS_PATH = "NSI_SERVER" Then
    '        AMIS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\AMIS\"
    '        CMIS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\CMIS\"
    '        CRIS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\CRIS\"
    '        CSMS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\CSMS\"
    '        HRMS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\HRMS\"
    '        OSMS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\OSMS\"
    '        SMIS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\SMIS\"
    '        PMIS_REPORT_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\PMIS\"
    '        HRMS_PICTURES_PATH = "\\SERVER\D\" & COMPANY_CODE & "\REPORTS\HRMS\images\"
    '    End If
End Sub

Sub SetCompanyProfile()
    MODULENAME = App.EXEName
    Dim rsProfile                                      As ADODB.Recordset
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
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim RSReprint                                      As New ADODB.Recordset
    Dim nard                                           As String

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
        '        With FrmReprintTransaction ' Look into AMIS folder AMIS form
        '            .LblTransactionNo = Xtransaction_no
        '            .lblTransaction_type = XApplication_type
        '            FrmReprintTransaction.Show 1
        '        End With
        REPRINT_CAPTION = "YES"
    End If
    Set RS = Nothing

    Exit Sub
RYAN:
    MsgBox Err.Description, "Error", "Please Contact NSI Administrator"
End Sub
Sub ReturnInvoiceNo(XVoucherno As String, Xjtype As String)
    ' update by BTT
    Dim RSHD                                           As New ADODB.Recordset
    Set RSHD = gconDMIS.Execute("Select voucherno,invoicetype,invoiceno,jtype from AMIS_journal_hd where voucherno ='" & XVoucherno & _
                                "' and jtype='" & Xjtype & "'")
    If Not (RSHD.EOF And RSHD.BOF) Then
        AMIS_Invoiceno = Null2String(RSHD!INVOICENO)
        AMIS_Invoicetype = Null2String(RSHD!InvoiceType)
    Else
        AMIS_Invoiceno = N2Str2Null("")
        AMIS_Invoicetype = N2Str2Null("")
    End If
    Set RSHD = Nothing
End Sub

Function UpdateBalanceSJ(CRJVoucherno As String, is_posted As Boolean) As Double
    Dim RSCRJ                                          As New ADODB.Recordset
    Dim TotalPayment                                   As Double
    Dim RSSJ                                           As New ADODB.Recordset
    Set RSCRJ = gconDMIS.Execute("SELECT * FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO ='" & CRJVoucherno & "'")
    If Not (RSCRJ.EOF And RSCRJ.BOF) Then
        Do While Not RSCRJ.EOF
            Set RSSJ = gconDMIS.Execute("Select voucherno,invoiceamt,invoiceno,invoicetype,balance from AMIS_journal_hd where jtype='SJ' and invoiceno='" & Null2String(RSCRJ!INVOICENO) & _
                                        "' and invoicetype='" & Null2String(RSCRJ!InvoiceType) & "'")
            If Not (RSSJ.EOF And RSSJ.BOF) Then
                If is_posted = True Then
                    UpdateBalanceSJ = NumericVal(RSSJ!BALANCE) - NumericVal(RSCRJ!invoiceamount)
                    gconDMIS.Execute ("Update amis_journal_hd set balance='" & NumericVal(UpdateBalanceSJ) & _
                                      "' where voucherno='" & RSSJ!VOUCHERNO & "' and jtype = 'SJ'")
                Else
                    gconDMIS.Execute ("Update amis_journal_hd set balance='" & NumericVal(RSSJ!INVOICEAMT) & _
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


Function EXTRACT_FILES(CUSTOMID As Long, FILENAME As String) As Boolean
    On Error GoTo Errorcode:
    '---------------------------------------
    '    FOR PARTS
    '---------------------------------------
    '    CHANGES IN MAC.xlt  = 102
    '    DPI.xlt             = 103
    '    MACMAC.xlt          = 104
    '    PartsRundown.xlt    = 105
    '    PO.xlt              = 106
    '    PQIR.xlt            = 107
    '---------------------------------------

    Dim b()                                            As Byte
    Dim s                                              As String
    Dim I                                              As Long

    Dim Temp                                           As String
    Dim StartPosition                                  As Long
    Dim mHandle                                        As Integer

    s = ""
    b = LoadResData(CUSTOMID, "CUSTOM")
    For I = 0 To UBound(b())
        s = s & Chr(b(I))
    Next I
    Erase b

    mHandle = FreeFile
    Open App.Path & "\" & FILENAME For Binary As #mHandle
    StartPosition = LOF(mHandle)
    Temp = s
    Put #mHandle, , Temp
    Put #mHandle, , StartPosition
    Close #mHandle
    EXTRACT_FILES = True
    Exit Function
Errorcode:
    Err.Clear
    EXTRACT_FILES = False
End Function

Sub FormExistsShow(FRMx As Form, Optional ByVal ismodal As Boolean)
    '    On Error GoTo ERRORCODE
    Dim m_Exists                                     As Boolean
    Dim frm                                          As Form
    For Each frm In Forms
        If (UCase(frm.Name) = UCase(FRMx.Name)) Then
            m_Exists = True
            Exit For
        End If
    Next
    Set frm = Nothing
    If m_Exists = True Then
        'frmx.WindowState = 0
        FRMx.ZOrder 0
    Else
    If ismodal = True Then
        FRMx.Show vbModal
    Else
        FRMx.Show
        'frmx.WindowState = 0
        FRMx.ZOrder 0
    End If
        
        

    End If

    Exit Sub
Errorcode:
    Err.Clear
End Sub

Function CheckIfRoIsAlreadyInvoice(XXX As String) As Boolean
    Dim RSTMP                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Not Null2String(RSTMP!INVOICE) = "" Then
            CheckIfRoIsAlreadyInvoice = True
        End If
    End If
    Set RSTMP = Nothing
End Function

Function CheckIfROStillExist(XXX As String) As Boolean
    Dim RSTMP                           As New ADODB.Recordset
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

Function CheckIfTheJobIsFinish(xRONO As String, xJOBCODE As String) As String
    Dim RSTMP                                   As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DONE FROM CSMS_RODET " & _
        " WHERE LIVIL = 1 " & _
        " AND DETCODE = " & N2Str2Null(xJOBCODE) & _
        " AND REP_OR = " & N2Str2Null(xRONO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!DONE) = "Y" Then
            CheckIfTheJobIsFinish = "Finish"
        Else
            CheckIfTheJobIsFinish = "Not Finish"
        End If
    End If
    Set RSTMP = Nothing
End Function

Function CheckIfAppointmentTimeIsAvailable(XDATE As String, xtime As String) As String
    Dim RSTMP                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CUSNAM FROM CSMS_APPOINTMENT WHERE TRANDATE = " & N2Str2Null(XDATE) & _
        " AND APPTTIME = " & N2Str2Null(xtime) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfAppointmentTimeIsAvailable = Null2String(RSTMP!CUSNAM)
    End If
    Set RSTMP = Nothing
End Function

Function CheckAppointmentStatus(XDATE As String, xtime As String) As String
    Dim RSTMP                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_APPOINTMENT WHERE TRANDATE = " & N2Str2Null(XDATE) & _
        " AND APPTTIME = " & N2Str2Null(xtime) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckAppointmentStatus = Null2String(RSTMP!Status)
    End If
    Set RSTMP = Nothing
End Function

Function CheckEstimateStatus(xESTNO As String) As String
    Dim RSTMP                          As New ADODB.Recordset
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

Function GetFreshServiceCounterStatus(xRONO As String) As String
    Dim RSTMP           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT STATUS FROM CSMS_REPAIRORDER WHERE RO_NO = " & N2Str2Null(xRONO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetFreshServiceCounterStatus = LTrim(RTrim(Null2String(RSTMP!Status)))
    End If
    Set RSTMP = Nothing
End Function
