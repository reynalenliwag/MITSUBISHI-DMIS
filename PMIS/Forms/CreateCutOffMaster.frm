VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMIS_Physical_CreateCutOffMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Cut-Off Master File"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CreateCutOffMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5730
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
      Left            =   4620
      MouseIcon       =   "CreateCutOffMaster.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "CreateCutOffMaster.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit Window"
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
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
      Left            =   3720
      MouseIcon       =   "CreateCutOffMaster.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "CreateCutOffMaster.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Create Cut-Off Master File"
      Top             =   1200
      Width           =   915
   End
   Begin VB.CheckBox chkLastMonth 
      Caption         =   "Last Month Cut-Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   -180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   30
      Width           =   5715
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         ScaleHeight     =   195
         ScaleWidth      =   5325
         TabIndex        =   2
         Top             =   750
         Width           =   5325
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
            Height          =   195
            Left            =   60
            TabIndex        =   3
            ToolTipText     =   "Process progress"
            Top             =   0
            Width           =   5265
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   5565
         TabIndex        =   4
         Top             =   660
         Width           =   5565
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   5
            Top             =   0
            Width           =   5475
            _ExtentX        =   9657
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
            MICON           =   "CreateCutOffMaster.frx":0C59
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "CreateCutOffMaster.frx":0C75
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "CreateCutOffMaster.frx":0C91
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
      Begin VB.Label labCPB 
         BackColor       =   &H00DEDFDE&
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
         TabIndex        =   7
         Top             =   30
         Width           =   5595
      End
   End
End
Attribute VB_Name = "frmPMIS_Physical_CreateCutOffMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Sub CreateCutOffMasterLastMonth()
    Dim MoveSql                                        As String
    Dim i                                              As Long

    Dim varPmasID                                      As String
    Dim varPmasSTOCKNO                                 As String
    Dim varPmasSTOCKDESC                               As String
    Dim varPmasINVCLASS                                As String
    Dim varPmasVEHTYPE                                 As String
    Dim varPmasMODELCODE                               As String
    Dim varPmasLOCATION                                As String
    Dim varPmasMAC                                     As Double
    Dim varPmasMAD                                     As Long
    Dim varPmasOLDNO                                   As String
    Dim varPmasNEWNO                                   As String
    Dim varPmasGENNO                                   As String
    Dim varPmasSRP                                     As Double
    Dim varPmasNOSHIP                                  As Double
    Dim varPmasLASTM_MAC                               As Double
    Dim varPmasLASTM_MAD                               As Double
    Dim varPmasLASTM_SELL                              As Double
    Dim varPmasLASTM_OH                                As Long
    Dim varPmasLASTM_OO                                As Long
    Dim varPmasOnhand                                  As Long
    Dim varPmasTrecqty                                 As Double
    Dim varPmasTISSQTY                                 As Double
    Dim varPmasOnOrder                                 As Long
    Dim varPmasTpoqty                                  As Long
    Dim varPmasPRQTY                                   As Long
    Dim varPmasTPRQTY                                  As Long
    Dim varPmasLAST_RECQ                               As Long
    Dim varPmasLAST_RECD                               As String
    Dim varPmasLASTY_OH                                As Long
    Dim varPmasLASTY_MAC                               As Double
    Dim varPmasLASTY_OO                                As Long
    Dim varPmasLASTY_ADJ                               As Long
    Dim varPmasHOLD                                    As Long
    Dim varPmasSUPCODE                                 As String
    Dim varPmasVARIANCE                                As Long
    Dim varPmasSUBINVCLASS                             As String
    Dim varPmasPHYCOUNT                                As Long
    Dim varPmasADJPHYCOUNT                             As Long
    Dim varPmasCUTOFFQTY                               As Long
    Dim varPmasCUTOFFMAC                               As Double
    Dim varPmasRECEIPTS                                As Long
    Dim varPmasISSUANCES                               As Long
    Dim varPmasUSERCODE                                As String
    Dim varPmasLASTUPDATE                              As String
    Dim varPmasDNP                                     As Double
    Dim varPmasVALID_ICC                               As String
    Dim varPmasDATE_ENTERED                            As String

    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select * from PMIS_STOCKMAS WHERE [TYPE] = '" & C_TYPE & "' order by STOCKNO asc", gconDMIS

    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        MsgSpeech "Creating Cut Off Master File"
        Me.Caption = "Creating Cut Off Master File"
        Screen.MousePointer = 11
        DoEvents
        i = 0
        gconINVENTORY.Execute ("delete * from CUTOFF")
        Do While Not RSPARTMAS.EOF
            varPmasID = i + 1
            labProcessing.Caption = "Processing Part Number: " & Null2String(RSPARTMAS!STOCKNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(RSPARTMAS!STOCKNO)
            varPmasSTOCKDESC = N2Str2Null(RSPARTMAS!STOCKDESC)
            varPmasINVCLASS = N2Str2Null(RSPARTMAS!InvClass)
            varPmasMAD = N2Str2IntZero(RSPARTMAS!mad)
            varPmasVEHTYPE = N2Str2Null(RSPARTMAS!vehtype)
            varPmasMODELCODE = N2Str2Null(RSPARTMAS!MODELCODE)
            varPmasLOCATION = N2Str2Null(RSPARTMAS!Location)
            varPmasMAC = N2Str2Zero(RSPARTMAS!MAC)
            varPmasOLDNO = N2Str2Null(RSPARTMAS!oldno)
            varPmasNEWNO = N2Str2Null(RSPARTMAS!NEWNO)
            varPmasGENNO = N2Str2Null(RSPARTMAS!GENNO)
            varPmasSRP = N2Str2Zero(RSPARTMAS!SRP)
            varPmasNOSHIP = N2Str2Zero(RSPARTMAS!NOSHIP)
            varPmasLASTM_MAC = N2Str2Zero(RSPARTMAS!LASTM_MAC)
            varPmasLASTM_MAD = N2Str2Zero(RSPARTMAS!LASTM_MAD)
            varPmasLASTM_SELL = N2Str2Zero(RSPARTMAS!LASTM_SELL)
            varPmasLASTM_OH = N2Str2IntZero(RSPARTMAS!LASTM_OH)
            varPmasLASTM_OO = N2Str2IntZero(RSPARTMAS!LASTM_OO)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(RSPARTMAS!ONHAND)
            varPmasTrecqty = N2Str2IntZero(RSPARTMAS!TRECQTY)
            varPmasTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            varPmasOnOrder = N2Str2IntZero(RSPARTMAS!ONORDER)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(RSPARTMAS!tpoqty)
            varPmasPRQTY = N2Str2IntZero(RSPARTMAS!PRQTY)
            varPmasTPRQTY = N2Str2IntZero(RSPARTMAS!TPRQTY)
            varPmasLAST_RECQ = N2Str2IntZero(RSPARTMAS!last_recq)
            varPmasLAST_RECD = N2Date2Null(RSPARTMAS!LAST_RECD)
            varPmasLASTY_OH = N2Str2IntZero(RSPARTMAS!LASTY_OH)
            varPmasLASTY_MAC = N2Str2Zero(RSPARTMAS!LASTY_MAC)
            varPmasLASTY_OO = N2Str2IntZero(RSPARTMAS!LASTY_OO)
            varPmasLASTY_ADJ = N2Str2IntZero(RSPARTMAS!LASTY_ADJ)
            varPmasHOLD = N2Str2IntZero(RSPARTMAS!hold)
            varPmasSUPCODE = N2Str2Null(RSPARTMAS!SupCode)
            varPmasVARIANCE = N2Str2IntZero(RSPARTMAS!VARIANCE)
            varPmasSUBINVCLASS = N2Str2Null(RSPARTMAS!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(RSPARTMAS!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(RSPARTMAS!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(RSPARTMAS!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(RSPARTMAS!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(RSPARTMAS!RECEIPTS)
            varPmasISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)
            varPmasUSERCODE = N2Str2Null(RSPARTMAS!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(RSPARTMAS!LASTUPDATE)
            varPmasDNP = N2Str2Zero(RSPARTMAS!dnp)
            varPmasVALID_ICC = N2Str2Null(RSPARTMAS!VALID_ICC)
            varPmasDATE_ENTERED = N2Str2Null(RSPARTMAS!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,STOCKNO,STOCKDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                On Error GoTo ErrorCode
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                               " CUTOFFQTY = " & N2Str2IntZero(RSPARTMAS!ONHAND) & "," & _
                               " CUTOFFMAC =" & N2Str2Zero(RSPARTMAS!MAC) & _
                               " WHERE STOCKNO = " & varPmasSTOCKNO
            End If
            i = i + 1
            progCPB.Value = (i / RSPARTMAS.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSPARTMAS.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Cut-Off Master File Successfully Created"
        Me.Caption = "Cut-Off Master File Successfully Created"
        Screen.MousePointer = 0
        DoEvents
    End If
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub CreateCutOffMaster()
    Dim MoveSql                                        As String
    Dim i                                              As Long

    Dim varPmasID                                      As String
    Dim varPmasSTOCKNO                                 As String
    Dim varPmasSTOCKDESC                               As String
    Dim varPmasINVCLASS                                As String
    Dim varPmasVEHTYPE                                 As String
    Dim varPmasMODELCODE                               As String
    Dim varPmasLOCATION                                As String
    Dim varPmasMAC                                     As Double
    Dim varPmasMAD                                     As Long
    Dim varPmasOLDNO                                   As String
    Dim varPmasNEWNO                                   As String
    Dim varPmasGENNO                                   As String
    Dim varPmasSRP                                     As Double
    Dim varPmasNOSHIP                                  As Double
    Dim varPmasLASTM_MAC                               As Double
    Dim varPmasLASTM_MAD                               As Double
    Dim varPmasLASTM_SELL                              As Double
    Dim varPmasLASTM_OH                                As Long
    Dim varPmasLASTM_OO                                As Long
    Dim varPmasOnhand                                  As Long
    Dim varPmasTrecqty                                 As Double
    Dim varPmasTISSQTY                                 As Double
    Dim varPmasOnOrder                                 As Long
    Dim varPmasTpoqty                                  As Long
    Dim varPmasPRQTY                                   As Long
    Dim varPmasTPRQTY                                  As Long
    Dim varPmasLAST_RECQ                               As Long
    Dim varPmasLAST_RECD                               As String
    Dim varPmasLASTY_OH                                As Long
    Dim varPmasLASTY_MAC                               As Double
    Dim varPmasLASTY_OO                                As Long
    Dim varPmasLASTY_ADJ                               As Long
    Dim varPmasHOLD                                    As Long
    Dim varPmasSUPCODE                                 As String
    Dim varPmasVARIANCE                                As Long
    Dim varPmasSUBINVCLASS                             As String
    Dim varPmasPHYCOUNT                                As Long
    Dim varPmasADJPHYCOUNT                             As Long
    Dim varPmasCUTOFFQTY                               As Long
    Dim varPmasCUTOFFMAC                               As Double
    Dim varPmasRECEIPTS                                As Long
    Dim varPmasISSUANCES                               As Long
    Dim varPmasUSERCODE                                As String
    Dim varPmasLASTUPDATE                              As String
    Dim varPmasDNP                                     As Double
    Dim varPmasVALID_ICC                               As String
    Dim varPmasDATE_ENTERED                            As String
    Dim RCOUNT                                         As Long
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select * from PMIS_STOCKMAS WHERE [TYPE] = '" & C_TYPE & "' AND ACTIVE = 'Y' order by STOCKNO asc", gconDMIS, adOpenKeyset, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        MsgSpeech "Creating Cut-Off Master File"
        Me.Caption = "Creating Cut-Off Master File"
        Screen.MousePointer = 11
        DoEvents
        i = 0
        gconINVENTORY.Execute "delete * from CUTOFF"
        RCOUNT = RSPARTMAS.RecordCount
        Do While Not RSPARTMAS.EOF
            varPmasID = i + 1
            labProcessing.Caption = "PROCESSING " & DESC_TYPE & " NUMBER: " & Null2String(RSPARTMAS!STOCKNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(RSPARTMAS!STOCKNO)
            varPmasSTOCKDESC = N2Str2Null(RSPARTMAS!STOCKDESC)
            varPmasINVCLASS = N2Str2Null(RSPARTMAS!InvClass)
            varPmasMAD = N2Str2IntZero(RSPARTMAS!mad)
            varPmasVEHTYPE = N2Str2Null(RSPARTMAS!vehtype)
            varPmasMODELCODE = N2Str2Null(RSPARTMAS!MODELCODE)
            varPmasLOCATION = N2Str2Null(RSPARTMAS!Location)
            varPmasMAC = N2Str2Zero(RSPARTMAS!MAC)
            varPmasOLDNO = N2Str2Null(RSPARTMAS!oldno)
            varPmasNEWNO = N2Str2Null(RSPARTMAS!NEWNO)
            varPmasGENNO = N2Str2Null(RSPARTMAS!GENNO)
            varPmasSRP = N2Str2Zero(RSPARTMAS!SRP)
            varPmasNOSHIP = N2Str2Zero(RSPARTMAS!NOSHIP)
            varPmasLASTM_MAC = N2Str2Zero(RSPARTMAS!LASTM_MAC)
            varPmasLASTM_MAD = N2Str2Zero(RSPARTMAS!LASTM_MAD)
            varPmasLASTM_SELL = N2Str2Zero(RSPARTMAS!LASTM_SELL)
            varPmasLASTM_OH = N2Str2IntZero(RSPARTMAS!LASTM_OH)
            varPmasLASTM_OO = N2Str2IntZero(RSPARTMAS!LASTM_OO)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(RSPARTMAS!ONHAND)
            varPmasTrecqty = N2Str2IntZero(RSPARTMAS!TRECQTY)
            varPmasTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            varPmasOnOrder = N2Str2IntZero(RSPARTMAS!ONORDER)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(RSPARTMAS!tpoqty)
            varPmasPRQTY = N2Str2IntZero(RSPARTMAS!PRQTY)
            varPmasTPRQTY = N2Str2IntZero(RSPARTMAS!TPRQTY)
            varPmasLAST_RECQ = N2Str2IntZero(RSPARTMAS!last_recq)
            varPmasLAST_RECD = N2Date2Null(RSPARTMAS!LAST_RECD)
            varPmasLASTY_OH = N2Str2IntZero(RSPARTMAS!LASTY_OH)
            varPmasLASTY_MAC = N2Str2Zero(RSPARTMAS!LASTY_MAC)
            varPmasLASTY_OO = N2Str2IntZero(RSPARTMAS!LASTY_OO)
            varPmasLASTY_ADJ = N2Str2IntZero(RSPARTMAS!LASTY_ADJ)
            varPmasSUPCODE = N2Str2Null(RSPARTMAS!SupCode)
            varPmasVARIANCE = N2Str2IntZero(RSPARTMAS!VARIANCE)
            varPmasSUBINVCLASS = N2Str2Null(RSPARTMAS!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(RSPARTMAS!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(RSPARTMAS!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(RSPARTMAS!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(RSPARTMAS!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(RSPARTMAS!RECEIPTS)
            varPmasISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)
            varPmasUSERCODE = N2Str2Null(RSPARTMAS!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(RSPARTMAS!LASTUPDATE)
            varPmasDNP = N2Str2Zero(RSPARTMAS!dnp)
            varPmasVALID_ICC = N2Str2Null(RSPARTMAS!VALID_ICC)
            varPmasDATE_ENTERED = N2Str2Null(RSPARTMAS!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,PARTNO,PARTDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                               " CUTOFFQTY = " & N2Str2IntZero(RSPARTMAS!ONHAND) & "," & _
                               " CUTOFFMAC =" & N2Str2Zero(RSPARTMAS!MAC) & _
                               " WHERE STOCKNO = " & varPmasSTOCKNO
            End If
            i = i + 1
            progCPB.Value = (i / RCOUNT) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            RSPARTMAS.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Cut-Off Master File Successfully Created"
        Me.Caption = "Cut-Off Master File Successfully Created"
        MsgBox "Cut-Off Master File Successfully Created", vbInformation
        Screen.MousePointer = 0
        DoEvents
    End If
    Exit Sub

    Screen.MousePointer = 0
    ShowVBError
    Resume Next
End Sub

Private Sub cmdCreate_Click()
    cmdCreate.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    CreateCutOffMaster
    NEW_LogAudit "G", "PHYSICAL COUNT", "", "", "", "", "CUT-OFF MASTER FILE", ""
    cmdExit.Enabled = True
    DoEvents
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMIS_Physical_CreateCutOffMaster = Nothing
    UnloadForm Me
End Sub

