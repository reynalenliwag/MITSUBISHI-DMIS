VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISCreateCutOffMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Cut-Off Master File"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Accessories_CreateCutOffMaster.frx":0000
   LinkTopic       =   "Form1"
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
      MouseIcon       =   "Accessories_CreateCutOffMaster.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Accessories_CreateCutOffMaster.frx":045C
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
      MouseIcon       =   "Accessories_CreateCutOffMaster.frx":07C2
      MousePointer    =   99  'Custom
      Picture         =   "Accessories_CreateCutOffMaster.frx":0914
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
            MICON           =   "Accessories_CreateCutOffMaster.frx":0C59
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
         Picture         =   "Accessories_CreateCutOffMaster.frx":0C75
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "Accessories_CreateCutOffMaster.frx":0C91
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
Attribute VB_Name = "frmPMISCreateCutOffMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Sub CreateCutOffMasterLastMonth()
    Dim MoveSql                                                       As String
    Dim I                                                             As Integer

    Dim varPmasID                                                     As String
    Dim varPmasSTOCKNO                                                As String
    Dim varPmasSTOCKDESC                                              As String
    Dim varPmasINVCLASS                                               As String
    Dim varPmasVEHTYPE                                                As String
    Dim varPmasMODELCODE                                              As String
    Dim varPmasLOCATION                                               As String
    Dim varPmasMAC                                                    As Double
    Dim varPmasMAD                                                    As Integer
    Dim varPmasOLDNO                                                  As String
    Dim varPmasNEWNO                                                  As String
    Dim varPmasGENNO                                                  As String
    Dim varPmasSRP                                                    As Double
    Dim varPmasNOSHIP                                                 As Double
    Dim varPmasLASTM_MAC                                              As Double
    Dim varPmasLASTM_MAD                                              As Double
    Dim varPmasLASTM_SELL                                             As Double
    Dim varPmasLASTM_OH                                               As Integer
    Dim varPmasLASTM_OO                                               As Integer
    Dim varPmasOnhand                                                 As Integer
    Dim varPmasTrecqty                                                As Double
    Dim varPmasTISSQTY                                                As Double
    Dim varPmasOnOrder                                                As Integer
    Dim varPmasTpoqty                                                 As Integer
    Dim varPmasPRQTY                                                  As Integer
    Dim varPmasTPRQTY                                                 As Integer
    Dim varPmasLAST_RECQ                                              As Integer
    Dim varPmasLAST_RECD                                              As String
    Dim varPmasLASTY_OH                                               As Integer
    Dim varPmasLASTY_MAC                                              As Double
    Dim varPmasLASTY_OO                                               As Integer
    Dim varPmasLASTY_ADJ                                              As Integer
    Dim varPmasHOLD                                                   As Integer
    Dim varPmasSUPCODE                                                As String
    Dim varPmasVARIANCE                                               As Integer
    Dim varPmasSUBINVCLASS                                            As String
    Dim varPmasPHYCOUNT                                               As Integer
    Dim varPmasADJPHYCOUNT                                            As Integer
    Dim varPmasCUTOFFQTY                                              As Integer
    Dim varPmasCUTOFFMAC                                              As Double
    Dim varPmasRECEIPTS                                               As Integer
    Dim varPmasISSUANCES                                              As Integer
    Dim varPmasUSERCODE                                               As String
    Dim varPmasLASTUPDATE                                             As String
    Dim varPmasDNP                                                    As Double
    Dim varPmasVALID_ICC                                              As String
    Dim varPmasDATE_ENTERED                                           As String

    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from PMIS_STOCKMAS WHERE [TYPE] = 'P' order by STOCKNO asc", gconDMIS
    
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        MsgSpeech "Creating Cut Off Master File"
        Me.Caption = "Creating Cut Off Master File"
        Screen.MousePointer = 11
        DoEvents
        I = 0
        gconINVENTORY.Execute ("delete * from CUTOFF")
        Do While Not rsPartMas.EOF
            varPmasID = I + 1
            labProcessing.Caption = "Processing Part Number: " & Null2String(rsPartMas!STOCKNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(rsPartMas!STOCKNO)
            varPmasSTOCKDESC = N2Str2Null(rsPartMas!STOCKDESC)
            varPmasINVCLASS = N2Str2Null(rsPartMas!InvClass)
            varPmasMAD = N2Str2IntZero(rsPartMas!mad)
            varPmasVEHTYPE = N2Str2Null(rsPartMas!vehtype)
            varPmasMODELCODE = N2Str2Null(rsPartMas!modelcode)
            varPmasLOCATION = N2Str2Null(rsPartMas!Location)
            varPmasMAC = N2Str2Zero(rsPartMas!Mac)
            varPmasOLDNO = N2Str2Null(rsPartMas!oldno)
            varPmasNEWNO = N2Str2Null(rsPartMas!newno)
            varPmasGENNO = N2Str2Null(rsPartMas!genno)
            varPmasSRP = N2Str2Zero(rsPartMas!SRP)
            varPmasNOSHIP = N2Str2Zero(rsPartMas!noship)
            varPmasLASTM_MAC = N2Str2Zero(rsPartMas!lastm_mac)
            varPmasLASTM_MAD = N2Str2Zero(rsPartMas!lastm_mad)
            varPmasLASTM_SELL = N2Str2Zero(rsPartMas!lastm_sell)
            varPmasLASTM_OH = N2Str2IntZero(rsPartMas!lastm_oh)
            varPmasLASTM_OO = N2Str2IntZero(rsPartMas!lastm_oo)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(rsPartMas!ONHAND)
            varPmasTrecqty = N2Str2IntZero(rsPartMas!trecqty)
            varPmasTISSQTY = N2Str2IntZero(rsPartMas!tissqty)
            varPmasOnOrder = N2Str2IntZero(rsPartMas!onorder)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(rsPartMas!tpoqty)
            varPmasPRQTY = N2Str2IntZero(rsPartMas!prqty)
            varPmasTPRQTY = N2Str2IntZero(rsPartMas!tprqty)
            varPmasLAST_RECQ = N2Str2IntZero(rsPartMas!last_recq)
            varPmasLAST_RECD = N2Date2Null(rsPartMas!last_recd)
            varPmasLASTY_OH = N2Str2IntZero(rsPartMas!lasty_oh)
            varPmasLASTY_MAC = N2Str2Zero(rsPartMas!lasty_mac)
            varPmasLASTY_OO = N2Str2IntZero(rsPartMas!lasty_oo)
            varPmasLASTY_ADJ = N2Str2IntZero(rsPartMas!lasty_adj)
            varPmasHOLD = N2Str2IntZero(rsPartMas!hold)
            varPmasSUPCODE = N2Str2Null(rsPartMas!SupCode)
            varPmasVARIANCE = N2Str2IntZero(rsPartMas!variance)
            varPmasSUBINVCLASS = N2Str2Null(rsPartMas!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(rsPartMas!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(rsPartMas!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(rsPartMas!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(rsPartMas!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(rsPartMas!receipts)
            varPmasISSUANCES = N2Str2IntZero(rsPartMas!issuances)
            varPmasUSERCODE = N2Str2Null(rsPartMas!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(rsPartMas!lastupdate)
            varPmasDNP = N2Str2Zero(rsPartMas!DNP)
            varPmasVALID_ICC = N2Str2Null(rsPartMas!valid_icc)
            varPmasDATE_ENTERED = N2Str2Null(rsPartMas!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,STOCKNO,STOCKDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                On Error GoTo ERRORCODE
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update PMIS_STOCKMAS set " & _
                               " CUTOFFQTY = " & N2Str2IntZero(rsPartMas!ONHAND) & "," & _
                               " CUTOFFMAC =" & N2Str2Zero(rsPartMas!Mac) & _
                               " WHERE STOCKNO = " & varPmasSTOCKNO
            End If
            I = I + 1
            progCPB.Value = (I / rsPartMas.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsPartMas.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Cut-Off Master File Successfully Created"
        Me.Caption = "Cut-Off Master File Successfully Created"
        Screen.MousePointer = 0
        DoEvents
    End If
    Exit Sub

ERRORCODE:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub CreateCutOffMaster()
    Dim MoveSql                                                       As String
    Dim I                                                             As Integer

    Dim varPmasID                                                     As String
    Dim varPmasSTOCKNO                                                As String
    Dim varPmasSTOCKDESC                                              As String
    Dim varPmasINVCLASS                                               As String
    Dim varPmasVEHTYPE                                                As String
    Dim varPmasMODELCODE                                              As String
    Dim varPmasLOCATION                                               As String
    Dim varPmasMAC                                                    As Double
    Dim varPmasMAD                                                    As Integer
    Dim varPmasOLDNO                                                  As String
    Dim varPmasNEWNO                                                  As String
    Dim varPmasGENNO                                                  As String
    Dim varPmasSRP                                                    As Double
    Dim varPmasNOSHIP                                                 As Double
    Dim varPmasLASTM_MAC                                              As Double
    Dim varPmasLASTM_MAD                                              As Double
    Dim varPmasLASTM_SELL                                             As Double
    Dim varPmasLASTM_OH                                               As Integer
    Dim varPmasLASTM_OO                                               As Integer
    Dim varPmasOnhand                                                 As Integer
    Dim varPmasTrecqty                                                As Double
    Dim varPmasTISSQTY                                                As Double
    Dim varPmasOnOrder                                                As Integer
    Dim varPmasTpoqty                                                 As Integer
    Dim varPmasPRQTY                                                  As Integer
    Dim varPmasTPRQTY                                                 As Integer
    Dim varPmasLAST_RECQ                                              As Integer
    Dim varPmasLAST_RECD                                              As String
    Dim varPmasLASTY_OH                                               As Integer
    Dim varPmasLASTY_MAC                                              As Double
    Dim varPmasLASTY_OO                                               As Integer
    Dim varPmasLASTY_ADJ                                              As Integer
    Dim varPmasHOLD                                                   As Integer
    Dim varPmasSUPCODE                                                As String
    Dim varPmasVARIANCE                                               As Integer
    Dim varPmasSUBINVCLASS                                            As String
    Dim varPmasPHYCOUNT                                               As Integer
    Dim varPmasADJPHYCOUNT                                            As Integer
    Dim varPmasCUTOFFQTY                                              As Integer
    Dim varPmasCUTOFFMAC                                              As Double
    Dim varPmasRECEIPTS                                               As Integer
    Dim varPmasISSUANCES                                              As Integer
    Dim varPmasUSERCODE                                               As String
    Dim varPmasLASTUPDATE                                             As String
    Dim varPmasDNP                                                    As Double
    Dim varPmasVALID_ICC                                              As String
    Dim varPmasDATE_ENTERED                                           As String
    Dim RCOUNT                                                        As Long
    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from PMIS_PARTMAS WHERE [TYPE] = 'P' AND ACTIVE = 'Y' order by PARTNO asc", gconDMIS, adOpenKeyset, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        MsgSpeech "Creating Cut-Off Master File"
        Me.Caption = "Creating Cut-Off Master File"
        Screen.MousePointer = 11
        DoEvents
        I = 0
        gconINVENTORY.Execute "delete * from CUTOFF"
        RCOUNT = rsPartMas.RecordCount
        Do While Not rsPartMas.EOF
            varPmasID = I + 1
            labProcessing.Caption = "Processing Part Number: " & Null2String(rsPartMas!PARTNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(rsPartMas!PARTNO)
            varPmasSTOCKDESC = N2Str2Null(rsPartMas!PARTDESC)
            varPmasINVCLASS = N2Str2Null(rsPartMas!InvClass)
            varPmasMAD = N2Str2IntZero(rsPartMas!mad)
            varPmasVEHTYPE = N2Str2Null(rsPartMas!vehtype)
            varPmasMODELCODE = N2Str2Null(rsPartMas!modelcode)
            varPmasLOCATION = N2Str2Null(rsPartMas!Location)
            varPmasMAC = N2Str2Zero(rsPartMas!Mac)
            varPmasOLDNO = N2Str2Null(rsPartMas!oldno)
            varPmasNEWNO = N2Str2Null(rsPartMas!newno)
            varPmasGENNO = N2Str2Null(rsPartMas!genno)
            varPmasSRP = N2Str2Zero(rsPartMas!SRP)
            varPmasNOSHIP = N2Str2Zero(rsPartMas!noship)
            varPmasLASTM_MAC = N2Str2Zero(rsPartMas!lastm_mac)
            varPmasLASTM_MAD = N2Str2Zero(rsPartMas!lastm_mad)
            varPmasLASTM_SELL = N2Str2Zero(rsPartMas!lastm_sell)
            varPmasLASTM_OH = N2Str2IntZero(rsPartMas!lastm_oh)
            varPmasLASTM_OO = N2Str2IntZero(rsPartMas!lastm_oo)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(rsPartMas!ONHAND)
            varPmasTrecqty = N2Str2IntZero(rsPartMas!trecqty)
            varPmasTISSQTY = N2Str2IntZero(rsPartMas!tissqty)
            varPmasOnOrder = N2Str2IntZero(rsPartMas!onorder)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(rsPartMas!tpoqty)
            varPmasPRQTY = N2Str2IntZero(rsPartMas!prqty)
            varPmasTPRQTY = N2Str2IntZero(rsPartMas!tprqty)
            varPmasLAST_RECQ = N2Str2IntZero(rsPartMas!last_recq)
            varPmasLAST_RECD = N2Date2Null(rsPartMas!last_recd)
            varPmasLASTY_OH = N2Str2IntZero(rsPartMas!lasty_oh)
            varPmasLASTY_MAC = N2Str2Zero(rsPartMas!lasty_mac)
            varPmasLASTY_OO = N2Str2IntZero(rsPartMas!lasty_oo)
            varPmasLASTY_ADJ = N2Str2IntZero(rsPartMas!lasty_adj)
            varPmasSUPCODE = N2Str2Null(rsPartMas!SupCode)
            varPmasVARIANCE = N2Str2IntZero(rsPartMas!variance)
            varPmasSUBINVCLASS = N2Str2Null(rsPartMas!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(rsPartMas!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(rsPartMas!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(rsPartMas!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(rsPartMas!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(rsPartMas!receipts)
            varPmasISSUANCES = N2Str2IntZero(rsPartMas!issuances)
            varPmasUSERCODE = N2Str2Null(rsPartMas!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(rsPartMas!lastupdate)
            varPmasDNP = N2Str2Zero(rsPartMas!DNP)
            varPmasVALID_ICC = N2Str2Null(rsPartMas!valid_icc)
            varPmasDATE_ENTERED = N2Str2Null(rsPartMas!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,PARTNO,PARTDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update PMIS_PARTMAS set " & _
                               " CUTOFFQTY = " & N2Str2IntZero(rsPartMas!ONHAND) & "," & _
                               " CUTOFFMAC =" & N2Str2Zero(rsPartMas!Mac) & _
                               " WHERE PARTNO = " & varPmasSTOCKNO
            End If
            I = I + 1
            progCPB.Value = (I / RCOUNT) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsPartMas.MoveNext
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
    Set frmPMISCreateCutOffMaster = Nothing
    UnloadForm Me
End Sub

