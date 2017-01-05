VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISMAT_CreateCutOffMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Cut-Off Master File"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_CreateCutOffMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5880
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
      Left            =   4800
      MouseIcon       =   "MAT_CreateCutOffMaster.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "MAT_CreateCutOffMaster.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
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
      Left            =   3900
      MouseIcon       =   "MAT_CreateCutOffMaster.frx":07C2
      Picture         =   "MAT_CreateCutOffMaster.frx":0ACC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1110
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
         Left            =   90
         ScaleHeight     =   195
         ScaleWidth      =   3615
         TabIndex        =   2
         Top             =   750
         Width           =   3615
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
            Height          =   225
            Left            =   60
            TabIndex        =   3
            ToolTipText     =   "Process progress"
            Top             =   -30
            Width           =   3525
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   3765
         TabIndex        =   4
         Top             =   660
         Width           =   3765
         Begin wizButton.cmd cmd1 
            Height          =   345
            Left            =   30
            TabIndex        =   5
            Top             =   0
            Width           =   3705
            _ExtentX        =   6535
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
            MICON           =   "MAT_CreateCutOffMaster.frx":0E11
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
         Picture         =   "MAT_CreateCutOffMaster.frx":0E2D
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "MAT_CreateCutOffMaster.frx":0E49
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
Attribute VB_Name = "frmPMISMAT_CreateCutOffMaster"
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
    rsPartMas.Open "select * from CSMS_MATMAS order by STOCKNO asc", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        MsgSpeech "Creating Cut Off Master File"
        Me.Caption = "Creating Cut Off Master File"
        Screen.MousePointer = 11
        DoEvents
        I = 0
        gconINVENTORY.Execute ("delete * from CUTOFF")
        Do While Not rsPartMas.EOF
            varPmasID = I + 1
            labProcessing.Caption = "Processing Material Code: " & Null2String(rsPartMas!STOCKNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(rsPartMas!STOCKNO)
            varPmasSTOCKDESC = N2Str2Null(rsPartMas!STOCKDESC)
            varPmasINVCLASS = N2Str2Null(rsPartMas!InvClass)
            varPmasMAD = N2Str2IntZero(rsPartMas!MAD)
            varPmasVEHTYPE = N2Str2Null(rsPartMas!VEHTYPE)
            varPmasMODELCODE = N2Str2Null(rsPartMas!MODELCODE)
            varPmasLOCATION = N2Str2Null(rsPartMas!Location)
            varPmasMAC = N2Str2Zero(rsPartMas!Mac)
            varPmasOLDNO = N2Str2Null(rsPartMas!OLDNO)
            varPmasNEWNO = N2Str2Null(rsPartMas!NEWNO)
            varPmasGENNO = N2Str2Null(rsPartMas!GENNO)
            varPmasSRP = N2Str2Zero(rsPartMas!SRP)
            varPmasNOSHIP = N2Str2Zero(rsPartMas!NOSHIP)
            varPmasLASTM_MAC = N2Str2Zero(rsPartMas!LASTM_MAC)
            varPmasLASTM_MAD = N2Str2Zero(rsPartMas!LASTM_MAD)
            varPmasLASTM_SELL = N2Str2Zero(rsPartMas!LASTM_SELL)
            varPmasLASTM_OH = N2Str2IntZero(rsPartMas!LASTM_OH)
            varPmasLASTM_OO = N2Str2IntZero(rsPartMas!LASTM_OO)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(rsPartMas!ONHAND)
            varPmasTrecqty = N2Str2IntZero(rsPartMas!TRECQTY)
            varPmasTISSQTY = N2Str2IntZero(rsPartMas!TISSQTY)
            varPmasOnOrder = N2Str2IntZero(rsPartMas!ONORDER)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(rsPartMas!TPOQTY)
            varPmasPRQTY = N2Str2IntZero(rsPartMas!PRQTY)
            varPmasTPRQTY = N2Str2IntZero(rsPartMas!TPRQTY)
            varPmasLAST_RECQ = N2Str2IntZero(rsPartMas!LAST_RECQ)
            varPmasLAST_RECD = N2Date2Null(rsPartMas!LAST_RECD)
            varPmasLASTY_OH = N2Str2IntZero(rsPartMas!LASTY_OH)
            varPmasLASTY_MAC = N2Str2Zero(rsPartMas!LASTY_MAC)
            varPmasLASTY_OO = N2Str2IntZero(rsPartMas!LASTY_OO)
            varPmasLASTY_ADJ = N2Str2IntZero(rsPartMas!LASTY_ADJ)
            varPmasHOLD = N2Str2IntZero(rsPartMas!hold)
            varPmasSUPCODE = N2Str2Null(rsPartMas!SUPCODE)
            varPmasVARIANCE = N2Str2IntZero(rsPartMas!VARIANCE)
            varPmasSUBINVCLASS = N2Str2Null(rsPartMas!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(rsPartMas!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(rsPartMas!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(rsPartMas!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(rsPartMas!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(rsPartMas!RECEIPTS)
            varPmasISSUANCES = N2Str2IntZero(rsPartMas!ISSUANCES)
            varPmasUSERCODE = N2Str2Null(rsPartMas!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(rsPartMas!LASTUPDATE)
            varPmasDNP = N2Str2Zero(rsPartMas!DNP)
            varPmasVALID_ICC = N2Str2Null(rsPartMas!VALID_ICC)
            varPmasDATE_ENTERED = N2Str2Null(rsPartMas!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,STOCKNO,STOCKDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                On Error GoTo ERRORCODE
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update CSMS_MATMAS set " & _
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
        MsgSpeech "Cut Off Master File Successfully Created"
        Me.Caption = "Cut Off Master File Successfully Created"
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

    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from CSMS_MATMAS WHERE ACTIVE = 'Y' order by PARTNO asc", gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        MsgSpeech "Creating Cut Off Master File"
        Me.Caption = "Creating Cut Off Master File"
        Screen.MousePointer = 11
        DoEvents
        I = 0
        gconINVENTORY.Execute "delete * from CUTOFF"
        Do While Not rsPartMas.EOF
            varPmasID = I + 1
            labProcessing.Caption = "Processing Material Code: " & Null2String(rsPartMas!PARTNO)
            DoEvents
            varPmasSTOCKNO = N2Str2Null(rsPartMas!PARTNO)
            varPmasSTOCKDESC = N2Str2Null(rsPartMas!PARTDESC)
            varPmasINVCLASS = N2Str2Null(rsPartMas!InvClass)
            varPmasMAD = N2Str2IntZero(rsPartMas!MAD)
            varPmasVEHTYPE = N2Str2Null(rsPartMas!VEHTYPE)
            varPmasMODELCODE = N2Str2Null(rsPartMas!MODELCODE)
            varPmasLOCATION = N2Str2Null(rsPartMas!Location)
            varPmasMAC = N2Str2Zero(rsPartMas!Mac)
            varPmasOLDNO = N2Str2Null(rsPartMas!OLDNO)
            varPmasNEWNO = N2Str2Null(rsPartMas!NEWNO)
            varPmasGENNO = N2Str2Null(rsPartMas!GENNO)
            varPmasSRP = N2Str2Zero(rsPartMas!SRP)
            varPmasNOSHIP = N2Str2Zero(rsPartMas!NOSHIP)
            varPmasLASTM_MAC = N2Str2Zero(rsPartMas!LASTM_MAC)
            varPmasLASTM_MAD = N2Str2Zero(rsPartMas!LASTM_MAD)
            varPmasLASTM_SELL = N2Str2Zero(rsPartMas!LASTM_SELL)
            varPmasLASTM_OH = N2Str2IntZero(rsPartMas!LASTM_OH)
            varPmasLASTM_OO = N2Str2IntZero(rsPartMas!LASTM_OO)
            If varPmasLASTM_OO < 0 Then varPmasLASTM_OO = 0
            varPmasOnhand = N2Str2IntZero(rsPartMas!ONHAND)
            varPmasTrecqty = N2Str2IntZero(rsPartMas!TRECQTY)
            varPmasTISSQTY = N2Str2IntZero(rsPartMas!TISSQTY)
            varPmasOnOrder = N2Str2IntZero(rsPartMas!ONORDER)
            If varPmasOnOrder < 0 Then varPmasOnOrder = 0
            varPmasTpoqty = N2Str2IntZero(rsPartMas!TPOQTY)
            varPmasPRQTY = N2Str2IntZero(rsPartMas!PRQTY)
            varPmasTPRQTY = N2Str2IntZero(rsPartMas!TPRQTY)
            varPmasLAST_RECQ = N2Str2IntZero(rsPartMas!LAST_RECQ)
            varPmasLAST_RECD = N2Date2Null(rsPartMas!LAST_RECD)
            varPmasLASTY_OH = N2Str2IntZero(rsPartMas!LASTY_OH)
            varPmasLASTY_MAC = N2Str2Zero(rsPartMas!LASTY_MAC)
            varPmasLASTY_OO = N2Str2IntZero(rsPartMas!LASTY_OO)
            varPmasLASTY_ADJ = N2Str2IntZero(rsPartMas!LASTY_ADJ)
            'varPmasHOLD = N2Str2IntZero(rsPARTMAS!hold)
            varPmasSUPCODE = N2Str2Null(rsPartMas!SUPCODE)
            varPmasVARIANCE = N2Str2IntZero(rsPartMas!VARIANCE)
            varPmasSUBINVCLASS = N2Str2Null(rsPartMas!SubInvClas)
            varPmasPHYCOUNT = N2Str2IntZero(rsPartMas!PHYCOUNT)
            varPmasADJPHYCOUNT = N2Str2IntZero(rsPartMas!ADJPHYCNT)
            varPmasCUTOFFQTY = N2Str2IntZero(rsPartMas!CUTOFFQTY)
            varPmasCUTOFFMAC = N2Str2Zero(rsPartMas!CUTOFFMAC)
            varPmasRECEIPTS = N2Str2IntZero(rsPartMas!RECEIPTS)
            varPmasISSUANCES = N2Str2IntZero(rsPartMas!ISSUANCES)
            varPmasUSERCODE = N2Str2Null(rsPartMas!USERCODE)
            varPmasLASTUPDATE = N2Date2Null(rsPartMas!LASTUPDATE)
            varPmasDNP = N2Str2Zero(rsPartMas!DNP)
            varPmasVALID_ICC = N2Str2Null(rsPartMas!VALID_ICC)
            varPmasDATE_ENTERED = N2Str2Null(rsPartMas!DATE_ENTERED)
            If varPmasSTOCKNO <> "NULL" Then
                MoveSql = "INSERT INTO CUTOFF " & _
                          "(ID,PARTNO,PARTDESC,INVCLASS,VEHTYPE,MODELCODE,LOCATION,MAC,MAD,OLDNO,NEWNO,GENNO,SRP,NOSHIP,LASTM_MAC,LASTM_MAD,LASTM_SELL,LASTM_OH,LASTM_OO,ONHAND,TRECQTY,TISSQTY,ONORDER,TPOQTY,PRQTY,TPRQTY,LAST_RECQ,LAST_RECD,LASTY_OH,LASTY_MAC,LASTY_OO,LASTY_ADJ,HOLD,SUPCODE,VARIANCE,SUBINVCLAS,PHYCOUNT,ADJPHYCNT,CUTOFFQTY,CUTOFFMAC,RECEIPTS,ISSUANCES,USERCODE,LASTUPDATE,DNP,VALID_ICC,DATE_ENTERED)" & _
                        " values (" & varPmasID & ", " & varPmasSTOCKNO & "," & varPmasSTOCKDESC & "," & varPmasINVCLASS & "," & varPmasVEHTYPE & "," & varPmasMODELCODE & "," & varPmasLOCATION & "," & varPmasMAC & "," & varPmasMAD & "," & varPmasOLDNO & "," & varPmasNEWNO & "," & varPmasGENNO & "," & varPmasSRP & "," & varPmasNOSHIP & "," & varPmasLASTM_MAC & "," & varPmasLASTM_MAD & "," & varPmasLASTM_SELL & "," & varPmasLASTM_OH & "," & varPmasLASTM_OO & "," & varPmasOnhand & "," & varPmasTrecqty & "," & varPmasTISSQTY & "," & varPmasOnOrder & "," & varPmasTpoqty & "," & varPmasPRQTY & "," & varPmasTPRQTY & "," & varPmasLAST_RECQ & "," & varPmasLAST_RECD & "," & varPmasLASTY_OH & "," & varPmasLASTY_MAC & "," & varPmasLASTY_OO & "," & varPmasLASTY_ADJ & "," & varPmasHOLD & "," & _
                        " " & varPmasSUPCODE & "," & varPmasVARIANCE & "," & varPmasSUBINVCLASS & "," & varPmasPHYCOUNT & "," & varPmasADJPHYCOUNT & "," & varPmasCUTOFFQTY & "," & varPmasCUTOFFMAC & "," & varPmasRECEIPTS & "," & varPmasISSUANCES & "," & varPmasUSERCODE & "," & varPmasLASTUPDATE & "," & varPmasDNP & "," & varPmasVALID_ICC & ", " & varPmasDATE_ENTERED & ")"
                On Error GoTo ERRORCODE
                gconINVENTORY.Execute MoveSql
                gconDMIS.Execute "update CSMS_MATMAS set " & _
                               " CUTOFFQTY = " & N2Str2IntZero(rsPartMas!ONHAND) & "," & _
                               " CUTOFFMAC =" & N2Str2Zero(rsPartMas!Mac) & _
                               " WHERE PARTNO = " & varPmasSTOCKNO
            End If
            I = I + 1
            progCPB.Value = (I / rsPartMas.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsPartMas.MoveNext
        Loop
        labProcessing.Caption = ""
        DoEvents
        MsgSpeech "Cut Off Master File Successfully Created"
        Me.Caption = "Cut Off Master File Successfully Created"
        Screen.MousePointer = 0
        DoEvents
    End If
    Exit Sub

ERRORCODE:
    Screen.MousePointer = 0
    ShowVBError
    MsgBox varPmasSTOCKNO
    Resume Next
    Exit Sub
End Sub

Private Sub cmdCreate_Click()


    cmdCreate.Enabled = False
    cmdExit.Enabled = False
    DoEvents
    CreateCutOffMaster
    LogAudit "G", "CREATE CUT OFF MASTER"

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

