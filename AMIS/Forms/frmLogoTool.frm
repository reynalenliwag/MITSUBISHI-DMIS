VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogoTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Logo Tool"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   4080
      Width           =   3615
   End
   Begin VB.CommandButton cmdCrop 
      Caption         =   "CROP"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Frame fraContainer 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdClose 
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   5040
         Width           =   3615
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "UPLOAD"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   4560
         Width           =   3615
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "BROWSE"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Preview"
         Height          =   2655
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3615
         Begin VB.PictureBox picContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   840
            Picture         =   "frmLogoTool.frx":0000
            ScaleHeight     =   1785
            ScaleWidth      =   1905
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   480
            Width           =   1935
            Begin VB.Image picTemp 
               Height          =   1695
               Left            =   0
               ToolTipText     =   "picture editor"
               Top             =   0
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "To be uploaded"
         Height          =   2655
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         Begin VB.PictureBox picPreview 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1785
            Left            =   840
            ScaleHeight     =   1755
            ScaleWidth      =   1875
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   480
            Width           =   1905
         End
      End
      Begin MSComctlLib.Slider sldZoom 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "zoom in or zoom out picture"
         Top             =   3360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         Max             =   200
         SelStart        =   100
         TickStyle       =   3
         TickFrequency   =   5
         Value           =   100
      End
      Begin MSComctlLib.Slider sldVPos 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "move the picture vertically"
         Top             =   4080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Max             =   100
         TickStyle       =   3
         TickFrequency   =   5
      End
      Begin MSComctlLib.Slider sldHPos 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "move the picture horizontally"
         Top             =   4800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Max             =   100
         TickStyle       =   3
         TickFrequency   =   5
      End
      Begin MSComDlg.CommonDialog cdBrowse 
         Left            =   240
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Horizontal Position :"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Vertical Position :"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Zoom :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLogoTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim uploadedPixX As Integer
Dim uploadedPixY As Integer

Dim CMD                   As New ADODB.Command
Dim bytData()             As Byte
Dim strDescription        As String
Dim rsPic                 As New ADODB.Recordset

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    With cdBrowse
        .DialogTitle = "Load Picture"
        .filter = "Picture Files|*.jpg;*.jpeg;*.bmp;*.gif;*.png"
        .DefaultExt = "*.jpg"
        .FilterIndex = 1
        .InitDir = "C:\"
    End With
    
    FillPicture
    
    sldZoom.Value = 100
    sldVPos.Value = 0
    sldHPos.Value = 0
    cmdBrowse.Enabled = True
    cmdReset.Enabled = False
    cmdCrop.Enabled = False
    cmdUpload.Enabled = False
    Screen.MousePointer = vbDefault
    'REYNALEN LIWAG
End Sub

Public Sub FillPicture()
    Dim rsPic As New ADODB.Recordset
    Set rsPic = New ADODB.Recordset
    rsPic.Open "SELECT LOGO FROM ALL_PROFILE WHERE MODULENAME = 'AMIS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    If Not (rsPic.BOF And rsPic.EOF) Then
        Set picContainer.DataSource = rsPic
        'picContainer.DataField = Null2String(rsPic!LOGO)
    End If
End Sub

Private Sub cmdBrowse_Click()
On Error Resume Next
    cdBrowse.ShowOpen
    If cdBrowse.FileName <> "" Then
        picTemp.Picture = LoadPicture(cdBrowse.FileName)
        uploadedPixX = picTemp.Width
        uploadedPixY = picTemp.Height
        cmdBrowse.Enabled = False
        sldZoom.Enabled = True
        sldVPos.Enabled = True
        sldHPos.Enabled = True
        cmdCrop.Enabled = True
        cmdReset.Enabled = True
    End If
    
'    Open cdBrowse.FileName For Binary As #1
'    ReDim bytData(FileLen(cdBrowse.FileName))
End Sub

Private Sub cmdCrop_Click()
    picPreview.Cls
    picPreview.PaintPicture picTemp.Picture, picTemp.Left, picTemp.Top, picTemp.Width, picTemp.Height, 0, 0, picTemp.Width / (sldZoom.Value / 100), picTemp.Height / (sldZoom.Value / 100)
    cmdUpload.Enabled = True
    Exit Sub
End Sub

Private Sub cmdReset_Click()
    picContainer.Picture = Nothing
    picPreview.Picture = Nothing
    Call Form_Load
End Sub

Private Sub cmdUpload_Click()
    Set rsPic = New ADODB.Recordset
    rsPic.Open "ALL_PROFILE", gconDMIS, adOpenKeyset, adLockPessimistic, adCmdTable
    
    With rsPic
        .Update
        .Fields("LOGO").AppendChunk bytData(FileLen(cdBrowse.FileName))
        .Update
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub sldZoom_Change()
    resamplePic
End Sub

Private Sub sldHPos_Change()
    resamplePic
End Sub

Private Sub sldVPos_Change()
    resamplePic
End Sub

Private Sub resamplePic()
    On Error GoTo err
    picTemp.Stretch = True
    picTemp.Top = -((((uploadedPixY * (sldZoom.Value / 100)) - picContainer.Height) / 100) * sldVPos.Value)
    picTemp.Left = -((((uploadedPixX * (sldZoom.Value / 100)) - picContainer.Width) / 100) * sldHPos.Value)
    picTemp.Height = uploadedPixY * (sldZoom.Value / 100)
    picTemp.Width = uploadedPixX * (sldZoom.Value / 100)
    Exit Sub
err:
    MsgBox ("Error on cropping the image")
End Sub

