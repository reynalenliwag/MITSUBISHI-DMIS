VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBayMonitoring 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   15210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBayMonitoring.frx":0000
   ScaleHeight     =   11265
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicsecondPage 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9225
      Left            =   120
      Picture         =   "frmBayMonitoring.frx":66ECD
      ScaleHeight     =   9225
      ScaleWidth      =   15345
      TabIndex        =   209
      Top             =   1320
      Width           =   15345
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   29
         Left            =   12300
         Picture         =   "frmBayMonitoring.frx":748D4
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   394
         Top             =   6450
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   29
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   29
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":7BD8C
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   29
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":7C4BC
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   29
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":7CB5E
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   29
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":7D510
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   29
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":7DF25
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   29
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":7E601
            MousePointer    =   99  'Custom
            TabIndex        =   406
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   29
            Left            =   1110
            TabIndex        =   405
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   29
            Left            =   780
            TabIndex        =   404
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   1110
            TabIndex        =   403
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   30
            Left            =   45
            TabIndex        =   402
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   29
            Left            =   585
            TabIndex        =   401
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   29
            Left            =   510
            TabIndex        =   400
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   59
            Left            =   30
            TabIndex        =   399
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   29
            Left            =   1110
            TabIndex        =   398
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   1110
            TabIndex        =   397
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   29
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":7E753
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   58
            Left            =   480
            TabIndex        =   396
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   1110
            TabIndex        =   395
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   28
         Left            =   9390
         Picture         =   "frmBayMonitoring.frx":7EFEC
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   381
         Top             =   6420
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   28
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   28
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":864A4
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   28
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":86BD4
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   28
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":87276
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   28
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":87C28
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   28
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":8863D
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   28
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":88D19
            MousePointer    =   99  'Custom
            TabIndex        =   393
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   28
            Left            =   1110
            TabIndex        =   392
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   780
            TabIndex        =   391
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   1110
            TabIndex        =   390
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   29
            Left            =   45
            TabIndex        =   389
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   28
            Left            =   585
            TabIndex        =   388
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   28
            Left            =   510
            TabIndex        =   387
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   57
            Left            =   30
            TabIndex        =   386
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   28
            Left            =   1110
            TabIndex        =   385
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   1110
            TabIndex        =   384
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   28
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":88E6B
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   56
            Left            =   480
            TabIndex        =   383
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   1110
            TabIndex        =   382
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   27
         Left            =   6480
         Picture         =   "frmBayMonitoring.frx":89704
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   368
         Top             =   6450
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   27
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   27
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":90BBC
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   27
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":912EC
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   27
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":9198E
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   27
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":92340
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   27
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":92D55
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   27
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":93431
            MousePointer    =   99  'Custom
            TabIndex        =   380
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   27
            Left            =   1110
            TabIndex        =   379
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   780
            TabIndex        =   378
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   1110
            TabIndex        =   377
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   28
            Left            =   45
            TabIndex        =   376
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   27
            Left            =   585
            TabIndex        =   375
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   27
            Left            =   510
            TabIndex        =   374
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   55
            Left            =   30
            TabIndex        =   373
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   27
            Left            =   1110
            TabIndex        =   372
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   1110
            TabIndex        =   371
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   27
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":93583
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   54
            Left            =   480
            TabIndex        =   370
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   1110
            TabIndex        =   369
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   26
         Left            =   3570
         Picture         =   "frmBayMonitoring.frx":93E1C
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   355
         Top             =   6450
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   26
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   26
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":9B2D4
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   26
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":9BA04
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   26
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":9C0A6
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   26
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":9CA58
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   26
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":9D46D
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   26
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":9DB49
            MousePointer    =   99  'Custom
            TabIndex        =   367
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   26
            Left            =   1110
            TabIndex        =   366
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   780
            TabIndex        =   365
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   1110
            TabIndex        =   364
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   27
            Left            =   45
            TabIndex        =   363
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   26
            Left            =   585
            TabIndex        =   362
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   26
            Left            =   510
            TabIndex        =   361
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   53
            Left            =   30
            TabIndex        =   360
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   26
            Left            =   1110
            TabIndex        =   359
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   1110
            TabIndex        =   358
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   26
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":9DC9B
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   52
            Left            =   480
            TabIndex        =   357
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   1110
            TabIndex        =   356
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   25
         Left            =   690
         Picture         =   "frmBayMonitoring.frx":9E534
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   342
         Top             =   6450
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   25
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   25
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":A59EC
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   25
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":A611C
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   25
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":A67BE
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   25
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":A7170
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   25
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":A7B85
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   25
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":A8261
            MousePointer    =   99  'Custom
            TabIndex        =   354
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   25
            Left            =   1110
            TabIndex        =   353
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   780
            TabIndex        =   352
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   25
            Left            =   1110
            TabIndex        =   351
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   26
            Left            =   45
            TabIndex        =   350
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   25
            Left            =   585
            TabIndex        =   349
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   25
            Left            =   510
            TabIndex        =   348
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   51
            Left            =   30
            TabIndex        =   347
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   25
            Left            =   1110
            TabIndex        =   346
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   25
            Left            =   1110
            TabIndex        =   345
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   25
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":A83B3
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   50
            Left            =   480
            TabIndex        =   344
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   25
            Left            =   1110
            TabIndex        =   343
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   24
         Left            =   12150
         Picture         =   "frmBayMonitoring.frx":A8C4C
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   329
         Top             =   3300
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   24
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   24
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":B0104
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   24
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":B0834
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   24
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":B0ED6
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   24
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":B1888
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   24
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":B229D
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   24
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":B2979
            MousePointer    =   99  'Custom
            TabIndex        =   341
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   24
            Left            =   1110
            TabIndex        =   340
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   780
            TabIndex        =   339
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   1110
            TabIndex        =   338
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   25
            Left            =   45
            TabIndex        =   337
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   24
            Left            =   585
            TabIndex        =   336
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   24
            Left            =   510
            TabIndex        =   335
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   49
            Left            =   30
            TabIndex        =   334
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   24
            Left            =   1110
            TabIndex        =   333
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   1110
            TabIndex        =   332
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   24
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":B2ACB
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   48
            Left            =   480
            TabIndex        =   331
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   1110
            TabIndex        =   330
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   23
         Left            =   9360
         Picture         =   "frmBayMonitoring.frx":B3364
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   316
         Top             =   3330
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   23
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   23
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":BA81C
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   23
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":BAF4C
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   23
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":BB5EE
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   23
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":BBFA0
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   23
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":BC9B5
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   23
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":BD091
            MousePointer    =   99  'Custom
            TabIndex        =   328
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   23
            Left            =   1110
            TabIndex        =   327
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   780
            TabIndex        =   326
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   1110
            TabIndex        =   325
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   24
            Left            =   45
            TabIndex        =   324
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   23
            Left            =   585
            TabIndex        =   323
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   23
            Left            =   510
            TabIndex        =   322
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   47
            Left            =   30
            TabIndex        =   321
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   23
            Left            =   1110
            TabIndex        =   320
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   1110
            TabIndex        =   319
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   23
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":BD1E3
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   46
            Left            =   480
            TabIndex        =   318
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   1110
            TabIndex        =   317
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   22
         Left            =   6480
         Picture         =   "frmBayMonitoring.frx":BDA7C
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   303
         Top             =   3300
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   22
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   22
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":C4F34
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   22
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":C5664
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   22
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":C5D06
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   22
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":C66B8
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   22
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":C70CD
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   22
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":C77A9
            MousePointer    =   99  'Custom
            TabIndex        =   315
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   22
            Left            =   1110
            TabIndex        =   314
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   780
            TabIndex        =   313
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   1110
            TabIndex        =   312
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   23
            Left            =   45
            TabIndex        =   311
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   22
            Left            =   585
            TabIndex        =   310
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   22
            Left            =   510
            TabIndex        =   309
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   45
            Left            =   30
            TabIndex        =   308
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   22
            Left            =   1110
            TabIndex        =   307
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   1110
            TabIndex        =   306
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   22
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":C78FB
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   44
            Left            =   480
            TabIndex        =   305
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   1110
            TabIndex        =   304
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   21
         Left            =   3600
         Picture         =   "frmBayMonitoring.frx":C8194
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   290
         Top             =   3300
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   21
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   21
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":CF64C
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   21
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":CFD7C
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   21
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":D041E
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   21
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":D0DD0
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   21
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":D17E5
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   21
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":D1EC1
            MousePointer    =   99  'Custom
            TabIndex        =   302
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   21
            Left            =   1110
            TabIndex        =   301
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   780
            TabIndex        =   300
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   1110
            TabIndex        =   299
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   22
            Left            =   45
            TabIndex        =   298
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   21
            Left            =   585
            TabIndex        =   297
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   21
            Left            =   510
            TabIndex        =   296
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   43
            Left            =   30
            TabIndex        =   295
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   21
            Left            =   1110
            TabIndex        =   294
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   1110
            TabIndex        =   293
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   21
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":D2013
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   42
            Left            =   480
            TabIndex        =   292
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   1110
            TabIndex        =   291
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   20
         Left            =   660
         Picture         =   "frmBayMonitoring.frx":D28AC
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   277
         Top             =   3270
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   20
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   20
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":D9D64
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   20
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":DA494
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   20
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":DAB36
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   20
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":DB4E8
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   20
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":DBEFD
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   20
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":DC5D9
            MousePointer    =   99  'Custom
            TabIndex        =   289
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   20
            Left            =   1110
            TabIndex        =   288
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   780
            TabIndex        =   287
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   1110
            TabIndex        =   286
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   21
            Left            =   45
            TabIndex        =   285
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   20
            Left            =   585
            TabIndex        =   284
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   20
            Left            =   510
            TabIndex        =   283
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   41
            Left            =   30
            TabIndex        =   282
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   20
            Left            =   1110
            TabIndex        =   281
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   1110
            TabIndex        =   280
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   20
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":DC72B
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   40
            Left            =   480
            TabIndex        =   279
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   1110
            TabIndex        =   278
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   19
         Left            =   12180
         Picture         =   "frmBayMonitoring.frx":DCFC4
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   264
         Top             =   180
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   19
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   19
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":E447C
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   19
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":E4BAC
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   19
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":E524E
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   19
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":E5C00
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   19
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":E6615
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   19
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":E6CF1
            MousePointer    =   99  'Custom
            TabIndex        =   276
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   19
            Left            =   1110
            TabIndex        =   275
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   780
            TabIndex        =   274
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   1110
            TabIndex        =   273
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   20
            Left            =   45
            TabIndex        =   272
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   19
            Left            =   585
            TabIndex        =   271
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   19
            Left            =   510
            TabIndex        =   270
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   39
            Left            =   30
            TabIndex        =   269
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   19
            Left            =   1110
            TabIndex        =   268
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   1110
            TabIndex        =   267
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   19
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":E6E43
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   38
            Left            =   480
            TabIndex        =   266
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   1110
            TabIndex        =   265
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   18
         Left            =   9330
         Picture         =   "frmBayMonitoring.frx":E76DC
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   251
         Top             =   180
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   18
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   18
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":EEB94
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   18
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":EF2C4
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   18
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":EF966
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   18
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":F0318
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   18
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":F0D2D
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   18
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":F1409
            MousePointer    =   99  'Custom
            TabIndex        =   263
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   18
            Left            =   1110
            TabIndex        =   262
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   780
            TabIndex        =   261
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   1110
            TabIndex        =   260
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   19
            Left            =   45
            TabIndex        =   259
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   18
            Left            =   585
            TabIndex        =   258
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   18
            Left            =   510
            TabIndex        =   257
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   37
            Left            =   30
            TabIndex        =   256
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   18
            Left            =   1110
            TabIndex        =   255
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   1110
            TabIndex        =   254
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   18
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":F155B
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   36
            Left            =   480
            TabIndex        =   253
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   1110
            TabIndex        =   252
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   17
         Left            =   6450
         Picture         =   "frmBayMonitoring.frx":F1DF4
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   238
         Top             =   180
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   17
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   17
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":F92AC
            Top             =   330
            Width           =   795
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   17
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":F99DC
            Top             =   330
            Width           =   480
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   17
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":FA07E
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   17
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":FAA30
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   17
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":FB445
            Top             =   150
            Width           =   720
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   17
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":FBB21
            MousePointer    =   99  'Custom
            TabIndex        =   250
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   17
            Left            =   1110
            TabIndex        =   249
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   780
            TabIndex        =   248
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   1110
            TabIndex        =   247
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   18
            Left            =   45
            TabIndex        =   246
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   17
            Left            =   585
            TabIndex        =   245
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   17
            Left            =   510
            TabIndex        =   244
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   35
            Left            =   30
            TabIndex        =   243
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   17
            Left            =   1110
            TabIndex        =   242
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   1110
            TabIndex        =   241
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   17
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":FBC73
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   34
            Left            =   480
            TabIndex        =   240
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   1110
            TabIndex        =   239
            Top             =   1620
            Width           =   690
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   16
         Left            =   3570
         Picture         =   "frmBayMonitoring.frx":FC50C
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   223
         Top             =   180
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   16
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   1110
            TabIndex        =   235
            Top             =   1620
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   33
            Left            =   480
            TabIndex        =   234
            Top             =   1620
            Width           =   540
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   16
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":1039C4
            Stretch         =   -1  'True
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   1110
            TabIndex        =   233
            Top             =   2100
            Width           =   345
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   16
            Left            =   1110
            TabIndex        =   232
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   32
            Left            =   30
            TabIndex        =   231
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   16
            Left            =   510
            TabIndex        =   230
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   16
            Left            =   585
            TabIndex        =   229
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   17
            Left            =   45
            TabIndex        =   228
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   1110
            TabIndex        =   227
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label lblrostatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   780
            TabIndex        =   226
            Top             =   390
            Width           =   1530
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   16
            Left            =   1110
            TabIndex        =   225
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   16
            Left            =   240
            MouseIcon       =   "frmBayMonitoring.frx":10425D
            MousePointer    =   99  'Custom
            TabIndex        =   224
            Top             =   2400
            Width           =   2265
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   16
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":1043AF
            Top             =   150
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   16
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":104A8B
            Top             =   180
            Width           =   720
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   16
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":1054A0
            Top             =   150
            Width           =   720
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   16
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":105E52
            Top             =   330
            Width           =   480
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   16
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":1064F4
            Top             =   330
            Width           =   795
         End
      End
      Begin VB.PictureBox thepic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   15
         Left            =   660
         Picture         =   "frmBayMonitoring.frx":106C24
         ScaleHeight     =   2775
         ScaleWidth      =   2535
         TabIndex        =   210
         Top             =   180
         Width           =   2535
         Begin VB.Timer Timer1 
            Index           =   15
            Interval        =   500
            Left            =   1980
            Top             =   1890
         End
         Begin VB.Label lblbaydesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "thebay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   315
            Index           =   15
            Left            =   180
            MouseIcon       =   "frmBayMonitoring.frx":10E0DC
            MousePointer    =   99  'Custom
            TabIndex        =   222
            Top             =   2400
            Width           =   2205
         End
         Begin VB.Label lbltheRO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theRO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   15
            Left            =   1080
            TabIndex        =   221
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label lblrostatus 
            BackStyle       =   0  'Transparent
            Caption         =   "TheROStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   15
            Left            =   810
            TabIndex        =   220
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label lblbaystatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "bayStatus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1080
            TabIndex        =   219
            Top             =   1140
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bay Status :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   16
            Left            =   45
            TabIndex        =   218
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ro #:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   15
            Left            =   585
            TabIndex        =   217
            Top             =   1380
            Width           =   435
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   15
            Left            =   510
            TabIndex        =   216
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cus Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   31
            Left            =   30
            TabIndex        =   215
            Top             =   2070
            Width           =   990
         End
         Begin VB.Label lblplate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "theplate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   15
            Left            =   1080
            TabIndex        =   214
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label lblCustName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cust"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1080
            TabIndex        =   213
            Top             =   2100
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   15
            Left            =   90
            Picture         =   "frmBayMonitoring.frx":10E22E
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   225
            Index           =   30
            Left            =   480
            TabIndex        =   212
            Top             =   1620
            Width           =   540
         End
         Begin VB.Label LblModel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblModel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1080
            TabIndex        =   211
            Top             =   1620
            Width           =   690
         End
         Begin VB.Image Working 
            Height          =   720
            Index           =   15
            Left            =   120
            Picture         =   "frmBayMonitoring.frx":10EAC7
            Top             =   210
            Width           =   720
         End
         Begin VB.Image FinishNa 
            Height          =   720
            Index           =   15
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":10F1A3
            Top             =   240
            Width           =   720
         End
         Begin VB.Image Iddle 
            Height          =   720
            Index           =   15
            Left            =   60
            Picture         =   "frmBayMonitoring.frx":10FBB8
            Top             =   210
            Width           =   720
         End
         Begin VB.Image billed 
            Height          =   480
            Index           =   15
            Left            =   210
            Picture         =   "frmBayMonitoring.frx":11056A
            Top             =   360
            Width           =   480
         End
         Begin VB.Image handpoint 
            Height          =   540
            Index           =   15
            Left            =   1890
            Picture         =   "frmBayMonitoring.frx":110C0C
            Top             =   330
            Width           =   795
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   15210
      TabIndex        =   190
      Top             =   10560
      Width           =   15210
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   9600
         Top             =   180
      End
      Begin wizButton.cmd cmd1 
         Height          =   375
         Left            =   13830
         TabIndex        =   191
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmBayMonitoring.frx":11133C
      End
      Begin wizButton.cmd cmd2 
         Height          =   375
         Left            =   9480
         TabIndex        =   197
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmBayMonitoring.frx":111358
      End
      Begin wizButton.cmd cmdnext 
         Height          =   375
         Left            =   12480
         TabIndex        =   236
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         TX              =   "16 - 30"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmBayMonitoring.frx":111374
      End
      Begin wizButton.cmd cmdBack 
         Height          =   375
         Left            =   11010
         TabIndex        =   237
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         TX              =   "1 - 15"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmBayMonitoring.frx":111390
      End
      Begin VB.Label lblbi 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9180
         MouseIcon       =   "frmBayMonitoring.frx":1113AC
         MousePointer    =   99  'Custom
         TabIndex        =   202
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblFi 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5550
         MouseIcon       =   "frmBayMonitoring.frx":1114FE
         MousePointer    =   99  'Custom
         TabIndex        =   200
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblpark 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1260
         MouseIcon       =   "frmBayMonitoring.frx":111650
         MousePointer    =   99  'Custom
         TabIndex        =   198
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Billed -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         MouseIcon       =   "frmBayMonitoring.frx":1117A2
         MousePointer    =   99  'Custom
         TabIndex        =   196
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle Time - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         MouseIcon       =   "frmBayMonitoring.frx":1118F4
         MousePointer    =   99  'Custom
         TabIndex        =   195
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Finish Job - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4290
         MouseIcon       =   "frmBayMonitoring.frx":111A46
         MousePointer    =   99  'Custom
         TabIndex        =   194
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Working - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2340
         MouseIcon       =   "frmBayMonitoring.frx":111B98
         MousePointer    =   99  'Custom
         TabIndex        =   193
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Park-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   660
         MouseIcon       =   "frmBayMonitoring.frx":111CEA
         MousePointer    =   99  'Custom
         TabIndex        =   192
         Top             =   240
         Width           =   555
      End
      Begin VB.Image Image6 
         Height          =   360
         Left            =   7950
         Picture         =   "frmBayMonitoring.frx":111E3C
         Stretch         =   -1  'True
         Top             =   180
         Width           =   435
      End
      Begin VB.Image Image5 
         Height          =   450
         Left            =   5970
         Picture         =   "frmBayMonitoring.frx":1124DE
         Stretch         =   -1  'True
         Top             =   90
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   3690
         Picture         =   "frmBayMonitoring.frx":112E90
         Stretch         =   -1  'True
         Top             =   90
         Width           =   540
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   1620
         Picture         =   "frmBayMonitoring.frx":1138A5
         Stretch         =   -1  'True
         Top             =   120
         Width           =   510
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":113F81
         Stretch         =   -1  'True
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblidle 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7620
         MouseIcon       =   "frmBayMonitoring.frx":11481A
         MousePointer    =   99  'Custom
         TabIndex        =   201
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblworking 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         MouseIcon       =   "frmBayMonitoring.frx":11496C
         MousePointer    =   99  'Custom
         TabIndex        =   199
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer theTimer 
      Interval        =   500
      Left            =   210
      Top             =   2820
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   14
      Left            =   12390
      Picture         =   "frmBayMonitoring.frx":114ABE
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   177
      Top             =   7710
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   14
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   14
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":11BF76
         Top             =   390
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   14
         Left            =   240
         Picture         =   "frmBayMonitoring.frx":11C6A6
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   14
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":11CD48
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   14
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":11D6FA
         Top             =   180
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   14
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":11E10F
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   1080
         TabIndex        =   189
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   29
         Left            =   510
         TabIndex        =   188
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   14
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":11E7EB
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   1080
         TabIndex        =   187
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   14
         Left            =   1080
         TabIndex        =   186
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   28
         Left            =   60
         TabIndex        =   185
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   14
         Left            =   540
         TabIndex        =   184
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   14
         Left            =   615
         TabIndex        =   183
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   14
         Left            =   75
         TabIndex        =   182
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   1080
         TabIndex        =   181
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   780
         TabIndex        =   180
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   14
         Left            =   1080
         TabIndex        =   179
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   14
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":11F084
         MousePointer    =   99  'Custom
         TabIndex        =   178
         Top             =   2400
         Width           =   2325
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   13
      Left            =   9486
      Picture         =   "frmBayMonitoring.frx":11F1D6
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   164
      Top             =   7710
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   13
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   13
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":12668E
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   13
         Left            =   240
         Picture         =   "frmBayMonitoring.frx":126DBE
         Top             =   300
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   13
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":127460
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   13
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":127E12
         Top             =   180
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   13
         Left            =   180
         Picture         =   "frmBayMonitoring.frx":128827
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   1080
         TabIndex        =   176
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   27
         Left            =   510
         TabIndex        =   175
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   13
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":128F03
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   1080
         TabIndex        =   174
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   13
         Left            =   1080
         TabIndex        =   173
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   26
         Left            =   60
         TabIndex        =   172
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   13
         Left            =   540
         TabIndex        =   171
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro#:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   13
         Left            =   660
         TabIndex        =   170
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   13
         Left            =   75
         TabIndex        =   169
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   1080
         TabIndex        =   168
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   13
         Left            =   780
         TabIndex        =   167
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   13
         Left            =   1080
         TabIndex        =   166
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   13
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":12979C
         MousePointer    =   99  'Custom
         TabIndex        =   165
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   12
      Left            =   6584
      Picture         =   "frmBayMonitoring.frx":1298EE
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   151
      Top             =   7710
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   12
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RO#:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   15
         Left            =   615
         TabIndex        =   208
         Top             =   1380
         Width           =   435
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   12
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":130DA6
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   12
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":1314D6
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   12
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":131B78
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   12
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":13252A
         Top             =   180
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   12
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":132F3F
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1110
         TabIndex        =   163
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   25
         Left            =   510
         TabIndex        =   162
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   12
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":13361B
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1110
         TabIndex        =   161
         Top             =   2100
         Width           =   1425
      End
      Begin VB.Label lblplate 
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   12
         Left            =   1110
         TabIndex        =   160
         Top             =   1860
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   24
         Left            =   60
         TabIndex        =   159
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   12
         Left            =   540
         TabIndex        =   158
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   285
         Index           =   12
         Left            =   2550
         TabIndex        =   157
         Top             =   1530
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   12
         Left            =   75
         TabIndex        =   156
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1110
         TabIndex        =   155
         Top             =   1140
         Width           =   2145
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   780
         TabIndex        =   154
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   12
         Left            =   1110
         TabIndex        =   153
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   12
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":133EB4
         MousePointer    =   99  'Custom
         TabIndex        =   152
         Top             =   2430
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   11
      Left            =   3682
      Picture         =   "frmBayMonitoring.frx":134006
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   138
      Top             =   7710
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   11
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   11
         Left            =   1920
         Picture         =   "frmBayMonitoring.frx":13B4BE
         Top             =   390
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   11
         Left            =   240
         Picture         =   "frmBayMonitoring.frx":13BBEE
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   11
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":13C290
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   11
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":13CC42
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   11
         Left            =   150
         Picture         =   "frmBayMonitoring.frx":13D657
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   150
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   23
         Left            =   510
         TabIndex        =   149
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   11
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":13DD33
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   148
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   147
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   22
         Left            =   60
         TabIndex        =   146
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   11
         Left            =   540
         TabIndex        =   145
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   11
         Left            =   615
         TabIndex        =   144
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   11
         Left            =   75
         TabIndex        =   143
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   142
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   780
         TabIndex        =   141
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   140
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   11
         Left            =   90
         MouseIcon       =   "frmBayMonitoring.frx":13E5CC
         MousePointer    =   99  'Custom
         TabIndex        =   139
         Top             =   2430
         Width           =   2415
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   10
      Left            =   780
      Picture         =   "frmBayMonitoring.frx":13E71E
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   125
      Top             =   7710
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   10
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   10
         Left            =   1920
         Picture         =   "frmBayMonitoring.frx":145BD6
         Top             =   390
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   10
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":146306
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   10
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1469A8
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   10
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":14735A
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   10
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":147D6F
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   1080
         TabIndex        =   137
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   21
         Left            =   510
         TabIndex        =   136
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   10
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":14844B
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   1080
         TabIndex        =   135
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   10
         Left            =   1080
         TabIndex        =   134
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   20
         Left            =   60
         TabIndex        =   133
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   10
         Left            =   540
         TabIndex        =   132
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   10
         Left            =   615
         TabIndex        =   131
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   10
         Left            =   75
         TabIndex        =   130
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   1080
         TabIndex        =   129
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   10
         Left            =   810
         TabIndex        =   128
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   10
         Left            =   1080
         TabIndex        =   127
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   10
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":148CE4
         MousePointer    =   99  'Custom
         TabIndex        =   126
         Top             =   2430
         Width           =   2415
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   9
      Left            =   12360
      Picture         =   "frmBayMonitoring.frx":148E36
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   112
      Top             =   4605
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   9
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   9
         Left            =   1860
         Picture         =   "frmBayMonitoring.frx":1502EE
         Top             =   390
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   9
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":150A1E
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   9
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1510C0
         Top             =   240
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   9
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":151A72
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   9
         Left            =   150
         Picture         =   "frmBayMonitoring.frx":152487
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   124
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   19
         Left            =   480
         TabIndex        =   123
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   9
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":152B63
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   122
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   121
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   18
         Left            =   30
         TabIndex        =   120
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   9
         Left            =   510
         TabIndex        =   119
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   9
         Left            =   585
         TabIndex        =   118
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   9
         Left            =   45
         TabIndex        =   117
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   116
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   9
         Left            =   810
         TabIndex        =   115
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   9
         Left            =   1080
         TabIndex        =   114
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   9
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":1533FC
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   2400
         Width           =   2355
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   8
      Left            =   9465
      Picture         =   "frmBayMonitoring.frx":15354E
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   99
      Top             =   4605
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   8
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   8
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":15AA06
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   8
         Left            =   240
         Picture         =   "frmBayMonitoring.frx":15B136
         Top             =   390
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   8
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":15B7D8
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   8
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":15C18A
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   8
         Left            =   150
         Picture         =   "frmBayMonitoring.frx":15CB9F
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   1080
         TabIndex        =   111
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   17
         Left            =   480
         TabIndex        =   110
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   8
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":15D27B
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   1080
         TabIndex        =   109
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   1080
         TabIndex        =   108
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   16
         Left            =   30
         TabIndex        =   107
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   510
         TabIndex        =   106
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   585
         TabIndex        =   105
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   45
         TabIndex        =   104
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   1080
         TabIndex        =   103
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   8
         Left            =   810
         TabIndex        =   102
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   1080
         TabIndex        =   101
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   8
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":15DB14
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   7
      Left            =   6570
      Picture         =   "frmBayMonitoring.frx":15DC66
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   86
      Top             =   4605
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   7
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   7
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":16511E
         Top             =   330
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   7
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":16584E
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   7
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":165EF0
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   7
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":1668A2
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   7
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1672B7
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1080
         TabIndex        =   98
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   14
         Left            =   480
         TabIndex        =   97
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   7
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":167993
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1080
         TabIndex        =   96
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   1080
         TabIndex        =   95
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   13
         Left            =   30
         TabIndex        =   94
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   7
         Left            =   510
         TabIndex        =   93
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   7
         Left            =   585
         TabIndex        =   92
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   91
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1080
         TabIndex        =   90
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   780
         TabIndex        =   89
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   1080
         TabIndex        =   88
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   7
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":16822C
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   6
      Left            =   3675
      Picture         =   "frmBayMonitoring.frx":16837E
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   73
      Top             =   4605
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   6
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   6
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":16F836
         Top             =   300
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   6
         Left            =   180
         Picture         =   "frmBayMonitoring.frx":16FF66
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   6
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":170608
         Top             =   180
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   6
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":170FBA
         Top             =   180
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   6
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1719CF
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   85
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   12
         Left            =   525
         TabIndex        =   84
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   6
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1720AB
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   83
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   82
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   11
         Left            =   75
         TabIndex        =   81
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   6
         Left            =   555
         TabIndex        =   80
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   6
         Left            =   630
         TabIndex        =   79
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   78
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   77
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   780
         TabIndex        =   76
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   75
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   6
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":172944
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   5
      Left            =   780
      Picture         =   "frmBayMonitoring.frx":172A96
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   60
      Top             =   4605
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   5
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   5
         Left            =   1920
         Picture         =   "frmBayMonitoring.frx":179F4E
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   5
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":17A67E
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   5
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":17AD20
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   5
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":17B6D2
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   5
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":17C0E7
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1080
         TabIndex        =   72
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   10
         Left            =   480
         TabIndex        =   71
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   5
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":17C7C3
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1080
         TabIndex        =   70
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   5
         Left            =   1080
         TabIndex        =   69
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   9
         Left            =   30
         TabIndex        =   68
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   5
         Left            =   510
         TabIndex        =   67
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro#:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   5
         Left            =   630
         TabIndex        =   66
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   5
         Left            =   45
         TabIndex        =   65
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1080
         TabIndex        =   64
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   810
         TabIndex        =   63
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   5
         Left            =   1080
         TabIndex        =   62
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   5
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":17D05C
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   4
      Left            =   12330
      Picture         =   "frmBayMonitoring.frx":17D1AE
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   47
      Top             =   1500
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   4
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   4
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":184666
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   4
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":184D96
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   4
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":185438
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   4
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":185DEA
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   4
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":1867FF
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   59
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   8
         Left            =   480
         TabIndex        =   58
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":186EDB
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   57
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   56
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   7
         Left            =   30
         TabIndex        =   55
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   4
         Left            =   510
         TabIndex        =   54
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   4
         Left            =   585
         TabIndex        =   53
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   52
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   51
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   780
         TabIndex        =   50
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   49
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   4
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":187774
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   3
      Left            =   9441
      Picture         =   "frmBayMonitoring.frx":1878C6
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   39
      Top             =   1500
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   3
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   3
         Left            =   15
         TabIndex        =   207
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   3
         Left            =   555
         TabIndex        =   206
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   3
         Left            =   480
         TabIndex        =   205
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   5
         Left            =   0
         TabIndex        =   204
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   6
         Left            =   450
         TabIndex        =   203
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   3
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":18ED7E
         Top             =   360
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   3
         Left            =   240
         Picture         =   "frmBayMonitoring.frx":18F4AE
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   3
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":18FB50
         Top             =   240
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   3
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":190502
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   3
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":190F17
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   46
         Top             =   1620
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   3
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1915F3
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   45
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   44
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   43
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   810
         TabIndex        =   42
         Top             =   420
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   1080
         TabIndex        =   41
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   3
         Left            =   60
         MouseIcon       =   "frmBayMonitoring.frx":191E8C
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   2400
         Width           =   2445
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   2
      Left            =   6554
      Picture         =   "frmBayMonitoring.frx":191FDE
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   26
      Top             =   1500
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   2
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   2
         Left            =   1920
         Picture         =   "frmBayMonitoring.frx":199496
         Top             =   300
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   2
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":199BC6
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   2
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":19A268
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   2
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":19AC1A
         Top             =   210
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   2
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":19B62F
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   38
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   4
         Left            =   510
         TabIndex        =   37
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   2
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":19BD0B
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   36
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   35
         Top             =   1890
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   34
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   2
         Left            =   540
         TabIndex        =   33
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   2
         Left            =   615
         TabIndex        =   32
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   2
         Left            =   75
         TabIndex        =   31
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   30
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   810
         TabIndex        =   29
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   2
         Left            =   1080
         TabIndex        =   28
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   2
         Left            =   150
         MouseIcon       =   "frmBayMonitoring.frx":19C5A4
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   2400
         Width           =   2355
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   1
      Left            =   3667
      Picture         =   "frmBayMonitoring.frx":19C6F6
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   13
      Top             =   1500
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   1
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   1
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":1A3BAE
         Top             =   330
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   1
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":1A42DE
         Top             =   330
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":1A4980
         Top             =   150
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   1
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":1A5332
         Top             =   180
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   1
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":1A5D47
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmBayMonitoring.frx":1A6423
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   24
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblrostatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   780
         TabIndex        =   23
         Top             =   390
         Width           =   1530
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   22
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   21
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   1
         Left            =   585
         TabIndex        =   20
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   1
         Left            =   510
         TabIndex        =   19
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   18
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   17
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   16
         Top             =   2100
         Width           =   345
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   1
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":1A6575
         Stretch         =   -1  'True
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1110
         TabIndex        =   14
         Top             =   1620
         Width           =   690
      End
   End
   Begin VB.PictureBox thepic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Index           =   0
      Left            =   780
      Picture         =   "frmBayMonitoring.frx":1A6E0E
      ScaleHeight     =   2775
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   1500
      Width           =   2535
      Begin VB.Timer Timer1 
         Index           =   0
         Interval        =   500
         Left            =   1980
         Top             =   1890
      End
      Begin VB.Image handpoint 
         Height          =   540
         Index           =   0
         Left            =   1890
         Picture         =   "frmBayMonitoring.frx":1AE2C6
         Top             =   330
         Width           =   795
      End
      Begin VB.Image billed 
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "frmBayMonitoring.frx":1AE9F6
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Iddle 
         Height          =   720
         Index           =   0
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":1AF098
         Top             =   210
         Width           =   720
      End
      Begin VB.Image FinishNa 
         Height          =   720
         Index           =   0
         Left            =   60
         Picture         =   "frmBayMonitoring.frx":1AFA4A
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Working 
         Height          =   720
         Index           =   0
         Left            =   120
         Picture         =   "frmBayMonitoring.frx":1B045F
         Top             =   210
         Width           =   720
      End
      Begin VB.Label LblModel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblModel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   15
         Left            =   480
         TabIndex        =   11
         Top             =   1620
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   90
         Picture         =   "frmBayMonitoring.frx":1B0B3B
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label lblplate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theplate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cus Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   1860
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ro #:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   0
         Left            =   585
         TabIndex        =   6
         Top             =   1380
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bay Status :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label lblbaystatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bayStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblrostatus 
         BackStyle       =   0  'Transparent
         Caption         =   "TheROStatus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   810
         TabIndex        =   3
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lbltheRO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "theRO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   1380
         Width           =   525
      End
      Begin VB.Label lblbaydesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "thebay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   0
         Left            =   180
         MouseIcon       =   "frmBayMonitoring.frx":1B13D4
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   2400
         Width           =   2205
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   9765
      Left            =   120
      Picture         =   "frmBayMonitoring.frx":1B1526
      ScaleHeight     =   9765
      ScaleWidth      =   15135
      TabIndex        =   407
      Top             =   1290
      Width           =   15135
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   9915
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   1290
      Width           =   15765
   End
   Begin VB.Menu mnuOPTION 
      Caption         =   "OPTION"
      Visible         =   0   'False
      Begin VB.Menu MNUREMOVE 
         Caption         =   "Remove Vehice"
      End
   End
End
Attribute VB_Name = "frmBayMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theRostatus                                        As String
Dim theplate_no                                        As String
Dim thecust                                            As String
Dim theBayDesc                                         As String
Dim thecolor                                           As Integer
Dim CounterMe                                          As Integer
Dim theRo                                              As String

Sub fillstatus()
    Dim rsstatus                                       As ADODB.Recordset
    Set rsstatus = gconDMIS.Execute("SELECT STATUS,COUNT(STATUS) total FROM CSMS_REPAIRORDER WHERE RO_NO IN(SELECT RO FROM CSMS_BAYMONITORING)GROUP BY STATUS")

    While Not (rsstatus.EOF Or rsstatus.BOF)
        If Trim(rsstatus!Status) = "Finish Job" Then
            lblFi.Caption = (rsstatus!Total)
        End If
        If Trim(rsstatus!Status) = "Billed" Then
            lblbi.Caption = (rsstatus!Total)
        End If
        If Trim(rsstatus!Status) = "Idle Time" Then
            lblidle.Caption = (rsstatus!Total)
        End If
        If Trim(rsstatus!Status) = "Working" Then
            lblworking.Caption = (rsstatus!Total)
        End If
        If Trim(rsstatus!Status) = "Park" Then
            lblpark.Caption = (rsstatus!Total)
        End If

        rsstatus.MoveNext
    Wend
End Sub

Sub loadthebay()
    initWorkingImage
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim cnt                                            As Integer

    SQL = "SELECT bay_code,bay_description FROM CSMS_baymonitoring "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cnt = 0

    Do While Not RS.EOF
        cnt = cnt + 1
        theBayDesc = Null2String(RS!bay_description)
        getthecode Null2String(RS!bay_code), cnt - 1
        RS.MoveNext
    Loop
    
    
    
    Set RS = Nothing
End Sub

Sub getthecode(BayCode As String, cnt As Integer)
    'initilized the bay
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String
    Dim vBaycode                                       As Integer

    SQL = "SELECT * FROM CSMS_Baymonitoring where bay_code='" & BayCode & "'"
    vBaycode = BayCode - 1
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lbltheRO(vBaycode).Caption = Null2String(RS!ro)
        laodRoStatus Null2String(RS!ro)
        LoadtheROInformation Null2String(RS!ro), vBaycode
        lblbaydesc(vBaycode) = theBayDesc
        If Null2String(RS!Bay_status) = "Allocated" Then
            thepic(vBaycode).Visible = True
        End If
        lblrostatus(vBaycode).Caption = theRostatus
        If Trim(lblrostatus(vBaycode).Caption) = "Working" Then
            lblrostatus(vBaycode).ForeColor = &HC0C000
            Working(vBaycode).Visible = True
        Else
            lblrostatus(vBaycode).ForeColor = &HC0C000
            Working(vBaycode).Visible = False
        End If
        If Trim(lblrostatus(vBaycode).Caption) = "Billed" Then
            lblrostatus(vBaycode).ForeColor = &H800080
            billed(vBaycode).Visible = True
        Else
            billed(vBaycode).Visible = False
            'thecolor = vBaycode
        End If
        If Trim(lblrostatus(vBaycode).Caption) = "Finish Job" Then
            lblrostatus(vBaycode).ForeColor = &HC00000
            FinishNa(vBaycode).Visible = True
        Else
            FinishNa(vBaycode).Visible = False
        End If
        If Trim(lblrostatus(vBaycode).Caption) = "Park" Then
            'lblrostatus(vBaycode).ForeColor = &HC00000
            Image1(vBaycode).Visible = True
        Else
            Image1(vBaycode).Visible = False
        End If
        If Trim(lblrostatus(vBaycode).Caption) = "Idle Time" Then
            lblrostatus(vBaycode).ForeColor = &HC0C000
            Iddle(vBaycode).Visible = True
        Else
            Iddle(vBaycode).Visible = False
        End If
        lblbaystatus(vBaycode).Caption = Null2String(RS!Bay_status)
    End If
    Set RS = Nothing
End Sub

Sub laodRoStatus(XXX As String)
    'get the status of the RO
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT status from CSMS_Repairorder where RO_no='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then
        theRostatus = Null2String(RS!Status)
    Else
        theRostatus = ""
    End If
    RemoveFromBay XXX
    Set RS = Nothing
End Sub

Sub LoadtheROInformation(XXX As String, BayCode As Integer)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    'SQL = "SELECT Plate_no,Lastname,model From CSMS_vw_repairorder where Ro_no='" & XXX & "'"
    
    SQL = "SELECT Plate_no,customer,model From CSMS_vw_repairorder where isnull(Ro_no,'')='" & XXX & "'"


    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then
        
            If XXX = "" Then
                lblplate(BayCode).Caption = ""
                lblCustName(BayCode).Caption = ""
                LblModel(BayCode).Caption = ""
            Else
                lblplate(BayCode).Caption = Null2String(RS!PLATE_NO)
                lblCustName(BayCode).Caption = LCase(Null2String(RS!customer))
                LblModel(BayCode).Caption = LCase(Null2String(RS!Model))
            
            End If
    Else
        lblplate(BayCode).Caption = ""
        lblCustName(BayCode).Caption = ""
        LblModel(BayCode).Caption = ""
    End If
    Set RS = Nothing
End Sub

Sub PastTheInformation(XXX As Integer)
    With frmCSMSViewRO
        .labRO = lbltheRO(XXX).Caption
        '.lblplate = lblplate(XXX).Caption
    End With

    frmCSMSViewRO.Show 1
End Sub

Sub initWorkingImage()
    Dim X                                              As Integer
    For X = 0 To 29
        Working(X).Visible = False
        Image1(X).Visible = False
        FinishNa(X).Visible = False
        Iddle(X).Visible = False
        billed(X).Visible = False
        handpoint(X).Visible = False
    Next
End Sub

Sub RemoveFromBay(yyy As String)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT Status from CSMS_repairorder where ro_no='" & yyy & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If Null2String(Trim(RS!Status)) = "Released" Or Null2String(Trim(RS!Status)) = "Finish Job" Or Null2String(Trim(RS!Status)) = "Billed" Then
            'If Null2String(Trim(rs!Status)) = "Released" Then
            gconDMIS.Execute "Update CSMS_Baymonitoring set ro=null,bay_status='Available' where ro='" & yyy & "'"
        End If
    End If
End Sub

Sub SearchMe(Status As String)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT * from CSMS_BayMonitoring where ro is not null"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    Do While Not RS.EOF
        checktheStatus Null2String(RS!ro), Status
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub checktheStatus(XXX As String, yyy As String)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim BayCode                                        As String
    Dim thelocation                                    As Integer

    SQL = "SELECT status from CSMS_Repairorder where RO_no='" & XXX & "' and status='" & yyy & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Dim rsBay                                      As New ADODB.Recordset
        'get thelocation by bay code
        BayCode = "SELECT bay_code from CSMS_baymonitoring where ro='" & XXX & "'"

        Set rsBay = New ADODB.Recordset
        Set rsBay = gconDMIS.Execute(BayCode)
        If Not RS.EOF And Not RS.BOF Then
            thelocation = Null2String(rsBay!bay_code)
            handpoint(thelocation - 1).Visible = True
        Else
            thelocation = ""
            handpoint(thelocation - 1).Visible = False
        End If

    End If
    'RemoveFromBay xxx
    Set RS = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmd1_Click()
    Unload Me
End Sub

Private Sub cmd2_Click()
    initWorkingImage
    loadthebay
    fillstatus
End Sub

Private Sub cmdBack_Click()
PicsecondPage.Visible = False
End Sub

Private Sub cmdNext_Click()
PicsecondPage.Visible = True
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    initWorkingImage
    PicsecondPage.Visible = False
    loadthebay
End Sub

Private Sub Label3_Click()
    SearchMe "Park"
End Sub

Private Sub Label6_Click()
    SearchMe "Working"
    loadthebay
End Sub

Private Sub Label7_Click()
    SearchMe "Finish Job"
End Sub

Private Sub Label8_Click()
    SearchMe "Idle Time"
End Sub

Private Sub Label9_Click()
    SearchMe "Billed"
End Sub

Private Sub lblbaydesc_Click(Index As Integer)
    If lbltheRO(Index).Caption = "" Then
        MsgBox "Selected Bay is not Occupied", vbInformation, "Empty Bay"
        Exit Sub
    End If
    Select Case Index
        Case 0:
            PastTheInformation 0
        Case 1:
            PastTheInformation 1
        Case 2:
            PastTheInformation 2
        Case 3:
            PastTheInformation 3
        Case 4:
            PastTheInformation 4
        Case 5:
            PastTheInformation 5
        Case 6:
            PastTheInformation 6
        Case 7:
            PastTheInformation 7
        Case 8:
            PastTheInformation 8
        Case 9:
            PastTheInformation 9
        Case 10:
            PastTheInformation 10
        Case 11:
            PastTheInformation 11
        Case 12:
            PastTheInformation 12
        Case 13:
            PastTheInformation 13
        Case 14:
            PastTheInformation 14
        Case 15:
            PastTheInformation 15
        Case 16:
            PastTheInformation 16
        Case 17:
           PastTheInformation 17
        Case 18:
          PastTheInformation 18
        Case 19:
          PastTheInformation 19
        Case 20:
          PastTheInformation 20
        Case 21:
          PastTheInformation 21
        Case 22:
          PastTheInformation 22
        Case 23:
          PastTheInformation 23
        Case 24:
          PastTheInformation 24
        Case 25:
          PastTheInformation 25
          Case 26:
          PastTheInformation 26
        Case 27:
          PastTheInformation 28
        Case 29:
          PastTheInformation 29
         Case 30:
          PastTheInformation 30
    
    
    End Select
End Sub

Private Sub mnuremove_Click()
    If MsgBox("Remove This Vehicle to this BAY", vbQuestion + vbYesNo, "Are Youi Sure") = vbNo Then
        Exit Sub
    Else
        gconDMIS.Execute "Update CSMS_Baymonitoring set ro = null,bay_status = 'Available' where ro = '" & theRo & "'"
        Call cmd2_Click
    End If
End Sub

Private Sub thepic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If lbltheRO(Index).Caption = "" Then
            'Do Nothing
        Else
            theRo = lbltheRO(Index).Caption
            PopupMenu mnuOPTION
        End If
    End If
End Sub

Private Sub theTimer_Timer()
    If lblrostatus(thecolor).ForeColor = &HFFFFFF Then
        lblrostatus(thecolor).ForeColor = &HFFFF00
    Else
        lblrostatus(thecolor).ForeColor = &HFFFFFF
    End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
    If lblrostatus(thecolor).ForeColor = &HFFFFFF Then
        lblrostatus(thecolor).ForeColor = &HFFFF00
    Else
        lblrostatus(thecolor).ForeColor = &HFFFFFF
    End If
End Sub

Private Sub Timer2_Timer()
    CounterMe = CounterMe + 1
    If CounterMe = 60 Then
        CounterMe = 0
        cmd2_Click
    End If
End Sub

