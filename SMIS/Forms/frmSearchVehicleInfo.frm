VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_SearchVehicleInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Vehicle Information"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8760
   ClipControls    =   0   'False
   ForeColor       =   &H00FCFCFC&
   Icon            =   "frmSearchVehicleInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture11 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8760
      TabIndex        =   26
      Top             =   6105
      Width           =   8760
      Begin VB.Label labP 
         Caption         =   "Posted Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   32
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5670
         TabIndex        =   31
         Top             =   -30
         Width           =   285
      End
      Begin VB.Label labU 
         Caption         =   "Unposted Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3240
         TabIndex        =   30
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2910
         TabIndex        =   29
         Top             =   -30
         Width           =   285
      End
      Begin VB.Label labC 
         Caption         =   "Cancelled Transaction"
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
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         TabIndex        =   27
         Top             =   -30
         Width           =   285
      End
   End
   Begin TabDlg.SSTab SearchTab 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   10716
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "By &Make"
      TabPicture(0)   =   "frmSearchVehicleInfo.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By M&odel"
      TabPicture(1)   =   "frmSearchVehicleInfo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "By &Description"
      TabPicture(2)   =   "frmSearchVehicleInfo.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "By &Prod No."
      TabPicture(3)   =   "frmSearchVehicleInfo.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "By Conduction Sticker #"
      TabPicture(4)   =   "frmSearchVehicleInfo.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "PO Number"
      TabPicture(5)   =   "frmSearchVehicleInfo.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picpo"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "By RR #"
      TabPicture(6)   =   "frmSearchVehicleInfo.frx":03B2
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "lstRR"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Picture12"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "txtRR"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      Begin VB.TextBox txtRR 
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
         Left            =   1440
         TabIndex        =   41
         Top             =   120
         Width           =   6550
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   38
         Top             =   120
         Width           =   1335
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   39
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.PictureBox picpo 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   33
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtponumber 
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
            Left            =   1380
            TabIndex        =   36
            Top             =   30
            Width           =   6550
         End
         Begin VB.PictureBox pc1 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   34
            Top             =   30
            Width           =   1335
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   35
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstponumber 
            Height          =   5025
            Left            =   30
            TabIndex        =   37
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":03CE
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PO NUMBER"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "MAKE"
               Object.Width           =   3616
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MODEL"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "DESCRIPTION"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "PROD NO"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "STATUS"
               Object.Width           =   353
            EndProperty
         End
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   21
         Top             =   90
         Width           =   7965
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   24
            Top             =   30
            Width           =   1335
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   25
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtIgnitionKey 
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
            Left            =   1380
            TabIndex        =   23
            Top             =   30
            Width           =   6550
         End
         Begin MSComctlLib.ListView lstIgnitionKey 
            Height          =   5025
            Left            =   30
            TabIndex        =   22
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":06E8
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IGNKEY"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "MAKE"
               Object.Width           =   3616
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MODEL"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "DESCRIPTION"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "PROD NO"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "STATUS"
               Object.Width           =   353
            EndProperty
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   16
         Top             =   90
         Width           =   7965
         Begin MSComctlLib.ListView lstProdNo 
            Height          =   5025
            Left            =   30
            TabIndex        =   20
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":0A02
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PROD NO"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "MAKE"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MODEL"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "DESCRIPTION"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "IGNKEY"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "STATUS"
               Object.Width           =   353
            EndProperty
         End
         Begin VB.TextBox txtProdNo 
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
            Left            =   1380
            TabIndex        =   19
            Top             =   30
            Width           =   6550
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   17
            Top             =   30
            Width           =   1335
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   18
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   11
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtDesc 
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
            Left            =   1380
            TabIndex        =   14
            Top             =   30
            Width           =   6550
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   12
            Top             =   30
            Width           =   1335
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   13
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstDesc 
            Height          =   5025
            Left            =   30
            TabIndex        =   15
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":0D1C
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "DESCRIPTION"
               Object.Width           =   6625
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "MODEL"
               Object.Width           =   5381
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MAKE"
               Object.Width           =   1853
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "PROD NO"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "IGNKEY"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   6
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtModel 
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
            Left            =   1380
            TabIndex        =   9
            Top             =   30
            Width           =   6550
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   7
            Top             =   30
            Width           =   1335
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   8
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstModel 
            Height          =   5025
            Left            =   30
            TabIndex        =   10
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":1036
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "MODEL"
               Object.Width           =   3617
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DESCRIPTION"
               Object.Width           =   6624
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MAKE"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "PROD NO"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "IGNKEY"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   1
         Top             =   90
         Width           =   7965
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   3
            Top             =   30
            Width           =   1335
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   4
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtMake 
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
            Left            =   1380
            TabIndex        =   2
            Top             =   30
            Width           =   6550
         End
         Begin MSComctlLib.ListView lstMake 
            Height          =   5025
            Left            =   30
            TabIndex        =   5
            Top             =   480
            Width           =   7900
            _ExtentX        =   13944
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmSearchVehicleInfo.frx":1350
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "MAKE"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DESCRIPTION"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MODEL"
               Object.Width           =   3616
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "PROD NO"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "IGNKEY"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lstRR 
         Height          =   5025
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   7900
         _ExtentX        =   13944
         _ExtentY        =   8864
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15920873
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmSearchVehicleInfo.frx":166A
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "RR NO"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPT"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MODEL"
            Object.Width           =   5380
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MAKE"
            Object.Width           =   5380
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "IGNKEY"
            Object.Width           =   5380
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "STATUS"
            Object.Width           =   353
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "PROD NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSMIS_SearchVehicleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset
Dim Y, k                                                              As Long
Attribute k.VB_VarUserMemId = 1073938433

Sub ShowStatus()
    Dim RSSTATUS                                                      As ADODB.Recordset
    labC = "Cancelled Transaction  "
    labP = "Posted Transaction  "
    labU = "Unposted Transaction  "
    Dim STATUS_C                                                      As Integer
    Dim STATUS_P                                                      As Integer
    Dim STATUS_U                                                      As Integer

    Set RSSTATUS = gconDMIS.Execute("SELECT COUNT(*) T, STATUS FROM SMIS_MRRINV_TABLE  GROUP BY STATUS")
    If Not RSSTATUS.EOF Or Not RSSTATUS.BOF Then
        While Not RSSTATUS.EOF

            If Null2String(RSSTATUS!STATUS) = "" Or UCase(Null2String(RSSTATUS!STATUS)) = "U" Then
                STATUS_U = STATUS_U + RSSTATUS!T
            ElseIf UCase(Null2String(RSSTATUS!STATUS)) = "C" Then
                STATUS_C = STATUS_C + RSSTATUS!T
            Else
                STATUS_P = STATUS_P + RSSTATUS!T
            End If
            RSSTATUS.MoveNext
        Wend
        labU = "Unposted Transaction (" & STATUS_U & ")"
        labC = "Cancelled Transaction (" & STATUS_C & ")"
        labP = "Posted Transaction (" & STATUS_P & ")"
    End If
End Sub

Sub SetColorX(colorx As OLE_COLOR, lstitem As ListItem)
    Dim i
    lstitem.ForeColor = colorx
    For i = 1 To lstitem.ListSubItems.Count - 1
        lstitem.ListSubItems(i).ForeColor = colorx
    Next

End Sub

Private Sub Form_Activate()
    Select Case SEARCH_TAB
        Case 0
            On Error Resume Next
            txtMake.SetFocus

        Case 1
            On Error Resume Next
            txtModel.SetFocus

        Case 2
            On Error Resume Next
            txtDesc.SetFocus

        Case 3
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            On Error Resume Next
            txtIgnitionKey.SetFocus
        Case 5
            On Error Resume Next
            txtponumber.SetFocus
        Case 6
            On Error Resume Next
            txtRR.SetFocus
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
            Case 0:
                If Trim(txtMake) <> "" Then
                    On Error Resume Next
                    txtMake.SetFocus
                Else
                    Unload Me
                End If
            Case 1:
                If Trim(txtModel) <> "" Then
                    On Error Resume Next
                    txtModel.SetFocus
                Else
                    Unload Me
                End If

            Case 2:
                If Trim(txtDesc) <> "" Then
                    On Error Resume Next
                    txtDesc.SetFocus
                Else
                    Unload Me
                End If
            Case 3:
                If Trim(txtProdNo) <> "" Then
                    On Error Resume Next
                    txtProdNo.SetFocus
                Else
                    Unload Me
                End If
            Case 4:
                If Trim(txtIgnitionKey) <> "" Then
                    On Error Resume Next
                    txtIgnitionKey.SetFocus
                Else
                    Unload Me
                End If
            Case 5:
                If Trim(txtponumber) <> "" Then
                    On Error Resume Next
                    txtponumber.SetFocus
                Else
                    Unload Me
                End If
            Case 6:
                If Trim(txtRR) <> "" Then
                    On Error Resume Next
                    txtRR.SetFocus
                Else
                    Unload Me
                End If
        End Select
    End If
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyO: SearchTab.Tab = 0
            Case vbKeyM: SearchTab.Tab = 1
            Case vbKeyD: SearchTab.Tab = 2
            Case vbKeyP: SearchTab.Tab = 3
            Case vbKeyI: SearchTab.Tab = 4
        End Select
        SEARCH_TAB = SearchTab.Tab: SearchTab_Click (SEARCH_TAB)
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"



    SearchTab_Click SearchTab.Tab
    SearchTab.Tab = SEARCH_TAB
    ShowStatus
End Sub

Private Sub lstDesc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDesc
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With

End Sub

Private Sub lstIgnitionKey_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstIgnitionKey
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstIgnitionKey_DblClick()

    If lstIgnitionKey.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstIgnitionKey.SelectedItem.ListSubItems(6).Text))
    Unload Me
End Sub

Private Sub lstIgnitionKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstIgnitionKey_DblClick
    End If
End Sub

Private Sub lstIgnitionKey_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtIgnitionKey.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub lstMake_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstMake
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With

End Sub

Private Sub lstModel_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstModel
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With

End Sub

Private Sub LstModel_DblClick()
    If lstModel.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstModel.SelectedItem.ListSubItems(6).Text))
    Unload Me
End Sub

Private Sub LstModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstModel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LstModel_DblClick
    End If
End Sub

Private Sub lstponumber_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstponumber
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstponumber_DblClick()
    If lstponumber.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstponumber.SelectedItem.ListSubItems(6).Text))
    Unload Me
End Sub

Private Sub lstponumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstponumber_DblClick
    End If
End Sub

Private Sub lstponumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtponumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub lstProdNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstProdNo
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstProdNo_DblClick()
    If lstProdNo.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstProdNo.SelectedItem.ListSubItems(6).Text))
    Unload Me
End Sub

Private Sub lstProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstProdNo_DblClick
    End If
End Sub

Private Sub lstProdNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtProdNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub


'************************ UPDATED BY DHANG_ERZ 102715
'                               BY RR#

Private Sub lstRR_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     With lstRR
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With

End Sub
Private Sub lstRR_DblClick()
    If lstRR.SelectedItem Is Nothing Then Exit Sub
'    If FORMULTIPLE = True Then
'        frmSMIS_Trans_MRR_Multi.SearchID (Trim(Me.lstRR.SelectedItem.ListSubItems(7).Text))
'    Else
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstRR.SelectedItem.ListSubItems(7).Text))
'    End If
    Unload Me
End Sub
'Private Sub lstponumber_DblClick()
'    If lstponumber.SelectedItem Is Nothing Then Exit Sub
'    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstponumber.SelectedItem.ListSubItems(6).Text))
'    Unload Me
'End Sub


Private Sub lstRR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstRR_DblClick
    End If
End Sub

Private Sub lstRR_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtRR.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub txtIgnitionKey_Change()
    On Error GoTo ErrorCode:

    If txtIgnitionKey = "" Then
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper(IgnKey), make, descript, model, ProdNo ,status,ID  from SMIS_MrrInv_Table order by IgnKey asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsMRRINV
        End If
    Else
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper(IgnKey), make, descript, model, ProdNo, status,ID  from SMIS_MrrInv_Table WHERE IgnKey like '" & Trim(Repleys(Me.txtIgnitionKey)) & "%' order by IgnKey asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsMRRINV
        End If
    End If

    Dim i
    For i = 1 To lstIgnitionKey.ListItems.Count
        If lstIgnitionKey.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstIgnitionKey.ListItems(i)
        ElseIf lstIgnitionKey.ListItems(i).ListSubItems(5).Text = "" Or lstIgnitionKey.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstIgnitionKey.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstIgnitionKey
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtIgnitionKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtIgnitionKey.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then

        If lstIgnitionKey.ListItems.Count > 0 And lstIgnitionKey.Enabled = True Then: lstIgnitionKey.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtModel.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstModel.ListItems.Count > 0 And lstModel.Enabled = True Then: lstModel.SetFocus

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtModel_Change()
    On Error GoTo ErrorCode:

    If txtModel = "" Then
        Me.lstModel.Sorted = False: Me.lstModel.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper(model), descript, make, ProdNo, IgnKey ,status,id from SMIS_MrrInv_Table order by [model] asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstModel.ListItems, rsMRRINV
        End If
    Else
        Me.lstModel.Sorted = False: Me.lstModel.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( model), descript, make, ProdNo, IgnKey, status,id from SMIS_MrrInv_Table WHERE model like '" & Trim(ReplaceQuote(Me.txtModel)) & "%' order by model asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstModel.ListItems, rsMRRINV
        End If
    End If


    Dim i
    For i = 1 To lstModel.ListItems.Count
        If lstModel.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstModel.ListItems(i)
        ElseIf lstModel.ListItems(i).ListSubItems(5).Text = "" Or lstModel.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstModel.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstModel



    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub LstMake_DblClick()
    If lstMake.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstMake.SelectedItem.ListSubItems(6).Text))


    Unload Me
End Sub

Private Sub LstMake_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstMake_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LstMake_DblClick
    End If
End Sub

Private Sub txtMake_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtMake.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then

        If lstMake.ListItems.Count > 0 And lstMake.Enabled = True Then: lstMake.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtMake_Change()
    On Error GoTo ErrorCode:

    If txtMake = "" Then
        Me.lstMake.Sorted = False: Me.lstMake.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper(make) , descript, model, ProdNo, IgnKey ,status,id from SMIS_MrrInv_Table order by make asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstMake.ListItems, rsMRRINV
        End If
    Else
        Me.lstMake.Sorted = False: Me.lstMake.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper(make) ,descript,model, ProdNo, IgnKey, status,id from SMIS_MrrInv_Table WHERE make like '" & Trim(ReplaceQuote(Me.txtMake)) & "%' order by make asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstMake.ListItems, rsMRRINV
        End If
    End If

    Dim i
    For i = 1 To lstMake.ListItems.Count
        If lstMake.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstMake.ListItems(i)
        ElseIf lstMake.ListItems(i).ListSubItems(5).Text = "" Or lstMake.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstMake.ListItems(i)
        End If
    Next

    LV_AutoSizeColumn lstMake

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub LstDesc_DblClick()
    If lstDesc.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstDesc.SelectedItem.ListSubItems(6).Text))
    Unload Me
End Sub

Private Sub LstDesc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtDesc.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LstDesc_DblClick
    End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtDesc.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstDesc.ListItems.Count > 0 And lstDesc.Enabled = True Then: lstDesc.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtDesc_Change()
    On Error GoTo ErrorCode:

    If txtDesc = "" Then
        Me.lstDesc.Sorted = False: Me.lstDesc.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( descript),model,make, ProdNo, IgnKey, status,id from SMIS_MrrInv_Table order by descript asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstDesc.ListItems, rsMRRINV
        End If
    Else
        Me.lstDesc.Sorted = False: Me.lstDesc.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( descript), model, make, ProdNo, IgnKey, status,id from SMIS_MrrInv_Table WHERE descript like '" & Trim(Me.txtDesc) & "%' order by descript asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstDesc.ListItems, rsMRRINV
        End If
    End If


    Dim i
    For i = 1 To lstDesc.ListItems.Count
        If lstDesc.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstDesc.ListItems(i)
        ElseIf lstDesc.ListItems(i).ListSubItems(5).Text = "" Or lstDesc.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstDesc.ListItems(i)
        End If
    Next

    LV_AutoSizeColumn lstDesc

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab
    DoEvents
    Select Case SEARCH_TAB
        Case 0
            txtMake.Enabled = True: lstMake.Enabled = True
            Me.Caption = "Search Item by Make"
            txtMake_Change
            On Error Resume Next
            txtMake.SetFocus

        Case 1
            txtModel.Enabled = True: lstModel.Enabled = True
            Me.Caption = "Search Item by Vehicle Model"
            txtModel_Change
            On Error Resume Next
            txtModel.SetFocus

        Case 2
            txtDesc.Enabled = True: lstDesc.Enabled = True
            Me.Caption = "Search Item by Description"
            txtDesc_Change
            On Error Resume Next
            txtDesc.SetFocus

        Case 3
            txtProdNo.Enabled = True: lstProdNo.Enabled = True
            Me.Caption = "Search Item by Product Number"
            txtProdNo_Change
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            txtIgnitionKey.Enabled = True: lstIgnitionKey.Enabled = True
            Me.Caption = "Search Item by Conduction Sticker Number"
            On Error Resume Next
            txtIgnitionKey_Change
            txtIgnitionKey.SetFocus
        Case 5
            txtponumber.Enabled = True: lstponumber.Enabled = True
            Me.Caption = "Search Item by PO Number"
            On Error Resume Next
            txtponumber_Change
            txtponumber.SetFocus
        Case 6
            txtRR.Enabled = True: lstRR.Enabled = True
            Me.Caption = "Search Item by RR Number"
            On Error Resume Next
            txtRR_Change
'            txtRR.SetFocus
    End Select
End Sub

Private Sub txtponumber_Change()
    On Error GoTo ErrorCode:

    If txtponumber = "" Then
        Me.lstponumber.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( PONO), make, descript, model, IgnKey ,status,ID  from SMIS_MrrInv_Table order by PONO asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstponumber.ListItems, rsMRRINV
        End If
    Else
        Me.lstponumber.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( PONO), make,descript,make, IgnKey, status,ID from SMIS_MrrInv_Table WHERE PONO like '" & Trim(Me.txtponumber) & "%' order by PONO asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstponumber.ListItems, rsMRRINV
        End If
    End If

    Dim i
    For i = 1 To lstponumber.ListItems.Count
        If lstponumber.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstponumber.ListItems(i)
        ElseIf lstProdNo.ListItems(i).ListSubItems(5).Text = "" Or lstponumber.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstponumber.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstponumber



    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtProdNo_Change()
    On Error GoTo ErrorCode:

    If txtProdNo = "" Then
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( ProdNo), make, descript, model, IgnKey ,status,ID  from SMIS_MrrInv_Table order by prodno asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsMRRINV
        End If
    Else
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select upper( ProdNo), make,descript,make, IgnKey, status,ID from SMIS_MrrInv_Table WHERE prodno like '" & Trim(Me.txtProdNo) & "%' order by prodno asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsMRRINV
        End If
    End If

    Dim i
    For i = 1 To lstProdNo.ListItems.Count
        If lstProdNo.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstProdNo.ListItems(i)
        ElseIf lstProdNo.ListItems(i).ListSubItems(5).Text = "" Or lstProdNo.ListItems(i).ListSubItems(5).Text = "U" Then
            SetColorX vbBlue, lstProdNo.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstProdNo



    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtProdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtProdNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstProdNo.ListItems.Count > 0 And lstProdNo.Enabled = True Then: lstProdNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

'********************************** UPDATED BY DHANG_ERZ 102715
'                                   BY RR#

Private Sub txtRR_Change()
    On Error GoTo ErrorCode:

    If txtRR = "" Then
        Me.lstRR.Sorted = False: Me.lstRR.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select CODE, DESCRIPT, MODEL,  MAKE, IGNKEY, STATUS, PRODNO, ID  from SMIS_MrrInv_Table order by refPONO asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstRR.ListItems, rsMRRINV
        End If
    Else
        Me.lstRR.Sorted = False: Me.lstRR.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select CODE, DESCRIPT, MODEL,  MAKE, IGNKEY, STATUS, PRODNO, ID from SMIS_MrrInv_Table WHERE CODE like '%" & Trim(Me.txtRR) & "%' order by refPONO asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstRR.ListItems, rsMRRINV
        End If
    End If

    Dim i
    For i = 1 To lstRR.ListItems.Count
        If lstRR.ListItems(i).ListSubItems(5).Text = "C" Then
            SetColorX vbRed, lstRR.ListItems(i)
        ElseIf lstRR.ListItems(i).ListSubItems(5).Text = "" Or lstRR.ListItems(i).ListSubItems(1).Text = "U" Then
            SetColorX vbBlue, lstRR.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstRR

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

