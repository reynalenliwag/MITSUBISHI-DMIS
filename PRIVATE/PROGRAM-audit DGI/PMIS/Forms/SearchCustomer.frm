VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSMISSearchCustomer 
   BackColor       =   &H00FCFCFC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Customer"
   ClientHeight    =   6075
   ClientLeft      =   2835
   ClientTop       =   3270
   ClientWidth     =   10215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "SearchCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SearchTab 
      Height          =   6015
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   10610
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16579836
      TabCaption(0)   =   "By &Customer Name"
      TabPicture(0)   =   "SearchCustomer.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By &Product No."
      TabPicture(1)   =   "SearchCustomer.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "By &Invoice Number"
      TabPicture(2)   =   "SearchCustomer.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "By &Plate Number"
      TabPicture(3)   =   "SearchCustomer.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "By &Vehicle Model"
      TabPicture(4)   =   "SearchCustomer.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "By &Sales Executive"
      TabPicture(5)   =   "SearchCustomer.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture11"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture11 
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   22
         Top             =   60
         Width           =   10035
         Begin VB.PictureBox Picture12 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   24
            Top             =   30
            Width           =   1335
            Begin VB.Label Label6 
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
         Begin VB.TextBox txtSalesAE 
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
            Width           =   8535
         End
         Begin MSComctlLib.ListView ListSalesAE 
            Height          =   5055
            Left            =   30
            TabIndex        =   30
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777165
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":00B4
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SALES AE"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PLATE NUMBER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture9 
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   18
         Top             =   60
         Width           =   10035
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   20
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
               TabIndex        =   21
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtVehicleModel 
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
            Width           =   8535
         End
         Begin MSComctlLib.ListView ListVehicleModel 
            Height          =   5055
            Left            =   30
            TabIndex        =   29
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":03CE
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VEHICLE MODEL"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PLATE NUMBER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture7 
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   14
         Top             =   60
         Width           =   10035
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   16
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
               TabIndex        =   17
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtPlateNumber 
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
            TabIndex        =   15
            Top             =   30
            Width           =   8535
         End
         Begin MSComctlLib.ListView ListPlateNumber 
            Height          =   5055
            Left            =   30
            TabIndex        =   28
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":06E8
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   10
         Top             =   60
         Width           =   10035
         Begin VB.TextBox txtVI_NoNumber 
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
            TabIndex        =   13
            Top             =   30
            Width           =   8535
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   11
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
               TabIndex        =   12
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView ListVI_NoNumber 
            Height          =   5055
            Left            =   30
            TabIndex        =   27
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":0A02
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   6
         Top             =   60
         Width           =   10035
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   8
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
               TabIndex        =   9
               Top             =   0
               Width           =   1125
            End
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
            TabIndex        =   7
            Top             =   30
            Width           =   8535
         End
         Begin MSComctlLib.ListView ListProdNo 
            Height          =   5055
            Left            =   30
            TabIndex        =   26
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":0D1C
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   5595
         Left            =   30
         ScaleHeight     =   5535
         ScaleWidth      =   9975
         TabIndex        =   3
         Top             =   60
         Width           =   10035
         Begin VB.TextBox txtCustomerName 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   0
            Top             =   30
            Width           =   8535
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   4
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
               TabIndex        =   5
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView ListCustomerName 
            Height          =   5055
            Left            =   30
            TabIndex        =   1
            Top             =   450
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   8916
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
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "SearchCustomer.frx":1036
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6526
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PROD. NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "INVOICE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "ESTI AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSMISSearchCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree As New ADODB.Recordset
Dim Y, k As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Select Case SEARCH_TAB
          Case 0: If Trim(txtCustomerName) <> "" Then txtCustomerName.SetFocus Else Unload Me
          Case 1: If Trim(txtProdNo) <> "" Then txtProdNo.SetFocus Else Unload Me
          Case 2: If Trim(txtVI_NoNumber) <> "" Then txtVI_NoNumber.SetFocus Else Unload Me
          Case 3: If Trim(txtPlateNumber) <> "" Then txtPlateNumber.SetFocus Else Unload Me
          Case 4: If Trim(txtVehicleModel) <> "" Then txtVehicleModel.SetFocus Else Unload Me
          Case 5: If Trim(txtSalesAE) <> "" Then txtSalesAE.SetFocus Else Unload Me
   End Select
End If
If Shift = 2 Then
   Select Case KeyCode
          Case vbKeyC: SearchTab.Tab = 0
          Case vbKeyE: SearchTab.Tab = 1
          Case vbKeyI: SearchTab.Tab = 2
          Case vbKeyP: SearchTab.Tab = 3
          Case vbKeyV: SearchTab.Tab = 4
          Case vbKeyS: SearchTab.Tab = 5
   End Select
   SEARCH_TAB = SearchTab.Tab: SearchTab_Click (SEARCH_TAB)
End If
End Sub

Private Sub Form_Load()
CenterMe Screen, Me, 0
SearchTab.Tab = SEARCH_TAB
End Sub

Private Sub ListCustomerName_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
Unload Me
End Sub

Private Sub ListCustomerName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtCustomerName.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub ListCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
   Unload Me
End If
End Sub

Private Sub Listvi_noNumber_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListVI_NoNumber.SelectedItem.SubItems(2)))
Unload Me
End Sub

Private Sub ListPlateNumber_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
Unload Me
End Sub

Private Sub ListProdNo_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListProdNo.SelectedItem))
Unload Me
End Sub

Private Sub ListProdNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtProdNo.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub ListProdNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListProdNo.SelectedItem))
   Unload Me
End If
End Sub

Private Sub Listvi_noNumber_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtVI_NoNumber.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub Listvi_noNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListVI_NoNumber.SelectedItem.SubItems(2)))
   Unload Me
End If
End Sub

Private Sub ListPlateNumber_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtPlateNumber.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub ListPlateNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
   Unload Me
End If
End Sub

Private Sub listSalesAE_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListSalesAE.SelectedItem.SubItems(2)))
Unload Me
End Sub

Private Sub ListVehicleModel_DblClick()
frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
Unload Me
End Sub

Private Sub ListVehicleModel_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtVehicleModel.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub ListVehicleModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
   Unload Me
End If
End Sub

Private Sub listSalesAE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
   txtSalesAE.SetFocus
   SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub listSalesAE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   frmSMISVehicleInvoice.SearchProdNo (Trim(Me.ListSalesAE.SelectedItem.SubItems(2)))
   Unload Me
End If
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
SEARCH_TAB = SearchTab.Tab
DoEvents
txtCustomerName.Enabled = False: txtProdNo.Enabled = False
txtVI_NoNumber.Enabled = False: txtPlateNumber.Enabled = False
txtVehicleModel.Enabled = False: txtSalesAE.Enabled = False
ListCustomerName.Enabled = False: ListProdNo.Enabled = False
ListVI_NoNumber.Enabled = False: ListPlateNumber.Enabled = False
ListVehicleModel.Enabled = False: ListSalesAE.Enabled = False
Select Case SEARCH_TAB
       Case 0
            txtCustomerName.Enabled = True: ListCustomerName.Enabled = True
            Me.Caption = "Search Item by Customer Name"
            On Error Resume Next
            txtCustomerName.SetFocus
       Case 1
            txtProdNo.Enabled = True: ListProdNo.Enabled = True
            Me.Caption = "Search Item by Product Number"
            On Error Resume Next
            txtProdNo.SetFocus
       Case 2
            txtVI_NoNumber.Enabled = True: ListVI_NoNumber.Enabled = True
            Me.Caption = "Search Item by Vehicle Invoice Number"
            On Error Resume Next
            txtVI_NoNumber.SetFocus
       Case 3
            txtPlateNumber.Enabled = True: ListPlateNumber.Enabled = True
            Me.Caption = "Search Item by Plate Number Order"
            On Error Resume Next
            txtPlateNumber.SetFocus
       Case 4
            txtVehicleModel.Enabled = True: ListVehicleModel.Enabled = True
            Me.Caption = "Search Item by Vehicle Model"
            On Error Resume Next
            txtVehicleModel.SetFocus
       Case 5
            txtSalesAE.Enabled = True: ListSalesAE.Enabled = True
            Me.Caption = "Search Item by Sales Account Executive"
            On Error Resume Next
            txtSalesAE.SetFocus
End Select
End Sub

Private Sub txtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtCustomerName.Text) = "" Then
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      KeyCode = 0
   End If
End If
If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
   ListCustomerName.SetFocus
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCustomerName_Change()
If txtCustomerName = "" Then
   Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from PurchAgree inner join customer on PurchAgree.code = customer.code WHERE PurchAgree.DEALER_TYPE = " & DEALER_TYPE & " order by customer.lastname + ', ' + customer.firstname asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListCustomerName.ListItems, rsPurchAgree
   End If
Else
   Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from PurchAgree inner join customer on PurchAgree.code = customer.code Where PurchAgree.DEALER_TYPE = " & DEALER_TYPE & " AND customer.lastname + ', ' + customer.firstname like '" & Trim(Me.txtCustomerName) & "%' order by customer.lastname + ', ' + customer.firstname asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListCustomerName.ListItems, rsPurchAgree
   End If
End If
End Sub

Private Sub txtProdNo_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtProdNo.Text) = "" Then
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      KeyCode = 0
   End If
End If
If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
   ListProdNo.SetFocus
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtProdNo_Change()
If txtProdNo = "" Then
   Me.ListProdNo.Sorted = False: Me.ListProdNo.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.ProdNo,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code WHERE PurchAgree.DEALER_TYPE = " & DEALER_TYPE & " order by ProdNo asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListProdNo.ListItems, rsPurchAgree
   End If
Else
   Dim ProdNo As String
   ProdNo = UCase(txtProdNo.Text)
   Me.ListProdNo.Sorted = False: Me.ListProdNo.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.ProdNo,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code Where purchagree.DEALER_TYPE = " & DEALER_TYPE & " AND purchagree.ProdNo like '" & Trim(Me.txtProdNo) & "%' order by ProdNo asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListProdNo.ListItems, rsPurchAgree
   End If
End If
End Sub

Private Sub txtvi_noNumber_Change()
If txtVI_NoNumber = "" Then
   Me.ListVI_NoNumber.Sorted = False: Me.ListVI_NoNumber.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.vi_no,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code WHERE purchagree.DEALER_TYPE = " & DEALER_TYPE & " order by vi_no asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListVI_NoNumber.ListItems, rsPurchAgree
   End If
Else
   Dim vi_noNumber, vi_noNumber2, vi_noNumber3 As String
   vi_noNumber = UCase(txtVI_NoNumber.Text)
   If vi_noNumber <> "" Then
      If IsNumeric(vi_noNumber) = True Then
         vi_noNumber = Format(Right(vi_noNumber, 6), "000000")
      Else
         For k = 1 To Len(vi_noNumber)
             vi_noNumber2 = Mid(vi_noNumber, k, 1)
             If IsNumeric(vi_noNumber2) = True Then vi_noNumber3 = vi_noNumber3 + vi_noNumber2
         Next
         vi_noNumber = Format(vi_noNumber3, "000000")
      End If
   End If
   If IsNumeric(vi_noNumber) = True Then
      Me.ListVI_NoNumber.Sorted = False: Me.ListVI_NoNumber.ListItems.Clear
      Set rsPurchAgree = New ADODB.Recordset
      Set rsPurchAgree = gconSMIS.Execute("select purchagree.vi_no,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.plate_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code Where purchagree.DEALER_TYPE = " & DEALER_TYPE & " AND purchagree.vi_no like '" & Trim(Me.txtVI_NoNumber) & "%' order by vi_no asc")
      If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
         Listview_Loadval Me.ListVI_NoNumber.ListItems, rsPurchAgree
      End If
   End If
End If
End Sub

Private Sub txtPlateNumber_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtPlateNumber.Text) = "" Then
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      KeyCode = 0
   End If
End If
If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
   ListPlateNumber.SetFocus
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtPlateNumber_Change()
If txtPlateNumber = "" Then
   Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.plate_no,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code WHERE purchagree.DEALER_TYPE = " & DEALER_TYPE & " order by plate_no asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListPlateNumber.ListItems, rsPurchAgree
   End If
Else
   Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.plate_no,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.model,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code Where purchagree.DEALER_TYPE = " & DEALER_TYPE & " AND purchagree.plate_no like '" & Trim(Me.txtPlateNumber) & "%' order by plate_no asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListPlateNumber.ListItems, rsPurchAgree
   End If
End If
End Sub

Private Sub txtVehicleModel_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtVehicleModel.Text) = "" Then
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      KeyCode = 0
   End If
End If
If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
   ListVehicleModel.SetFocus
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVehicleModel_Change()
If txtVehicleModel = "" Then
   Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.model,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code WHERE purchagree.DEALER_TYPE = " & DEALER_TYPE & " order by model asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListVehicleModel.ListItems, rsPurchAgree
   End If
Else
   Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.model,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.salesAE,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code Where purchagree.DEALER_TYPE = " & DEALER_TYPE & " AND purchagree.model like '" & Trim(Me.txtVehicleModel) & "%' order by model asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListVehicleModel.ListItems, rsPurchAgree
   End If
End If
End Sub

Private Sub txtSalesAE_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtSalesAE.Text) = "" Then
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      KeyCode = 0
   End If
End If
If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
   ListSalesAE.SetFocus
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtSalesAE_Change()
If txtSalesAE = "" Then
   Me.ListSalesAE.Sorted = False: Me.ListSalesAE.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.salesAE,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code WHERE purchagree.DEALER_TYPE = " & DEALER_TYPE & " order by salesAE asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListSalesAE.ListItems, rsPurchAgree
   End If
Else
   Me.ListSalesAE.Sorted = False: Me.ListSalesAE.ListItems.Clear
   Set rsPurchAgree = New ADODB.Recordset
   Set rsPurchAgree = gconSMIS.Execute("select purchagree.salesAE,customer.lastname + ', ' + customer.firstname AS NIYM,purchagree.ProdNo,purchagree.vi_no,purchagree.plate_no,purchagree.model,purchagree.netsalesprice from customer inner join PurchAgree on customer.code = PurchAgree.code Where purchagree.DEALER_TYPE = " & DEALER_TYPE & " AND purchagree.salesAE like '" & Trim(Me.txtSalesAE) & "%' order by salesAE asc")
   If Not (rsPurchAgree.EOF And rsPurchAgree.BOF) Then
      Listview_Loadval Me.ListSalesAE.ListItems, rsPurchAgree
   End If
End If
End Sub

Sub clearListView()
For Y = 1 To Me.ListCustomerName.ListItems.Count
    If Me.ListCustomerName.ListItems.Count <= 0 Then Exit For
    Me.ListCustomerName.Sorted = False
    Me.ListCustomerName.ListItems.Remove Me.ListCustomerName.SelectedItem.Index
Next Y
For Y = 1 To Me.ListProdNo.ListItems.Count
    If Me.ListProdNo.ListItems.Count <= 0 Then Exit For
    Me.ListProdNo.Sorted = False
    Me.ListProdNo.ListItems.Remove Me.ListProdNo.SelectedItem.Index
Next Y
For Y = 1 To Me.ListVI_NoNumber.ListItems.Count
    If Me.ListVI_NoNumber.ListItems.Count <= 0 Then Exit For
    Me.ListVI_NoNumber.Sorted = False
    Me.ListVI_NoNumber.ListItems.Remove Me.ListVI_NoNumber.SelectedItem.Index
Next Y
For Y = 1 To Me.ListPlateNumber.ListItems.Count
    If Me.ListPlateNumber.ListItems.Count <= 0 Then Exit For
    Me.ListPlateNumber.Sorted = False
    Me.ListPlateNumber.ListItems.Remove Me.ListPlateNumber.SelectedItem.Index
Next Y
For Y = 1 To Me.ListVehicleModel.ListItems.Count
    If Me.ListVehicleModel.ListItems.Count <= 0 Then Exit For
    Me.ListVehicleModel.Sorted = False
    Me.ListVehicleModel.ListItems.Remove Me.ListVehicleModel.SelectedItem.Index
Next Y
For Y = 1 To Me.ListSalesAE.ListItems.Count
    If Me.ListSalesAE.ListItems.Count <= 0 Then Exit For
    Me.ListSalesAE.Sorted = False
    Me.ListSalesAE.ListItems.Remove Me.ListSalesAE.SelectedItem.Index
Next Y
End Sub

Function SetCustName(XXX As Variant) As String
Dim rsCustomerName As ADODB.Recordset
Set rsCustomerName = New ADODB.Recordset
    rsCustomerName.Open "select * from Customer where code = '" & XXX & "'", gconSMIS
If Not rsCustomerName.EOF And Not rsCustomerName.BOF Then
   SetCustName = UCase(Null2String(rsCustomerName!Lastname) & ", " & Null2String(rsCustomerName!Firstname))
End If
End Function
