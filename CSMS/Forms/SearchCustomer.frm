VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSSearchCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6255
   ClientLeft      =   2835
   ClientTop       =   3390
   ClientWidth     =   10230
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "SearchCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10230
   Begin TabDlg.SSTab SearchTab 
      Height          =   6015
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   10610
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "By &Customer Name"
      TabPicture(0)   =   "SearchCustomer.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By &Repair Order"
      TabPicture(1)   =   "SearchCustomer.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "By &Invoice Number"
      TabPicture(2)   =   "SearchCustomer.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "By &Plate Number"
      TabPicture(3)   =   "SearchCustomer.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "By &Vehicle Model"
      TabPicture(4)   =   "SearchCustomer.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture9"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "By &Service Adviser"
      TabPicture(5)   =   "SearchCustomer.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture11"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox Picture11 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5595
         ScaleWidth      =   10035
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
         Begin VB.TextBox txtServiceAdviser 
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
         Begin MSComctlLib.ListView ListServiceAdviser 
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":03B2
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. ADVISER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6527
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "REPAIR ORDER"
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
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5595
         ScaleWidth      =   10035
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":06CC
            NumItems        =   8
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
               Text            =   "REPAIR ORDER"
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
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5595
         ScaleWidth      =   10035
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":09E6
            NumItems        =   8
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
               Text            =   "REPAIR ORDER"
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
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5595
         ScaleWidth      =   10035
         TabIndex        =   10
         Top             =   60
         Width           =   10035
         Begin VB.TextBox txtInvoiceNumber 
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
         Begin MSComctlLib.ListView ListInvoiceNumber 
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":0D00
            NumItems        =   8
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
               Text            =   "REPAIR ORDER"
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
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74970
         ScaleHeight     =   5595
         ScaleWidth      =   10035
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
         Begin VB.TextBox txtRepairOrder 
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
         Begin MSComctlLib.ListView ListRepairOrder 
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":101A
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "REPAIR ORDER"
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
               SubItemIndex    =   3
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PLATE NUMBER"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "VEHICLE MODEL"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "S. ADVISER"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   30
         ScaleHeight     =   5595
         ScaleWidth      =   10035
         TabIndex        =   3
         Top             =   60
         Width           =   10035
         Begin VB.TextBox txtCustomerName 
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
            BackColor       =   16777215
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
            MouseIcon       =   "SearchCustomer.frx":1334
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   6526
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "REPAIR ORDER"
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
               Text            =   "RO AMOUNT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "STATUS"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmCSMSSearchCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREPOR                                            As New ADODB.Recordset
Dim y                                                  As Long
Dim k                                                  As Long

Sub clearListView()
    For y = 1 To Me.ListCustomerName.ListItems.Count
        If Me.ListCustomerName.ListItems.Count <= 0 Then Exit For
        Me.ListCustomerName.Sorted = False
        Me.ListCustomerName.ListItems.Remove Me.ListCustomerName.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListRepairOrder.ListItems.Count
        If Me.ListRepairOrder.ListItems.Count <= 0 Then Exit For
        Me.ListRepairOrder.Sorted = False
        Me.ListRepairOrder.ListItems.Remove Me.ListRepairOrder.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListInvoiceNumber.ListItems.Count
        If Me.ListInvoiceNumber.ListItems.Count <= 0 Then Exit For
        Me.ListInvoiceNumber.Sorted = False
        Me.ListInvoiceNumber.ListItems.Remove Me.ListInvoiceNumber.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListPlateNumber.ListItems.Count
        If Me.ListPlateNumber.ListItems.Count <= 0 Then Exit For
        Me.ListPlateNumber.Sorted = False
        Me.ListPlateNumber.ListItems.Remove Me.ListPlateNumber.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListVehicleModel.ListItems.Count
        If Me.ListVehicleModel.ListItems.Count <= 0 Then Exit For
        Me.ListVehicleModel.Sorted = False
        Me.ListVehicleModel.ListItems.Remove Me.ListVehicleModel.SelectedItem.INDEX
    Next y
    For y = 1 To Me.ListServiceAdviser.ListItems.Count
        If Me.ListServiceAdviser.ListItems.Count <= 0 Then Exit For
        Me.ListServiceAdviser.Sorted = False
        Me.ListServiceAdviser.ListItems.Remove Me.ListServiceAdviser.SelectedItem.INDEX
    Next y
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF3 Then
        Select Case SEARCH_TAB
            Case 0: txtCustomerName.SetFocus
            Case 1: txtRepairOrder.SetFocus
            Case 2: txtInvoiceNumber.SetFocus
            Case 3: txtPlateNumber.SetFocus
            Case 4: txtVehicleModel.SetFocus
            Case 5: txtServiceAdviser.SetFocus
        End Select
    End If

    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
            Case 0: If Trim(txtCustomerName) <> "" Then On Error Resume Next: txtCustomerName.SetFocus Else Unload Me
            Case 1: If Trim(txtRepairOrder) <> "" Then On Error Resume Next: txtRepairOrder.SetFocus Else Unload Me
            Case 2: If Trim(txtInvoiceNumber) <> "" Then On Error Resume Next: txtInvoiceNumber.SetFocus Else Unload Me
            Case 3: If Trim(txtPlateNumber) <> "" Then On Error Resume Next: txtPlateNumber.SetFocus Else Unload Me
            Case 4: If Trim(txtVehicleModel) <> "" Then On Error Resume Next: txtVehicleModel.SetFocus Else Unload Me
            Case 5: If Trim(txtServiceAdviser) <> "" Then On Error Resume Next: txtServiceAdviser.SetFocus Else Unload Me
        End Select
    End If
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyC: SearchTab.Tab = 0
            Case vbKeyR: SearchTab.Tab = 1
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
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SearchTab.Tab = SEARCH_TAB
    If SEARCH_TAB = 0 Then txtCustomerName.Text = SEARCHCUSTOMERNAME
    If SEARCH_TAB = 3 Then txtPlateNumber.Text = SEARCHPLATENO
End Sub

Private Sub ListCustomerName_DblClick()
    SEARCHCUSTOMERNAME = txtCustomerName.Text
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
    Unload Me
End Sub

Private Sub ListCustomerName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtCustomerName.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListCustomerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SEARCHCUSTOMERNAME = txtCustomerName.Text
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListCustomerName.SelectedItem.SubItems(1)))
        Unload Me
    End If
End Sub

Private Sub ListInvoiceNumber_DblClick()
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListInvoiceNumber.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListPlateNumber_DblClick()
    SEARCHPLATENO = txtPlateNumber.Text
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListRepairOrder_DblClick()
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListRepairOrder.SelectedItem))
    Unload Me
End Sub

Private Sub ListRepairOrder_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtRepairOrder.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListRepairOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListRepairOrder.SelectedItem))
        Unload Me
    End If
End Sub

Private Sub ListInvoiceNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtInvoiceNumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListInvoiceNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListInvoiceNumber.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListPlateNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtPlateNumber.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListPlateNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SEARCHPLATENO = txtPlateNumber.Text
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListPlateNumber.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListServiceAdviser_DblClick()
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListServiceAdviser.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListVehicleModel_DblClick()
    frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
    Unload Me
End Sub

Private Sub ListVehicleModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtVehicleModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListVehicleModel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListVehicleModel.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub ListServiceAdviser_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtServiceAdviser.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub ListServiceAdviser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmCSMSDataEntry.SearchRepairOrder (Trim(Me.ListServiceAdviser.SelectedItem.SubItems(2)))
        Unload Me
    End If
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab
    DoEvents
    txtCustomerName.Enabled = False: txtRepairOrder.Enabled = False
    txtInvoiceNumber.Enabled = False: txtPlateNumber.Enabled = False
    txtVehicleModel.Enabled = False: txtServiceAdviser.Enabled = False
    ListCustomerName.Enabled = False: ListRepairOrder.Enabled = False
    ListInvoiceNumber.Enabled = False: ListPlateNumber.Enabled = False
    ListVehicleModel.Enabled = False: ListServiceAdviser.Enabled = False
    Select Case SEARCH_TAB
        Case 0
            txtCustomerName.Enabled = True: ListCustomerName.Enabled = True
            Me.Caption = "Search Item by Customer Name"
            On Error Resume Next
            txtCustomerName.SetFocus
        Case 1
            txtRepairOrder.Enabled = True: ListRepairOrder.Enabled = True
            Me.Caption = "Search Item by Repair Order Number"
            On Error Resume Next
            txtRepairOrder.SetFocus
        Case 2
            txtInvoiceNumber.Enabled = True: ListInvoiceNumber.Enabled = True
            Me.Caption = "Search Item by Invoice Number"
            On Error Resume Next
            txtInvoiceNumber.SetFocus
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
            txtServiceAdviser.Enabled = True: ListServiceAdviser.Enabled = True
            Me.Caption = "Search Item by Service Adviser"
            On Error Resume Next
            txtServiceAdviser.SetFocus
    End Select
End Sub

Private Sub txtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCustomerName.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListCustomerName.ListItems.Count > 0 And ListCustomerName.Enabled = True Then
            ListCustomerName.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCustomerName_Change()
    If txtCustomerName = "" Then
        ListCustomerName.Enabled = False
        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.niym,CSMS_RepOr.rep_or, " & _
            "CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by, " & _
            " CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join " & _
            " CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO WHERE " & _
            " CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.niym asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsREPOR
            ListCustomerName.Enabled = True
        End If
    Else

        Me.ListCustomerName.Sorted = False: Me.ListCustomerName.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.niym,CSMS_RepOr.rep_or, " & _
            " CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by, " & _
            " CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join " & _
            " CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where " & _
            " CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' " & _
            " AND niym like '" & Repleys(Trim(Me.txtCustomerName)) & "%' order by CSMS_RepOr.niym asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListCustomerName.ListItems, rsREPOR
            ListCustomerName.Enabled = True
        End If
    End If
End Sub

Private Sub txtRepairOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtRepairOrder.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListRepairOrder.ListItems.Count > 0 And ListRepairOrder.Enabled = True Then
            ListRepairOrder.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtRepairOrder_Change()
    If txtRepairOrder = "" Then
        ListRepairOrder.Enabled = False
        Me.ListRepairOrder.Sorted = False: Me.ListRepairOrder.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        'Set rsREPOR = gconDMIS.Execute("select CSMS_RepOr.rep_or,CSMS_RepOr.niym,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TransType = 'R' order by CSMS_RepOr.rep_or asc")
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.rep_or,CSMS_RepOr.niym,CSMS_RepOr.invoice,CSMS_RepairOrder.Status,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TransType = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.rep_or asc")

        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListRepairOrder.ListItems, rsREPOR
            ListRepairOrder.Enabled = True
        End If
    Else
        Dim RepairOrder As String, RepairOrder2 As String, RepairOrder3 As String

        RepairOrder = UCase(txtRepairOrder.Text)
        If RepairOrder <> "" Then
            If IsNumeric(RepairOrder) = True Then
                RepairOrder = Format(Left(RepairOrder, 1), "R-") & Format(Right(RepairOrder, 6), "00000000")
            Else
                For k = 1 To Len(RepairOrder)
                    RepairOrder2 = Mid(RepairOrder, k, 1)
                    If IsNumeric(RepairOrder2) = True Then RepairOrder3 = RepairOrder3 + RepairOrder2
                Next
                RepairOrder3 = Format(RepairOrder3, "00000000"): RepairOrder = Format(Left(RepairOrder3, 1), "R-") & Format(Right(RepairOrder3, 6), "00000000")
            End If
        End If
        If Left(RepairOrder, 2) = "R-" Then
            Me.ListRepairOrder.Sorted = False: Me.ListRepairOrder.ListItems.Clear
            Set rsREPOR = New ADODB.Recordset
            Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.rep_or,CSMS_RepOr.niym,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' AND CSMS_RepOr.rep_or like'" & Repleys(RepairOrder) & "%' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.rep_or asc")
            If Not (rsREPOR.EOF And rsREPOR.BOF) Then
                Listview_Loadval Me.ListRepairOrder.ListItems, rsREPOR
                ListRepairOrder.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub txtInvoiceNumber_Change()
    If txtInvoiceNumber = "" Then
        ListInvoiceNumber.Enabled = False
        Me.ListInvoiceNumber.Sorted = False: Me.ListInvoiceNumber.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.invoice,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO WHERE CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.invoice asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListInvoiceNumber.ListItems, rsREPOR
            ListInvoiceNumber.Enabled = True
        End If
    Else
        Dim InvoiceNumber, InvoiceNumber2, InvoiceNumber3 As String
        InvoiceNumber = UCase(txtInvoiceNumber.Text)
        If InvoiceNumber <> "" Then
            If IsNumeric(InvoiceNumber) = True Then
                InvoiceNumber = Format(Right(InvoiceNumber, 6), "000000")
            Else
                For k = 1 To Len(InvoiceNumber)
                    InvoiceNumber2 = Mid(InvoiceNumber, k, 1)
                    If IsNumeric(InvoiceNumber2) = True Then InvoiceNumber3 = InvoiceNumber3 + InvoiceNumber2
                Next
                InvoiceNumber = Format(InvoiceNumber3, "000000")
            End If
        End If
        If IsNumeric(InvoiceNumber) = True Then
            Me.ListInvoiceNumber.Sorted = False: Me.ListInvoiceNumber.ListItems.Clear
            Set rsREPOR = New ADODB.Recordset
            Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.invoice,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' AND CSMS_Repor.invoice like'" & InvoiceNumber & "%' order by CSMS_Repor.invoice asc")
            If Not (rsREPOR.EOF And rsREPOR.BOF) Then
                Listview_Loadval Me.ListInvoiceNumber.ListItems, rsREPOR
                ListInvoiceNumber.Enabled = True
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
        If ListPlateNumber.ListItems.Count > 0 And ListPlateNumber.Enabled = True Then
            ListPlateNumber.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtPlateNumber_Change()
    If txtPlateNumber = "" Then
        ListPlateNumber.Enabled = False
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.plate_no,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO WHERE CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.plate_no asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsREPOR
            ListPlateNumber.Enabled = True
        End If
    Else
        Me.ListPlateNumber.Sorted = False: Me.ListPlateNumber.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.plate_no,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.model,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' AND CSMS_RepOr.plate_no like '" & Repleys(Trim(Me.txtPlateNumber)) & "%' order by CSMS_RepOr.plate_no asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListPlateNumber.ListItems, rsREPOR
            ListPlateNumber.Enabled = True
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
        If ListVehicleModel.ListItems.Count > 0 And ListVehicleModel.Enabled = True Then
            ListVehicleModel.SetFocus
        End If

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtVehicleModel_Change()
    If txtVehicleModel = "" Then
        ListVehicleModel.Enabled = False
        Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.model,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO WHERE CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by model asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListVehicleModel.ListItems, rsREPOR
            ListVehicleModel.Enabled = True
        End If
    Else
        Me.ListVehicleModel.Sorted = False: Me.ListVehicleModel.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.model,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.recd_by,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' AND CSMS_Repor.model like '" & Repleys(Trim(Me.txtVehicleModel)) & "%' order by CSMS_Repor.model asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListVehicleModel.ListItems, rsREPOR
            ListVehicleModel.Enabled = True
        End If
    End If
End Sub

Private Sub txtServiceAdviser_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtServiceAdviser.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If ListServiceAdviser.ListItems.Count > 0 And ListServiceAdviser.Enabled = True Then
            ListServiceAdviser.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtServiceAdviser_Change()
    If txtServiceAdviser = "" Then
        ListServiceAdviser.Enabled = False
        Me.ListServiceAdviser.Sorted = False: Me.ListServiceAdviser.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.recd_by,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO Where CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R'  AND CSMS_RepOr.recd_by like '" & Trim(Me.txtServiceAdviser) & "%' order by CSMS_RepOr.recd_by asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListServiceAdviser.ListItems, rsREPOR
            ListServiceAdviser.Enabled = True
        End If
    Else
        Me.ListServiceAdviser.Sorted = False: Me.ListServiceAdviser.ListItems.Clear
        Set rsREPOR = New ADODB.Recordset
        Set rsREPOR = gconDMIS.Execute("select TOP 100 CSMS_RepOr.recd_by,CSMS_RepOr.niym,CSMS_RepOr.rep_or,CSMS_RepOr.invoice,CSMS_RepOr.plate_no,CSMS_RepOr.model,CSMS_RepOr.ro_amount,CSMS_RepairOrder.STATUS from CSMS_RepOr Inner Join CSMS_RepairOrder on CSMS_RepOr.Rep_Or = CSMS_RepairOrder.RO_NO WHERE CSMS_RepOr.TRANSTYPE = 'R' AND CSMS_REPAIRORDER.TRANSTYPE = 'R' order by CSMS_RepOr.recd_by asc")
        If Not (rsREPOR.EOF And rsREPOR.BOF) Then
            Listview_Loadval Me.ListServiceAdviser.ListItems, rsREPOR
            ListServiceAdviser.Enabled = True
        End If
    End If
End Sub

