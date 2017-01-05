VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSMIS_SearchVehicleSalesOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Vehicle Sales Order"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ForeColor       =   &H00FCFCFC&
   Icon            =   "SearchVehicleSalesOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8250
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
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   8250
      TabIndex        =   28
      Top             =   6075
      Width           =   8250
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
         Height          =   195
         Left            =   5760
         TabIndex        =   34
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5430
         TabIndex        =   33
         Top             =   30
         Width           =   285
      End
      Begin VB.Label labU 
         Caption         =   "Un Posted Transaction"
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
         Height          =   195
         Left            =   3000
         TabIndex        =   32
         Top             =   60
         Width           =   2415
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2670
         TabIndex        =   31
         Top             =   30
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
         Height          =   195
         Left            =   360
         TabIndex        =   30
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   30
         TabIndex        =   29
         Top             =   30
         Width           =   285
      End
   End
   Begin TabDlg.SSTab SearchTab 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   10716
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
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
      TabCaption(0)   =   "By &Date"
      TabPicture(0)   =   "SearchVehicleSalesOrder.frx":01CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By &Customer"
      TabPicture(1)   =   "SearchVehicleSalesOrder.frx":01E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "By &VSO No"
      TabPicture(2)   =   "SearchVehicleSalesOrder.frx":0202
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "By &Prod No."
      TabPicture(3)   =   "SearchVehicleSalesOrder.frx":021E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "By Conduction Sticker#"
      TabPicture(4)   =   "SearchVehicleSalesOrder.frx":023A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Picture9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture9 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   90
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   18
         Top             =   90
         Width           =   7965
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
            TabIndex        =   19
            Top             =   30
            Width           =   6585
         End
         Begin MSComctlLib.ListView lstIgnitionKey 
            Height          =   5025
            Left            =   0
            TabIndex        =   25
            Top             =   510
            Width           =   7935
            _ExtentX        =   13996
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
            MouseIcon       =   "SearchVehicleSalesOrder.frx":0256
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CS#"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sales Agent"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   14
         Top             =   90
         Width           =   7965
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
            TabIndex        =   17
            Top             =   30
            Width           =   6495
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   15
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
               TabIndex        =   16
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstProdNo 
            Height          =   5025
            Left            =   0
            TabIndex        =   24
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
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
            MouseIcon       =   "SearchVehicleSalesOrder.frx":0570
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Prod No"
               Object.Width           =   3616
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sales Agent"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   10
         Top             =   90
         Width           =   7965
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
            Width           =   6495
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
         Begin MSComctlLib.ListView lstInvoiceNo 
            Height          =   5025
            Left            =   30
            TabIndex        =   23
            Top             =   450
            Width           =   7845
            _ExtentX        =   13838
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
            MouseIcon       =   "SearchVehicleSalesOrder.frx":088A
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SO No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sales Agent"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   6
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtCustomer 
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
            Width           =   6495
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
         Begin MSComctlLib.ListView lstCustomer 
            Height          =   5025
            Left            =   0
            TabIndex        =   22
            Top             =   450
            Width           =   7845
            _ExtentX        =   13838
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
            MouseIcon       =   "SearchVehicleSalesOrder.frx":0BA4
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Customer Name"
               Object.Width           =   4498
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sales Agent"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   1
         Top             =   90
         Width           =   7965
         Begin VB.PictureBox Picture12 
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   6030
            ScaleHeight     =   345
            ScaleWidth      =   1785
            TabIndex        =   26
            Top             =   30
            Width           =   1845
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "MM/DD/YYYY"
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
               TabIndex        =   27
               Top             =   0
               Width           =   1635
            End
         End
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
         Begin VB.TextBox txtDate 
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
            Width           =   4635
         End
         Begin MSComctlLib.ListView lstDate 
            Height          =   5025
            Left            =   30
            TabIndex        =   5
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
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
            MouseIcon       =   "SearchVehicleSalesOrder.frx":0EBE
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2471
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Color"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "CS#"
               Object.Width           =   1852
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Sales Agent"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Status"
               Object.Width           =   0
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSMIS_SearchVehicleSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSalesInvoice                                                    As ADODB.Recordset
Dim SAFILTER                                                          As String

Sub ShowStatus()
    Dim RSSTATUS                                                      As ADODB.Recordset
    labC = "Cancelled Transaction  "
    labP = "Posted Transaction  "
    labU = "Unposted Transaction  "
    Dim STATUS_C                                                      As Integer
    Dim STATUS_P                                                      As Integer
    Dim STATUS_U                                                      As Integer

    Set RSSTATUS = gconDMIS.Execute("SELECT COUNT(*) T, SOSTATUS FROM SMIS_SALESORDER  GROUP BY SOSTATUS")
    If Not RSSTATUS.EOF Or Not RSSTATUS.BOF Then
        While Not RSSTATUS.EOF

            If Null2String(RSSTATUS!SOSTATUS) = "" Or UCase(Null2String(RSSTATUS!SOSTATUS)) = "U" Then
                STATUS_U = STATUS_U + RSSTATUS!T
            ElseIf UCase(Null2String(RSSTATUS!SOSTATUS)) = "C" Then
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
            txtDate.SetFocus
        Case 1
            On Error Resume Next
            txtCustomer.SetFocus
        Case 2
            On Error Resume Next
            txtInvoiceNumber.SetFocus
        Case 3
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            On Error Resume Next
            txtIgnitionKey.SetFocus
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
            Case 0:
                If Trim(txtDate) <> "" Then
                    On Error Resume Next
                    txtDate.SetFocus
                Else
                    Unload Me
                End If
            Case 1:
                If Trim(txtCustomer) <> "" Then
                    On Error Resume Next
                    txtCustomer.SetFocus
                Else
                    Unload Me
                End If
            Case 2:
                If Trim(txtInvoiceNumber) <> "" Then
                    On Error Resume Next
                    txtInvoiceNumber.SetFocus
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
    If LOGSAE = "" Then
        SAFILTER = ""
    Else
        SAFILTER = " AND USERCODE='" & LOGSAE & "'"
    End If

    SearchTab.Tab = SEARCH_TAB
    SearchTab_Click SearchTab.Tab
    ShowStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SAFILTER = ""

End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCustomer
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

Private Sub lstCustomer_DblClick()
    If lstCustomer.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_SalesOrder.SearchID (Trim(Me.lstCustomer.SelectedItem.ListSubItems(7).Text))
    Unload Me
End Sub

Private Sub lstCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstCustomer_DblClick
End Sub

Private Sub lstDate_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDate
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

Private Sub lstDate_DblClick()
    If lstDate.SelectedItem Is Nothing Then Exit Sub

    frmSMIS_Trans_SalesOrder.SearchID (Trim(Me.lstDate.SelectedItem.ListSubItems(7).Text))
    Unload Me
End Sub

Private Sub lstDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstDate_DblClick
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
    frmSMIS_Trans_SalesOrder.SearchID (Trim(Me.lstIgnitionKey.SelectedItem.ListSubItems(7).Text))
    Unload Me
End Sub

Private Sub lstIgnitionKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: lstIgnitionKey_DblClick
End Sub

Private Sub lstIgnitionKey_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtIgnitionKey.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub lstInvoiceNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstInvoiceNo
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

Private Sub lstInvoiceNo_DblClick()
    If lstInvoiceNo.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_SalesOrder.SearchID (Trim(Me.lstInvoiceNo.SelectedItem.ListSubItems(7).Text))
    Unload Me
End Sub

Private Sub lstInvoiceNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstInvoiceNo_DblClick
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
    frmSMIS_Trans_SalesOrder.SearchID (Trim(Me.lstProdNo.SelectedItem.ListSubItems(7).Text))

    Unload Me
End Sub

Private Sub lstProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: lstProdNo_DblClick
End Sub

Private Sub lstProdNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtProdNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab

    Select Case SEARCH_TAB
        Case 0
            txtDate.Enabled = True: lstDate.Enabled = True
            Me.Caption = "Search Invoice by Date"
            txtDate_Change
            On Error Resume Next
            txtDate.SetFocus
        Case 1
            txtCustomer.Enabled = True: lstCustomer.Enabled = True
            Me.Caption = "Search Invoice by  Customer"
            txtCustomer_Change
            On Error Resume Next
            txtCustomer.SetFocus
        Case 2
            txtInvoiceNumber.Enabled = True: lstInvoiceNo.Enabled = True
            Me.Caption = "Search Invoice by  VSO Number"
            txtInvoiceNumber_Change
            On Error Resume Next
            txtInvoiceNumber.SetFocus
        Case 3
            txtProdNo.Enabled = True: lstProdNo.Enabled = True
            Me.Caption = "Search Invoice by  Product Number"
            txtProdNo_Change
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            txtIgnitionKey.Enabled = True: lstIgnitionKey.Enabled = True
            Me.Caption = "Search Invoice by Conduction Sticker Number"
            txtIgnitionKey_Change
            On Error Resume Next
            txtIgnitionKey.SetFocus
    End Select
End Sub

Private Sub txtCustomer_Change()
    'SELECT VI_NO, CustName, InvoicedDate, Model, ModelDescription, ProdNo, ConductionSticker, EngineNo, FrameNo, Vino, Plate_No, IGNKEY_NO  FROM SMIS_SalesOrder
    On Error GoTo ErrorCode:
    Dim i
    If txtCustomer = "" Then
        Me.lstCustomer.Sorted = False: Me.lstCustomer.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  upper(custname), MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID,STATUS FROM SMIS_SalesOrder where  custname is Not null  " & SAFILTER & " order by CUSTNAME asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstCustomer.ListItems, rsSalesInvoice
        End If
    Else
        Me.lstCustomer.Sorted = False: Me.lstCustomer.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  upper(custname), MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID,STATUS  FROM SMIS_SalesOrder  WHERE  custname is Not null AND CUSTNAME like '" & Trim(ReplaceQuote(Me.txtCustomer)) & "%' " & SAFILTER & " order by CUSTNAME asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstCustomer.ListItems, rsSalesInvoice
        End If
    End If
    For i = 1 To lstCustomer.ListItems.Count
        If lstCustomer.ListItems(i).ListSubItems(6).Text = "C" Or lstCustomer.ListItems(i).ListSubItems(8).Text = "C" Then
            SetColorX vbRed, lstCustomer.ListItems(i)
        ElseIf lstCustomer.ListItems(i).ListSubItems(6).Text = "" Or lstCustomer.ListItems(i).ListSubItems(6).Text = "U" Then
            SetColorX vbBlue, lstCustomer.ListItems(i)
        End If
    Next



    LV_AutoSizeColumn lstCustomer

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCustomer.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then: lstCustomer.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtDate_Change()
    'SELECT VI_NO, CustName, InvoicedDate, Model, ModelDescription, ProdNo, ConductionSticker, EngineNo, FrameNo, Vino, Plate_No, IGNKEY_NO  FROM SMIS_SalesOrder
    On Error GoTo ErrorCode:

    If txtDate = "" Then
        Me.lstDate.Sorted = False: Me.lstDate.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  convert(varchar,deyt, 101), MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID, status FROM SMIS_SalesOrder WHERE DEYT IS NOT NULL " & SAFILTER & " order by deyt asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstDate.ListItems, rsSalesInvoice
        End If
    Else
        Me.lstDate.Sorted = False: Me.lstDate.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  convert(varchar,deyt, 101), MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID,status  FROM SMIS_SalesOrder  WHERE  convert(varchar,deyt, 101) like '" & Trim(ReplaceQuote(Me.txtDate)) & "%' " & SAFILTER & " order by deyt asc")

        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstDate.ListItems, rsSalesInvoice
        End If
    End If
    Dim i
    For i = 1 To lstDate.ListItems.Count
        If lstDate.ListItems(i).ListSubItems(6).Text = "C" Or lstDate.ListItems(i).ListSubItems(8).Text = "C" Then
            SetColorX vbRed, lstDate.ListItems(i)
        ElseIf lstDate.ListItems(i).ListSubItems(6).Text = "" Or lstDate.ListItems(i).ListSubItems(6).Text = "U" Then
            SetColorX vbBlue, lstDate.ListItems(i)
        End If
    Next

    LV_AutoSizeColumn lstDate
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtDate.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstDate.ListItems.Count > 0 And lstDate.Enabled = True Then: lstDate.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtIgnitionKey_Change()
    On Error GoTo ErrorCode:
    lstIgnitionKey.ListItems.Clear
    If txtIgnitionKey = "" Then
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  IGNKEY_NO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS , ID,STATUS  FROM SMIS_SalesOrder where ignkey_no is not null " & SAFILTER & " order by IGNKEY_NO asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsSalesInvoice
        End If
    Else
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  IGNKEY_NO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS , ID,STATUS FROM SMIS_SalesOrder  WHERE  ignkey_no is not null  and IGNKEY_NO like '" & Trim(ReplaceQuote(Me.txtIgnitionKey)) & "%' " & SAFILTER & " order by ProdNo asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsSalesInvoice
        End If
    End If


    Dim i
    For i = 1 To lstIgnitionKey.ListItems.Count
        If lstIgnitionKey.ListItems(i).ListSubItems(6).Text = "C" Or lstIgnitionKey.ListItems(i).ListSubItems(8).Text = "C" Then
            SetColorX vbRed, lstIgnitionKey.ListItems(i)
        ElseIf lstIgnitionKey.ListItems(i).ListSubItems(6).Text = "" Or lstIgnitionKey.ListItems(i).ListSubItems(6).Text = "U" Then
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

Private Sub txtInvoiceNumber_Change()
    On Error GoTo ErrorCode:
    If txtInvoiceNumber = "" Then
        Me.lstInvoiceNo.Sorted = False: Me.lstInvoiceNo.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  SO_NO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS, ID ,STATUS  FROM SMIS_SalesOrder WHERE SO_NO IS NOT NULL " & SAFILTER & " order by SO_NO ASC")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstInvoiceNo.ListItems, rsSalesInvoice
        End If
    Else
        Me.lstInvoiceNo.Sorted = False: Me.lstInvoiceNo.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  SO_NO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS, ID,STATUS  FROM SMIS_SalesOrder  WHERE  SO_NO like '%" & Trim(ReplaceQuote(Me.txtInvoiceNumber)) & "%' " & SAFILTER & " order by SO_NO asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstInvoiceNo.ListItems, rsSalesInvoice
        End If
    End If
    Dim i
    For i = 1 To lstInvoiceNo.ListItems.Count
        If lstInvoiceNo.ListItems(i).ListSubItems(6).Text = "C" Or lstInvoiceNo.ListItems(i).ListSubItems(8).Text = "C" Then
            SetColorX vbRed, lstInvoiceNo.ListItems(i)
        ElseIf lstInvoiceNo.ListItems(i).ListSubItems(6).Text = "" Or lstInvoiceNo.ListItems(i).ListSubItems(6).Text = "U" Then
            SetColorX vbBlue, lstInvoiceNo.ListItems(i)
        End If
    Next
    LV_AutoSizeColumn lstInvoiceNo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtInvoiceNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtInvoiceNumber.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstInvoiceNo.ListItems.Count > 0 And lstInvoiceNo.Enabled = True Then: lstInvoiceNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtProdNo_Change()
    'SELECT VI_NO, CustName, InvoicedDate, Model, ModelDescription, ProdNo, ConductionSticker, EngineNo, FrameNo, Vino, Plate_No, IGNKEY_NO  FROM SMIS_SalesOrder
    On Error GoTo ErrorCode:

    If txtProdNo = "" Then
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  PRODNO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID ,STATUS FROM SMIS_SalesOrder where PRODNO is not null " & SAFILTER & " order by deyt asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsSalesInvoice
        End If
    Else
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsSalesInvoice = New ADODB.Recordset
        Set rsSalesInvoice = gconDMIS.Execute("SELECT  PRODNO, MODELDESCRIPTION, MODEL, COLOR, IGNKEY_NO  ,SALESAE, SOSTATUS,ID ,STATUS  FROM SMIS_SalesOrder  WHERE  PRODNO is not null  and ProdNo like '" & Trim(ReplaceQuote(Me.txtProdNo)) & "%' " & SAFILTER & " order by ProdNo asc")
        If Not (rsSalesInvoice.EOF And rsSalesInvoice.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsSalesInvoice
        End If
    End If


    Dim i
    For i = 1 To lstProdNo.ListItems.Count
        If lstProdNo.ListItems(i).ListSubItems(6).Text = "C" Or lstProdNo.ListItems(i).ListSubItems(8).Text = "C" Then
            SetColorX vbRed, lstProdNo.ListItems(i)
        ElseIf lstProdNo.ListItems(i).ListSubItems(6).Text = "" Or lstProdNo.ListItems(i).ListSubItems(6).Text = "U" Then
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

