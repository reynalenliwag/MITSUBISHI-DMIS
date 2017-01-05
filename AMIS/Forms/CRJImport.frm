VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmCRJImport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Import Process"
   ClientHeight    =   7980
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14100
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CRJImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   14100
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   3300
      ScaleHeight     =   645
      ScaleWidth      =   5925
      TabIndex        =   30
      Top             =   6840
      Width           =   5925
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   5460
         MouseIcon       =   "CRJImport.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   5100
         MouseIcon       =   "CRJImport.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4740
         MouseIcon       =   "CRJImport.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4380
         MouseIcon       =   "CRJImport.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4020
         MouseIcon       =   "CRJImport.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   3660
         MouseIcon       =   "CRJImport.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   3300
         MouseIcon       =   "CRJImport.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2940
         MouseIcon       =   "CRJImport.frx":1850
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2580
         MouseIcon       =   "CRJImport.frx":1B5A
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   2220
         MouseIcon       =   "CRJImport.frx":1E64
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":216E
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1500
         MouseIcon       =   "CRJImport.frx":2478
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1140
         MouseIcon       =   "CRJImport.frx":2782
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   780
         MouseIcon       =   "CRJImport.frx":2A8C
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   420
         MouseIcon       =   "CRJImport.frx":2D96
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   60
         MouseIcon       =   "CRJImport.frx":30A0
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   5100
         MouseIcon       =   "CRJImport.frx":33AA
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4740
         MouseIcon       =   "CRJImport.frx":36B4
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4380
         MouseIcon       =   "CRJImport.frx":39BE
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4020
         MouseIcon       =   "CRJImport.frx":3CC8
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   3660
         MouseIcon       =   "CRJImport.frx":3FD2
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3300
         MouseIcon       =   "CRJImport.frx":42DC
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   2940
         MouseIcon       =   "CRJImport.frx":45E6
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2580
         MouseIcon       =   "CRJImport.frx":48F0
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2220
         MouseIcon       =   "CRJImport.frx":4BFA
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":4F04
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1500
         MouseIcon       =   "CRJImport.frx":520E
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1140
         MouseIcon       =   "CRJImport.frx":5518
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   780
         MouseIcon       =   "CRJImport.frx":5822
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   420
         MouseIcon       =   "CRJImport.frx":5B2C
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   60
         MouseIcon       =   "CRJImport.frx":5E36
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   60
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdClearJournals 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Selected Date"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   150
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowTrans 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Transactions"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MouseIcon       =   "CRJImport.frx":6140
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Process Import of Cash Receipts"
      Top             =   120
      Width           =   2010
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   25
      Top             =   6540
      Width           =   4155
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Un-Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   6210
      Value           =   -1  'True
      Width           =   4155
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   6810
      Width           =   3195
      Begin VB.CommandButton cmdShowImp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2250
         Picture         =   "CRJImport.frx":6292
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   30
         Width           =   915
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "CRJImport.frx":7314
         Left            =   900
         List            =   "CRJImport.frx":7316
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   30
         Width           =   1335
      End
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3330
      ScaleHeight     =   465
      ScaleWidth      =   4905
      TabIndex        =   11
      Top             =   7500
      Width           =   4905
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   30
         MouseIcon       =   "CRJImport.frx":7318
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   90
         Width           =   315
      End
      Begin VB.Label lblNoTransaction 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "- No Transaction"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   16
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   3750
         MouseIcon       =   "CRJImport.frx":7622
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   90
         Width           =   315
      End
      Begin VB.Label lblImported 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "- Imported"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   90
         Width           =   810
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":792C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   90
         Width           =   315
      End
      Begin VB.Label lblNotYetImported 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "- Not Yet Imported"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2250
         TabIndex        =   12
         Top             =   90
         Width           =   1410
      End
   End
   Begin VB.PictureBox picBatchImport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   4980
      ScaleHeight     =   1365
      ScaleWidth      =   4125
      TabIndex        =   3
      Top             =   2970
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdBatchImporting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "BATCH IMPORT"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   900
         Width           =   3485
      End
      Begin VB.CommandButton cmdCloseRange 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   405
         Left            =   570
         TabIndex        =   6
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   115933185
         CurrentDate     =   40603
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   405
         Left            =   2520
         TabIndex        =   7
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   115933185
         CurrentDate     =   40603
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4155
         _Version        =   655364
         _ExtentX        =   7329
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Select Date Range for the Month"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8388608
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   2190
         TabIndex        =   8
         Top             =   480
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdBatchImport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Batch"
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
      Left            =   11865
      MouseIcon       =   "CRJImport.frx":7C36
      MousePointer    =   99  'Custom
      Picture         =   "CRJImport.frx":7D88
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Process Importing of Purchases"
      Top             =   7080
      Width           =   735
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Import"
      Enabled         =   0   'False
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
      Left            =   12585
      MouseIcon       =   "CRJImport.frx":8E0A
      MousePointer    =   99  'Custom
      Picture         =   "CRJImport.frx":8F5C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   7080
      Width           =   720
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
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
      Left            =   13290
      MouseIcon       =   "CRJImport.frx":91F7
      MousePointer    =   99  'Custom
      Picture         =   "CRJImport.frx":9349
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   7080
      Width           =   720
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   4800
      TabIndex        =   28
      Top             =   6450
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   556
      Picture         =   "CRJImport.frx":96AF
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "CRJImport.frx":96CB
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
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   405
      Left            =   1920
      TabIndex        =   29
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99221505
      CurrentDate     =   40063
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4965
      Left            =   150
      TabIndex        =   62
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      Appearance      =   0
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontName =   "Segoe UI"
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4965
      Left            =   4770
      TabIndex        =   63
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      Appearance      =   0
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontName =   "Segoe UI"
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4965
      Left            =   9390
      TabIndex        =   64
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      Appearance      =   0
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontName =   "Segoe UI"
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNDEPOSITED OR'S"
      DataField       =   "&H8000000D&"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   150
      TabIndex        =   69
      Top             =   660
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEPOSITED OR'S"
      DataField       =   "&H8000000D&"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   4770
      TabIndex        =   68
      Top             =   660
      Width           =   4575
   End
   Begin VB.Label lblTransactionDate 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   67
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4830
      TabIndex        =   66
      Top             =   6180
      Width           =   4515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PMIS/CSMS/SMIS"
      DataField       =   "&H8000000D&"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   9390
      TabIndex        =   65
      Top             =   660
      Width           =   4575
   End
End
Attribute VB_Name = "frmCRJImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TransactionID                                           As String
Dim rsCSMIOS_REPOR                                          As ADODB.Recordset
Dim rsCSMIOS_SUBLET                                         As ADODB.Recordset
Dim rsCSMIOS_PMS                                            As ADODB.Recordset
Dim rsCSMIOS_LABOR                                          As ADODB.Recordset
Dim rsCSMIOS_PARTS                                          As ADODB.Recordset
Dim rsCSMIOS_MATERIALS                                      As ADODB.Recordset
Dim rsSMIS_PURCHAGREE                                       As ADODB.Recordset
Dim rsCSMIOS_ACCESSORIES                                    As New ADODB.Recordset
Dim rsCSMIOS_TINSMITH                                       As New ADODB.Recordset
''Added by norman
Dim rsCSMIOS_GJ As ADODB.Recordset
Dim rsCSMIOS_SR As ADODB.Recordset
Dim rsCSMIOS_BP As ADODB.Recordset
Dim rsCSMIOS_INSPMS As ADODB.Recordset

'Discount
Dim CSMIOS_PMS_DISCOUNT                                     As Double

'Warranty
Dim WARRANTY_CSMIOS_PARTS_COST                              As Double
Dim WARRANTY_CSMIOS_MATERIALS_COST                          As Double
Dim WARRANTY_CSMIOS_ACCESSORIES_COST                        As Double
Dim WARRANTY_JNO                                            As String
Dim WARRANTY_VOUCHERNO                                      As String
Dim WARRANTY_ItemCnt                                        As Integer
Dim WARRANTY_J_JITEMNO                                      As String
Dim WARRANTY_DIRECT_EXPENSE_LABOR_COST                      As Double
Dim WARRANTY_J_AMOUNTTOPAY                                  As Double
Dim WARRANTY_J_INVOICEAMT                                   As Double
Dim WARRANTY_J_BALANCE                                      As Double
Dim WARRANTY_J_AMOUNTPAID                                   As Double

'Accessoreis
Dim CSMIOS_ACCESSORIES_DISCOUNT                             As Double

Dim CSMIOS_REP_OR                                           As String
Dim CSMIOS_ACCT_NO                                          As String
Dim CSMIOS_PARTICIPAT                                       As String
Dim CSMIOS_PLATE_NO                                         As String
Dim CSMIOS_NIYM                                             As String
Dim CSMIOS_TERM                                             As String
Dim CSMIOS_DTE_REL                                          As String
Dim CSMIOS_INVOICE                                          As String
Dim CSMIOS_VAT_EXEMPT                                       As Boolean
Dim CSMIOS_RO_AMOUNT                                        As Double

Dim CSMIOS_LABOR                                            As Double
Dim CSMIOS_PARTS                                            As Double
Dim CSMIOS_MATERIALS                                        As Double
Dim CSMIOS_ACCESSORIES                                      As Double

Dim CSMIOS_LABOR_COST                                       As Double
Dim CSMIOS_PARTS_COST                                       As Double
Dim CSMIOS_MATERIALS_COST                                   As Double
Dim CSMIOS_ACCESSORIES_COST                                 As Double

Dim CSMIOS_PMS_COST                                         As Double

Dim CSMIOS_TINSPAINT                                        As Double
Dim CSMIOS_SUBLET                                           As Double
Dim CSMIOS_AIRCON                                           As Double

Dim CSMIOS_TINSPAINT_DISCOUNT                               As Double
Dim CSMIOS_SUBLET_DISCOUNT                                  As Double

Dim CSMIOS_LABOR_DISCOUNT                                   As Double
Dim CSMIOS_PARTS_DISCOUNT                                   As Double
Dim CSMIOS_MATERIALS_DISCOUNT                               As Double

Dim WARRANTY_DIRECT_EXPENSE_LABOR                           As Double
Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS                      As Double
Dim WARRANTY_DIRECT_EXPENSE_GOL                             As Double

Dim COMPANY_DIRECT_EXPENSE_LABOR                            As Double
Dim COMPANY_DIRECT_EXPENSE_SPAREPARTS                       As Double
Dim COMPANY_DIRECT_EXPENSE_GOL                              As Double

Dim SALES_DIRECT_EXPENSE_LABOR                              As Double
Dim SALES_DIRECT_EXPENSE_SPAREPARTS                         As Double
Dim SALES_DIRECT_EXPENSE_GOL                                As Double

Dim INSURANCE_DIRECT_EXPENSE_LABOR                          As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS                     As Double
Dim INSURANCE_DIRECT_EXPENSE_GOL                            As Double
Dim INSURANCE_DIRECT_EXPENSE_ACCESSORIES                    As Double
Dim CSMIOS_PMS                                              As Double
''Added by norman
Dim INSURANCE_DIRECT_EXPENSE_SR                             As Double
Dim INSURANCE_DIRECT_EXPENSE_LABOR_BP                       As Double
Dim INSURANCE_DIRECT_EXPENSE_LABOR_PMS                      As Double
''--------------------------

Dim TOTAL_INSURANCE_AMOUNT                                  As Double
Dim TOTAL_DISCOUNT_AMOUNT                                   As Double
Dim CSMIOS_SUBLET_COST                                      As Double
Dim ALL_DEBIT, ALL_CREDIT                                   As Double
Dim CSMIOS_TINSPAINT_COST                                   As Double

'Internal
Dim INTERNAL_LABOR_AMT                                      As Double
Dim INTERNAL_PARTS_AMT                                      As Double
Dim INTERNAL_MATERIALS_AMT                                  As Double
Dim INTERNAL_LABOR_COST                                     As Double
Dim INTERNAL_PARTS_COST                                     As Double
Dim INTERNAL_MATERIALS_COST                                 As Double

Dim J_ACCT_CODE                                             As String
Dim J_ACCT_NAME                                             As String
Dim J_GROSS                                                 As Double
Dim J_TAX                                                   As Double
Dim J_NET                                                   As Double

Dim J_DEBIT                                                 As Double
Dim J_CREDIT                                                As Double

Dim TOTAL_DEBIT                                             As Double
Dim TOTAL_CREDIT                                            As Double
'Dim CSMIOS_VAT_EXEMPT                              As Boolean

Dim ItemCnt                                                 As Integer
Dim CSMS_ACCCOST                                            As Double
Dim rsINTERNAL_RO_DET                                       As ADODB.Recordset

Dim J_JDATE                                                 As String
Dim J_VOUCHERNO                                             As String
Dim J_JTYPE                                                 As String
Dim J_JNO                                                   As String
Dim J_REMARKS                                               As String
Dim J_REMARKS2                                              As String
Dim J_VENDORCODE                                            As String
Dim J_CUSTOMERCODE                                          As String
Dim J_OUTBALANCE                                            As Double
Dim J_AMOUNTTOPAY                                           As Double
Dim J_INVOICEAMT                                            As Double
Dim J_BALANCE                                               As Double
Dim J_AMOUNTPAID                                            As Double
Dim J_CHECKNO                                               As String
Dim J_INVOICEDATE                                           As String
Dim J_DUEDATE                                               As String
Dim J_PAYTYPE                                               As String
Dim J_INVOICETYPE                                           As String
Dim J_INVOICENO                                             As String
Dim J_CHECKDATE                                             As String
Dim J_BANKCODE                                              As String
Dim J_REFNO                                                 As String
Dim J_REFDATE                                               As String
Dim J_TERMS                                                 As String
Dim J_DEALER                                                As String
Dim J_PAIDSTATUS                                            As String
Dim J_RECEIVESTATUS                                         As String

'Dim J_ACCT_CODE, J_ACCT_NAME                       As String
'Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
Dim J_STATUS                                                As String
Dim J_JITEMNO                                               As String
Dim rsJournal_HDDup                                         As ADODB.Recordset
Dim LIM                                                     As Integer
Dim TRANSACTIONDATE                                         As String
Dim Indx                                                    As Integer
Dim xTranDate                                               As String
Dim DEF_INVOICETYPE                                         As String
Dim DEF_INVOICENO                                           As String
Dim xCUT_OFF_DATE                                           As String
Dim rsCSMIOS_FINANCE                                        As ADODB.Recordset
Dim FINANCE                                                 As String
Dim rsCSMIOS_CHKINS                                         As ADODB.Recordset
Dim INSURANCE                                               As String
Dim xCDEPOSITCHECK                                          As String
Dim rsUEA                                                   As New ADODB.Recordset
Dim rsDEFERRED                                              As ADODB.Recordset
Dim rsCHKSCHEDULED As New Recordset

''SALES IMPORT VARIABLE
Dim rsPMIOS_ORD_HD                                          As ADODB.Recordset
Dim rsTdayTran                                              As ADODB.Recordset
Dim rsCSMIOS_PARTS_NON                                      As ADODB.Recordset
Dim rsCSMIOS_PARTS_BP                                       As ADODB.Recordset
Dim rsCSMIOS_MATERIALS_BP                                   As ADODB.Recordset
Dim rsCSMIOS_INSPARTS_BP                                    As ADODB.Recordset
Dim rsCSMIOS_INSMATERIALS_BP                                As ADODB.Recordset
Dim rsCSMIOS_WARRANTYGJ                                     As ADODB.Recordset
Dim rsCSMIOS_WARRANTYSR                                     As ADODB.Recordset
Dim rsCSMIOS_WARRANTYBP                                     As ADODB.Recordset
Dim rsCSMIOS_WARRANTYPARTSBP                                As ADODB.Recordset
Dim rsCSMIOS_WARRANTYMATERIALSBP                            As ADODB.Recordset
Dim rsTransferOut                                           As ADODB.Recordset
Dim rsVENDOR                                                As New ADODB.Recordset
Dim DOTAX_SERVICE1                                          As ADODB.Recordset
'DISCOUNT
'WARRANTY
Dim WARRANTY_CSMIOS_PARTS_COST_BP                           As Double
Dim WARRANTY_CSMIOS_MATERIALS_COST_BP                       As Double
Dim WARRANTY_DIRECT_EXPENSE_LABOR_COST_BP                   As Double
Dim WARRANTY_DIRECT_EXPENSE_LABOR_COST_SR                   As Double
'ACCESSOREIS
Dim CSMIOS_VAT_EXEMPT1                                       As Boolean
Dim CSMIOS_PARTS_BP                                         As Double
Dim CSMIOS_PARTS_NON                                        As Double
Dim CSMIOS_MATERIALS_BP                                     As Double
Dim CSMIOS_PARTS_COST_BP                                    As Double
Dim CSMIOS_MATERIALS_COST_BP                                As Double
'DIM CSMIOS_RO_AMOUNT                                  AS DOUBLE
Dim CSMIOS_PARTS_NON_DISCOUNT                               As Double
Dim WARRANTY_DIRECT_EXPENSE_LABOR_BP                        As Double
Dim WARRANTY_DIRECT_EXPENSE_LABOR_SR                        As Double
Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS_BP                   As Double
Dim WARRANTY_DIRECT_EXPENSE_GOL_BP                          As Double
Dim COMPANY_DIRECT_EXPENSE_ACCESSORIES                      As Double
Dim SALES_DIRECT_EXPENSE_ACCESSORIES                        As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS_BP                  As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS_SR                  As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS_BP_SR               As Double
Dim INSURANCE_DIRECT_EXPENSE_GOL_BP                         As Double
'INTERNAL
Dim INTERNAL_BP_LABOR_RS                                    As New ADODB.Recordset
Dim INTERNAL_BP_LABOR                                       As Double
Dim INTERNAL_ACCESSORIES_AMT                                As Double
Dim INTERNAL_PARTS_INV                                      As Double
Dim INTERNAL_ACCESSORIES_COST                               As Double
Dim J_ITEMCNT                                               As Integer
Dim SubletType                                              As String
Dim DetCodeLen                                              As Integer
'DISCOUNT INSURANCE :NRE
Dim CHECK_INSURANCE_DIS                                     As New ADODB.Recordset
Dim CHECK_INSURANCE                                         As Double
Dim DISCOUNT_INS_LABOR_GJ                                   As Double
Dim DISCOUNT_INS_SUBLET                                     As Double
Dim DISCOUNT_INS_LABOR_BP                                   As Double
Dim DISCOUNT_INS_PARTS_GJ                                   As Double
Dim DISCOUNT_INS_PARTS_BP                                   As Double
Dim DISCOUNT_INS_MATERIALS                                  As Double
Dim DISCOUNT_INS_ACCESSORIES                                As Double
Dim TranType                                                As String
Dim rsDOTCIS                                                As New ADODB.Recordset
Dim xIDDOTAX                                                As String
Dim xCNAME                                                  As String
Dim INTERNAL_LABOR_AMT_SR                                   As Double
Dim INTERNAL_LABOR_COST_SR                                  As Double
Dim rsPARTSHARI                                             As New ADODB.Recordset
Dim rsPARTSNHARI                                            As New ADODB.Recordset
Dim INSURANCE_PARTS_HARI                                    As Double
Dim INSURANCE_PARTS_NONHARI                                 As Double

''UNDEPOSITED VARIABLES
    Dim J_CUSTOMERCODE2                                     As String
    Dim J_REFERENCENO                                       As String
    Dim J_ENTITY                                            As String
    Dim J_BANKCHARGES                                       As Double
    'DETAIL
    Dim J_JITEMNO_2                                         As String
    Dim J_ALLENTITY                                         As String
    Dim J_INVOICENUM                                        As String

    Dim CMIS_OR_NUM                                         As String
    Dim CMIS_OR_DATE                                        As String
    Dim CMIS_OR_AMT                                         As String
    Dim CMIS_DISCOUNT                                       As String
    Dim CMIS_TAX                                            As String
    Dim CMIS_CUSCDE                                         As String
    Dim CMIS_CUSNAME                                        As String
    Dim CMIS_DEPOSIT                                        As String
    Dim CMIS_BANKCODE                                       As String
    Dim CMIS_BANK                                           As String
    Dim CMIS_CARDBANK                                       As String
    Dim CMIS_TSEKE                                          As String
    Dim CMIS_CHECKDATE                                      As String
    Dim CMIS_STATUS                                         As String
    Dim CMIS_TYPE_PAYMENT                                   As String
    Dim CMIS_DT_TRANTYPE                                    As String
    Dim CMIS_DT_REFERENCE                                   As String
    Dim CMIS_DT_CUSCDE                                      As String
    Dim CMIS_DT_DESCRIPT                                    As String
    Dim CMIS_DT_REFERENCENO                                 As String
    Dim CMIS_DT_DOCDTE                                      As String
    Dim CMIS_DT_PAIDFOR                                     As String
    Dim CMIS_ENTITYCODE                                     As String
    Dim CMIS_ALLENTITY                                      As String
    Dim CMIS_CASHAMOUNT                                     As Double
    Dim CMIS_CHKAMOUNT                                      As Double
    Dim CMIS_CARDAMOUNT                                     As Double
    Dim CMIS_DT_AMOUNT                                      As Double
    Dim CMIS_DT_PAYMENT                                     As Double
    Dim CMIS_DT_DISCOUNT                                    As Double
    Dim CMIS_DT_TAX                                         As Double
    Dim CMIS_IS_VAT                                         As Boolean
    Dim i                                                   As Long

    Dim rsOFF_HD                                            As ADODB.Recordset
    Dim rsOFF_DT                                            As ADODB.Recordset
    Dim rsSJ_DATA                                           As ADODB.Recordset
    Dim rsCheckJournal_HD                                   As ADODB.Recordset

    Dim PV_MRRNO                                            As String
    Dim PV_INVNO                                            As String
    Dim PV_PRODNO                                           As String
    Dim J_JVOUCHERNO                                        As String
    Dim PV_STATUS, PV_ITEMNO                                As String
    Dim PV_AMOUNT                                           As Double
    Dim SJ_PV_ITEMNO                                        As Integer
    Dim GridImport                                          As Integer
    Dim J_INVOICETYPE2                                      As String
    Dim DEFERRED_OUTPUT                                     As Double
    Dim CMIS_DT_REFCODE                                     As String
    Dim CUSDEPOT_REFERENCE                                  As String
    Dim xCHKIFNONVAT                                        As Boolean
    
    Dim rsCreditCardCost                                    As ADODB.Recordset
    Dim rsCreditCardCostDistinc                             As ADODB.Recordset
    
    
    Dim xInvCMC As String
    Dim xInvDate As String
    Dim xInvTerm As String
    Dim rsGetDetailsFromSMIS As ADODB.Recordset

Function SetOTHChartCodes(XXX As String) As String
    Dim rsSBOOK_CHARTCODES                                  As ADODB.Recordset
    Set rsSBOOK_CHARTCODES = New ADODB.Recordset
    Set rsSBOOK_CHARTCODES = gconDMIS.Execute("Select * from CMIS_SBOOK where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOK_CHARTCODES.EOF And Not rsSBOOK_CHARTCODES.BOF Then
        SetOTHChartCodes = Null2String(rsSBOOK_CHARTCODES!CHARTCODES)
    End If
    Set rsSBOOK_CHARTCODES = Nothing
End Function

Function ReturnSITerm(XXX As String) As String
    Dim rsREPOR_INVOICE                                     As ADODB.Recordset
    Set rsREPOR_INVOICE = New ADODB.Recordset
    Set rsREPOR_INVOICE = gconDMIS.Execute("Select TERM from CSMS_Repor Where INVOICE = '" & XXX & "'")
    If Not rsREPOR_INVOICE.EOF And Not rsREPOR_INVOICE.BOF Then
        ReturnSITerm = Null2String(rsREPOR_INVOICE!TERM)
    End If
    Set rsREPOR_INVOICE = Nothing
End Function

Function ReturnTranType(XXX As String, YYY As String) As String
    Dim rsPMIS_Invoice                                      As ADODB.Recordset
    Set rsPMIS_Invoice = New ADODB.Recordset
    rsPMIS_Invoice.Open "SELECT TRANTYPE FROM PMIS_VW_ISS_HISTORY WHERE TRANNO = '" & XXX & "' AND TYPE='" & YYY & "' AND TRANTYPE<>'RIV'", gconDMIS, adOpenKeyset
    If Not rsPMIS_Invoice.EOF And Not rsPMIS_Invoice.BOF Then
        ReturnTranType = Null2String(rsPMIS_Invoice!TranType)
    End If
    Set rsPMIS_Invoice = Nothing
End Function

Function SetTransaction(XXX As Variant) As String
    Dim rsSBOOKTransaction                                  As ADODB.Recordset
    Set rsSBOOKTransaction = New ADODB.Recordset
    Set rsSBOOKTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'A' and CODE = '" & XXX & "'")
    If Not rsSBOOKTransaction.EOF And Not rsSBOOKTransaction.BOF Then
        SetTransaction = Null2String(rsSBOOKTransaction!DESCNAME)
    End If
    Set rsSBOOKTransaction = Nothing
End Function

Function SetOtherTransaction(XXX As Variant) As String
    Dim rsSBOOKOtherTransaction                             As ADODB.Recordset
    Set rsSBOOKOtherTransaction = New ADODB.Recordset
    Set rsSBOOKOtherTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOKOtherTransaction.EOF And Not rsSBOOKOtherTransaction.BOF Then
        SetOtherTransaction = Null2String(rsSBOOKOtherTransaction!DESCNAME)
    End If
    Set rsSBOOKOtherTransaction = Nothing
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                     As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
End Function

Function GetCRJVoucherNo() As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'CRJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetCRJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetCRJVoucherNo = "000001"
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    
    If COMPANY_CODE = "DJM" Or COMPANY_CODE = "MGS" Then
        Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'GJ' Order by MAX_VOUCHERNO desc")
    Else
        Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'DRJ' Order by MAX_VOUCHERNO desc")
    End If
    
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function GetSJVoucherNo() As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetSJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetSJVoucherNo = "000001"
    End If
End Function

'Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
'    Dim rsCheckSJ_Journal_HD                      As ADODB.Recordset
'    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
'    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
'    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
'        CheckSJExisting = True
'    Else
'        CheckSJExisting = False
'    End If
'    Set rsCheckSJ_Journal_HD = Nothing
'End Function

Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String, Optional VarTranType As String) As Boolean
    Dim rsCheckSJ_Journal_HD                                As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    If VarTranType = "" Then
        Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND PayType = " & N2Str2Null(VarTranType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckSJExisting = True
    Else
        CheckSJExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function

Function CheckRefNoExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
    Dim rsCheckSJ_Journal_HD                                As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND RefNo = " & N2Str2Null(VarInvoiceNo))
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckRefNoExisting = True
    Else
        CheckRefNoExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function

Function CheckCRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckCRJ_Journal_HD                               As ADODB.Recordset
    Set rsCheckCRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        If COMPANY_CODE = "HSM" Then
            Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND LEFT(InvoiceNo,2) = 'NV' AND (CASE WHEN LEN(INVOICENO)= 10 THEN RIGHT(InvoiceNo,8) WHEN LEN(INVOICENO)= 9 THEN RIGHT(InvoiceNo,7) WHEN LEN(INVOICENO)= 8 THEN RIGHT(InvoiceNo,6) WHEN LEN(INVOICENO)= 7 THEN RIGHT(InvoiceNo,5) WHEN LEN(INVOICENO)= 5 THEN RIGHT(InvoiceNo,3) ELSE RIGHT(InvoiceNo,5) END) = " & N2Str2Null(VarInvoiceNo))
        Else
            Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,len(InvoiceNo) - 2) = " & N2Str2Null(VarInvoiceNo))
        End If
    Else
        Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckCRJ_Journal_HD.EOF And Not rsCheckCRJ_Journal_HD.BOF Then
        CheckCRJExisting = True
    Else
        CheckCRJExisting = False
    End If
    Set rsCheckCRJ_Journal_HD = Nothing
End Function

Function CheckDRJExisting3(REFNO As String, DepositDate As Date) As Boolean
    Dim rsCheckDRJ_Journal_HD                               As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    
    Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype,INVOICENO from AMIS_Journal_HD where JTYPE = 'GJ' AND INVOICENO =  '" & REFNO & "' AND JDATE = '" & DepositDate & "'")
    
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExisting3 = True
    Else
        CheckDRJExisting3 = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function CheckDRJExisting2(VarInvoiceNo As String, BANKCODE As String) As Boolean
    Dim rsCheckDRJ_Journal_HD                               As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    
    If COMPANY_CODE = "HCA" Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype,refno from AMIS_Journal_HD where JTYPE = 'DRJ' AND BANKCODE =  '" & BANKCODE & "' AND   RefNo= " & N2Str2Null(VarInvoiceNo))
    ElseIf COMPANY_CODE = "DJM" Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype,refno from AMIS_Journal_HD where JTYPE = 'GJ' AND BANKCODE =  '" & BANKCODE & "'")
    End If
    
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExisting2 = True
    Else
        CheckDRJExisting2 = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function CheckDRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckDRJ_Journal_HD                          As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        If COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
            Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype,refno from AMIS_Journal_HD where JTYPE = 'DRJ' AND " & _
                                                    "INVOICENO=(CASE WHEN LEFT(InvoiceNo,2) = 'NV' THEN 'NV' + " & N2Str2Null(VarInvoiceNo) & " Else " & N2Str2Null(VarInvoiceNo) & " END) ")
        Else
            If COMPANY_CODE = "DGI" Then
                Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("SELECT VOUCHERNO,JTYPE,INVOICENO FROM AMIS_JOURNAL_HD WHERE JTYPE = 'DRJ' AND LEFT(INVOICENO,2) = 'NV' AND (CASE WHEN LEN(INVOICENO) = 9 THEN RIGHT(INVOICENO,7) ELSE RIGHT(INVOICENO,6) END)  = " & N2Str2Null(VarInvoiceNo))
            Else
                Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo))
            End If
        End If
    Else
        If COMPANY_CODE = "DJM" Then
            Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'GJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
        Else
            Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
        End If
    End If
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExisting = True
    Else
        CheckDRJExisting = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function CheckDRJExistingM(VarInvoiceNo As String, VarVAT As Variant, VarAmount As Double) As Boolean
    Dim rsCheckDRJ_Journal_HD                               As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo) & " AND INVOICEAMT =" & NumericVal(VarAmount) & "")
    Else
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo) & " And INVOICEAMT = " & NumericVal(VarAmount) & "")
    End If
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExistingM = True
    Else
        CheckDRJExistingM = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function ReturnAR_AccountCode(XXX As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AR' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAR_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnClearing_AccountCode(XXX As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'CLEARING' AND TRANTYPE1 = '" & Trim(XXX) & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnClearing_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAccountCode(XXX As String, Optional YYY As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
     If YYY = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "' AND TRANTYPE3 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnDeposit_AccountCode(XXX As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'DEPOSIT' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDeposit_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetCustomerName(VVV As Variant) As String
    Dim rsCustomer2                                         As ADODB.Recordset
    Set rsCustomer2 = New ADODB.Recordset
    rsCustomer2.Open "Select CustCode,acctname from ALL_CUSTMASTER_AMIS where CustCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer2.EOF And Not rsCustomer2.BOF Then
        SetCustomerName = UCase(Null2String(rsCustomer2!AcctName))
    Else
        SetCustomerName = ""
    End If
End Function

Function ReturnSales_AccountCode(InvType As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & InvType & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            If InvType = "SALES" Then
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & Right(OTHERTYPE, 5) & "%'")
            Else
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%'")
            End If
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnSales_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnCostofSales(InvType As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & InvType & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            If InvType = "SALES" Then
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & Right(OTHERTYPE, 5) & "%'")
            Else
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%'")
            End If
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnCostofSales = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInventory(InvType As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & InvType & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 = '" & OTHERTYPE & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInventory = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnOutputTax()
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnOutputTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnDiscount_AccountCode(InvType As String, Optional OTHERTYPE As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & InvType & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & InvType & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDiscount_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function PosibleDoubleInternal(XXX As String) As Boolean
    Dim RS                                                  As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT * FROM AMIS_journal_hd where refno='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then
        PosibleDoubleInternal = True
    Else
        PosibleDoubleInternal = False
    End If
    Set RS = Nothing
End Function

Function ReturnInternalAccountCode(XXX As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select ChartCodes from CMIS_SBOOK where BOOK = 'S' and CODE = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInternalAccountCode = Null2String(rsChartAccount!CHARTCODES)
    End If
    Set rsChartAccount = Nothing
End Function

Function GetJNo() As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(JNo AS int) AS MAX_JNO from AMIS_Journal_HD Order by MAX_JNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetJNo = Format(NumericVal(rsJournal_HD!MAX_JNO) + 1, "000000")
    Else
        GetJNo = "000001"
    End If
End Function

Function ReturnCode(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset
    Dim MARK                                                As String
    'This will return the code of the Selling Dealer

    MARK = (Replace(XXX, " ", ""))

    SQL = "SELECT Custcode, replace(custname,' ','') from ALL_Custmaster_AMIS where REPLACE(custname,' ','') like '%" & MARK & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        ReturnCode = Null2String(RS!CUSTCODE)
    Else
        ReturnCode = ""
    End If
    Set RS = Nothing
End Function

Function CheckIfPMS_Ik_to_5k(XXX As String) As Boolean
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    SQL = "select Status1 from CSMS_ro_det where jobtype='PMS' and livil='1' and wcode='W' and Rep_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If Null2String(RS!status1) = "Y" Then
            CheckIfPMS_Ik_to_5k = True
        Else
            CheckIfPMS_Ik_to_5k = False
        End If
    End If
    Set RS = Nothing
End Function

Function ReturnDeferredOutPutTax()
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'DEFERRED OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDeferredOutPutTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetRONiym(XXX As String)
    Dim rsRO_Niym                                           As ADODB.Recordset
    Set rsRO_Niym = New ADODB.Recordset
    Set rsRO_Niym = gconDMIS.Execute("Select NIYM from CSMS_REPOR WHERE INVOICE = '" & XXX & "'")
    If Not rsRO_Niym.EOF And Not rsRO_Niym.BOF Then
        SetRONiym = Null2String(rsRO_Niym!Niym)
    End If
End Function

'FUNCTION TO RETURN THE INFO OF THE PROCEEDS PAYMENT
Function ProceedsInfo(xRefInvNo As String)
    Dim xSalesCusCode As String
    Dim xCashierCuscode As String
    Dim rsProceeds As ADODB.Recordset
    Set rsProceeds = New ADODB.Recordset
    Set rsProceeds = gconDMIS.Execute("SELECT SM.CUSTOMERCODE AS SCUSCODE,HD.CUSCDE AS CCUSCODE,DT.OR_NUM,DT.REFERENCE FROM CMIS_OFF_HD HD INNER JOIN CMIS_OFF_DT DT ON HD.OR_NUM=DT.OR_NUM INNER JOIN SMIS_MRRINV_TABLE SM ON DT.REFERENCE=SM.VI_NO  WHERE SM.VI_NO = '" & xRefInvNo & "' AND DT.TRANTYPE = 'VI'")
    
    If Not rsProceeds.EOF And Not rsProceeds.BOF Then
        xSalesCusCode = Null2String(rsProceeds!SCUSCODE)
        xCashierCuscode = Null2String(rsProceeds!CCUSCODE)
        If xSalesCusCode = xCashierCuscode Then
            ProceedInfo = xCashierCuscode
        Else
            ProceedsInfo = xSalesCusCode
        End If
    End If
End Function

Function GetFao(xORNUM As String)
    Dim rsGetFao As ADODB.Recordset
    Set rsGetFao = New ADODB.Recordset
    Set rsGetFao = gconDMIS.Execute("SELECT FAO FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "'")
    
    If Not rsGetFao.EOF And Not rsGetFao.BOF Then
        GetFao = Null2String(rsGetFao!FAO)
    End If
End Function

'FUNCTION TO TRIM A CERTAIN CHARACTER FROM A STRING
Function TrimChar(ByVal Text As String, ByVal Characters As String) As String
  
  'Trim the right
  Do While Right(Text, 1) Like "[" & Characters & "]"
    Text = Left(Text, Len(Text) - 1)
  Loop

  'Trim the left
  'Do While Left(Text, 1) Like "[" & Characters & "]"
  '  Text = Mid(Text, 2)
  'Loop
  
  TrimChar = Text
End Function

Public Function GetInvoiceDetails(ByVal INV As String)
    Set rsGetDetailsFromSMIS = New ADODB.Recordset
    Set rsGetDetailsFromSMIS = gconDMIS.Execute("SELECT 'VI' + SI.VI_NO AS INVNO, SI.InvoicedDate, SO.Term FROM SMIS_MrrInv_Table SI INNER JOIN SMIS_SalesOrder SO ON SI.ignkey = SO.IGNKEY_NO  WHERE SI.STATUS = 'P' AND SO.STATUS = 'P' AND SI.VI_NO = '" & Null2String(rsOFF_DT!INVOICENO) & "'")
    xInvCMC = Null2String(rsGetDetailsFromSMIS!INVNO)
    xInvDate = Null2String(rsGetDetailsFromSMIS!InvoicedDate)
    xInvTerm = Null2String(rsGetDetailsFromSMIS!TERM)
End Function

Function ImportUnDeposit() As Boolean
    On Error GoTo ErrorCode
    Dim rsCreditCard                                        As ADODB.Recordset
    Dim BANKCHARGES                                         As Double
    Dim EWT                                                 As Double
    Dim TOTALCHARGES                                        As Double
    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 1 AND OR_DATE > '" & GetCUTOFF_DATE & "' AND  OR_DATE <= '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 0 AND OR_DATE > '" & GetCUTOFF_DATE & "' AND  OR_DATE <= '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            End If
            
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                
                'CHECK IF WITH SJ
                Dim SJINVOICETYPE As String
                Dim SJINVOICENO As String
                Dim CHKCUTOFFDATE As New ADODB.Recordset
                Dim rsCHECKOR As ADODB.Recordset
                Dim rsCheckSJ As ADODB.Recordset
                Set rsCHECKOR = New ADODB.Recordset
                Set rsCHECKOR = gconDMIS.Execute("Select * from CMIS_OFF_dt Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND TRANTYPE IN ('SI','PI','MI','AI','VI')")
                If Not rsCHECKOR.EOF And Not rsCHECKOR.BOF Then
                    Do While Not rsCHECKOR.EOF
                        Set CHKCUTOFFDATE = New ADODB.Recordset
                        If Null2String(rsCHECKOR!TranType) = "SI" Then
                            Set CHKCUTOFFDATE = gconDMIS.Execute("SELECT * FROM CSMS_Repor WHERE INVOICE = '" & Null2String(rsCHECKOR!INVOICENO) & "'  and DTE_COMP > (SELECT CUT_OFF_DATE from ALL_Profile where ModuleName = 'AMIS')")
                        ElseIf Null2String(rsCHECKOR!TranType) = "VI" Then
                            Set CHKCUTOFFDATE = gconDMIS.Execute("SELECT * FROM SMIS_SalesOrder WHERE VI_NO = '" & Null2String(rsCHECKOR!INVOICENO) & "'  and INVOICEDDATE > (SELECT CUT_OFF_DATE from ALL_Profile where ModuleName = 'AMIS')")
                        Else
                            Set CHKCUTOFFDATE = gconDMIS.Execute("SELECT * FROM PMIS_vw_ISS_HISTORY WHERE TRANNO = '" & Null2String(rsCHECKOR!INVOICENO) & "' AND TYPE ='" & Left(Null2String(rsCHECKOR!TranType), 1) & "' and TRANDATE  > (SELECT CUT_OFF_DATE from ALL_Profile where ModuleName = 'AMIS')")
                        End If
                        
                        If Not CHKCUTOFFDATE.EOF And Not CHKCUTOFFDATE.BOF Then
                            Set rsCheckSJ = New ADODB.Recordset
                            Set rsCheckSJ = gconDMIS.Execute("Select * from amis_journal_hd Where invoicetype = '" & Null2String(rsCHECKOR!TranType) & "' and invoiceno = '" & Null2String(rsCHECKOR!INVOICENO) & "'   and jtype IN ('SJ','COB')")
                            If Not rsCheckSJ.EOF And Not rsCheckSJ.BOF Then
                            Else
                                GoTo SKIP_OR 'WHEN NO SJ
                            End If
                        Else
                            GoTo IMPORT_OR
                        End If
                        rsCHECKOR.MoveNext
                    Loop
                End If
                
IMPORT_OR:
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                
'                If CMIS_OR_NUM = "OR003926" Then Stop
'                If CMIS_OR_NUM = "CR010469" Then Stop
                
                CMIS_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BANKCODE)
                CMIS_BANK = Null2String(rsOFF_HD!Bank)
                CMIS_CARDBANK = Null2String(rsOFF_HD!cardbnkcde)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BANKCODE)
                CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                CMIS_ALLENTITY = CMIS_ENTITYCODE + CMIS_CUSCDE
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                
                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If

                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
                
                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetCRJVoucherNo())
                J_JTYPE = "'CRJ'"
                J_ITEMCOUNT = 0
                J_REMARKS = ""
                
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = '" & CMIS_OR_NUM & "' ORDER BY INVOICETYPE ASC")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        J_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_INVOICENUM = Null2String(rsOFF_DT!INVOICENO)
                        J_INVOICETYPE2 = Null2String(rsOFF_DT!TranType)
                        
                        If COMPANY_CODE = "DJM" Then
                            If Null2String(rsOFF_DT!TranType) = "OTH" Then
                                J_REMARKS2 = Null2String(rsOFF_DT!DESCRIPT) & " -" & GetFao(CMIS_OR_NUM)
                            ElseIf Null2String(rsOFF_DT!TranType) = "VI" Then
                                J_REMARKS2 = Null2String(rsOFF_DT!TranType) & Null2String(rsOFF_DT!INVOICENO) & " - " & SetCustomerName(ProceedsInfo(Null2String(rsOFF_DT!INVOICENO)))
                            ElseIf Null2String(rsOFF_DT!TranType) = "SI" Then
                                J_REMARKS2 = "SB" & Null2String(rsOFF_DT!INVOICENO)
                            Else
                                J_REMARKS2 = Null2String(rsOFF_DT!TranType) & Null2String(rsOFF_DT!INVOICENO)
                            End If
                            
                            CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)
                        ElseIf COMPANY_CODE = "CMC" Then
                            If Null2String(rsOFF_DT!TranType) = "OTH" Then
                                J_REMARKS = Null2String(rsOFF_DT!DESCRIPT) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                            ElseIf Null2String(rsOFF_DT!TranType) = "VI" Then
                                J_REMARKS = ""
                                GetInvoiceDetails (Null2String(rsOFF_DT!INVOICENO))
                            Else
                                J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                            End If
                        Else
                            If Null2String(rsOFF_DT!TranType) = "OTH" Then
                                If COMPANY_CODE = "DGI" Then
                                    J_REMARKS = "OTHER TRANSACTION : " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment)) & " [" & Null2String(rsOFF_DT!DESCRIPT) & "] "
                                ElseIf COMPANY_CODE = "DJM" Then
                                    J_REMARKS = Null2String(rsOFF_DT!DESCRIPT)
                                Else
                                    J_REMARKS = Null2String(rsOFF_DT!DESCRIPT) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                                End If
                            Else
                                J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                            End If
                        End If
                        
                        If Not rsOFF_DT.EOF Then
                            If COMPANY_CODE = "DJM" Then
                                If Null2String(rsOFF_DT!TranType) = "SI" Then
                                    J_REMARKS = J_REMARKS + J_REMARKS2 & ","
                                Else
                                    J_REMARKS = J_REMARKS + J_REMARKS2 & ","
                                End If
                            Else
                                J_REMARKS = J_REMARKS & Chr(9)
                            End If
                        End If
                        
                        rsOFF_DT.MoveNext
                    Loop
                    
                    
                    If COMPANY_CODE = "DJM" Then
                        J_REMARKS = TrimChar(J_REMARKS, ",")
                        J_REMARKS = TrimChar(J_REMARKS, "-")
                        
                        If J_INVOICETYPE2 = "OTH" Then
                        ElseIf J_INVOICETYPE2 = "VI" Then
                        Else
                            J_REMARKS = J_REMARKS & " - " & CMIS_CUSNAME
                        End If
                    Else
                        J_REMARKS = N2Str2Null(J_REMARKS)
                    End If
                Else
                    J_REMARKS = "NULL"
                End If
                
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)
                J_DEPOSIT = 0
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
                
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CMIS_OR_AMT), 2)
                J_BALANCE = 0
                J_AMOUNTPAID = 0
                
                J_STATUS = "'N'"
                
                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 10))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 10))
                End If
                
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_ENTITY = N2Str2Null(CMIS_ENTITYCODE)
                J_ALLENTITY = N2Str2Null(CMIS_ALLENTITY)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
                
                'PAYMENT TYPE 1 = CASH
                'PAYMENT TYPE 2 = CHECK
                'PAYMENT TYPE 3 = CARD
                
                'CASH ON HAND | CHECK
                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    
                    If COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CIB", "BDO"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CIB", "BDO")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                    End If
                    
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                    Else
                        J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                    End If
                    
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    
                    If J_DEBIT > 0 Then
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                End If
                
                'CARD
                If CMIS_TYPE_PAYMENT = "3" Then
                    J_ITEMCOUNT = 0
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HCR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                    ElseIf COMPANY_CODE = "HNE" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDIT CARDS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDIT CARDS")))
                    ElseIf COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HMR" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HGS" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                    ElseIf COMPANY_CODE = "DJM" Then
                        If CMIS_DT_PAIDFOR = "482" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                        End If
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND", CMIS_CARDBANK))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND", CMIS_CARDBANK)))
                    End If
                    
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    
                    If CMIS_CARDAMOUNT > 0 Then
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                End If
                
                'CUSTOMER'S DEPOSIT WITH CARD TRANSACTION
                'DESCRIPTION: CHECKING FOR CUSTOMER'S DEPOSIT AND ADD UPON FULL PAYMENT OF THE INVOICE AMOUNT
                If J_INVOICENUM <> "" Then
                    Dim rsCheckDeposit As ADODB.Recordset
                    Set rsCheckDeposit = New ADODB.Recordset
                    
                    If COMPANY_CODE = "DPI" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DMI" Then
                        'UPDATE BY RRL 11/12/2015
                        'DESCRIPTION: ADDED OR_NUM FOR MORE ACCURATE INFO OF GETTING CUSTOMER DEPOSIT
                        rsCheckDeposit.Open "Select DP.PAIDFOR,DT.AMOUNT,DT.INVOICENO from CMIS_DepositDT DT INNER JOIN CMIS_DEPOSITS DP ON DP.ID=DT.DEPOSIT_ID where DT.InvoiceNo = '" & J_INVOICENUM & "' AND DT.OR_NUM='" & CMIS_OR_NUM & "' and DP.cuscde = " & J_CUSTOMERCODE & " And DT.InvoiceType = '" & J_INVOICETYPE2 & "'", gconDMIS, adOpenForwardOnly
                    Else
                        rsCheckDeposit.Open "Select DP.PAIDFOR,DT.AMOUNT,DT.INVOICENO from CMIS_DepositDT DT INNER JOIN CMIS_DEPOSITS DP ON DP.ID=DT.DEPOSIT_ID where DT.InvoiceNo = '" & J_INVOICENUM & "' and DP.cuscde = " & J_CUSTOMERCODE & " And DT.InvoiceType = '" & J_INVOICETYPE2 & "'", gconDMIS, adOpenForwardOnly
                    End If
                    
                    If Not rsCheckDeposit.EOF And Not rsCheckDeposit.BOF Then
                        Do While Not rsCheckDeposit.EOF
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR))))
                            J_DEBIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                            J_DEPOSIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                            
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                            
                            gconDMIS.Execute SQL_STATEMENT
                            
                            Dim rsGETACCTID As ADODB.Recordset
                            Set rsGETACCTID = New ADODB.Recordset
                            Set rsGETACCTID = gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "")
                            
                            xJournalDETID = ""
                            
                            If Not rsGETACCTID.EOF And Not rsGETACCTID.BOF Then
                                xJournalDETID = Null2String(rsGETACCTID!ID)
                            End If
                            
                            Dim rsDEPOSITDETAILS As ADODB.Recordset
                            Dim rsCHKDEPOSITAP As ADODB.Recordset
                            Set rsDEPOSITDETAILS = New ADODB.Recordset
                            Set rsCHKIFCLOSE = New ADODB.Recordset
                            Set rsDEPOSITDETAILS = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_NUM = '" & CMIS_OR_NUM & "' AND INVOICENO = '" & Null2String(rsCheckDeposit!INVOICENO) & "'  AND DESCRIPT = 'DEPOSIT APPLIED'")
                            If Not rsDEPOSITDETAILS.EOF And Not rsDEPOSITDETAILS.BOF Then
                                Set rsCHKIFCLOSE = New ADODB.Recordset
                                Set rsCHKIFCLOSE = gconDMIS.Execute("SELECT INVOICETYPEWITHNO,JTYPE,VOUCHERNO,JDATE,ACCOUNTNAME,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMT,BALANCE,VENDORCODE,ENTITYCODE " & _
                                                                    " FROM (" & _
                                                                    " SELECT CASE WHEN LEN(VOUCHERNO)=10 THEN LEFT(VOUCHERNO,3) ELSE LEFT(VOUCHERNO,2) END AS JTYPE,RIGHT(VOUCHERNO,6) AS VOUCHERNO,JDATE,VENDOR_CODE AS VENDORCODE,VENDOR_NAME AS ACCOUNTNAME,INVOICENO AS INVOICETYPEWITHNO,INVOICETYPE,INVOICENO,INVOICEDATE,AMOUNT2PAY AS INVOICEAMT,ENTITYCODE,AMOUNT2PAY - ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS AD WHERE (AP.INVOICENO=AD.INVOICENO AND AP.VENDOR_CODE=AD.VENDORCODE AND AP.ACCT_CODE=AD.ACCT_CODE)),0) AS BALANCE " & _
                                                                    " FROM AMIS_AP AP WHERE ACCT_CODE = " & N2Str2Null(J_ACCT_CODE) & ") T " & _
                                                                    " WHERE BALANCE <> 0 AND VENDORCODE = '" & CMIS_CUSCDE & "' And INVOICENO = '" & Null2String(rsDEPOSITDETAILS!Reference) & "'")
                                
                                If Not rsCHKIFCLOSE.EOF And Not rsCHKIFCLOSE.BOF Then
                                    SQL_STATEMENT = "INSERT INTO AMIS_DETAILS(VENDORCODE,VOUCHERNO,JTYPE,JDATE,INVOICENO,INVOICETYPE,AMOUNTPAID,ACCT_CODE,PV_VOUCHERNO,INVOICEDATE,ENTITYCODE,REFCODE,STATUS,JOURNAL_DET_ID) " & _
                                                    "VALUES('" & CMIS_CUSCDE & "'," & J_VOUCHERNO & "," & J_JTYPE & "," & J_JDATE & ",'" & Null2String(rsCHKIFCLOSE!INVOICETYPEWITHNO) & "','" & Null2String(rsCHKIFCLOSE!INVOICETYPE) & "'," & J_DEBIT & "," & J_ACCT_CODE & _
                                                    ",'" & Null2String(rsCHKIFCLOSE!JTYPE) + "-" + Null2String(rsCHKIFCLOSE!VOUCHERNO) & "','" & Null2String(rsCHKIFCLOSE!JDATE) & "','" & Null2String(rsCHKIFCLOSE!ENTITYCODE) & "', '" & Null2String(rsCHKIFCLOSE!ENTITYCODE) + CMIS_CUSCDE & "','N'," & xJournalDETID & " )"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                            End If

                            rsCheckDeposit.MoveNext
                        Loop
                    End If
                    Set rsCheckDeposit = Nothing
                End If
                
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 1 AND ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 0 AND ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                End If
                
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst: SJ_PV_ITEMNO = 0
                    Do While Not rsOFF_DT.EOF
                        CMIS_DT_TRANTYPE = Null2String(rsOFF_DT!TranType)
                        CMIS_DT_REFERENCE = Null2String(rsOFF_DT!INVOICENO)
                        CMIS_DT_CUSCDE = N2Str2Null(rsOFF_DT!CUSCDE)
                        CMIS_DT_DESCRIPT = Null2String(rsOFF_DT!DESCRIPT)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!amount)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!amount)
                        CMIS_DT_DOCDTE = Null2String(rsOFF_DT!DOCDTE)
                        CMIS_DT_PAYMENT = N2Str2Zero(rsOFF_DT!payment)
                        CMIS_DT_DISCOUNT = N2Str2Zero(rsOFF_DT!DISCOUNT)
                        CMIS_DT_TAX = N2Str2Zero(rsOFF_DT!tax)
                        CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)
                        CMIS_DT_REFCODE = Null2String(rsOFF_DT!ENTITY)
                        
                        'DESCRIPTION: CREDIT CARD DETAIL
                        CMIS_DT_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_JVOUCHERNO = J_VOUCHERNO
                        SJ_PV_ITEMNO = SJ_PV_ITEMNO + 1
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                        PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)
                        PV_AMOUNT = CMIS_DT_PAYMENT
                        PV_STATUS = "'N'"
                        J_BANKCHARGES = 0
                        PV_INVDATE = N2Date2Null(rsOFF_DT!ORDATE)
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        
                        Dim PV_ACCTCODE                     As String
                        Set rsSJ_DATA = New ADODB.Recordset
                        Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            DEF_INVOICETYPE = N2Str2Null(rsSJ_DATA!INVOICETYPE)
                            DEF_INVOICENO = N2Str2Null(rsSJ_DATA!INVOICENO)
                            J_JVOUCHERNO = J_VOUCHERNO
                            PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                            PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)
                            PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)
                            
                            If CMIS_DT_TAX = 0 Then
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT), 2)
                            Else
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                            End If
                            PV_STATUS = "'N'"
                        Else
                            'DESCRIPTION: INSERT DETAIL FOR A/R CREDIT CARD / CREDIT CARD PAYMENT RECEIVE
                            If CheckIfBank(CMIS_CUSCDE) = True And (CMIS_DT_PAIDFOR = "427") Then
                                Set rsCreditCard = New ADODB.Recordset
                                rsCreditCard.Open "SELECT * FROM CMIS_CARDBANK WHERE CUSCDE ='" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                                If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                                    BANKCHARGES = NumericVal(rsCreditCard!BANKCHARGES) / 100
                                    EWT = NumericVal(rsCreditCard!EWT) / 100
                                    TOTALCHARGES = 1 - (BANKCHARGES + EWT)
                                End If
                            ElseIf CheckIfBank(CMIS_CUSCDE) = True And (CMIS_DT_PAIDFOR = "478") Then
                                Set rsCreditCard = New ADODB.Recordset
                                rsCreditCard.Open "SELECT * FROM CMIS_CARDBANK WHERE CUSCDE ='" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                                If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                                    BANKCHARGES = NumericVal(rsCreditCard!BANKCHARGES) / 100
                                    EWT = NumericVal(rsCreditCard!EWT) / 100
                                    TOTALCHARGES = 1 - (BANKCHARGES + EWT)
                                End If
                            End If
                            
                            DEF_INVOICETYPE = N2Str2Null(DEF_INVOICETYPE)
                            DEF_INVOICENO = N2Str2Null(DEF_INVOICENO)
                        End If
                        
                        If CMIS_DT_TRANTYPE = "SI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HMR" Then
                                Set rsDEFERRED = New ADODB.Recordset
                                rsDEFERRED.Open "SELECT DT.CREDIT FROM AMIS_JOURNAL_HD HD " & _
                                "INNER JOIN AMIS_JOURNAL_DET DT " & _
                                "ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE " & _
                                "WHERE HD.INVOICENO=" & DEF_INVOICENO & " AND HD.INVOICETYPE=" & DEF_INVOICETYPE & " AND HD.CUSTOMERCODE = " & CMIS_DT_CUSCDE & " AND DT.ACCT_CODE='21-05001-00'", gconDMIS, adOpenForwardOnly
                                If Not rsDEFERRED.EOF And Not rsDEFERRED.BOF Then
                                    DEFERRED_OUTPUT = NumericVal(rsDEFERRED!Credit)
                                End If
                                
                                J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                J_DEBIT = DEFERRED_OUTPUT
                                J_CREDIT = 0
                                J_TAX = 0
                                J_GROSS = 0
                                J_NET = 0
                                J_STATUS = "'N'"
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                
                                Set rsUEA = New ADODB.Recordset
                                Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                If Not rsUEA.EOF And Not rsUEA.BOF Then
                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                Else
                                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                                Call CRJ_DEFFERED_TAX
                            ElseIf COMPANY_CODE = "DJM" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                
                                'J_DEBIT = Round(CMIS_DT_PAYMENT + CMIS_DT_TAX, 2)
                                'JULIE 12192015 - TO INCLUDE CUSTOMER DEPOSIT
                                If COMPANY_CODE = "DJM" Then
                                    J_DEBIT = Round(NumericVal(NumericVal(Round(CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX, 2) / 1.12) * 0.12), 2)
                                ElseIf COMPANY_CODE = "CMC" Then
                                    J_DEBIT = Round(NumericVal(NumericVal(Round(CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX, 2) / 1.12) * 0.02), 2)
                                Else
                                    PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                                    J_DEBIT = Round(NumericVal(NumericVal(Round(CMIS_DT_PAYMENT + CMIS_DT_TAX, 2) / 1.12) * 0.12), 2)
                                End If
                                
                                J_CREDIT = 0
                                J_TAX = 0
                                J_GROSS = 0
                                J_NET = 0
                                J_STATUS = "'N'"
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                
                                Set rsUEA = New ADODB.Recordset
                                Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                If Not rsUEA.EOF And Not rsUEA.BOF Then
                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                Else
                                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                                Call CRJ_DEFFERED_TAX
'                            ElseIf COMPANY_CODE = "HCA" Then
'                                Dim xOPT As Double
'                                xOPT = 0
'                                Set rsDEFERRED = New ADODB.Recordset
'                                rsDEFERRED.Open "SELECT DT.CREDIT AS DOTAMOUNT,CA.TRANTYPE1 AS DOTTYPE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DT ON HD.VOUCHERNO=DT.VOUCHERNO AND HD.JTYPE=DT.JTYPE " & _
'                                "INNER JOIN AMIS_CHARTACCOUNT CA ON CA.ACCTCODE = DT.ACCT_CODE " & _
'                                "WHERE HD.INVOICENO=" & DEF_INVOICENO & " AND HD.INVOICETYPE=" & DEF_INVOICETYPE & " AND HD.CUSTOMERCODE = " & CMIS_DT_CUSCDE & " AND CA.TRANTYPE1 IN ('DOTPARTS','DOTSERVICE') ", gconDMIS, adOpenForwardOnly
'                                If Not rsDEFERRED.EOF And Not rsDEFERRED.BOF Then
'                                    Do While Not rsDEFERRED.EOF
'                                        Dim DOTTYPE As String
'                                        DOTTYPE = (rsDEFERRED!DOTTYPE)
'                                        xOPT = xOPT + N2Str2Zero(rsDEFERRED!DOTAMOUNT)
'                                        If DOTTYPE = "DOTSERVICE" Then
'                                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("DOTSERVICE"))
'                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("DOTSERVICE")))
'                                            J_DEBIT = N2Str2Zero(Round(rsDEFERRED!DOTAMOUNT, 2))
'                                            J_CREDIT = 0
'                                            J_TAX = 0
'                                            J_GROSS = 0
'                                            J_NET = 0
'                                            J_STATUS = "'N'"
'                                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
'
'                                            Set rsUEA = New ADODB.Recordset
'                                            Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                            If Not rsUEA.EOF And Not rsUEA.BOF Then
'                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                            Else
'                                                J_ITEMCOUNT = J_ITEMCOUNT + 1
'                                                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
'                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
'                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
'                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
'                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
'                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
'                                                gconDMIS.Execute SQL_STATEMENT
'                                            End If
'
'                                            'HCA AMIS_DETAILS
'                                            Set DOTAXDS = New ADODB.Recordset
'                                            Set DOTAXDS = gconDMIS.Execute("Select* from amis_ap ap inner join amis_chartaccount ca on ap.acct_code = ca.acctcode  Where ap.INVOICENO = '" & CMIS_DT_REFERENCE & "' AND ap.INVOICETYPE = '" & CMIS_DT_TRANTYPE & "' AND ap.VENDOR_CODE = " & CMIS_DT_CUSCDE & " and ca.IS_SCHEDULE_ACCNT = 1 AND TRANTYPE1 = 'DOTSERVICE'  ")
'                                            If Not DOTAXDS.EOF And Not DOTAXDS.BOF Then
'                                                xJournalDETID = ""
'                                                xJournalDETID = N2Str2IntZero(gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "").Fields(0).Value)
'                                                SQL_STATEMENT = "INSERT INTO AMIS_DETAILS(VENDORCODE,VOUCHERNO,JTYPE,JDATE,INVOICENO,INVOICETYPE,AMOUNTPAID,ACCT_CODE,PV_VOUCHERNO,INVOICEDATE,ENTITYCODE,REFCODE,STATUS,JOURNAL_DET_ID)" & _
'                                                "VALUES(" & CMIS_DT_CUSCDE & "," & J_VOUCHERNO & "," & J_JTYPE & "," & J_JDATE & ",'" & CMIS_DT_REFERENCE & "','" & CMIS_DT_TRANTYPE & "'," & J_DEBIT & "," & J_ACCT_CODE & "," & N2Str2Null(DOTAXDS!VOUCHERNO) & "," & N2Str2Null(DOTAXDS!JDATE) & "," & N2Str2Null(DOTAXDS!ENTITYCODE) & ", " & N2Str2Null(DOTAXDS!REFCODE) & ",'N'," & xJournalDETID & " )"
'                                                gconDMIS.Execute SQL_STATEMENT
'                                            End If
'                                        End If
'
'                                        If DOTTYPE = "DOTPARTS" Then
'                                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("DOTPARTS"))
'                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("DOTPARTS")))
'                                            J_DEBIT = N2Str2Zero(Round(rsDEFERRED!DOTAMOUNT, 2))
'                                            J_CREDIT = 0
'                                            J_TAX = 0
'                                            J_GROSS = 0
'                                            J_NET = 0
'                                            J_STATUS = "'N'"
'                                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
'                                            Set rsUEA = New ADODB.Recordset
'                                            Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                            If Not rsUEA.EOF And Not rsUEA.BOF Then
'                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                            Else
'                                                J_ITEMCOUNT = J_ITEMCOUNT + 1
'                                                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
'                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
'                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
'                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
'                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
'                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
'                                                gconDMIS.Execute SQL_STATEMENT
'                                            End If
'
'                                            'HCA AMIS_DETAILS
'                                            Set DOTAXDS = New ADODB.Recordset
'                                            Set DOTAXDS = gconDMIS.Execute("Select* from amis_ap ap inner join amis_chartaccount ca on ap.acct_code = ca.acctcode  Where ap.INVOICENO = '" & CMIS_DT_REFERENCE & "' AND ap.INVOICETYPE = '" & CMIS_DT_TRANTYPE & "' AND ap.VENDOR_CODE = " & CMIS_DT_CUSCDE & " and ca.IS_SCHEDULE_ACCNT = 1 AND TRANTYPE1 = 'DOTSERVICE'  ")
'                                            If Not DOTAXDS.EOF And Not DOTAXDS.BOF Then
'                                                xJournalDETID = ""
'                                                xJournalDETID = N2Str2IntZero(gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "").Fields(0).Value)
'                                                SQL_STATEMENT = "INSERT INTO AMIS_DETAILS(VENDORCODE,VOUCHERNO,JTYPE,JDATE,INVOICENO,INVOICETYPE,AMOUNTPAID,ACCT_CODE,PV_VOUCHERNO,INVOICEDATE,ENTITYCODE,REFCODE,STATUS,JOURNAL_DET_ID)" & _
'                                                "VALUES(" & CMIS_DT_CUSCDE & "," & J_VOUCHERNO & "," & J_JTYPE & "," & J_JDATE & ",'" & CMIS_DT_REFERENCE & "','" & CMIS_DT_TRANTYPE & "'," & J_DEBIT & "," & J_ACCT_CODE & "," & N2Str2Null(DOTAXDS!VOUCHERNO) & "," & N2Str2Null(DOTAXDS!JDATE) & "," & N2Str2Null(DOTAXDS!ENTITYCODE) & ", " & N2Str2Null(DOTAXDS!REFCODE) & ",'N'," & xJournalDETID & " )"
'                                                gconDMIS.Execute SQL_STATEMENT
'                                            End If
'                                        End If
'                                        rsDEFERRED.MoveNext
'                                    Loop
'                                End If
                            End If
                            
                            If COMPANY_CODE = "DJM" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HMR" Or COMPANY_CODE = "DSSC" Then
                                If COMPANY_CODE = "HCA" And xOPT = 0 Then
                                Else
                                    If J_DEBIT = 0 Then
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                        J_CREDIT = CMIS_DT_PAYMENT
                                        
                                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HMR" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" And (CMIS_DT_TAX > 0) Then
                                            J_CREDIT = Round(CMIS_DT_PAYMENT + CMIS_DT_TAX, 2)
                                        ElseIf COMPANY_CODE = "HCA" Then
                                            J_CREDIT = xOPT
                                            J_DEBIT = 0
                                        End If
                                        
                                        If COMPANY_CODE = "HCA" Then
                                        ElseIf COMPANY_CODE = "DJM" Then
                                            'JULIE 12192015 - TO INCLUDE CUSTOMER DEPOSIT
                                            J_CREDIT = Round(NumericVal(NumericVal(Round(CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX, 2) / 1.12) * 0.12), 2)
                                        Else
                                            J_CREDIT = Round(NumericVal(NumericVal(J_CREDIT / 1.12) * 0.12), 2)
                                        End If
                                        
                                        J_TAX = 0
                                        J_DEBIT = 0
                                        J_GROSS = 0
                                        J_NET = 0
                                        J_STATUS = "'N'"
                                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                        Set rsUEA = New ADODB.Recordset
                                        Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        'RO  - SERVICE REPAIR ORDER
                        If CMIS_DT_TRANTYPE = "RO" Or CMIS_DT_TRANTYPE = "SI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HMR" Then
                                INSURANCE = ""
                                Set rsCSMIOS_CHKINS = New ADODB.Recordset
                                rsCSMIOS_CHKINS.Open "SELECT PARTICIPAT FROM CSMS_REPOR WHERE INVOICE = '" & CMIS_DT_REFERENCE & "' AND PARTICIPAT = '" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                                
                                If Not rsCSMIOS_CHKINS.EOF And Not rsCSMIOS_CHKINS.BOF Then
                                    INSURANCE = N2Str2Null(rsCSMIOS_CHKINS!PARTICIPAT)
                                End If
                                
                                If INSURANCE = "" Then
                                    If COMPANY_CODE = "HCE" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                    Else
                                        If ReturnSITerm(CMIS_DT_REFERENCE) = "CHG" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                        Else
                                            If COMPANY_CODE = "FMC" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                            Else
                                                J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                            End If
                                        End If
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("INSURANCE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("INSURANCE")))
                                End If
                            ElseIf COMPANY_CODE = "FMC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
                                If ReturnSITerm(CMIS_DT_REFERENCE) = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            ElseIf COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("TRADE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("TRADE")))
                            Else
                                If COMPANY_CODE = "HSM" Then
                                    INSURANCE = ""
                                    
                                    Set rsCSMIOS_CHKINS = New ADODB.Recordset
                                    rsCSMIOS_CHKINS.Open "SELECT ISNULL(PARTICIPAT,'') AS PARTICIPAT FROM CSMS_REPOR WHERE INVOICE = '" & CMIS_DT_REFERENCE & "' AND PARTICIPAT = '" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                                
                                    If Not rsCSMIOS_CHKINS.EOF And Not rsCSMIOS_CHKINS.BOF Then
                                        INSURANCE = Null2String(rsCSMIOS_CHKINS!PARTICIPAT)
                                    End If
                                    
                                    If INSURANCE = "" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("INSURANCE"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("INSURANCE")))
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                End If
                                
                            End If
                        End If
                        
                        If CMIS_DT_TRANTYPE = "PI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMR" Then
                                If ReturnTranType(CMIS_DT_REFERENCE, "P") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            ElseIf COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("TRADE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("TRADE")))
                            ElseIf COMPANY_CODE = "HCC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            ElseIf COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        End If
                        
                        'AI - ACCESSORIES INVOICE
                        If CMIS_DT_TRANTYPE = "AI" Then
                            If COMPANY_CODE = "DJM" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                            ElseIf COMPANY_CODE = "CMC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        End If
                        
                        'MI - MATERIALS INVOICE
                        If CMIS_DT_TRANTYPE = "MI" Then
                            If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DJM" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        End If
                        
                        'VI  - VEHICLE INVOICE
                        If CMIS_DT_TRANTYPE = "VI" Then
                            If COMPANY_CODE = "HCC" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HCA" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                            ElseIf COMPANY_CODE = "DJM" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Then
                                FINANCE = ""
                                Set rsCSMIOS_FINANCE = New ADODB.Recordset
                                rsCSMIOS_FINANCE.Open "SELECT FINANCINGCODE FROM SMIS_SalesOrder WHERE FINANCINGCODE = '" & CMIS_CUSCDE & "' AND VI_NO = '" & CMIS_DT_REFERENCE & "'", gconDMIS, adOpenForwardOnly
                                If Not rsCSMIOS_FINANCE.EOF And Not rsCSMIOS_FINANCE.BOF Then
                                    FINANCE = N2Str2Null(rsCSMIOS_FINANCE!FINANCINGCODE)
                                End If
                                
                                If FINANCE = "" Then
                                    If COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                                    ElseIf COMPANY_CODE = "CMC" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                                    Else
                                        If CMIS_DT_PAIDFOR = "" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                                        End If
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("FINANCING"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("FINANCING")))
                                End If
                            ElseIf COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("TRADE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("TRADE")))
                            Else
                                If CMIS_DT_PAIDFOR = "" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                                End If
                            End If
                        End If
                        
                        'EST - SERVICE ESTIMATE
                        If CMIS_DT_TRANTYPE = "EST" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDeposit_AccountCode("SERVICE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeposit_AccountCode("SERVICE")))
                        End If
                        
                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then    ' BTT
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("ACCESSORIES")))
                            End If
                        ElseIf COMPANY_CODE = "DAI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCC" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMR" Then
                            If CMIS_DT_TRANTYPE = "AI" Then
                                If COMPANY_CODE = "DAI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HCR" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                ElseIf ReturnTranType(CMIS_DT_REFERENCE, "A") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                ElseIf COMPANY_CODE = "HCC" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            End If
                        ElseIf COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Then
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            End If
                        ElseIf COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("TRADE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("TRADE")))
                            End If
                        ElseIf COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        End If
                        
                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HNE" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMR" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                If ReturnTranType(CMIS_DT_REFERENCE, "A") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            End If
                        'UPDATE BY KATH 08.08.15
                        ElseIf COMPANY_CODE = "DAI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        ElseIf COMPANY_CODE = "MGS" Or COMPANY_CODE = "HGS" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("TRADE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("TRADE")))
                            End If
                        ElseIf COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            End If
                        End If
                        
                        'OTHER TRANSACTION
                        If CMIS_DT_TRANTYPE = "OTH" Then
                            If COMPANY_CODE = "DJM" Then
                                If CMIS_DT_PAIDFOR = "482" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                                End If
                            Else
                                J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                                J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                            End If
                        End If
                        
                        If CMIS_DT_TRANTYPE = "OI" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            J_ACCT_CODE = N2Str2Null(rsSJ_DATA!J_CLASS)
                            J_ACCT_NAME = N2Str2Null(Setacctname(rsSJ_DATA!J_CLASS))
                        End If
                        
                        'CUSTOMER'S DEPOSIT
                        If J_INVOICENUM <> "" Then
                            Dim rsCheckDeposit2             As ADODB.Recordset
                            Set rsCheckDeposit2 = New ADODB.Recordset
                            rsCheckDeposit2.Open "Select * from CMIS_Deposits where InvoiceNo = '" & J_INVOICENUM & "' and invoicetype = '" & J_INVOICETYPE2 & "'", gconDMIS, adOpenForwardOnly
                            If Not rsCheckDeposit2.EOF And Not rsCheckDeposit2.BOF Then
                                J_DEPOSIT = Round(NumericVal(rsCheckDeposit2!amount), 2)
                            End If
                            Set rsCheckDeposit2 = Nothing
                        End If
                        
                        If COMPANY_CODE = "DJM" Then
                            J_GROSS = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                        ElseIf COMPANY_CODE = "CMC" Then
                            J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                        Else
                            J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                        End If
                        J_TAX = 0
                        
                        If CMIS_DT_TRANTYPE = "OTH" Then
                            If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CMIS_DT_PAIDFOR = "417" Then
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT / 1.12), 2)
                            ElseIf COMPANY_CODE = "DSSC" Then
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                            ElseIf COMPANY_CODE = "DJM" And (CMIS_DT_PAIDFOR = "427" Or CMIS_DT_PAIDFOR = "478" Or CMIS_DT_PAIDFOR = "479") Then
                                J_NET = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                            ElseIf COMPANY_CODE = "DJM" And (CMIS_DT_PAIDFOR = "419") Then
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                            Else
                                If COMPANY_CODE = "DJM" Then
                                    J_NET = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                                Else
                                    J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                                End If
                            End If
                        Else
                            If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DSSC" Then
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(CMIS_DT_TAX), 2) + Round(NumericVal(J_DEPOSIT), 2)
                            Else
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(J_DEPOSIT), 2)
                            End If
                        End If
                        
                        J_DEBIT = 0
                        
                        'CREDIT CARD
                        If J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD")) Or J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND")) Then
                            If CheckIfBank(CMIS_CUSCDE) = True And (CMIS_DT_PAIDFOR = "427" Or CMIS_DT_PAIDFOR = "478") Then
                                PV_MRRNO = "'CI'"
                                PV_INVDATE = GetCustomerORDate(CMIS_DT_REFERENCE)
                                
                                If (COMPANY_CODE = "DJM" Or COMPANY_CODE = "DAI") And (CMIS_DT_PAIDFOR = "478" Or CMIS_DT_PAIDFOR = "427") Then
                                    PV_AMOUNT = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                                ElseIf COMPANY_CODE = "DAI" Then
                                    PV_AMOUNT = Round(GetCustomerORAmount(CMIS_DT_REFERENCE) - (CMIS_DT_DISCOUNT) - (CMIS_DT_TAX), 2)
                                Else
                                    PV_AMOUNT = Round(GetCustomerORAmount(CMIS_DT_REFERENCE), 2)
                                End If
                                
                                If GetifNonVat(CMIS_DT_REFERENCE) = False Then
                                Else
                                    PV_INVNO = CMIS_DT_REFERENCE
                                    PV_INVNO = "NV" + PV_INVNO
                                    PV_INVNO = String2N(PV_INVNO)
                                End If
                                
                                J_CREDIT = Round(PV_AMOUNT, 2)
                                J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                                
                                Set rsCHKSCHEDULED = New ADODB.Recordset
                                rsCHKSCHEDULED.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE AcctCode = " & J_ACCT_CODE & " AND IS_SCHEDULE_ACCNT = 1", gconDMIS, adOpenForwardOnly
                                If Not rsCHKSCHEDULED.EOF And Not rsCHKSCHEDULED.BOF Then
                                    SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                    "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE,J_CLASS)" & _
                                    " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                    ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_INVDATE & ", " & PV_AMOUNT & _
                                    ", " & PV_STATUS & ",'" & CMIS_CUSCDE & "'," & J_ACCT_CODE & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                                
                                Set rsUEA = New ADODB.Recordset
                                Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                If Not rsUEA.EOF And Not rsUEA.BOF Then
                                    If J_CREDIT > 0 Then
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                    Else
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                    End If
                                Else
                                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity,invoiceno)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & "," & N2Str2Null(CMIS_DT_REFERENCE) & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                            End If
                            
                            If COMPANY_CODE = "HLB" Then
                                J_CREDIT = Round(GetCustomerORAmount(CMIS_OR_NUM), 2)
                            ElseIf (COMPANY_CODE = "DJM" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "CMC") And (CMIS_DT_PAIDFOR = "427" Or CMIS_DT_PAIDFOR = "478" Or CMIS_DT_PAIDFOR = "479" Or CMIS_DT_PAIDFOR = "482" Or CMIS_DT_PAIDFOR = "471" Or CMIS_DT_PAIDFOR = "CR001" Or CMIS_DT_PAIDFOR = "427A") Then
                                J_CREDIT = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                            '021816 -- DESCRIPTION FOR 999 PLEASE?? -SARAH HERE
                            ElseIf COMPANY_CODE = "DPI" And CMIS_DT_PAIDFOR = "999" Then
                                J_CREDIT = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE), 2)
                            ElseIf COMPANY_CODE = "CMC" And (CMIS_DT_PAIDFOR = "427" Or CMIS_DT_PAIDFOR = "427A" Or CMIS_DT_PAIDFOR = "471") Then
                                J_CREDIT = Round(GetCustomerORAmount_478(CMIS_DT_REFERENCE) + (CMIS_DT_DISCOUNT) + (CMIS_DT_TAX), 2)
                            Else
                                J_CREDIT = Round(GetCustomerORAmount(CMIS_DT_REFERENCE), 2)
                            End If
                        Else
                            If COMPANY_CODE = "HCA" Or COMPANY_CODE = "DSSC" Then
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                            ElseIf (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CMIS_DT_PAIDFOR = "417" Then
                                J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT / 1.12), 2)
                            ElseIf COMPANY_CODE = "DJM" Then
                                If CMIS_DT_PAIDFOR = "425" Or CMIS_DT_PAIDFOR = "426" Or CMIS_DT_PAIDFOR = "428" Or CMIS_DT_PAIDFOR = "422" Then
                                    J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT / 1.12), 2)
                                ElseIf CMIS_DT_PAIDFOR = "SII" Then
                                    If CMIS_TYPE_PAYMENT = 1 Then
                                        J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT / 1.12), 2)
                                    ElseIf CMIS_TYPE_PAYMENT = 2 Then
                                        J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                                    ElseIf CMIS_TYPE_PAYMENT = 3 Then
                                        J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                                    End If
                                ElseIf CMIS_DT_PAIDFOR = "417" Then
                                    If CMIS_DT_TAX > 0 Then
                                        J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX) / 1.12, 2)
                                    Else
                                        J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT / 1.12), 2)
                                    End If
                                Else
                                    'JULIE 12192015 - TO INCLUDE CUSTOMER DEPOSIT
                                    'OLD J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                                    J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                                End If
                            ElseIf COMPANY_CODE = "CMC" Then
                                J_CREDIT = Round(NumericVal(CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                            Else
                                J_CREDIT = Round(PV_AMOUNT, 2)
                            End If
                        End If
                        
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                        
                        If CheckIfBank(CMIS_CUSCDE) = True And (CMIS_DT_PAIDFOR = "427" Or CMIS_DT_PAIDFOR = "478" Or CMIS_DT_PAIDFOR = "479") Then
                        Else
                            If J_ACCT_CODE = "NULL" Or J_ACCT_NAME = "NULL" Then
                                MsgBox "Not yet configure in CMIS Other Transactions. OR# " & CMIS_OR_NUM, vbInformation, "Other Transactions"
                                Exit Function
                            Else
                                Dim xUEA As Double
                                Set rsUEA = New ADODB.Recordset
                                Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                If Not rsUEA.EOF And Not rsUEA.BOF Then
                                    If J_CREDIT > 0 Then
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                    Else
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                    End If
                                Else
                                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity,invoiceno)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & "," & N2Str2Null(J_INVOICENO & "/" & CMIS_DT_REFERENCE) & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                                
                                If (CMIS_DT_TRANTYPE = "OTH" And CMIS_DT_PAIDFOR = "415") Or (CMIS_DT_TRANTYPE = "OTH" And CMIS_DT_PAIDFOR = "421" And COMPANY_CODE <> "HMH") Then
                                    Dim rsCHKSCHEDULE As New ADODB.Recordset
                                    Set rsCHKSCHEDULE = New ADODB.Recordset
                                    Set rsCHKSCHEDULE = gconDMIS.Execute("SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = " & J_ACCT_CODE & "  AND IS_SCHEDULE_ACCNT = 1")
                                    If Not rsCHKSCHEDULE.EOF And Not rsCHKSCHEDULE.BOF Then
                                        If COMPANY_CODE = "DGI" Then
                                            SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,ENTITYCODE,REFCODE,DUEDATE,JOURNAL_DET_ID) " & _
                                            "VALUES(" & N2Str2Null(J_JTYPE & "-" & J_VOUCHERNO) & "," & J_INVOICETYPE & "," & N2Str2Null(J_INVOICENO & "/" & CMIS_DT_REFERENCE) & ",'" & Right(CMIS_DT_REFCODE, 6) & "','" & SetVendorName(Right(CMIS_DT_REFCODE, 6)) & "'," & J_CREDIT & ",'0'," & J_CREDIT & "," & J_ACCT_CODE & "," & J_JDATE & "," & LOGDATE & "," & J_JDATE & ",'V'," & N2Str2Null("C" + J_CUSTOMERCODE) & "," & J_JDATE & "," & TransactionID & ")"
                                        ElseIf COMPANY_CODE = "HCE" Then
                                            SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,ENTITYCODE,REFCODE,DUEDATE,JOURNAL_DET_ID) " & _
                                            "VALUES(" & N2Str2Null(J_JTYPE & "-" & J_VOUCHERNO) & "," & J_INVOICETYPE & "," & N2Str2Null(J_INVOICENO & "/" & CMIS_DT_REFERENCE) & ",'" & CMIS_CUSCDE & "','" & CMIS_CUSNAME & "'," & J_CREDIT & ",'0'," & J_CREDIT & "," & J_ACCT_CODE & "," & J_JDATE & "," & LOGDATE & "," & J_JDATE & ",'V'," & N2Str2Null("C" + J_CUSTOMERCODE) & "," & J_JDATE & "," & TransactionID & ")"
                                        Else
                                            SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,ENTITYCODE,REFCODE,DUEDATE,JOURNAL_DET_ID) " & _
                                            "VALUES(" & N2Str2Null(J_JTYPE & "-" & J_VOUCHERNO) & "," & J_INVOICETYPE & "," & N2Str2Null(J_INVOICENO & "/" & CMIS_DT_REFERENCE) & ",'" & Right(CMIS_DT_REFCODE, 9) & "','" & SetVendorName(Right(CMIS_DT_REFCODE, 9)) & "'," & J_CREDIT & ",'0'," & J_CREDIT & "," & J_ACCT_CODE & "," & J_JDATE & "," & LOGDATE & "," & J_JDATE & ",'V'," & N2Str2Null("C" + J_CUSTOMERCODE) & "," & J_JDATE & "," & TransactionID & ")"
                                        End If
                                        gconDMIS.Execute SQL_STATEMENT
                                    End If
                                End If
                            End If
                        End If
                        
'START OF INSERTING TO CRJ DETAIL==============================================================================================================================================================================
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        Set rsSJ_DATA = New ADODB.Recordset
                        Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & "  and CUSTOMERCODE = '" & CMIS_CUSCDE & "' and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            DEF_INVOICETYPE = N2Str2Null(rsSJ_DATA!INVOICETYPE)
                            DEF_INVOICENO = N2Str2Null(rsSJ_DATA!INVOICENO)
                            J_JVOUCHERNO = J_VOUCHERNO
                            PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                            PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)
                            PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)
                            If CMIS_DT_TAX = 0 Then
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT), 2)
                            Else
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                            End If
                            PV_STATUS = "'N'"
                            
                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                            "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE,J_CLASS)" & _
                            " values (" & J_JVOUCHERNO & "," & J_JDATE & "," & N2Str2Null(rsSJ_DATA!VOUCHERNO) & ", " & PV_ITEMNO & _
                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                            ", " & PV_STATUS & "," & N2Str2Null(rsSJ_DATA!CustomerCode) & "," & J_ACCT_CODE & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(PV_MRRNO)
                        End If
                        
                        'WITH ARCREDITABLE2307
                        If CMIS_DT_TAX > 0 And (UCase(CMIS_DT_PAIDFOR) = "412S" Or UCase(CMIS_DT_PAIDFOR) = "412P" Or UCase(CMIS_DT_PAIDFOR) = "412V") And COMPANY_CODE = "HMH" Then
                            PV_AMOUNT = Round((CMIS_DT_PAYMENT + CMIS_DT_TAX), 2)
                        End If
                        
                        If COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "HGS" Or COMPANY_CODE = "HCA" Then
                        ElseIf COMPANY_CODE = "" Then
                        ElseIf UCase(CMIS_DT_PAIDFOR) = "412S" Or UCase(CMIS_DT_PAIDFOR) = "412P" Or UCase(CMIS_DT_PAIDFOR) = "412V" Or (UCase(CMIS_DT_PAIDFOR) = "422" And COMPANY_CODE = "HCE") Then
                            Set rsCHKSCHEDULED = New ADODB.Recordset
                            rsCHKSCHEDULED.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE AcctCode = " & J_ACCT_CODE & " AND IS_SCHEDULE_ACCNT = 1", gconDMIS, adOpenForwardOnly
                            If Not rsCHKSCHEDULED.EOF And Not rsCHKSCHEDULED.BOF Then
                                Dim AMIS_CRJ_DETAIL As ADODB.Recordset
                                Dim AMISCRJDET As String
                                AMISCRJDET = ""
                                Set AMIS_CRJ_DETAIL = New ADODB.Recordset
                                AMIS_CRJ_DETAIL.Open "SELECT * FROM AMIS_CRJ_DETAIL WHERE INVOICENO='" & CMIS_OR_NUM & "' AND J_CLASS = " & J_ACCT_CODE & " AND CUSTOMERCODE = " & J_CUSTOMERCODE & "", gconDMIS, adOpenForwardOnly
                                If Not AMIS_CRJ_DETAIL.EOF And Not AMIS_CRJ_DETAIL.BOF Then
                                    AMISCRJDET = N2Str2Null(AMIS_CRJ_DETAIL!J_CLASS)
                                End If
                                If AMISCRJDET = "" Then
                                    If xCDEPOSITCHECK = "" Then
                                        SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                        "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE,J_CLASS)" & _
                                        " values (" & J_JVOUCHERNO & "," & J_JDATE & ",NULL, " & PV_ITEMNO & _
                                        ", 'CI', '" & CMIS_OR_NUM & "', " & PV_INVDATE & ", " & PV_AMOUNT & _
                                        ", " & PV_STATUS & "," & J_CUSTOMERCODE & "," & J_ACCT_CODE & ")"
                                        gconDMIS.Execute SQL_STATEMENT
                                    Else
                                        SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                        "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE,J_CLASS)" & _
                                        " values (" & J_JVOUCHERNO & "," & J_JDATE & ",NULL, " & PV_ITEMNO & _
                                        ", 'CI'," & N2Str2Null(CMIS_OR_NUM & "/" & CMIS_DT_REFERENCE) & ", " & PV_INVDATE & ", " & PV_AMOUNT & _
                                        ", " & PV_STATUS & "," & J_CUSTOMERCODE & "," & J_ACCT_CODE & ")"
                                        gconDMIS.Execute SQL_STATEMENT
                                    End If
                                Else
                                    SQL_STATEMENT = "UPDATE AMIS_CRJ_DETAIL SET INVOICEAMOUNT = INVOICEAMOUNT + " & PV_AMOUNT & " WHERE INVOICENO='" & CMIS_OR_NUM & "' AND J_CLASS = " & J_ACCT_CODE & " AND CUSTOMERCODE = " & J_CUSTOMERCODE & " "
                                    gconDMIS.Execute SQL_STATEMENT
                                    
                                    SQL_STATEMENT = "UPDATE AMIS_AP SET AMOUNT2PAY = AMOUNT2PAY + " & PV_AMOUNT & ", BALANCE = BALANCE + " & PV_AMOUNT & " WHERE INVOICENO='" & CMIS_OR_NUM & "' AND ACCT_CODE = " & J_ACCT_CODE & " AND VENDOR_CODE = " & J_CUSTOMERCODE & " AND ENTITYCODE ='C'"
                                    gconDMIS.Execute SQL_STATEMENT
                                End If
                            End If
                        End If
                        
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
'END OF INSERTING TO CRJ DETAIL==============================================================================================================================================================================

'START FOR OTHER INCOME INCENTIVES OUTPUT TAX================================================================================================================================================================
                        If ((COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CMIS_DT_PAIDFOR = "417") Or (COMPANY_CODE = "DJM" And (CMIS_DT_PAIDFOR = "425" Or CMIS_DT_PAIDFOR = "426" Or CMIS_DT_PAIDFOR = "428" Or CMIS_DT_PAIDFOR = "422" Or CMIS_DT_PAIDFOR = "417" Or CMIS_DT_PAIDFOR = "SII")) Then
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = J_ITEMCOUNT
                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                            J_DEBIT = 0
                            
                            If COMPANY_CODE <> "DJM" Then
                                J_CREDIT = Round(NumericVal(Round((CMIS_DT_PAYMENT / 1.12), 2) * 0.12), 2)
                            Else
                                J_CREDIT = Round(NumericVal(((CMIS_DT_PAYMENT + CMIS_DT_TAX) / 1.12)) * 0.12, 2)
                            End If
                            
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Set rsUEA = New ADODB.Recordset
                            Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                            If Not rsUEA.EOF And Not rsUEA.BOF Then
                                If COMPANY_CODE = "DJM" Then
                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & ",CREDIT=CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                Else
                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                End If
                            Else
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                NEW_LogAudit "MM", "JOURNAL IMPORT", SQL_STATEMENT, J_VOUCHERNO, "", CSMIOS_REP_OR, J_JTYPE, J_JNO
                            End If
                        End If
'END FOR OTHER INCOME INCENTIVES OUTPUT TAX================================================================================================================================================================
                        
'START FOR SII DEFERRED OUTPUT TAX=========================================================================================================================================================================
                        If COMPANY_CODE = "DJM" And CMIS_DT_PAIDFOR = "SII" Then
                            J_DEBIT = Round(NumericVal(((CMIS_DT_PAYMENT + CMIS_DT_TAX) / 1.12)) * 0.12, 2)
                            If J_DEBIT > 0 Then
                                J_CREDIT = 0
                                J_ITEMCOUNT = J_ITEMCOUNT + 1
                                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                Set rsSII = New ADODB.Recordset
                                Set rsSII = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                If Not rsSII.EOF And Not rsSII.BOF Then
                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                Else
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity,invoiceno)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & "," & N2Str2Null(J_INVOICENO & "/" & CMIS_DT_REFERENCE) & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    
                                    If CMIS_CASHAMOUNT > 0 Then
                                        J_DEBIT = 0: J_CREDIT = 0
                                        J_CREDIT = CMIS_DT_PAYMENT + CMIS_DT_TAX
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR)) & "")
                                        NEW_LogAudit "MM", "JOURNAL IMPORT", SQL_STATEMENT, J_VOUCHERNO, "", CSMIOS_REP_OR, J_JTYPE, J_JNO
                                    ElseIf CMIS_CHKAMOUNT > 0 Then
                                        J_DEBIT = 0: J_CREDIT = 0
                                        J_CREDIT = CMIS_DT_PAYMENT + CMIS_DT_TAX
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR)) & "")
                                        NEW_LogAudit "MM", "JOURNAL IMPORT", SQL_STATEMENT, J_VOUCHERNO, "", CSMIOS_REP_OR, J_JTYPE, J_JNO
                                    ElseIf CMIS_CARDAMOUNT > 0 Then
                                        If CMIS_DT_TRANTYPE = "OTH" Then
                                        Else
                                        J_DEBIT = 0: J_CREDIT = 0
                                        J_DEBIT = CMIS_DT_PAYMENT - Round(NumericVal(Round((CMIS_DT_PAYMENT / 1.12), 2) * 0.12), 2)
                                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & N2Str2Null(ReturnAccountCode("CARD ON HAND")) & "")
                                        NEW_LogAudit "MM", "JOURNAL IMPORT", SQL_STATEMENT, J_VOUCHERNO, "", CSMIOS_REP_OR, J_JTYPE, J_JNO
                                        End If
                                    End If
                                    
                                End If
                            End If
                        End If
                        
                        rsOFF_DT.MoveNext
                    Loop
                End If
'END FOR SII DEFERRED OUTPUT TAX=========================================================================================================================================================================
                
                Set rsCreditCard = New ADODB.Recordset
                If COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HGS" Then
                    rsCreditCard.Open "select  ISNULL(SUM(DISCOUNT),0) AS BANKCHARGES,ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and descript = 'CREDIT CARD PAYMENT RCVD'", gconDMIS, adOpenForwardOnly
                ElseIf COMPANY_CODE = "HSM" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DPI" Then
                    rsCreditCard.Open "select  ISNULL(SUM(DISCOUNT),0) AS BANKCHARGES,ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and descript = 'CREDIT CARD PMENT RCVD'", gconDMIS, adOpenForwardOnly
                ElseIf COMPANY_CODE = "DJM" Then
                    rsCreditCard.Open "select  ISNULL(SUM(DISCOUNT),0) AS BANKCHARGES,ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' AND (PAIDFOR IN ('427','478') OR PAIDFOR IS NULL)", gconDMIS, adOpenForwardOnly
                Else
                    rsCreditCard.Open "select  ISNULL(SUM(DISCOUNT),0) AS BANKCHARGES,ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and descript = 'CREDIT CARD PYMENT RECEIVED'", gconDMIS, adOpenForwardOnly
                End If
                
                If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                    BANKCHARGES = 0
                    EWT = 0
                    BANKCHARGES = Round(rsCreditCard!BANKCHARGES, 2)
                    EWT = Round(rsCreditCard!EWT, 2)
                    
                    'BANK CHARGES
                    If BANKCHARGES > 0 Then
                        If COMPANY_CODE = "DJM" And (CMIS_DT_PAIDFOR = "478" Or CMIS_DT_PAIDFOR = "427") Then
                            Set rsCreditCardCost = New ADODB.Recordset
                            rsCreditCardCost.Open "SELECT (SELECT TRANTYPE FROM CMIS_OFF_DT WHERE OR_NUM=A.REFERENCE and descript <>'deposit applied') AS TRANTYPE_INV,DISCOUNT,TAX AS EWT, * FROM CMIS_OFF_DT A WHERE OR_NUM = '" & CMIS_OR_NUM & "'", gconDMIS, adOpenForwardOnly
                            
                            'QUERY REVISED 112316 resolution for tcn19888
                            'rsCreditCardCost.Open "SELECT TRANTYPE AS TRANTYPE_INV,DISCOUNT,TAX AS EWT,paidfor, * FROM CMIS_OFF_DT A WHERE descript <>'deposit applied' and OR_NUM = '" & CMIS_OR_NUM & "'", gconDMIS, adOpenForwardOnly
                            If Not rsCreditCardCost.EOF And Not rsCreditCardCost.BOF Then
                                
                                rsCreditCardCost.MoveFirst
                                Do While Not rsCreditCardCost.EOF
                                    If rsCreditCardCost!TRANTYPE_INV = "SI" Then
                                        'COST - GJ - CUSTOMER
                                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                                        J_JITEMNO = "'" & Format(J_ITEMCOUNT + 1, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "CUSTOMER"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "CUSTOMER")))
                                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                                        J_DEBIT = NumericVal(Round(rsCreditCardCost!DISCOUNT / 2, 2))
                                        J_CREDIT = 0
                                        J_TAX = 0
                                        J_GROSS = 0
                                        J_NET = 0
                                        J_STATUS = "'N'"
                                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                        
                                        Set rsCreditCardCostDistinc = New ADODB.Recordset
                                        Set rsCreditCardCostDistinc = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsCreditCardCostDistinc.EOF And Not rsCreditCardCostDistinc.BOF Then
                                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = 'CRJ' AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                                            
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                        
                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                        
                                        'COST - PARTS - WORKSHOP CUSTOMER
                                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                                        J_JITEMNO = "'" & Format(J_ITEMCOUNT + 1, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "CUSTOMER"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "CUSTOMER")))
                                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                                        J_DEBIT = NumericVal(Round(rsCreditCardCost!DISCOUNT / 2, 2))
                                        J_CREDIT = 0
                                        J_TAX = 0
                                        J_GROSS = 0
                                        J_NET = 0
                                        J_STATUS = "'N'"
                                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                        
                                        Set rsCreditCardCostDistinc = New ADODB.Recordset
                                        Set rsCreditCardCostDistinc = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsCreditCardCostDistinc.EOF And Not rsCreditCardCostDistinc.BOF Then
                                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = 'CRJ' AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                                            
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                        
                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                        
                                    ElseIf rsCreditCardCost!TRANTYPE_INV = "PI" Then
                                        'COST - PARTS - RETAIL CUSTOMER
                                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("COSTPARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("COSTPARTS")))
                                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                                        J_DEBIT = rsCreditCardCost!DISCOUNT
                                        J_CREDIT = 0
                                        J_TAX = 0
                                        J_GROSS = 0
                                        J_NET = 0
                                        J_STATUS = "'N'"
                                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                        
                                        Set rsCreditCardCostDistinc = New ADODB.Recordset
                                        Set rsCreditCardCostDistinc = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsCreditCardCostDistinc.EOF And Not rsCreditCardCostDistinc.BOF Then
                                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = 'CRJ' AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                                            
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                        
                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                    
                                    ElseIf rsCreditCardCost!TRANTYPE_INV = "OTH" Then
                                    'OTH added 112316 resolution for tcn19888
'                                        J_ITEMCOUNT = J_ITEMCOUNT + 1
'                                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
'                                        J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(Null2String(rsCreditCardCost!PAIDFOR)))
'                                        J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(Null2String(rsCreditCardCost!PAIDFOR))))
'                                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
'                                        J_DEBIT = rsCreditCardCost!DISCOUNT
'                                        J_CREDIT = 0
'                                        J_TAX = 0
'                                        J_GROSS = 0
'                                        J_NET = 0
'                                        J_STATUS = "'N'"
'                                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
'
'                                        Set rsCreditCardCostDistinc = New ADODB.Recordset
'                                        Set rsCreditCardCostDistinc = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                        If Not rsCreditCardCostDistinc.EOF And Not rsCreditCardCostDistinc.BOF Then
'                                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = 'CRJ' AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
'                                        Else
'                                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
'                                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
'                                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
'                                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
'                                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
'
'                                            gconDMIS.Execute SQL_STATEMENT
'                                        End If
'
'                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
'                                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                    End If
                                    
                                    rsCreditCardCost.MoveNext
                                Loop
                            End If
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("BANK CHARGES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("BANK CHARGES")))
                            J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                            
                            J_DEBIT = NumericVal(Round(BANKCHARGES, 2))
                            
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                            
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                            
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                        End If
                    End If
                    
                    'EWT
                    If EWT > 0 Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HNE" Or COMPANY_CODE = "FMC" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CWT"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CWT")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDITABLE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDITABLE")))
                        End If
                        J_DEBIT = EWT
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                    End If
                End If
                
                Dim rsEWT As New ADODB.Recordset
                Set rsEWT = New ADODB.Recordset
                If COMPANY_CODE = "MGS" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HGS" Then
                    Set rsEWT = gconDMIS.Execute("select  ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and ISNULL(descript,'') <> 'CREDIT CARD PAYMENT RCVD'")
                ElseIf COMPANY_CODE = "HSM" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DAI" Then
                    Set rsEWT = gconDMIS.Execute("select  ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and ISNULL(descript,'') <> 'CREDIT CARD PMENT RCVD'")
                ElseIf COMPANY_CODE = "DJM" Then
                    Set rsEWT = gconDMIS.Execute("select  ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and (PAIDFOR <> '427' and PAIDFOR <> '478')")
                ElseIf COMPANY_CODE = "CMC" Then
                    Set rsEWT = gconDMIS.Execute("select  ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsEWT = gconDMIS.Execute("select  ISNULL(SUM(TAX),0) as EWT from cmis_off_dt where or_num = '" & CMIS_OR_NUM & "' and ISNULL(descript,'') <> 'CREDIT CARD PYMENT RECEIVED'")
                End If
                
                If Not rsEWT.EOF And Not rsEWT.BOF Then
                    BANKCHARGES = 0
                    EWT = 0
                    EWT = Round(rsEWT!EWT, 2)

                    If EWT > 0 Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HNE" Or COMPANY_CODE = "FMC" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CWT"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CWT")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDITABLE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDITABLE")))
                        End If
                        J_DEBIT = EWT
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                    End If
                End If
                
'START FORCE BALANCE======================================================================================================================================
    Call StartForceBalance(J_VOUCHERNO)
        
'START FOR REMARKS========================================================================================================================================
                Dim xREMARKS As String
                Dim rsREMARKS As New ADODB.Recordset
                
                If COMPANY_CODE = "HSM" Or COMPANY_CODE = "DJM" Then
                Else
                    J_REMARKS = ""
                End If
                
                If COMPANY_CODE = "HSM" Or COMPANY_CODE = "DJM" Then
                ElseIf COMPANY_CODE = "CMC" Then
                    Set rsREMARKS = New ADODB.Recordset
                    Set rsREMARKS = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE  ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = '" & CMIS_OR_NUM & "'")
                    If Not rsREMARKS.EOF And Not rsREMARKS.BOF Then
                        Do While Not rsREMARKS.EOF
                            If CMIS_DT_TRANTYPE = "OTH" Then
                                J_REMARKS = Null2String(rsREMARKS!DESCRIPT) & ": " & Null2String(rsREMARKS!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsREMARKS!payment))
                            ElseIf CMIS_DT_TRANTYPE = "VI" Then
                                J_REMARKS = J_REMARKS & xInvCMC & " | " & xInvDate & " | " & xInvTerm
                            Else
                                J_REMARKS = J_REMARKS + SetTransaction(Null2String(rsREMARKS!TranType)) & ": " & Null2String(rsREMARKS!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsREMARKS!payment)) & " "
                            End If
                            rsREMARKS.MoveNext
                        Loop
                    End If
                Else
                    Set rsREMARKS = New ADODB.Recordset
                    Set rsREMARKS = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE  ISNULL(DESCRIPT,0) <> 'DEPOSIT APPLIED' AND OR_NUM = '" & CMIS_OR_NUM & "'")
                    If Not rsREMARKS.EOF And Not rsREMARKS.BOF Then
                        Do While Not rsREMARKS.EOF
                            If COMPANY_CODE <> "DJM" And CMIS_DT_TRANTYPE = "OTH" Then
                                J_REMARKS = Null2String(rsREMARKS!DESCRIPT) & ": " & Null2String(rsREMARKS!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsREMARKS!payment))
                            ElseIf COMPANY_CODE = "DJM" Then
                                J_REMARKS = J_REMARKS + Null2String(rsREMARKS!DESCRIPT)
                            Else
                                J_REMARKS = J_REMARKS + SetTransaction(Null2String(rsREMARKS!TranType)) & ": " & Null2String(rsREMARKS!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsREMARKS!payment)) & " "
                            End If
                            rsREMARKS.MoveNext
                        Loop
                    End If
                End If
'END FOR REMARKS========================================================================================================================================
                
'INSERT HEADER HERE=====================================================================================================================================
                If COMPANY_CODE = "HSM" Then
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,ReferenceNo,Bank,Entity_Class)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                    ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & "," & J_CUSTOMERCODE & "," & N2Str2Null(CMIS_CARDBANK) & "," & J_ENTITY & ")"
                Else
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,ReferenceNo,Bank,Entity_Class)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                    ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & ",'" & J_REMARKS & "'," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & "," & J_CUSTOMERCODE & "," & N2Str2Null(CMIS_CARDBANK) & "," & J_ENTITY & ")"
                End If
                
                gconDMIS.Execute SQL_STATEMENT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "CASH RECEIPTS JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                
                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
SKIP_OR:
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed - [" & CMIS_OR_NUM & "]"
        DoEvents
    Next
    Screen.MousePointer = 0
    
    ImportUnDeposit = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportUnDeposit = False
End Function

Sub StartForceBalance(xVOUCHERNO As String)
    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "FMC" Then
        Dim rsCheckDC As ADODB.Recordset
        DCTOTAL = 0
        Set rsCheckDC = New ADODB.Recordset
        Set rsCheckDC = gconDMIS.Execute("Select round(sum(Isnull(debit,0))-sum(isnull(credit,0)),2) as DCTOTAL from amis_journal_det where jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & " ")
        If Not rsCheckDC.EOF And Not rsCheckDC.BOF Then
            DCTOTAL = (rsCheckDC!DCTOTAL)
            If DCTOTAL = 0.02 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + (+0.02) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('OUTPUT TAX')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            ElseIf DCTOTAL = -0.02 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + (-0.02) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('OUTPUT TAX')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            ElseIf DCTOTAL = -0.01 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + (-0.01) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('OUTPUT TAX')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            End If
        End If
        
        DCTOTAL = 0
        Set rsCheckDC = New ADODB.Recordset
        Set rsCheckDC = gconDMIS.Execute("Select round(sum(Isnull(debit,0))-sum(isnull(credit,0)),2) as DCTOTAL from amis_journal_det where jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & " ")
        If Not rsCheckDC.EOF And Not rsCheckDC.BOF Then
            DCTOTAL = (rsCheckDC!DCTOTAL)
            If DCTOTAL = 0.02 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + (-0.02) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            ElseIf DCTOTAL = -0.02 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + (+0.02) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            ElseIf DCTOTAL = -0.01 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + (+0.01) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            ElseIf DCTOTAL = 0.01 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + (-0.01) WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            End If
        End If
    ElseIf COMPANY_CODE = "HCA" Or COMPANY_CODE = "HMR" Then
        Set rsCheckDC = New ADODB.Recordset
        Set rsCheckDC = gconDMIS.Execute("Select round(sum(Isnull(debit,0))-sum(isnull(credit,0)),2) as DCTOTAL from amis_journal_det where jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & " ")
        If Not rsCheckDC.EOF And Not rsCheckDC.BOF Then
            DCTOTAL = (rsCheckDC!DCTOTAL)
            If DCTOTAL = 0.01 Or DCTOTAL = 0.02 Or DCTOTAL = 0.03 Or DCTOTAL = 0.04 Or DCTOTAL = 0.05 Then
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnAccountCode("ROD"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("ROD")))
                J_DEBIT = EWT
                J_CREDIT = DCTOTAL
                J_DEBIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                
                gconDMIS.Execute SQL_STATEMENT
                
            ElseIf DCTOTAL = -0.01 Or DCTOTAL = -0.02 Or DCTOTAL = -0.03 Or DCTOTAL = -0.04 Or DCTOTAL = -0.05 Then
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnAccountCode("ROD"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("ROD")))
                J_DEBIT = EWT
                J_DEBIT = DCTOTAL * (-1)
                J_CREDIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                
                gconDMIS.Execute SQL_STATEMENT
            End If
        End If
        
    ElseIf COMPANY_CODE = "DJM" Then
        Set rsCheckDC = New ADODB.Recordset
        Set rsCheckDC = gconDMIS.Execute("Select SUM(ISNULL(DEBIT,0)) - SUM(ISNULL(CREDIT,0)) as DCTOTAL from amis_journal_det where jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & " ")
        If Not rsCheckDC.EOF And Not rsCheckDC.BOF Then
            DCTOTAL = (rsCheckDC!DCTOTAL)
            J_DEBIT = DCTOTAL
            
            'HIGHER DEBIT
            If DCTOTAL = 0.01 Or DCTOTAL = 0.02 Or DCTOTAL = 0.03 Or DCTOTAL = 0.04 Or DCTOTAL = 0.05 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT - " & J_DEBIT & " WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            'HIGHER CREDIT
            ElseIf DCTOTAL = -0.01 Or DCTOTAL = -0.02 Or DCTOTAL = -0.03 Or DCTOTAL = -0.04 Or DCTOTAL = -0.05 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT - " & J_DEBIT & " WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 IN ('CWT')) AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            End If
            
        End If
        
        Set rsCheckDC = New ADODB.Recordset
        Set rsCheckDC = gconDMIS.Execute("Select SUM(ISNULL(DEBIT,0)) - SUM(ISNULL(CREDIT,0)) as DCTOTAL from amis_journal_det where jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & " ")
        If Not rsCheckDC.EOF And Not rsCheckDC.BOF Then
            DCTOTAL = (rsCheckDC!DCTOTAL)
            J_DEBIT = DCTOTAL
            
            'HIGHER DEBIT
            If DCTOTAL = 0.01 Or DCTOTAL = 0.02 Or DCTOTAL = 0.03 Or DCTOTAL = 0.04 Or DCTOTAL = 0.05 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT - " & J_DEBIT & " WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = 'PARTS-GJ' AND TRANTYPE2='SERVICE' AND TRANTYPE3='COST OF SALES' AND TRANTYPE4='CUSTOMER') AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            'HIGHER CREDIT
            ElseIf DCTOTAL = -0.01 Or DCTOTAL = -0.02 Or DCTOTAL = -0.03 Or DCTOTAL = -0.04 Or DCTOTAL = -0.05 Then
                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT - " & J_DEBIT & " WHERE ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE1 = 'PARTS-GJ' AND TRANTYPE2='SERVICE' AND TRANTYPE3='COST OF SALES' AND TRANTYPE4='CUSTOMER') AND jtype = 'CRJ' and Voucherno =" & xVOUCHERNO & "")
            End If
            
        End If
    End If
End Sub

'DESCRIPTION: IMPORTING OF DEPOSITED PER BANKNAME DJM
'SJR 12072015
Function ImportDeposited_DJM() As Boolean
    On Error GoTo ErrorCode
    
    Dim J_JDATE                                             As String
    Dim J_VOUCHERNO                                         As String
    Dim J_JTYPE                                             As String
    Dim J_JNO                                               As String
    Dim J_REMARKS_CIB                                       As String
    Dim J_REMARKS_COH                                       As String
    Dim J_VENDORCODE                                        As String
    Dim J_CUSTOMERCODE                                      As String
    
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_CHECKDATE, J_BANKCODE                             As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS                                             As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String
    Dim J_ACCT_CODE, J_ACCT_NAME                            As String
    Dim J_STATUS                                            As String
    Dim J_JITEMNO                                           As String
    
    Dim CMIS_BANKACCTCODE                                   As String
    Dim CMIS_BANKCODE                                       As String
    Dim CMIS_DEPOSITED_DATE                                 As String
    Dim CMIS_OR_AMT                                         As String
    Dim CMIS_PAY_TYPE                                       As String
    Dim CMIS_DISCOUNT                                       As String
    Dim CMIS_TAX                                            As String
    Dim CMIS_CUSCDE                                         As String
    Dim CMIS_CUSNAME                                        As String
    Dim CMIS_DEPOSITED_AMT                                   As String
    Dim CMIS_TSEKE                                          As String
    Dim CMIS_CHECKDATE                                      As String
    Dim CMIS_STATUS                                         As String
    Dim CMIS_TYPE_PAYMENT                                   As String
    Dim CMIS_OR_NUM_DEP                                     As String
    
    Dim CMIS_INCASHCHK                                      As String
    Dim CMIS_INCASHAMT                                      As String
    Dim CMIS_CHECKNUM                                       As String
    Dim CMIS_CHKAMOUNT                                      As Double
    Dim CMIS_ORNUM                                          As String
    Dim J_REMARKS1                                          As String
    Dim J_REMARKS2                                          As String
    
    Dim INVOICENULL                                         As String
    Dim ENTITY_CODE                                         As String
    Dim ENTITY_NAME                                         As String
    Dim CMIS_BANK_DEPOSITED                                 As String
    
    Dim J_OUTBALANCE                                        As Double
    Dim J_AMOUNTTOPAY                                       As Double
    Dim J_INVOICEAMT                                        As Double
    Dim J_BALANCE                                           As Double
    Dim J_AMOUNTPAID                                        As Double
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET            As Double
    Dim CMIS_CASHAMOUNT                                     As Double
    Dim CMIS_CARDAMOUNT                                     As Double
    Dim TOTAL_DEBIT, TOTAL_CREDIT                           As Double
    
    Dim CMIS_IS_VAT                                         As Boolean
    
    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Dim rsOFF_HD                                            As ADODB.Recordset
    Dim rsOFF_DT                                            As ADODB.Recordset
    Dim rsREMARKS                                           As ADODB.Recordset
    
    Dim GridImport                                          As Integer
    
    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsBank = New ADODB.Recordset
            Set rsOFF_HD = New ADODB.Recordset
            Set rsOFF_HD2 = New ADODB.Recordset
            Set rsSum_Depslip = New ADODB.Recordset
            
'INSERT CASH IN BANK HERE
            Set rsOFF_HD = gconDMIS.Execute("SELECT DISTINCT DEPOSIT_TO, (SELECT BANKNAME FROM ALL_BANKDEPOSITS WHERE BANKCODE=CBD.DEPOSIT_TO) AS BANKNAME,DATDEPOSIT, SUM(DEPOSIT) AS SUMPERBANK FROM CMIS_BANKDEPO CBD WHERE DATDEPOSIT='" & CDate(dtpTranDate) & "' GROUP BY DEPOSIT_TO,DATDEPOSIT ORDER BY DEPOSIT_TO ASC")
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                
                Do While Not rsOFF_HD.EOF
                    CMIS_CUSCDE = Null2String(rsOFF_HD!DEPOSIT_TO)
                    CMIS_CUSNAME = Null2String(rsOFF_HD!BankName)
                    CMIS_BANKCODE = Null2String(rsOFF_HD!DEPOSIT_TO)
                    CMIS_DEPOSITED_DATE = Null2Date(rsOFF_HD!DATDEPOSIT)
                    CMIS_SUMPERBANK = Null2String(rsOFF_HD!SUMPERBANK)
                    CMIS_DEPOSITED_AMT = Null2String(rsOFF_HD!SUMPERBANK)
                    CMIS_STATUS = "N"
                    
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If
                    
                    J_JDATE = N2Date2Null(CMIS_DEPOSITED_DATE)
                    J_REFDATE = N2Date2Null(CMIS_DEPOSITED_DATE)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_REMARKS_CIB = "DEPOSITED TO: " & CMIS_CUSNAME & " | TOTAL AMOUNT DEPOSITED: " & ToDoubleNumber(CMIS_SUMPERBANK)
                    J_REFNO = "'NULL'"
                    J_JTYPE = "'GJ'"
                    
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = "NULL"
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = NumericVal(CMIS_SUMPERBANK)
                    J_BALANCE = 0
                    J_AMOUNTPAID = 0
                    J_INVOICENO = ""
                    J_STATUS = "'N'"
                    J_INVOICEDATE = N2Date2Null(CMIS_DEPOSITED_DATE)
                    J_INVOICENO = (Grid2.Cell(GridImport, 2).Text)
                    J_INVOICETYPE = "'CI'"
                    J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                    J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                    J_TERMS = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"
                    
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    J_JITEMNO = "0001"
                    J_ACCT_CODE = N2Str2Null(bankaccountcode(CMIS_CUSCDE))
                    J_ACCT_NAME = N2Str2Null(Setacctname(bankaccountcode(CMIS_CUSCDE)))
                    J_DEBIT = Round(NumericVal(CMIS_SUMPERBANK), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,invoiceno,INVOICETYPE,debit,credit,tax,grossamt,netamt,status,ADJ_remarks,referenceno)" & _
                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ",'" & J_INVOICENO & "'," & J_INVOICETYPE & ", " & J_DEBIT & _
                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & J_REMARKS_CIB & "'," & J_REFNO & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", " DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                    
                    rsOFF_HD.MoveNext
                Loop
            End If
        End If
        
        
'INSERT CASH ON HAND HERE
        Set rsCASHONHAND = New ADODB.Recordset
        Set rsCASHONHAND = gconDMIS.Execute("Select SUM(DEPOSIT) AS SUM_COH from CMIS_BANKDEPO WHERE DATDEPOSIT='" & CMIS_DEPOSITED_DATE & "'")
                
        If Not rsCASHONHAND.EOF And Not rsCASHONHAND.BOF Then
            J_JITEMNO = Format(NumericVal(J_JITEMNO), "0000")
            J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
            J_ACCT_CODE = Null2String(ReturnAccountCode("CASH ON HAND"))
            J_ACCT_NAME = N2Str2Null(Setacctname(J_ACCT_CODE))
            CMIS_DEPOSITED_AMT = Null2String(rsCASHONHAND!SUM_COH)
            
            J_REFNO = "NULL"
            CMIS_CUSCDE = ""
            CMIS_PAY_TYPE = ""
            J_REMARKS_COH = ""
            
            ENTITY_CODE = ""
            ENTITY_NAME = ""
            
            Set rsREMARKS = New ADODB.Recordset
            Set rsREMARKS = gconDMIS.Execute("Select * from CMIS_BANKDEPO WHERE DATDEPOSIT='" & CMIS_DEPOSITED_DATE & "'")
            
            If Not rsREMARKS.EOF And Not rsREMARKS.BOF Then
                Do While Not rsREMARKS.EOF
                    CMIS_OR_NUM_DEP = Null2String(rsREMARKS!OR_NUM)
                    
                    If J_REMARKS_COH = "" Then
                        J_REMARKS_COH = CMIS_OR_NUM_DEP
                    Else
                        J_REMARKS_COH = J_REMARKS_COH & "," & CMIS_OR_NUM_DEP
                    End If
                    
                rsREMARKS.MoveNext
                Loop
            End If
            
            J_GROSS = 0
            J_NET = 0
            J_STATUS = "'N'"
            J_DEBIT = 0
            J_CREDIT = 0
            J_CREDIT = CMIS_DEPOSITED_AMT
            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
            
            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                            "(jdate,voucherno,jtype,jno,jitemno, " & _
                            "acct_code,acct_name,invoiceno,INVOICETYPE, " & _
                            "debit,credit,tax,grossamt,netamt,status,adj_remarks,referenceno,entity)" & _
                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                            ", " & J_JNO & ", " & J_JITEMNO & ", '" & J_ACCT_CODE & "', " & J_ACCT_NAME & ", " & _
                            "'" & J_INVOICENO & "'," & J_INVOICETYPE & ", " & J_DEBIT & _
                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & _
                            "" & J_STATUS & ", '" & J_REMARKS_COH & "'," & J_REFNO & ", '" & ENTITY_CODE & "')"
            gconDMIS.Execute SQL_STATEMENT
            
            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
            NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
        End If
        
'INSERT HEADER HERE
        J_INVOICEAMT = CMIS_DEPOSITED_AMT
        J_BANKCODE = "'NULL'"
        J_REMARKS_CIB = "TOTAL AMOUNT DEPOSITED: " & ToDoubleNumber(J_INVOICEAMT)

        SQL_STATEMENT = "Insert into AMIS_Journal_HD " & _
                        " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoiceamt,refno,refdate,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,BankCode,invoiceno,remarks,PaidStatus,ReceiveStatus,referenceno)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & ",'" & J_CUSTOMERCODE & "', " & J_INVOICEDATE & "," & J_INVOICEAMT & "," & J_REFNO & "," & J_INVOICEDATE & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_BANKCODE & ",'" & J_INVOICENO & "','" & J_REMARKS_CIB & "'," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ", " & J_REFNO & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
        NEW_LogAudit "M", "DEPOSITED RECEIPTS JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
        
        Grid2.Cell(GridImport, 1).Text = 1
        
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    
    Screen.MousePointer = 0
    ImportDeposited_DJM = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportDeposited_DJM = False
End Function
'DESCRIPTION: IMPORTING OF DEPOSITED PER BANKNAME DJM
'SJR 12072015

'SJR 082114
'DESCRIPTION: IMPORTING OF DEPOSITED SLIP NUMBER FOR HCA
Function ImportDeposited_HCA() As Boolean
'HEADER
    On Error GoTo ErrorCode

    Dim J_JDATE                                             As String
    Dim J_VOUCHERNO                                         As String
    Dim J_JTYPE                                             As String
    Dim J_JNO                                               As String
    Dim J_REMARKS_CIB                                       As String
    Dim J_REMARKS_COH                                       As String
    Dim J_VENDORCODE                                        As String
    Dim J_CUSTOMERCODE                                      As String
    Dim J_OUTBALANCE                                        As Double
    Dim J_AMOUNTTOPAY                                       As Double
    Dim J_INVOICEAMT                                        As Double
    Dim J_BALANCE                                           As Double
    Dim J_AMOUNTPAID                                        As Double
    Dim J_CHECKNO                                           As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_CHECKDATE, J_BANKCODE                             As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS, J_DEALER                                   As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String

    'DETAIL
    Dim J_ACCT_CODE, J_ACCT_NAME                            As String
    Dim J_DEBIT                                             As Double
    Dim J_CREDIT                                            As Double
    Dim J_TAX                                               As Double
    Dim J_GROSS                                             As Double
    Dim J_NET                                               As Double
    Dim J_STATUS                                            As String
    Dim J_JITEMNO                                           As String

    Dim CMIS_DEPSLIPNUM                                     As String
    Dim CMIS_OR_DATE                                        As String
    Dim CMIS_OR_AMT                                         As String
    Dim CMIS_PAY_TYPE                                       As String
    Dim CMIS_DISCOUNT                                       As String
    Dim CMIS_TAX                                            As String
    Dim CMIS_CASHAMOUNT                                     As Double
    Dim CMIS_CARDAMOUNT                                     As Double
    Dim CMIS_CUSCDE                                         As String
    Dim CMIS_CUSNAME                                        As String
    Dim CMIS_DEPOSIT                                        As String
    Dim CMIS_BANKCODE                                       As String
    Dim CMIS_TSEKE                                          As String
    Dim CMIS_CHECKDATE                                      As String
    Dim CMIS_STATUS                                         As String
    Dim CMIS_TYPE_PAYMENT                                   As String
    Dim CMIS_OR_NUM_DEP                                     As String
    
    Dim CMIS_INCASHCHK                                      As String
    Dim CMIS_INCASHAMT                                      As String
    Dim CMIS_CHECKNUM                                       As String
    Dim CMIS_CHKAMOUNT                                      As Double
    Dim CMIS_ORNUM                                          As String
    Dim J_REMARKS1                                          As String
    Dim J_REMARKS2                                          As String
        
    Dim INVOICENULL                                         As String
    Dim ENTITY_CODE                                         As String
    Dim ENTITY_NAME                                         As String

    Dim CMIS_IS_VAT                                         As Boolean
    Dim CMIS_BANK_DEPOSITED                                 As String

    Dim TOTAL_DEBIT, TOTAL_CREDIT                           As Double
    
    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Dim rsOFF_HD                                            As ADODB.Recordset
    Dim rsOFF_DT                                            As ADODB.Recordset
    Dim rsEntity                                            As ADODB.Recordset
    Dim rsIncash                                            As ADODB.Recordset
    Dim rsCEParticulars                                     As ADODB.Recordset
    
    'GET CASH ON HAND HERE
    If COMPANY_CODE = "HCA" Then
        COA_CASH_ON_HAND = Null2String(ReturnAccountCode("CASH ON HAND"))
    Else
    End If
    
    Dim GridImport                                          As Integer
    
    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsBank = New ADODB.Recordset
            Set rsOFF_HD = New ADODB.Recordset
            Set rsOFF_HD2 = New ADODB.Recordset
            Set rsSum_Depslip = New ADODB.Recordset
            
            If COMPANY_CODE = "HCA" Then
                    Set rsOFF_HD = gconDMIS.Execute("Select distinct DEPSLIPNUM,DATDEPOSIT,DEPOSIT_TO,sum(deposit) as SUMDEPSLIP from CMIS_BankDepo Where DATDEPOSIT >= '" & GetCUTOFF_DATE & "' AND DATDEPOSIT <= '" & CDate(dtpTranDate) & "' AND DEPSLIPNUM = '" & Grid2.Cell(GridImport, 2).Text & "' and deposit_to='" & Grid2.Cell(GridImport, 4).Text & "' group by DEPSLIPNUM,DATDEPOSIT,DEPOSIT_TO Order by DEPSLIPNUM ASC")
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                    CMIS_DEPSLIPNUM = Null2String(rsOFF_HD!DEPSLIPNUM)
                    CMIS_OR_DATE = Null2Date(rsOFF_HD!DATDEPOSIT)
                    CMIS_SUMDEPSLIP = Null2String(rsOFF_HD!SUMDEPSLIP)
                    CMIS_CUSNAME = Null2String(rsOFF_HD!DEPOSIT_TO)
                    CMIS_DEPOSIT = Null2String(rsOFF_HD!SUMDEPSLIP)
                    CMIS_BANKCODE = Null2String(rsOFF_HD!DEPOSIT_TO)
'                    CMIS_INCASHCHK = Null2String(rsOFF_HD!INCASHCHK)
'                    CMIS_INCASHAMT = Null2String(rsOFF_HD!ChkAmount)
'                    CMIS_CHECKNUM = Null2String(rsOFF_HD!CHECKNUM)
'                    CMIS_ORNUM = Null2String(rsOFF_HD!ORNUM)
                    CMIS_STATUS = "N"
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
            Else
            
            End If
                
                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'DRJ'"

                Set rsOFF_DT = New ADODB.Recordset
                Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_BankDepo WHERE DEPSLIPNUM = '" & CMIS_DEPSLIPNUM & "' and deposit_to='" & CMIS_BANKCODE & "'")
                
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            J_REMARKS_CIB = "Deposited to: " & CMIS_CUSNAME & " / Total Deposited Amount: " & CMIS_SUMDEPSLIP
                            J_REFNO = N2Str2Null(rsOFF_DT!DEPSLIPNUM)
                            J_REFDATE = N2Date2Null(rsOFF_DT!DATDEPOSIT)
                            CMIS_PAY_TYPE = Null2String(rsOFF_DT!Type)
                            rsOFF_DT.MoveNext
                            If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                            End If
                    End If
                    
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = NumericVal(CMIS_SUMDEPSLIP)
                J_BALANCE = 0
                J_AMOUNTPAID = 0
                J_INVOICENO = ""
                J_STATUS = "'N'"
                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                J_INVOICENO = CMIS_DEPSLIPNUM
                INVOICENULL = "NULL"
                
                    If CMIS_PAY_TYPE = "1" Then
                    J_INVOICETYPE = "'CSH'"
                    ElseIf CMIS_PAY_TYPE = "2" Then
                    J_INVOICETYPE = "'CHK'"
                    Else
                    End If
                    
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'INSERT CASH IN BANK HERE
                If COMPANY_CODE = "HCA" Then
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    J_JITEMNO = "0001"
                    J_ACCT_CODE = N2Str2Null(bankaccountcode(CMIS_CUSNAME))
                    J_ACCT_NAME = N2Str2Null(Setacctname(bankaccountcode(CMIS_CUSNAME)))
                    J_REMARKS_CIB = "Deposited to: " & CMIS_CUSNAME & " / Total Deposited Amount: " & CMIS_SUMDEPSLIP
                    J_DEBIT = Round(NumericVal(CMIS_SUMDEPSLIP), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,invoiceno,INVOICETYPE,debit,credit,tax,grossamt,netamt,status,ADJ_remarks,referenceno)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ",'" & INVOICENULL & "'," & J_INVOICETYPE & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & J_REMARKS_CIB & "'," & J_REFNO & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", " DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    'INSERT CASH ON HAND HERE
                    Set rsCASHONHAND = New ADODB.Recordset
                    Set rsCASHONHAND = gconDMIS.Execute("Select * from CMIS_BankDepo " & _
                                                        "WHERE DEPSLIPNUM = '" & CMIS_DEPSLIPNUM & "' " & _
                                                        "and deposit_to='" & CMIS_BANKCODE & "'")

                    If COMPANY_CODE = "HCA" Then
                    If Not rsCASHONHAND.EOF And Not rsCASHONHAND.BOF Then
                        Do While Not rsCASHONHAND.EOF
                            J_JITEMNO = Format(NumericVal(J_JITEMNO), "0000")
                            J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
                            J_ACCT_CODE = N2Str2Null(COA_CASH_ON_HAND)
                            J_ACCT_NAME = N2Str2Null(Setacctname(COA_CASH_ON_HAND))
                            CMIS_DEPOSIT = Null2String(rsCASHONHAND!DEPOSIT)
                            CMIS_OR_NUM_DEP = Null2String(rsCASHONHAND!OR_NUM)
                            CMIS_CUSCDE = Null2String(rsCASHONHAND!BANKCODE)
                            CMIS_OR_NUM_DEP = Null2String(rsCASHONHAND!OR_NUM)
                            CMIS_PAY_TYPE = Null2String(rsCASHONHAND!Type)
                            
                            If CMIS_PAY_TYPE = True Then
                            Set rsCEParticulars = New ADODB.Recordset
                            Set rsCEParticulars = gconDMIS.Execute("SELECT ORNUM,CHKAMOUNT FROM CMIS_INCASH WHERE CHKNUMBER='" & Mid(CMIS_OR_NUM_DEP, 4) & "' AND BANKCODE='" & CMIS_CUSCDE & "'")
                            J_REMARKS_COH = ""
                            J_REMARKS1 = ""
                            J_REMARKS2 = ""
                            If Not rsCEParticulars.EOF And Not rsCEParticulars.BOF Then
                                Do While Not rsCEParticulars.EOF
                                    J_REMARKS1 = " ," & Null2String(rsCEParticulars!ORNUM) & ":" & Null2String(rsCEParticulars!ChkAmount)
                                    J_REMARKS2 = J_REMARKS2 & J_REMARKS1
                                rsCEParticulars.MoveNext
                                Loop
                            J_REMARKS_COH = "Check Encashment: " & CMIS_DEPOSIT & " / OR NO: " & J_REMARKS2
                            End If
                            Else
                            End If
                            
                            Set rsEntity = New ADODB.Recordset
                            Set rsEntity = gconDMIS.Execute("SELECT * FROM " & _
                                                            "(SELECT BANKCODE AS CODE, " & _
                                                            "BANKNAME AS ACCOUNTNAME " & _
                                                            "FROM CMIS_BANKS " & _
                                                            "UNION " & _
                                                            "SELECT CODE AS CODE, " & _
                                                            "ACCOUNTNAME AS ACCOUNTNAME " & _
                                                            "FROM ALL_ENTITY) AS TABLEENTIY " & _
                                                            "WHERE TABLEENTIY.CODE='" & CMIS_CUSCDE & "'")
                            
                            'SJR 082614
                            ENTITY_CODE = "": ENTITY_NAME = ""
                            ENTITY_CODE = "C" & Null2String(rsEntity!Code)
                            ENTITY_NAME = Null2String(rsEntity!ACCOUNTNAME)
                            
                                If CMIS_PAY_TYPE = "1" Then
                                J_INVOICETYPE = "'CSH'"
                                ElseIf CMIS_PAY_TYPE = "2" Then
                                J_INVOICETYPE = "'CHK'"
                                Else
                                End If
                                
                            J_REFNO = N2Str2Null(rsCASHONHAND!DEPSLIPNUM)
                            J_CREDIT = Round(NumericVal(CMIS_DEPOSIT), 2)
                            J_DEBIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno, " & _
                                    "acct_code,acct_name,invoiceno,INVOICETYPE, " & _
                                    "debit,credit,tax,grossamt,netamt,status,adj_remarks,referenceno,entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & _
                                    "'" & CMIS_OR_NUM_DEP & "'," & J_INVOICETYPE & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & _
                                    "" & J_STATUS & ", '" & J_REMARKS_COH & "'," & J_REFNO & ", '" & ENTITY_CODE & "')"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                        rsCASHONHAND.MoveNext
                        Loop
                    End If
'                    End If
                    
                'INSERT HEADERS HERE
                If COMPANY_CODE = "HCA" Then
                SQL_STATEMENT = "Insert into AMIS_Journal_HD " & _
                        " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoiceamt,refno,refdate,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,BankCode,remarks,PaidStatus,ReceiveStatus,referenceno)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & ",'" & CMIS_CUSNAME & "', " & J_INVOICEDATE & "," & J_INVOICEAMT & "," & J_REFNO & "," & J_INVOICEDATE & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_BANKCODE & ",'" & J_REMARKS_CIB & "'," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ", " & J_REFNO & ")"
                Else
                End If
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "DEPOSITED RECEIPTS JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

            End If
            Grid2.Cell(GridImport, 1).Text = 1
        End If
End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    Screen.MousePointer = 0
    ImportDeposited_HCA = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportDeposited_HCA = False
End Function
'DESCRIPTION: IMPORTING OF DEPOSITED SLIP NUMBER FOR HCA
'SJR 082114

Function ImportDeposited() As Boolean
'HEADER
    On Error GoTo ErrorCode

    Dim J_JDATE                                             As String
    Dim J_VOUCHERNO                                         As String
    Dim J_JTYPE                                             As String
    Dim J_JNO                                               As String
    Dim J_REMARKS                                           As String
    Dim J_VENDORCODE                                        As String
    Dim J_CUSTOMERCODE                                      As String
    Dim J_OUTBALANCE                                        As Double
    Dim J_AMOUNTTOPAY                                       As Double
    Dim J_INVOICEAMT                                        As Double
    Dim J_BALANCE                                           As Double
    Dim J_AMOUNTPAID                                        As Double
    Dim J_CHECKNO                                           As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_CHECKDATE, J_BANKCODE                             As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS, J_DEALER                                   As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String

    'DETAIL
    Dim J_ACCT_CODE, J_ACCT_NAME                            As String
    Dim J_DEBIT                                             As Double
    Dim J_CREDIT                                            As Double
    Dim J_TAX                                               As Double
    Dim J_GROSS                                             As Double
    Dim J_NET                                               As Double
    Dim J_STATUS                                            As String
    Dim J_JITEMNO                                           As String

    Dim rsJournal_HDDup                                     As ADODB.Recordset

    Dim CMIS_OR_NUM                                         As String
    Dim CMIS_OR_DATE                                        As String
    Dim CMIS_OR_AMT                                         As String
    Dim CMIS_DISCOUNT                                       As String
    Dim CMIS_TAX                                            As String
    Dim CMIS_CASHAMOUNT                                     As Double
    Dim CMIS_CHKAMOUNT                                      As Double
    Dim CMIS_CARDAMOUNT                                     As Double
    Dim CMIS_CUSCDE                                         As String
    Dim CMIS_CUSNAME                                        As String
    Dim CMIS_DEPOSIT                                        As String
    Dim CMIS_BANKCODE                                       As String
    Dim CMIS_TSEKE                                          As String
    Dim CMIS_CHECKDATE                                      As String
    Dim CMIS_STATUS                                         As String
    Dim CMIS_TYPE_PAYMENT                                   As String

    Dim CMIS_IS_VAT                                         As Boolean
    Dim CMIS_BANK_DEPOSITED                                 As String

    Dim TOTAL_DEBIT, TOTAL_CREDIT                           As Double

    Dim rsOFF_HD                                            As ADODB.Recordset
    Dim rsOFF_DT                                            As ADODB.Recordset
    Dim i                                                   As Long

    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
        COA_CASH_ON_HAND = Null2String(ReturnAccountCode("CASH ON HAND"))
        COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = Null2String(ReturnAccountCode("CARD"))
    Else
        COA_CASH_ON_HAND = Null2String(ReturnAccountCode("CASH ON HAND"))
        COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = Null2String(ReturnAccountCode("CARD ON HAND"))
    End If
    Dim GridImport                                          As Integer
    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If COMPANY_CODE = M_COMPANY_CODE Then
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 or DEPOSIT2 = 1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND OR_AMT = '" & NumericVal(Grid2.Cell(GridImport, 4).Text) & "' AND VAT = 1 Order by OR_NUM ASC")
                Else
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 or DEPOSIT2 = 1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 0 Order by OR_NUM ASC")
                End If
            Else
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT >= '" & GetCUTOFF_DATE & "' AND DATDEPOSIT <= '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 1 Order by OR_NUM ASC")
                Else
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT >= '" & GetCUTOFF_DATE & "' AND DATDEPOSIT <= '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 0 Order by OR_NUM ASC")
                End If
            End If
            
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!DATDEPOSIT)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                If COMPANY_CODE = M_COMPANY_CODE Then
                    CMIS_DEPOSIT1 = Null2String(rsOFF_HD!DEPOSIT1)
                    CMIS_DEPOSIT2 = Null2String(rsOFF_HD!DEPOSIT2)
                Else
                    CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                End If
                CMIS_BANKCODE = Null2String(rsOFF_HD!DEPOSIT_TO)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)

                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                    If CMIS_BANKCODE = "AUB" Then
                        CMIS_BANK_DEPOSITED = "'11-01007-00'"
                    Else
                        CMIS_BANK_DEPOSITED = Null2String(rsOFF_HD!BANKACCOUNTNO)
                    End If
                Else
                    CMIS_BANK_DEPOSITED = Null2String(rsOFF_HD!BANKACCOUNTNO)
                End If

                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0



                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                
                If COMPANY_CODE = "DJM" Then
                    J_JTYPE = "'GJ'"
                Else
                    J_JTYPE = "'DRJ'"
                End If

                'INSERTED SEPTEMBER 8, 2007
                Set rsOFF_DT = New ADODB.Recordset
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        If COMPANY_CODE = "HSM" Then
                            Dim rsREMARKS As New ADODB.Recordset
                            Set rsREMARKS = New ADODB.Recordset
                            Set rsREMARKS = gconDMIS.Execute("select remarks from amis_journal_hd where jtype='CRJ' and right(invoiceno,8)='" & CMIS_OR_NUM & "'")
                              If Null2String(rsOFF_DT!TranType) = "OTH" Then
                                  J_REMARKS = rsREMARKS(remarks!)
                              Else
                                  J_REMARKS = rsREMARKS(remarks!)
                              End If
                              rsOFF_DT.MoveNext
                              If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                        Else
                            If Null2String(rsOFF_DT!TranType) = "OTH" Then
                                J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                            Else
                                J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                            End If
                            rsOFF_DT.MoveNext
                            If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                        End If
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = NumericVal(CMIS_OR_AMT)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 10))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 10))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        If CMIS_TYPE_PAYMENT = "1" Then
                            J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    Else
                        If CMIS_CASHAMOUNT > 0 Then
                            J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", " DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_CASH_ON_HAND)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_CASH_ON_HAND))
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        If CMIS_TYPE_PAYMENT = "1" Then
                            J_CREDIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_CREDIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    Else
                        If CMIS_CASHAMOUNT > 0 Then
                            J_CREDIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_CREDIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                End If
                If CMIS_TYPE_PAYMENT = "3" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD))
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                End If

                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "DEPOSITED RECEIPTS JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

            End If
            Grid2.Cell(GridImport, 1).Text = 1
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    '=========================================================================================================

    Screen.MousePointer = 0
    '=========================================================================================================

    ImportDeposited = True
    Exit Function
    
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportDeposited = False
End Function

Sub ShowUnImportedPaidInvoices(VarTranType As String, VarTranno As String)
    Screen.MousePointer = 11
    Dim INVOICETYPE, InvoiceTypeCode                        As String
    Dim IS_Exist                                            As Byte
    LIM = 0

    If VarTranType = "PI" Or VarTranType = "MI" Or VarTranType = "AI" Then
        Set rsPMIOS_ORD_HD = New ADODB.Recordset
        Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where TYPE = '" & Left(VarTranType, 1) & "' AND (TranType = 'CSH' OR  TranType = 'CHG') and STATUS = 'P' AND tranno = '" & VarTranno & "' AND TRANDATE > '" & GetCUTOFF_DATE & "' AND TRANDATE <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by Tranno ASC")
        If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
            rsPMIOS_ORD_HD.MoveFirst:
            Do While Not rsPMIOS_ORD_HD.EOF
                LIM = LIM + 1
                If Null2String(rsPMIOS_ORD_HD!Type) = "P" Then
                    INVOICETYPE = "Parts"
                    InvoiceTypeCode = "PI"
                ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "A" Then
                    INVOICETYPE = "Accessories"
                    InvoiceTypeCode = "AI"
                ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "M" Then
                    INVOICETYPE = "Materials"
                    InvoiceTypeCode = "MI"
                Else
                    INVOICETYPE = "Unknown"
                    InvoiceTypeCode = ""
                End If
                If CheckSJExisting(InvoiceTypeCode, Null2String(rsPMIOS_ORD_HD!TRANNO), Null2String(rsPMIOS_ORD_HD!TranType)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & UCase(INVOICETYPE) & Chr(9) & Null2String(rsPMIOS_ORD_HD!TranType) & "-" & Null2String(rsPMIOS_ORD_HD!TRANNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPMIOS_ORD_HD!NETINVAMT)) & Chr(9) & Null2String(rsPMIOS_ORD_HD!CUSTNAME) & Chr(9) & UCase(INVOICETYPE)
                rsPMIOS_ORD_HD.MoveNext
                DoEvents
            Loop
        End If
    End If
    If VarTranType = "SI" Then
        'PURELY INTERNAL
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = '" & VarTranno & "' AND DTE_REL > '" & GetCUTOFF_DATE & "' AND DTE_REL <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
        Else
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = '" & VarTranno & "' AND DTE_COMP > '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
        End If
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckRefNoExisting("SI", Null2String(rsCSMIOS_REPOR!REP_OR)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice = '" & VarTranno & "' AND DTE_REL >= '" & GetCUTOFF_DATE & "' AND DTE_REL <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        Else
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice = '" & VarTranno & "' AND DTE_COMP >= '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        End If
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!INVOICE)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If

        'PURELY WARRANTY
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HCE" Then
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice = '" & VarTranno & "' AND DTE_REL > '" & GetCUTOFF_DATE & "' AND DTE_REL <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        Else
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice = '" & VarTranno & "' AND DTE_COMP > '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        End If
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!INVOICE)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If

    End If
    If VarTranType = "VI" Then
        Set rsSMIS_PURCHAGREE = New ADODB.Recordset
        Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & VarTranno & "' AND DATERELEASED > '" & GetCUTOFF_DATE & "' AND CONVERT(VARCHAR,DATERELEASED,101) <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
        If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
            rsSMIS_PURCHAGREE.MoveFirst:
            Do While Not rsSMIS_PURCHAGREE.EOF
                LIM = LIM + 1
                If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsSMIS_PURCHAGREE!IGNKEY_NO) & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!Code)) & Chr(9) & "VEHICLE"
                rsSMIS_PURCHAGREE.MoveNext
                DoEvents
            Loop
        End If
    End If
    Screen.MousePointer = 0
End Sub

Sub InitGrid3()
    With Grid3
        .Rows = 1
        .Cols = 7
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Inv. Type"
        .Cell(0, 3).Text = "Inv. No."
        .Cell(0, 4).Text = "Inv. Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 80
        .Column(4).Width = 80
        .Column(5).Width = 200
        .Column(6).Width = 50

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With
End Sub

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "OR Type"
        .Cell(0, 3).Text = "OR No."
        .Cell(0, 4).Text = "OR Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

    With Grid2
        If COMPANY_CODE = "HCA" Then
            .Rows = 1
            .Cell(0, 1).Text = "Imported"
            .Cell(0, 2).Text = "DepSlip No."
            .Cell(0, 3).Text = "Deposited"
            .Cell(0, 4).Text = "Customer"
    
            .Column(0).Width = 10
            .Column(1).Width = 50
            .Column(3).Width = 60
            .Column(4).Width = 100
    
            .Column(1).CellType = cellCheckBox
            .Column(4).Alignment = cellRightGeneral
    
            .Column(1).Locked = True
            .Column(2).Locked = True
            .Column(3).Locked = True
            .Column(4).Locked = True
            
        ElseIf COMPANY_CODE = "DJM" Then
            .Rows = 1
            .Cell(0, 1).Text = "Imported"
            .Cell(0, 2).Text = "Ref. No."
            .Cell(0, 3).Text = "Deposited"
            .Cell(0, 4).Text = ""
    
            .Column(0).Width = 10
            .Column(1).Width = 50
            .Column(3).Width = 100
            .Column(4).Width = 100
    
            .Column(1).CellType = cellCheckBox
            .Column(4).Alignment = cellRightGeneral
    
            .Column(1).Locked = True
            .Column(2).Locked = True
            .Column(3).Locked = True
            .Column(4).Locked = True
            
        Else
            .Rows = 1
            .Cell(0, 1).Text = "Imported"
            .Cell(0, 2).Text = "OR Type"
            .Cell(0, 3).Text = "OR No."
            .Cell(0, 4).Text = "OR Amt."
            .Cell(0, 5).Text = "Customer"
    
            .Column(0).Width = 10
            .Column(1).Width = 50
            .Column(2).Width = 80
            .Column(3).Width = 60
            .Column(4).Width = 80
            .Column(5).Width = 200
    
            .Column(1).CellType = cellCheckBox
            .Column(4).Alignment = cellRightGeneral
    
            .Column(1).Locked = True
            .Column(2).Locked = True
            .Column(3).Locked = True
            .Column(4).Locked = True
            .Column(5).Locked = True
        End If
    End With
    InitGrid3
End Sub

Sub InsertToJournalDet(vJ_JDATE As Variant, vJ_VOUCHERNO As Variant, vJ_JTYPE As Variant, vJ_JNO As Variant, vJ_JITEMNO As Variant, vJ_ACCT_CODE As Variant, vJ_ACCT_NAME As Variant, vJ_DEBIT As Variant, vJ_CREDIT As Variant, vJ_TAX As Variant, vJ_GROSS As Variant, vJ_NET As Variant, vJ_STATUS As Variant)
    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                     " values (" & vJ_JDATE & ", " & vJ_VOUCHERNO & ", " & vJ_JTYPE & ", " & vJ_JNO & ", " & vJ_JITEMNO & ", " & vJ_ACCT_CODE & ", " & vJ_ACCT_NAME & ", " & vJ_DEBIT & ", " & vJ_CREDIT & ", " & vJ_TAX & "," & vJ_GROSS & "," & vJ_NET & ", " & vJ_STATUS & ")"

    TransactionID = (FindTransactionID(N2Str2Null(vJ_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(vJ_VOUCHERNO), J_JTYPE, N2Str2Zero(vJ_JNO)
End Sub

Private Sub cboMonth_Click()
    Dim iCount, xCount                                      As Integer
    Dim Indx                                                As Integer
    iCount = 0
    Select Case cboMonth.Text
    Case "January": Indx = 31
    Case "February"
        If NumericVal(cboYear.Text) Mod 4 = 0 Then
            Indx = 29
        Else
            Indx = 28
        End If
    Case "March": Indx = 31
    Case "April": Indx = 30
    Case "May": Indx = 31
    Case "June": Indx = 30
    Case "July": Indx = 31
    Case "August": Indx = 31
    Case "September": Indx = 30
    Case "October": Indx = 31
    Case "November": Indx = 30
    Case "December": Indx = 31
    Case Else: Indx = -1
    End Select
    For xCount = 1 To 30
        lab.Item(xCount).Visible = False
    Next
    Do While iCount <= Indx - 1
        lab.Item(iCount).BackColor = &HE0E0E0
        lab.Item(iCount).Visible = True
        iCount = iCount + 1
    Loop
End Sub

Private Sub cboYear_Click()
    Dim iCount, xCount                                      As Integer
    Dim Indx                                                As Integer
    iCount = 0
    Select Case cboMonth.Text
    Case "January": Indx = 31
    Case "February"
        If NumericVal(cboYear.Text) Mod 4 = 0 Then
            Indx = 29
        Else
            Indx = 28
        End If
    Case "March": Indx = 31
    Case "April": Indx = 30
    Case "May": Indx = 31
    Case "June": Indx = 30
    Case "July": Indx = 31
    Case "August": Indx = 31
    Case "September": Indx = 30
    Case "October": Indx = 31
    Case "November": Indx = 30
    Case "December": Indx = 31
    Case Else: Indx = -1
    End Select
    For xCount = 1 To 30
        lab.Item(xCount).Visible = False
    Next
    Do While iCount <= Indx - 1
        lab.Item(iCount).BackColor = &HE0E0E0
        lab.Item(iCount).Visible = True
        iCount = iCount + 1
    Loop
End Sub

Private Sub cmdBatchImport_Click()
    BATCHIMPORT = True
    picBatchImport.Visible = True
    picBatchImport.ZOrder 0
    dtFrom.Value = firstDay(Now())
    dtTo.Value = Format(CRJLASTTRANS, "mm/dd/yyyy")
    cmdBatchImport.Enabled = False
    cmdBatchImporting.Enabled = True
    cmdExit.Enabled = False
    cmdClearJournals.Enabled = False
    cmdCheck.Enabled = False
    ShortcutCaption1.VisualTheme = xtpShortcutThemeOffice2007
End Sub

Private Sub cmdBatchImporting_Click()
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
    cmdBatchImporting.Enabled = False
    Dim i                                                   As Integer
    For i = 0 To (dtTo.Value - dtFrom.Value)
        dtpTranDate.Value = dtFrom.Value + i
        Call cmdShowTrans_Click
        Call cmdCheck_Click
    Next i

    If BATCHIMPORT = True Then
        MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    End If
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "IMPORT CASH RECEIPTS") = False Then Exit Sub
    Screen.MousePointer = 11
    Dim str_MSG                                             As String

    str_MSG = "Error in saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf & vbCrLf & vbCrLf
    str_MSG = str_MSG & "Event Name: " & UCase(Screen.ActiveControl.Name) & vbCrLf
    str_MSG = str_MSG & "Form Name: " & UCase(Screen.ActiveForm.Name) & vbCrLf
    str_MSG = str_MSG & "ERRORSOURCE" & vbCrLf

    If dtpTranDate.Value <= CDate(GetCUTOFF_DATE) Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Date selected is less than Cut-Off Date."
        cmdExit.Enabled = True
        Screen.MousePointer = 0
        Call cmdShowImp_Click
    Else
        If Option1.Value = True Then
            If COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "HGS" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HNE" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
                gconDMIS.BeginTrans
                If ImportUnDeposit = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Undeposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                gconDMIS.CommitTrans
                Call cmdShowImp_Click
                If BATCHIMPORT = False Then
                    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
                End If
                LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
                Exit Sub
            Else
                gconDMIS.BeginTrans
                If ImportPMISSales = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "PMIS Sales")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                If ImportCSMSSales = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "CSMS Sales")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                If ImportSMISSales = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "SMIS Sales")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If

            If COMPANY_CODE = M_COMPANY_CODE Then
                If ImportUnDepositM = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Undeposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                If ImportUnDeposit = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Undeposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            gconDMIS.CommitTrans
            Call cmdShowImp_Click
            MsgBox "Import Successfully Completed!", vbInformation, "Finish"
            LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
            Exit Sub
        End If
        
        If Option2.Value = True Then
            gconDMIS.BeginTrans
            If COMPANY_CODE = "HCA" Then
                If ImportDeposited_HCA = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Deposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            ElseIf COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
                If ImportDeposited_DJM = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Deposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                If ImportDeposited = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Deposited Cash Receipts")
                    str_MSG = Replace(str_MSG, "ERRORSOURCE", Err_handler)
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            gconDMIS.CommitTrans
            Call cmdShowImp_Click
            If BATCHIMPORT = False Then
                MsgBox "Import Successfully Completed!", vbInformation, "Finish"
            End If
            LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
            Exit Sub
        End If
    End If
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdClearJournals_Click()
    If Option1.Value = True Then
    
        If MsgBox("Clear Unposted Data?", vbQuestion + vbYesNo, "Confirm...") = vbNo Then
            Exit Sub
        Else
            Dim i                                           As Integer
            Dim xInvoiceType                                As String
            Dim xInvoiceNo                                  As String
            Dim rsJournalHD                                 As ADODB.Recordset
            Dim rsCRJ_Detail                                As ADODB.Recordset

            i = 0
            For i = 1 To Grid1.Rows - 1
                If Grid1.Cell(i, 2).Text = "VAT" Then
                    xInvoiceNo = Grid1.Cell(i, 3).Text
                   
                Else
                    xInvoiceNo = ("NV" & Grid1.Cell(i, 3).Text)
                
                End If
                 xInvoiceType = Grid1.Cell(i, 2).Text
                Set rsJournalHD = New ADODB.Recordset
                rsJournalHD.Open ("SELECT VOUCHERNO,JTYPE FROM AMIS_JOURNAL_HD WHERE STATUS='N' AND INVOICENO ='" & xInvoiceNo & "' AND JTYPE='CRJ'"), gconDMIS, adOpenForwardOnly
                If Not rsJournalHD.EOF And Not rsJournalHD.BOF Then
                    Set rsCRJ_Detail = New ADODB.Recordset
                    rsCRJ_Detail.Open "SELECT ID FROM AMIS_CRJ_DETAIL WHERE STATUS='N' AND VOUCHERNO='" & rsJournalHD!VOUCHERNO & "'", gconDMIS, adOpenForwardOnly
                    If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                        Do While Not rsCRJ_Detail.EOF
                            gconDMIS.Execute ("Delete from AMIS_CRJ_Detail Where STATUS = 'N' AND ID=" & rsCRJ_Detail!ID & "")
                            rsCRJ_Detail.MoveNext
                        Loop
                    End If
                    gconDMIS.Execute ("Delete from AMIS_Journal_Det Where STATUS = 'N' AND Jtype = 'CRJ' AND VOUCHERNO='" & rsJournalHD!VOUCHERNO & "'")
                    gconDMIS.Execute ("Delete from AMIS_Journal_HD Where STATUS = 'N' AND Jtype = 'CRJ' and VOUCHERNO='" & rsJournalHD!VOUCHERNO & "'")
                    gconDMIS.Execute ("Delete from AMIS_AP Where STATUS = 'N' AND VOUCHERNO = " & N2Str2Null("CRJ" + "-" + rsJournalHD!VOUCHERNO) & "")
                    gconDMIS.Execute ("Delete from AMIS_details Where STATUS = 'N' AND Jtype = 'CRJ' AND VOUCHERNO='" & rsJournalHD!VOUCHERNO & "'")
                    gconDMIS.Execute ("Delete from AMIS_ar Where STATUS = 'N'  AND SJVOUCHERNO=" & N2Str2Null("CRJ" + "-" + rsJournalHD!VOUCHERNO) & "")
                    gconDMIS.Execute ("Delete from AMIS_detail Where STATUS = 'N' AND Jtype = 'CRJ' AND VOUCHERNO='" & rsJournalHD!VOUCHERNO & "'")
                    
                End If
            Next i
            Set rsJournalHD = Nothing
            cmdShowTrans.Value = True
            Screen.MousePointer = 0
            MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
        End If
    End If
    
    If Option2.Value = True Then
        Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
        
        If COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
            Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "'")
        Else
            Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "'")
        End If
        
        If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
            Screen.MousePointer = 0
            If LOGLEVEL = "ADM" Then
                If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                    Screen.MousePointer = 11
                    
                    If COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
                        gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                        gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'GJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    Else
                        gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                        gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    End If
                    cmdShowTrans.Value = True
                    Screen.MousePointer = 0
                    MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                End If
            End If
            Exit Sub
        End If
    End If
    Call cmdShowImp_Click
    Call cmdShowTrans_Click
End Sub

Private Sub cmdCloseRange_Click()
    picBatchImport.Visible = False
    picBatchImport.ZOrder 1
    cmdBatchImport.Enabled = True
    cmdExit.Enabled = True
    BATCHIMPORT = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShowImp_Click()
    Screen.MousePointer = 11
    If cmdCheck.Value = True Then
    Else
        cmdCheck.Value = False
InitGrids:         DoEvents:
        Grid1.Rows = 1
        Grid2.Rows = 1
        Grid3.Rows = 1
    End If
    Dim tmpDate                                             As String
    Dim iCount                                              As Integer
    MonthIndex
    Dim rsCheckCRJData                                      As ADODB.Recordset
    Dim rsCheckCRJData2                                     As ADODB.Recordset

    Set rsCheckCRJData = New ADODB.Recordset

    gconDMIS.CommandTimeout = 900
    rsCheckCRJData.Open "SELECT DISTINCT DATE FROM AMIS_VW_IMPORTED_OR WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckCRJData.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckCRJData!Date)), "d") Then
                lab.Item(iCount).BackColor = &HC0FFC0
                DoEvents
            End If
            iCount = iCount + 1
        Loop

        rsCheckCRJData.MoveNext
        DoEvents
    Loop
    
    gconDMIS.CommandTimeout = 900
    Set rsCheckCRJData2 = New ADODB.Recordset
    rsCheckCRJData2.Open "SELECT DISTINCT DATE FROM AMIS_VW_UNIMPORTED_OR WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckCRJData2.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckCRJData2!Date)), "d") Then
                lab.Item(iCount).BackColor = &HFFFF&
                DoEvents
            End If
            iCount = iCount + 1
        Loop

        rsCheckCRJData2.MoveNext
        DoEvents
    Loop

    Screen.MousePointer = 0
End Sub

Private Sub cmdShowTrans_Click()
    Screen.MousePointer = 11
InitGrids:     DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
    Grid3.AutoRedraw = False
    Grid1.Rows = 1: Grid2.Rows = 1: Grid3.Rows = 1: KIM = 0: LIM = 0
    Dim ORType                                              As String
    Dim IS_Exist                                            As Byte
    Dim rsOR_UNDEPOSITED                                    As ADODB.Recordset
    Dim rsOR_DEPOSITED                                      As ADODB.Recordset
    Dim rsUNDEPOSITED_INVOICES                              As ADODB.Recordset

    If COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "DAI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Or COMPANY_CODE = "HGS" Then
        GoTo ShowTrans
    Else
        If CheckImportedOR(dtpTranDate) = True Then
            MsgBox "Previous transactions dated " & TRANSACTIONDATE & " are not yet imported.", vbExclamation, "Message"
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo ShowTransactions
        End If

ShowTransactions:
    End If
ShowTrans:

    Set rsOR_UNDEPOSITED = New ADODB.Recordset
    Set rsOR_UNDEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD where (PAIDNA = 1 OR STATUS = 'P') AND OR_DATE = '" & CDate(dtpTranDate) & "' and cancel = 0 and left(or_num,3) <> 'SOA' order by OR_NUM ASC")
    If Not rsOR_UNDEPOSITED.EOF And Not rsOR_UNDEPOSITED.BOF Then
        rsOR_UNDEPOSITED.MoveFirst: KIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsOR_UNDEPOSITED.EOF
            KIM = KIM + 1
           If CheckCRJExisting(Null2String(rsOR_UNDEPOSITED!OR_NUM), N2Str2Zero(rsOR_UNDEPOSITED!VAT)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            
            If N2Str2Zero(rsOR_UNDEPOSITED!VAT) = 1 Then
                ORType = "VAT"
            Else
                ORType = "NON VAT"
            End If
            
            Grid1.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_UNDEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_UNDEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_UNDEPOSITED!CUSNAME)
            Set rsUNDEPOSITED_INVOICES = New ADODB.Recordset
            Set rsUNDEPOSITED_INVOICES = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE (TRANTYPE = 'VI' OR TRANTYPE = 'SI' OR TRANTYPE = 'PI' OR TRANTYPE = 'AI' OR TRANTYPE = 'MI') AND OR_NUM = " & N2Str2Null(rsOR_UNDEPOSITED!OR_NUM) & " AND VAT = " & N2Str2Zero(rsOR_UNDEPOSITED!VAT))
            If Not rsUNDEPOSITED_INVOICES.EOF And Not rsUNDEPOSITED_INVOICES.BOF Then
                ShowUnImportedPaidInvoices Null2String(rsUNDEPOSITED_INVOICES!TranType), Null2String(rsUNDEPOSITED_INVOICES!INVOICENO)
            End If
            rsOR_UNDEPOSITED.MoveNext
            DoEvents
        Loop
        If KIM = 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If

    Set rsOR_DEPOSITED = New ADODB.Recordset
    If COMPANY_CODE = M_COMPANY_CODE Then
        Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 OR DEPOSIT2 =1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' and Cancel = 0 Order by OR_NUM ASC")
    
    ElseIf COMPANY_CODE = "HCA" Then
        Set rsOR_DEPOSITED = gconDMIS.Execute("Select DISTINCT DEPSLIPNUM, " & _
                                            "SUM(DEPOSIT) AS SUMOFDEPSLIP,DEPOSIT_TO " & _
                                            "from CMIS_BankDepo Where DEPOSIT > 0 " & _
                                            "AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' " & _
                                            "GROUP by DEPSLIPNUM,DEPOSIT_TO")
    ElseIf COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
        'PER BANK
        'Set rsOR_DEPOSITED = gconDMIS.Execute("SELECT DISTINCT DEPOSIT_TO, BANKNAME,DATDEPOSIT, SUM(OR_AMT) AS SUMPERBANK FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND CANCEL = 0 GROUP BY DEPOSIT_TO, BANKNAME, DATDEPOSIT")
        'PER DAY
        Set rsOR_DEPOSITED = gconDMIS.Execute("SELECT DISTINCT DATDEPOSIT,SUM(DEPOSIT) AS SUMPERBANK FROM CMIS_BANKDEPO WHERE DATDEPOSIT = '" & CDate(dtpTranDate) & "' GROUP BY DATDEPOSIT")
    Else
        Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' and Cancel = 0 Order by OR_NUM ASC")
    End If
    
    If COMPANY_CODE = "HCA" Then
        If Not rsOR_DEPOSITED.EOF And Not rsOR_DEPOSITED.BOF Then
            rsOR_DEPOSITED.MoveFirst: KIM = 0
            Grid2.AutoRedraw = False
            Do While Not rsOR_DEPOSITED.EOF
                KIM = KIM + 1
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If CheckDRJExistingM(Null2String(rsOR_DEPOSITED!SUMOFDEPSLIP), N2Str2Zero(rsOR_DEPOSITED!DEPSLIPNUM), NumericVal(rsOR_DEPOSITED!SUMOFDEPSLIP)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                Else
                    If CheckDRJExisting2(Null2String(rsOR_DEPOSITED!DEPSLIPNUM), Null2String(rsOR_DEPOSITED!DEPOSIT_TO)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                End If

                Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsOR_DEPOSITED!DEPSLIPNUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_DEPOSITED!SUMOFDEPSLIP)) & Chr(9) & Null2String(rsOR_DEPOSITED!DEPOSIT_TO)
                rsOR_DEPOSITED.MoveNext
                DoEvents
            Loop
            If KIM = 0 Then Grid2.RemoveItem 1
            Grid2.AutoRedraw = True
            Grid2.Refresh
        End If
        
    ElseIf COMPANY_CODE = "CMC" Or COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
        If Not rsOR_DEPOSITED.EOF And Not rsOR_DEPOSITED.BOF Then
            rsOR_DEPOSITED.MoveFirst: KIM = 0
            Grid2.AutoRedraw = False
            Do While Not rsOR_DEPOSITED.EOF
                KIM = KIM + 1
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If CheckDRJExistingM(Null2String(rsOR_DEPOSITED!SUMOFDEPSLIP), N2Str2Zero(rsOR_DEPOSITED!DEPSLIPNUM), NumericVal(rsOR_DEPOSITED!SUMOFDEPSLIP)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                Else
                    If CheckDRJExisting3("DEP-" & Format(Null2String(rsOR_DEPOSITED!DATDEPOSIT), "MMDDYY"), Null2String(rsOR_DEPOSITED!DATDEPOSIT)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                End If

                Grid2.AddItem IS_Exist & Chr(9) & "DEP-" & Format(Null2String(rsOR_DEPOSITED!DATDEPOSIT), "MMDDYY") & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_DEPOSITED!SUMPERBANK))
                rsOR_DEPOSITED.MoveNext
                DoEvents
            Loop
            If KIM = 0 Then Grid2.RemoveItem 1
            Grid2.AutoRedraw = True
            Grid2.Refresh
        End If
        
    Else
    
        If Not rsOR_DEPOSITED.EOF And Not rsOR_DEPOSITED.BOF Then
            rsOR_DEPOSITED.MoveFirst: KIM = 0
            Grid2.AutoRedraw = False
            Do While Not rsOR_DEPOSITED.EOF
                KIM = KIM + 1
                
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If CheckDRJExistingM(Null2String(rsOR_DEPOSITED!OR_NUM), N2Str2Zero(rsOR_DEPOSITED!VAT), NumericVal(rsOR_DEPOSITED!OR_AMT)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                Else
                    If CheckDRJExisting(Null2String(rsOR_DEPOSITED!OR_NUM), N2Str2Zero(rsOR_DEPOSITED!VAT)) = True Then
                        IS_Exist = 1
                    Else
                        IS_Exist = 0
                    End If
                End If
                
                If N2Str2Zero(rsOR_DEPOSITED!VAT) = 1 Then
                    ORType = "VAT"
                Else
                    ORType = "NON VAT"
                End If
                
                Grid2.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_DEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_DEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_DEPOSITED!CUSNAME)
                rsOR_DEPOSITED.MoveNext
                DoEvents
            Loop
            
            If KIM = 0 Then Grid2.RemoveItem 1
            Grid2.AutoRedraw = True
            Grid2.Refresh
        End If
    End If

    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
        cmdCheck.SetFocus
    End If
    
    If LIM = 0 Then Grid3.RemoveItem 1
    Grid3.AutoRedraw = True
    Grid3.Refresh
    Screen.MousePointer = 0
End Sub

Private Sub dtpTranDate_Change()
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
End Sub

Private Sub dtTo_LostFocus()
    If dtTo.Value < dtFrom.Value Then
        MsgBox "Please check selected date.", vbInformation, "Check Date"
        dtTo.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpTranDate = LOGDATE
    
    InitGrids
    InitCombo
    InitNoDays
    If COMPANY_CODE = "HBK" Then Option2.Enabled = False

    Screen.MousePointer = 0
    Exit Sub
    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Then
        cmdClearJournals.Visible = False
    End If

ErrorCode:
    Screen.MousePointer = 0
    MsgBox err.Number & vbCrLf & err.DESCRIPTION, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Function GetEntityCode(XXX As String) As String
    Dim rsEntity                                            As ADODB.Recordset
    Set rsEntity = gconDMIS.Execute("select * from All_Entity where Code = '" & XXX & "'")
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        GetEntityCode = rsEntity!ENTITYCODE
    End If
    Set rsEntity = Nothing
End Function

Function CheckImportedOR(xOR_Date As String) As Boolean
'Update: ACL 11162009
    Dim xNoDays                                             As Integer
    Dim rsCheckCRJ                                          As ADODB.Recordset
    Dim rsCHECKOR                                           As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='CRJ' AND STATUS <> 'C' AND JDATE < '" & CDate(xOR_Date) & "' ORDER BY JDATE DESC", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        Set rsCHECKOR = New ADODB.Recordset
        rsCHECKOR.Open "Select * FROM (SELECT TOP 1 OR_DATE AS TRANDATE FROM CMIS_OFF_HD WHERE (PAIDNA = 1 OR STATUS = 'P') AND OR_DATE > '" & Null2Date(rsCheckCRJ!JDATE) & "' AND CANCEL =0  ORDER BY OR_DATE ASC " & _
                       "UNION SELECT TOP 1 DATDEPOSIT AS TRANDATE FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND DATDEPOSIT > '" & Null2Date(rsCheckCRJ!JDATE) & "' AND CANCEL = 0 ORDER BY DATDEPOSIT ASC) X ORDER BY TRANDATE ASC", gconDMIS, adOpenKeyset
        If Not rsCHECKOR.EOF And Not rsCHECKOR.BOF Then
            TRANSACTIONDATE = Null2Date(rsCHECKOR!trandate)
            xNoDays = DateDiff("d", TRANSACTIONDATE, dtpTranDate)
            If xNoDays > 0 Then
                CheckImportedOR = True
            End If
        End If
    End If
    Set rsCheckCRJ = Nothing
End Function

Function CheckToClearOR(xOR_Date As String) As Boolean
    Dim rsCheckCRJ                                          As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate > '" & CDate(xOR_Date) & "' Order by JDate Asc", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        CheckToClearOR = True
    End If
End Function

Function CheckToClearDRJ(xOR_Date As String) As Boolean
    Dim rsCheckCRJ                                          As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate > '" & CDate(xOR_Date) & "' Order by JDate Asc", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        CheckToClearDRJ = True
    End If
End Function

Function CheckDepositedOR(xOR_Date As String) As Boolean
    Dim rsCheckDRJ                                          As ADODB.Recordset
    Set rsCheckDRJ = New ADODB.Recordset
    rsCheckDRJ.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='DRJ' AND JDATE<'" & CDate(xOR_Date) & "'", gconDMIS, adOpenKeyset
    If Not rsCheckDRJ.EOF And Not rsCheckDRJ.BOF Then
        CheckDepositedOR = True
    End If
    Set rsCheckDRJ = Nothing
End Function

Function CheckIfBank(xCUSCDE As String) As Boolean
    Dim rsCheckCode                                         As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCUSCDE) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                                 As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Private Sub lab_Click(Index As Integer)
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
    MonthIndex
    xTranDate = Indx & "/" & lab.Item(Index).Caption & "/" & cboYear.Text
    dtpTranDate.Value = xTranDate
    Call cmdShowTrans_Click
End Sub

Sub MonthIndex()
    Select Case cboMonth.Text
    Case "January": Indx = 1
    Case "February": Indx = 2
    Case "March": Indx = 3
    Case "April": Indx = 4
    Case "May": Indx = 5
    Case "June": Indx = 6
    Case "July": Indx = 7
    Case "August": Indx = 8
    Case "September": Indx = 9
    Case "October": Indx = 10
    Case "November": Indx = 11
    Case "December": Indx = 12
    Case Else: Indx = -1
    End Select
End Sub

Sub InitNoDays()
    Dim iCount                                              As Integer
    For iCount = 1 To 31
        lab.Item(iCount - 1).Caption = iCount
    Next
End Sub

Sub InitCombo()
    Dim NoDays                                              As Integer
    With cboYear
        .AddItem ("2005")
        .AddItem ("2006")
        .AddItem ("2007")
        .AddItem ("2008")
        .AddItem ("2009")
        .AddItem ("2010")
        .AddItem ("2011")
        .AddItem ("2012")
        .AddItem ("2013")
        .AddItem ("2014")
        .AddItem ("2015")
        .AddItem ("2016")
        .AddItem ("2017")
        .AddItem ("2018")
    End With
    cboYear.Text = Format(LOGDATE, "yyyy")

    With cboMonth
        .AddItem ("January")
        .AddItem ("February")
        .AddItem ("March")
        .AddItem ("April")
        .AddItem ("May")
        .AddItem ("June")
        .AddItem ("July")
        .AddItem ("August")
        .AddItem ("September")
        .AddItem ("October")
        .AddItem ("November")
        .AddItem ("December")
    End With
    cboMonth.ListIndex = Month(LOGDATE) - 1
End Sub

Function ImportUnDepositM() As Boolean
'HEADER
    On Error GoTo ErrorCode
    Dim J_JDATE                                             As String
    Dim J_VOUCHERNO                                         As String
    Dim J_JTYPE                                             As String
    Dim J_JNO                                               As String
    Dim J_REMARKS                                           As String
    Dim J_VENDORCODE                                        As String
    Dim J_CUSTOMERCODE                                      As String
    Dim J_CUSTOMERCODE2                                     As String
    Dim J_CHECKNO                                           As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_CHECKDATE, J_BANKCODE                             As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS, J_DEALER                                   As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String
    Dim J_REFERENCENO                                       As String
    Dim J_ENTITY                                            As String
    Dim J_OUTBALANCE                                        As Double
    Dim J_AMOUNTTOPAY                                       As Double
    Dim J_INVOICEAMT                                        As Double
    Dim J_BALANCE                                           As Double
    Dim J_AMOUNTPAID                                        As Double

    'DETAIL
    Dim J_ACCT_CODE                                         As String
    Dim J_ACCT_NAME                                         As String
    Dim J_STATUS                                            As String
    Dim J_JITEMNO                                           As String
    Dim J_JITEMNO_2                                         As String
    Dim J_ALLENTITY                                         As String
    Dim J_DEBIT                                             As Double
    Dim J_CREDIT                                            As Double
    Dim J_TAX                                               As Double
    Dim J_GROSS                                             As Double
    Dim J_NET                                               As Double

    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Dim CMIS_OR_NUM                                         As String
    Dim CMIS_OR_DATE                                        As String
    Dim CMIS_OR_AMT                                         As String
    Dim CMIS_DISCOUNT                                       As String
    Dim CMIS_TAX                                            As String
    Dim CMIS_CUSCDE                                         As String
    Dim CMIS_CUSNAME                                        As String
    Dim CMIS_DEPOSIT                                        As String
    Dim CMIS_BANKCODE                                       As String
    Dim CMIS_BANK                                           As String
    Dim CMIS_TSEKE                                          As String
    Dim CMIS_CHECKDATE                                      As String
    Dim CMIS_STATUS                                         As String
    Dim CMIS_TYPE_PAYMENT1                                  As String
    Dim CMIS_TYPE_PAYMENT2                                  As String
    Dim CMIS_TYPE_PAYMENT3                                  As String
    Dim CMIS_DT_TRANTYPE                                    As String
    Dim CMIS_DT_REFERENCE                                   As String
    Dim CMIS_DT_CUSCDE                                      As String
    Dim CMIS_DT_DESCRIPT                                    As String
    Dim CMIS_DT_REFERENCENO                                 As String
    Dim CMIS_DT_DOCDTE                                      As String
    Dim CMIS_DT_PAIDFOR                                     As String
    Dim CMIS_CASHAMOUNT                                     As Double
    Dim CMIS_CHKAMOUNT                                      As Double
    Dim CMIS_CARDAMOUNT                                     As Double
    Dim TOTAL_DEBIT                                         As Double
    Dim TOTAL_CREDIT                                        As Double
    Dim CMIS_DT_AMOUNT                                      As Double
    Dim CMIS_DT_PAYMENT                                     As Double
    Dim CMIS_DT_DISCOUNT                                    As Double
    Dim CMIS_DT_TAX                                         As Double
    Dim CMIS_IS_VAT                                         As Boolean
    Dim i                                                   As Long

    Dim rsOFF_HD                                            As ADODB.Recordset
    Dim rsOFF_DT                                            As ADODB.Recordset
    Dim rsSJ_DATA                                           As ADODB.Recordset
    Dim rsCheckJournal_HD                                   As ADODB.Recordset

    Dim PV_MRRNO                                            As String
    Dim PV_INVNO                                            As String
    Dim PV_PRODNO                                           As String
    Dim J_JVOUCHERNO                                        As String
    Dim PV_STATUS, PV_ITEMNO                                As String
    Dim PV_AMOUNT                                           As Double
    Dim SJ_PV_ITEMNO                                        As Integer
    Dim GridImport                                          As Integer

    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 1 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 0 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT1)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BANKCODE)
                CMIS_BANK = Null2String(rsOFF_HD!Bank)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT1 = Null2String(rsOFF_HD!TOF1)
                CMIS_TYPE_PAYMENT2 = Null2String(rsOFF_HD!TOF2)
                CMIS_TYPE_PAYMENT3 = Null2String(rsOFF_HD!TOF3)
                CMIS_BANKCODE = Null2String(rsOFF_HD!BANKCODE)
                CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                CMIS_ALLENTITY = CMIS_ENTITYCODE + CMIS_CUSCDE
                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetCRJVoucherNo())
                J_JTYPE = "'CRJ'"

                'INSERTED SEPTEMBER 8, 2007
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        J_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_INVOICENUM = Null2String(rsOFF_DT!INVOICENO)

                        If Null2String(rsOFF_DT!TranType) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)
                J_DEPOSIT = 0
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CMIS_OR_AMT), 2)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 10))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 10))
                End If
                
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF1) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF2) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF3) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_ENTITY = N2Str2Null(CMIS_ENTITYCODE)
                J_ALLENTITY = N2Str2Null(CMIS_ALLENTITY)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                'DESCRIPTION: CASH
                If CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = NumericVal(CMIS_CARDAMOUNT)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                End If

                J_JITEMNO = 0
                If CMIS_TYPE_PAYMENT1 = "1" Or CMIS_TYPE_PAYMENT2 = "2" Then
                    J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                End If
                If CMIS_TYPE_PAYMENT3 = "3" Then
                    J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                    End If
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                End If

                'DESCRIPTION: Checking for Customer Deposits and ADD upon full payment of the invoice amount
                If J_INVOICENUM <> "" Then
                    Dim rsCheckDeposit                      As ADODB.Recordset
                    Set rsCheckDeposit = New ADODB.Recordset
                    rsCheckDeposit.Open "Select * from CMIS_Deposits where InvoiceNo = '" & J_INVOICENUM & "'", gconDMIS, adOpenForwardOnly
                    If Not rsCheckDeposit.EOF And Not rsCheckDeposit.BOF Then
                        J_JITEMNO_2 = 1
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR))))
                        J_DEBIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                        J_DEPOSIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    Set rsCheckDeposit = Nothing
                End If

                'DESCRIPTION: Bank Charges for credit card payment receive
                If CheckIfBank(CMIS_CUSCDE) = True Then
                    If J_REFERENCENO <> "" Then
                        J_JITEMNO_2 = 1
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null("71-35000-20")
                        J_ACCT_NAME = N2Str2Null(Setacctname("71-35000-20"))
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            If CMIS_CASHAMOUNT > 0 Then
                                J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                            Else
                                J_DEBIT = Round((NumericVal(CMIS_CHKAMOUNT) / 0.97) * 0.025, 2)
                            End If
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        'DESCRIPTION: Creditable Withholding Tax for credit card payment receive
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null("11-07000-00")
                        J_ACCT_NAME = N2Str2Null(Setacctname("11-07000-00"))
                        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            If CMIS_CASHAMOUNT > 0 Then
                                J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                            Else
                                J_DEBIT = Round((NumericVal(CMIS_CHKAMOUNT) / 0.97) * 0.005, 2)
                            End If
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                End If
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                'DESCRIPTION: CREDIT
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 1 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 0 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst: SJ_PV_ITEMNO = 0
                    Do While Not rsOFF_DT.EOF
                        CMIS_DT_TRANTYPE = Null2String(rsOFF_DT!TranType)
                        CMIS_DT_REFERENCE = Null2String(rsOFF_DT!INVOICENO)
                        CMIS_DT_CUSCDE = Null2String(rsOFF_DT!CUSCDE)
                        CMIS_DT_DESCRIPT = Null2String(rsOFF_DT!DESCRIPT)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!amount)
                        CMIS_DT_DOCDTE = Null2String(rsOFF_DT!DOCDTE)
                        CMIS_DT_PAYMENT = N2Str2Zero(rsOFF_DT!payment)
                        CMIS_DT_DISCOUNT = N2Str2Zero(rsOFF_DT!DISCOUNT)
                        CMIS_DT_TAX = N2Str2Zero(rsOFF_DT!tax)
                        CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)

                        'DESCRIPTION: For credit card detail
                        CMIS_DT_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_JVOUCHERNO = J_VOUCHERNO
                        SJ_PV_ITEMNO = SJ_PV_ITEMNO + 1
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                        PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                        PV_AMOUNT = CMIS_DT_PAYMENT   ' AMOUNT
                        PV_STATUS = "'N'"
                        PV_INVDATE = N2Date2Null(rsOFF_DT!ORDATE)

                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        Set rsSJ_DATA = New ADODB.Recordset
                        Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            J_JVOUCHERNO = J_VOUCHERNO
                            PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                            PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                            PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)    ' DATE
                            PV_AMOUNT = CMIS_DT_PAYMENT + J_DEPOSIT    ' AMOUNT
                            PV_STATUS = "'N'"

                            'DESCRIPTION: Modify database structure - added fields: CusCde - Insert J_Class and Customer Code upon importing (,J_Class,CusCde) NOT YET
                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                            "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                            ", " & PV_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT

                            'DESCRIPTION: Customer Deposits
                            If J_INVOICENUM <> "" Then
                                Dim rsCheckDetail           As ADODB.Recordset
                                Set rsCheckDetail = New ADODB.Recordset
                                rsCheckDetail.Open "SELECT ReferenceNo FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                                If Not rsCheckDetail.EOF And Not rsCheckDetail.BOF Then
                                    Dim rsCheckDetail2      As ADODB.Recordset
                                    Set rsCheckDetail2 = New ADODB.Recordset
                                    rsCheckDetail2.Open "SELECT * FROM CMIS_OFF_DT WHERE REFERENCENO=" & N2Str2Null(rsCheckDetail!ReferenceNo) & " AND OR_NUM NOT IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                                    If Not rsCheckDetail2.EOF And Not rsCheckDetail2.BOF Then
                                        Dim rsCheckVoucher  As ADODB.Recordset
                                        Set rsCheckVoucher = New ADODB.Recordset
                                        rsCheckVoucher.Open "SELECT VOUCHERNO,InvoiceAmt from AMIS_JOURNAL_HD WHERE InvoiceNo =" & N2Str2Null(rsCheckDetail2!OR_NUM) & "", gconDMIS, adOpenForwardOnly
                                        If Not rsCheckVoucher.EOF And Not rsCheckVoucher.BOF Then
                                            PV_AMOUNT = rsCheckVoucher!InvoiceAmt / 0.97
                                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                            "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                            " values (" & N2Str2Null(rsCheckVoucher!VOUCHERNO) & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                            ", " & PV_STATUS & ")"
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                        Set rsCheckVoucher = Nothing
                                    End If
                                    Set rsCheckDetail2 = Nothing
                                End If
                                Set rsCheckDetail = Nothing
                            End If

                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(PV_MRRNO)

                            Set rsCheckJournal_HD = New ADODB.Recordset
                            Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
                            If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'Y' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"

                                    gconDMIS.Execute SQL_STATEMENT

                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)



                                Else
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'N' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
                                    gconDMIS.Execute SQL_STATEMENT

                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                End If
                            Else
                                Set rsCheckJournal_HD = New ADODB.Recordset
                                Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
                                If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                    If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then

                                        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                         " ReceiveStatus = 'Y' " & "," & _
                                                         " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                         " Balance = Balance - " & PV_AMOUNT & _
                                                         " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                    Else
                                        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                        " ReceiveStatus = 'N' " & "," & _
                                                        " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                        " Balance = Balance - " & PV_AMOUNT & _
                                                        " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
                                        gconDMIS.Execute SQL_STATEMENT

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                    End If
                                End If
                            End If
                        Else
                            'DESCRIPTION: Insert detail for A/R CREDIT CARD / Credit Card payment receive
                            Set rsCheckCreditCardSJ = gconDMIS.Execute("select * from CMIS_OFF_DT where ReferenceNo = " & N2Str2Null(CMIS_DT_REFERENCENO) & " and TRANTYPE <>'OTH'")
                            If Not rsCheckCreditCardSJ.EOF And Not rsCheckCreditCardSJ.BOF Then
                                PV_INVNO = N2Str2Null(rsCheckCreditCardSJ!INVOICENO)
                                PV_MRRNO = "'" & rsCheckCreditCardSJ!TranType & "'"
                                PV_AMOUNT = (NumericVal(PV_AMOUNT) / 0.97)
                                SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_INVDATE & ", " & PV_AMOUNT & _
                                                ", " & PV_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT

                            End If
                        End If

                        'J_JITEMNO = "'0002'"
                        J_JITEMNO = Format(NumericVal(J_JITEMNO) + 1, "0000")
                        'RO  - SERVICE REPAIR ORDER
                        If CMIS_DT_TRANTYPE = "RO" Or CMIS_DT_TRANTYPE = "SI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                        End If
                        'CSH - PARTS CASH INVOICE
                        'CHG - PARTS CHARGE INVOICE
                        'If CMIS_DT_TRANTYPE = "CSH" Or CMIS_DT_TRANTYPE = "CHG" Then
                        If CMIS_DT_TRANTYPE = "PI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            'J_ACCT_CODE = COA_AR_TRADE_PARTS
                            'J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_PARTS))
                        End If
                        'VI  - VEHICLE INVOICE
                        If CMIS_DT_TRANTYPE = "VI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                        End If
                        'EST - SERVICE ESTIMATE
                        If CMIS_DT_TRANTYPE = "EST" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDeposit_AccountCode("SERVICE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeposit_AccountCode("SERVICE")))
                        End If
                        '* IOC - INTER OFFICE COLLECTION
                        'BRA - BRANCH PAYMENT
                        'If CMIS_DT_TRANTYPE = "BRA" Then
                        '    J_ACCT_CODE = COA_BRANCH_LEGASPI
                        '    J_ACCT_NAME = N2Str2Null(Setacctname(COA_BRANCH_LEGASPI))
                        'End If
                        'WAR - A/R WARRANTY CLAIMS
                        'INV - INVENTORIES - GAS,OIL,LUBS
                        'CRD - CREDIT CARD PAYMENT
                        'OTH
                        If COMPANY_CODE = "HBK" Then  ' BTT
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null("11-02104-00")
                                J_ACCT_NAME = N2Str2Null(Setacctname("11-02104-00"))
                            End If
                        End If
                        If CMIS_DT_TRANTYPE = "OTH" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            'OTHER TRANSACTION
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                        End If

                        If J_INVOICENUM <> "" Then
                            Dim rsCheckDeposit2             As ADODB.Recordset
                            Set rsCheckDeposit2 = New ADODB.Recordset
                            rsCheckDeposit2.Open "Select * from CMIS_Deposits where InvoiceNo =' " & J_INVOICENUM & "'", gconDMIS, adOpenForwardOnly
                            If Not rsCheckDeposit2.EOF And Not rsCheckDeposit2.BOF Then
                                J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                                J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                                J_DEPOSIT = Round(NumericVal(rsCheckDeposit2!amount), 2)
                            End If
                            Set rsCheckDeposit2 = Nothing
                        End If
                        J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                        J_TAX = 0
                        J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(J_DEPOSIT), 2)
                        J_DEBIT = 0
                        If J_ACCT_CODE = N2Str2Null("11-02002-00") Then
                            J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                            J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                            J_CREDIT = Round(NumericVal(J_NET) / 0.97, 2)
                        Else
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                        End If

                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        'DESCRIPTION: Insert in AMIS_Reference to be used in Customer Ledger for Customer and Credit Card Company
                        SQL_STATEMENT = "insert into AMIS_REFERENCE (VoucherNo,Jtype,ReferenceNo,JDate) values (" & J_VOUCHERNO & "," & J_JTYPE & "," & J_CUSTOMERCODE2 & "," & J_JDATE & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                        rsOFF_DT.MoveNext
                    Loop
                End If
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,ReferenceNo,Bank,Entity_Class)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & "," & J_CUSTOMERCODE & "," & N2Str2Null(CMIS_BANK) & "," & J_ENTITY & ")"

                gconDMIS.Execute SQL_STATEMENT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Screen.MousePointer = 0

    ImportUnDepositM = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportUnDepositM = False
End Function

Function ReturnExpense(XXX As String, Optional YYY As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If Trim(YYY) = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnExpense = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function GetCustomerORAmount(XXX As String) As Double
    Dim rsGetCustomerORAmount                               As ADODB.Recordset
    Set rsGetCustomerORAmount = New ADODB.Recordset
    Dim ORDATE                                              As String
    Dim ORAMOUNT                                            As Double
    If Len(XXX) > 6 Then
        rsGetCustomerORAmount.Open "SELECT OR_AMT FROM CMIS_OFF_HD WHERE OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORAMOUNT = N2Str2Zero(rsGetCustomerORAmount!OR_AMT)
    Else
        rsGetCustomerORAmount.Open "SELECT OR_AMT FROM CMIS_OFF_HD WHERE OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
        If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then
            ORAMOUNT = NumericVal(rsGetCustomerORAmount!OR_AMT)
        Else
        Set rsGetCustomerORAmount = New ADODB.Recordset
            rsGetCustomerORAmount.Open "SELECT amount FROM CMIS_OFF_dt WHERE reference='" & XXX & "' and trantype = 'OTH'", gconDMIS, adOpenForwardOnly
            If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then ORAMOUNT = NumericVal(rsGetCustomerORAmount!amount)
        End If
    End If
    
    If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then
        GetCustomerORAmount = ORAMOUNT
    End If
    
    Set rsGetCustomerORAmount = Nothing
End Function

Function GetCustomerORAmount_478(XXX As String) As Double
    Dim rsGetCustomerORAmount                               As ADODB.Recordset
    Set rsGetCustomerORAmount = New ADODB.Recordset
    Dim ORDATE                                              As String
    Dim ORAMOUNT                                            As Double
 
    Set rsGetCustomerORAmount = New ADODB.Recordset
    
    If COMPANY_CODE = "DPI" Then
        '021816
        rsGetCustomerORAmount.Open "SELECT PAYMENT + DISCOUNT + TAX AS PAYMENT FROM CMIS_OFF_dt WHERE reference='" & XXX & "' and trantype = 'OTH'", gconDMIS, adOpenForwardOnly
    ElseIf COMPANY_CODE = "CMC" Then
        rsGetCustomerORAmount.Open "SELECT PAYMENT + DISCOUNT + TAX AS PAYMENT FROM CMIS_OFF_dt WHERE reference='" & XXX & "' and trantype = 'OTH' AND OR_NUM = '" & CMIS_OR_NUM & "'", gconDMIS, adOpenForwardOnly
    Else
        rsGetCustomerORAmount.Open "SELECT PAYMENT FROM CMIS_OFF_dt WHERE reference='" & XXX & "' and trantype = 'OTH'", gconDMIS, adOpenForwardOnly
    End If
    
    If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then ORAMOUNT = NumericVal(rsGetCustomerORAmount!payment)
    
    If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then
        GetCustomerORAmount_478 = ORAMOUNT
    End If
    Set rsGetCustomerORAmount = Nothing
End Function

Function GetCustomerORDate(XXX As String) As String
    Dim rsGetCustomerORDate                                 As ADODB.Recordset
    Set rsGetCustomerORDate = New ADODB.Recordset
    Dim ORDATE                                              As String
    If Len(XXX) > 5 Then
        rsGetCustomerORDate.Open "SELECT ORDATE FROM CMIS_OFF_DT WHERE INVOICENO='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORDATE = N2Date2Null(rsGetCustomerORDate!ORDATE)
    Else
        rsGetCustomerORDate.Open "SELECT OR_DATE FROM CMIS_OFF_HD WHERE OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORDATE = N2Date2Null(rsGetCustomerORDate!OR_DATE)
    End If
    If Not rsGetCustomerORDate.EOF And Not rsGetCustomerORDate.BOF Then
        GetCustomerORDate = ORDATE
    End If
    Set rsGetCustomerORAmount = Nothing
End Function
Function GetifNonVat(XXX As String) As Boolean
    Dim rsGetifNonVat                                As ADODB.Recordset
    Set rsGetifNonVat = New ADODB.Recordset
    rsGetifNonVat.Open "SELECT * FROM CMIS_OFF_HD WHERE VAT = 0 AND OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGetifNonVat.EOF And Not rsGetifNonVat.BOF Then
        GetifNonVat = True
    Else
        GetifNonVat = False
    End If
End Function

Function CreditCardPayment(XXX As String, YYY As String, zzz As String) As Double
    Dim xACCT_CODE                                          As String
    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
        xACCT_CODE = "CARD"
    Else
        xACCT_CODE = "CARD ON HAND"
    End If
    Dim rsAMOUNT                                            As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND ACCT_CODE = " & N2Str2Null(ReturnAccountCode(xACCT_CODE)) & "", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        CreditCardPayment = NumericVal(N2Str2Zero(rsAMOUNT!Credit))
    End If
    Set rsAMOUNT = Nothing
End Function

Function AmountLessVAT(XXX As String, YYY As String) As Double
    Dim rsAMOUNT                                            As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT SUM(ISNULL(CREDIT,0)) AS CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND (LEFT(ACCT_CODE,2)='41' OR LEFT(ACCT_CODE,5)='71-61')", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        AmountLessVAT = NumericVal(rsAMOUNT!Credit)
    End If
    Set rsAMOUNT = Nothing
End Function

Function DiscountLessVAT(XXX As String, YYY As String) As Double
    Dim rsAMOUNT                                            As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT SUM(ISNULL(DEBIT,0)) AS DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND LEFT(ACCT_CODE,1)='5'", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        DiscountLessVAT = N2Str2Zero(rsAMOUNT!Debit)
    End If
    Set rsAMOUNT = Nothing
End Function

Function CRJLASTTRANS() As String
'Update: ACL 05092011
    Dim rsCHECKOR                                           As ADODB.Recordset
    Set rsCHECKOR = New ADODB.Recordset
    rsCHECKOR.Open "Select MAX(TRANDATE) AS TRANDATE FROM (SELECT MAX(OR_DATE) AS TRANDATE FROM CMIS_OFF_HD WHERE (PAIDNA = 1 OR STATUS = 'P') AND CANCEL =0 " & _
                   "UNION SELECT MAX(DATDEPOSIT) AS TRANDATE FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND CANCEL = 0) X", gconDMIS, adOpenKeyset
    If Not rsCHECKOR.EOF And Not rsCHECKOR.BOF Then
        CRJLASTTRANS = Null2Date(rsCHECKOR!trandate)
    End If
    Set rsCHECKOR = Nothing
End Function

Function CheckIfPMS(XXX As String) As Boolean
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    SQL = "select Status1 from CSMS_ro_det where jobtype='PMS' and livil='1' and wcode='W' and Rep_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        CheckIfPMS = True
    Else
        CheckIfPMS = False
    End If
    Set RS = Nothing
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function
Function ImportPurelyInternal() As Boolean
    On Error GoTo ErrorCode

    Dim GridImport                                          As Integer
    Dim i                                                   As Integer
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & Grid2.Cell(GridImport, 2).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_REL)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)
                CSMIOS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)

                'INTERNAL - COMPANY
                '====================================================================================================================================================================================

                COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2) Else COMPANY_DIRECT_EXPENSE_LABOR = 0

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2) Else COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2) Else COMPANY_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                           "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                           "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                           "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                           "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'C') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!Accessories))
                Else
                    COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
                End If

                '====================================================================================================================================================================================

                'INTERNAL - SALES DEPARTMENT
                '====================================================================================================================================================================================

                SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                    SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                Else
                    SALES_DIRECT_EXPENSE_LABOR = 0
                End If

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                Else
                    SALES_DIRECT_EXPENSE_SPAREPARTS = 0
                End If

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                Else
                    SALES_DIRECT_EXPENSE_GOL = 0
                End If

                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                           "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                           "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                           "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                           "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'S') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!Accessories))
                Else
                    SALES_DIRECT_EXPENSE_ACCESSORIES = 0
                End If
                '====================================================================================================================================================================================

                '=========================================================================================================================================================
                'ENTRY FOR PURELY INTERNAL
                'UPDATED BY: JUN - UPDATE DUE TO ERROR IN SAVING JNO
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    WARRANTY_JNO = "000001"
                End If
                'UPDATED BY: JUN

                If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT = 0 Then

                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"

                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'": J_REMARKS = "NULL": J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0: J_AMOUNTTOPAY = 0
                    CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)

                    J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_AMOUNTPAID = 0
                    J_STATUS = "'N'"

                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)

                    J_CHECKNO = "NULL": J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL): J_PAYTYPE = "NULL": J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL": J_BANKCODE = "NULL": J_REFNO = N2Str2Null(CSMIOS_REP_OR): J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM): J_DEALER = "NULL": J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"

                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()), "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'LABOR

                    'to check if there is more than 1 purely charge to internal : update By BTT
                    If PosibleDoubleInternal(CSMIOS_REP_OR) = True Then
                        ImportPurelyInternal = True
                        Exit Function
                    End If

                    INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0:
                    INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0:

                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                       
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                            J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)

                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                       

                       
                            'COST OF SALES
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            End If
                            J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        
                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    'PARTS
                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' ) AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_PARTS_AMT > 0 Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                                End If
                                J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO

                            If COMPANY_CODE = "HPI" Then
                            Else
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    End If
                                End If
                                J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If

                    End If

                    'MATERIALS
                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                    
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_MATERIALS_AMT > 0 Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                End If

                                J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
 
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    End If
                            
                                J_GROSS = 0
                                J_TAX = 0
                                J_NET = 0
                                J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                End If
                            J_GROSS = 0: J_TAX = 0: J_NET = 0
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If

                    
                        'ACCESSORIES
                        Set rsINTERNAL_RO_DET = New ADODB.Recordset
                        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '4' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                            rsINTERNAL_RO_DET.MoveFirst
                            Do While Not rsINTERNAL_RO_DET.EOF
                                If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                    INTERNAL_ACCESSORIES_AMT = INTERNAL_ACCESSORIES_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                    INTERNAL_ACCESSORIES_COST = INTERNAL_ACCESSORIES_AMT + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                    If COMPANY_CODE = "HPI" Then
                                    Else
                                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                        J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                        J_DEBIT = Round(NumericVal(J_NET), 2)
                                        J_CREDIT = 0
                                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                        gconDMIS.Execute SQL_STATEMENT
                                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                    End If
                                End If
                                rsINTERNAL_RO_DET.MoveNext
                            Loop

                            If INTERNAL_MATERIALS_AMT > 0 Then
                            
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                    End If

                                    J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                    Else
                                        J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    End If
                                    J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    J_DEBIT = 0
                                    J_CREDIT = Round(NumericVal(J_NET), 2)
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO


                                
                                    'COST OF SALES
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                        If CSMIOS_TERM = "CSH" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        Else
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        End If
                                        
                                    J_GROSS = 0
                                    J_TAX = 0
                                    J_NET = 0
                                    J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                                    J_CREDIT = 0
                                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'INVENTORY
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                                    End If
                                J_GROSS = 0: J_TAX = 0: J_NET = 0
                                J_DEBIT = 0
                                J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        End If
                    'OUT PUT TAX
                        If INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST > 0 Then
                            If CSMIOS_VAT_EXEMPT = False Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(Round(((INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST)), 2) * 0.12), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                J_TAX = 0: J_GROSS = 0: J_NET = 0
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        End If


                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                    CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                    CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)

                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_REL)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)

                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                    ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    ':-)=================================
                    Grid2.Cell(GridImport, 1).Text = 1
                End If
                '=========================================================================================================================================================
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    ImportPurelyInternal = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportPurelyInternal = False
End Function

Function ImportPurelyInternalNew() As Boolean
    On Error GoTo ErrorCode

    ALL_DEBIT = 0: ALL_CREDIT = 0
    Set rsCSMIOS_REPOR = New ADODB.Recordset
    'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
    'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & Grid2.Cell(GridImport, 2).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
    Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & CSMIOS_REP_OR & "'")
    If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
        ItemCnt = 0
        CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
        CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
        CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
        CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
        CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
        CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "FMC" Or COMPANY_CODE = "HCE" Then
            CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_REL)
        Else
            CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
        End If
        CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)
        CSMIOS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
        CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)

        'INTERNAL - COMPANY
        '====================================================================================================================================================================================

        COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0: COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
        If COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
            Set rsCSMIOS_LABOR = New ADODB.Recordset
            Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS LABOR,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'C' and livil = 1 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
            COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
            End If
            
            Dim rsCSMIOS_SR_C As New ADODB.Recordset
            Dim COMPANY_DIRECT_EXPENSE_SR_C As Double
            Set rsCSMIOS_SR_C = New ADODB.Recordset
            Set rsCSMIOS_SR_C = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS SUBLET,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'C' and ISNULL(ROTYPE,0) = 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_SR_C.EOF And Not rsCSMIOS_SR_C.BOF Then
            COMPANY_DIRECT_EXPENSE_SR_C = Round(N2Str2Zero(rsCSMIOS_SR_C!SUBLET), 2)
            COMPANY_DIRECT_EXPENSE_LABOR = COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SR_C
            End If
            
            Set rsCSMIOS_PARTS = New ADODB.Recordset
            Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS PARTS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'C' and livil = 2 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
            COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
            End If
            
            Set rsCSMIOS_MATERIALS = New ADODB.Recordset
            Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS MATERIALS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET  Where WCODE = 'C' and livil = 3 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
            COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
            End If
            
            Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
            Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS MATERIALS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET  Where WCODE = 'C' and livil = 4 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!MATERIALS), 2)
            End If
        Else
            Set rsCSMIOS_LABOR = New ADODB.Recordset
            Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2) Else COMPANY_DIRECT_EXPENSE_LABOR = 0
    
            Set rsCSMIOS_PARTS = New ADODB.Recordset
            Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2) Else COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0
    
            Set rsCSMIOS_MATERIALS = New ADODB.Recordset
            Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2) Else COMPANY_DIRECT_EXPENSE_GOL = 0
    
            Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
            rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                       "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                       "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                       "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                       "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'C') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
            If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!Accessories))
            Else
                COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
            End If
        End If

        '====================================================================================================================================================================================

        'INTERNAL - SALES DEPARTMENT
        '====================================================================================================================================================================================
        If COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
            Set rsCSMIOS_LABOR = New ADODB.Recordset
            Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS LABOR,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'S' and livil = 1 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
            SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
            End If
            
            Dim rsCSMIOS_SR_S As New ADODB.Recordset
            Dim COMPANY_DIRECT_EXPENSE_SR_S As Double
            Set rsCSMIOS_SR_S = New ADODB.Recordset
            Set rsCSMIOS_SR_S = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS SUBLET,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'S'   and ISNULL(ROTYPE,0)='SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_SR_S.EOF And Not rsCSMIOS_SR_S.BOF Then
            COMPANY_DIRECT_EXPENSE_SR_S = Round(N2Str2Zero(rsCSMIOS_SR_S!SUBLET), 2)
            SALES_DIRECT_EXPENSE_LABOR = SALES_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SR_S
            End If
            
             Set rsCSMIOS_PARTS = New ADODB.Recordset
            Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS PARTS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET Where WCODE = 'S' and livil = 2 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
            SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
            End If
            
            Set rsCSMIOS_MATERIALS = New ADODB.Recordset
            Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS MATERIALS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET  Where WCODE = 'S' and livil = 3 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
            SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
            End If
            
            Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
            Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DET_AMT),2) AS MATERIALS,ROUND(sum(DISCOUNT_2),2) AS DISCOUNT from CSMS_RO_DET  Where WCODE = 'S' and livil = 4 and ISNULL(ROTYPE,0) <> 'SR' and REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!MATERIALS), 2)
            End If
        Else
            SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0
    
            Set rsCSMIOS_LABOR = New ADODB.Recordset
            Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
            Else
                SALES_DIRECT_EXPENSE_LABOR = 0
            End If
    
            Set rsCSMIOS_PARTS = New ADODB.Recordset
            Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
            Else
                SALES_DIRECT_EXPENSE_SPAREPARTS = 0
            End If
    
            Set rsCSMIOS_MATERIALS = New ADODB.Recordset
            Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
            If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
            Else
                SALES_DIRECT_EXPENSE_GOL = 0
            End If
    
            Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
            rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                       "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                       "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                       "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                       "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'S') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
            If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!Accessories))
            Else
                SALES_DIRECT_EXPENSE_ACCESSORIES = 0
            End If
        End If
        '====================================================================================================================================================================================

        '=========================================================================================================================================================
        'ENTRY FOR PURELY INTERNAL
        'UPDATED BY: JUN - UPDATE DUE TO ERROR IN SAVING JNO
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
            WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
        Else
            WARRANTY_JNO = "000001"
        End If
        'UPDATED BY: JUN
'        If CSMIOS_REP_OR = "R-00000144" Then Stop
        If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL + SALES_DIRECT_EXPENSE_LABOR + SALES_DIRECT_EXPENSE_SPAREPARTS + SALES_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT = 0 Then

            Set rsJournal_HDDup = New ADODB.Recordset
            Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
            If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
            J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
            J_VOUCHERNO = N2Str2Null(GetVoucherNo())
            J_JTYPE = "'SJ'": J_REMARKS = "NULL": J_VENDORCODE = "'999999'"
            J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

            J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0: J_AMOUNTTOPAY = 0
            CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)

            J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
            J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
            J_AMOUNTPAID = 0
            J_STATUS = "'N'"

            J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
            J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)

            J_CHECKNO = "NULL": J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL): J_PAYTYPE = "NULL": J_INVOICETYPE = "'SI'"
            J_CHECKDATE = "NULL": J_BANKCODE = "NULL": J_REFNO = N2Str2Null(CSMIOS_REP_OR): J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
            J_TERMS = N2Str2Null(CSMIOS_TERM): J_DEALER = "NULL": J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
            
            WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()), "000000"))
            WARRANTY_ItemCnt = 0
            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
            TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
            'LABOR

            'to check if there is more than 1 purely charge to internal : update By BTT
            If PosibleDoubleInternal(CSMIOS_REP_OR) = True Then Exit Function

            INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0:
            INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0:
If COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCE" Then
Call PURELYITERNAL
Else
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCE" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCE" Then
                    Else
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                    ElseIf COMPANY_CODE = "HMH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL")))
                    ElseIf COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "CUSTOMER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "CUSTOMER")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    End If
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    If COMPANY_CODE = "HPI" Then
                        J_CREDIT = 0
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                    Else
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    End If

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    End If
                
                If COMPANY_CODE = "DJM" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                Else
                    'COST OF SALES
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HMH" Then
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    End If
                ElseIf COMPANY_CODE = "HMR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "CUSTOMER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "CUSTOMER")))
                Else
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    End If
                End If
                    J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                    J_CREDIT = 0
                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
                If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "HSM" Then
                Else
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
            End If
            
''SUBLET
If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
INTERNAL_LABOR_AMT = 0: INTERNAL_LABOR_COST = 0
 Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S')AND ROTYPE = 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    If COMPANY_CODE = "HMH" Then
                    Else
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET", "CUSTOMER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET", "CUSTOMER")))
                    ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HMR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    End If
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    If COMPANY_CODE = "HPI" Then
                        J_CREDIT = 0
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                    Else
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    End If

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    End If
                
                If COMPANY_CODE = "DJM" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HSM" Then
                Else
                    'COST OF SALES
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HMH" Then
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    End If
                ElseIf COMPANY_CODE = "HMR" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "CUSTOMER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "CUSTOMER")))
                Else
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    End If
                End If
                    J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                    J_CREDIT = 0
                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DJM" Or COMPANY_CODE = "HSM" Then
                Else
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
            End If
End If
'END SUBLET
            'PARTS
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' ) AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                If COMPANY_CODE = "HMH" Then
                Else
                If INTERNAL_PARTS_AMT > 0 Then
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       If COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "CUSTOMER"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "CUSTOMER")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                        End If
                        J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    End If
                    If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HMH" Then
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL")))
                            End If
                        ElseIf COMPANY_CODE = "HMR" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "CUSTOMER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "CUSTOMER")))
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            End If
                        End If
                        J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                 If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                 Else
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    ElseIf COMPANY_CODE = "HCA" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARST-IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARST-IN-PROCESS")))
                    Else
                    
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                 End If
            End If

            'MATERIALS
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop

                If INTERNAL_MATERIALS_AMT > 0 Then
                    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCE" Then
                    Else
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HSM" Or COMPANY_CODE = "HMR" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "CUSTOMER"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "CUSTOMER")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                        End If

                        J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    

                    If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HMH" Then
                             If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                            Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                                End If
                        ElseIf COMPANY_CODE = "HMR" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "CUSTOMER"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "CUSTOMER")))
                        ElseIf COMPANY_CODE = "HCA" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                End If
                            Else
                                If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                End If
                            End If
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                    
                    If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                    Else
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HMH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM", "MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM", "MATERIALS")))
                    ElseIf COMPANY_CODE = "HCA" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    Else
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                        End If
                    End If
                    J_GROSS = 0: J_TAX = 0: J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                 End If
                End If
            End If

            
                'ACCESSORIES

             INTERNAL_ACCESSORIES_COST = 0
                Set rsINTERNAL_RO_DET = New ADODB.Recordset
                Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '4' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
                If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                    rsINTERNAL_RO_DET.MoveFirst
                    Do While Not rsINTERNAL_RO_DET.EOF
                        If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                            INTERNAL_ACCESSORIES_AMT = INTERNAL_ACCESSORIES_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                            INTERNAL_ACCESSORIES_COST = INTERNAL_ACCESSORIES_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                            If COMPANY_CODE = "HPI" Then
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                If COMPANY_CODE = "HMH" Then
                                J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                J_TAX = 0
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                Else
                                J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                J_TAX = 0
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                End If
                                J_DEBIT = Round(NumericVal(J_NET), 2)
                                J_CREDIT = 0
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                If COMPANY_CODE = "HCA" Then
                                        Set rsUEA = New ADODB.Recordset
                                        Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                            If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                            Else
                                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                    gconDMIS.Execute SQL_STATEMENT
                                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                            End If
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                        End If
                        rsINTERNAL_RO_DET.MoveNext
                    Loop

                    If INTERNAL_MATERIALS_AMT > 0 Then
                    If COMPANY_CODE = "HMH" Then
                    Else
                        
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "CUSTOMER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "CUSTOMER")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                            End If
                            J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            If CSMIOS_VAT_EXEMPT = True Then
                                J_TAX = 0
                            Else
                                J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            End If
                            J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                        

                        If COMPANY_CODE = "HNE" Then
                        Else
                            'COST OF SALES
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "ACCESSORIES", "CUSTOMER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "ACCESSORIES", "CUSTOMER")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    End If
                                End If
                            End If
                            J_GROSS = 0
                            J_TAX = 0
                            J_NET = 0
                            J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                            J_CREDIT = 0
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If
                        
                        'INVENTORY
                        If COMPANY_CODE = "HNE" Then
                        Else
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                            End If
                        End If
                        J_GROSS = 0: J_TAX = 0: J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(INTERNAL_ACCESSORIES_COST, 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                       End If
                    End If
                End If
End If
            'OUT PUT TAX
'            If COMPANY_CODE = "HPI" Then
'                If INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST > 0 Then
'                    If CSMIOS_VAT_EXEMPT = False Then
'                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
'                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
'                        If COMPANY_CODE = "HBK" Then
'                            If CSMIOS_TERM = "CHG" Then
'                                J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
'                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
'                            Else
'
'                                If COMPANY_CODE = "HBK" Then
'                                    J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
'                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
'                                Else
'                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
'                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
'                                End If
'                            End If
'                        Else
'                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
'                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
'                        End If
'                        J_DEBIT = 0
'                        J_CREDIT = Round(NumericVal(Round(((INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST)), 2) * 0.12), 2)
'                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
'                        J_TAX = 0: J_GROSS = 0: J_NET = 0
'                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
'                    End If
'                End If
'            End If


            CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
            CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
            CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
            CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
            CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)

            CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
            CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_REL)
            CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)

            J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)
            If COMPANY_CODE = "DJM" Then
                J_REMARKS = " SERVICE BILLING:" + " " + CSMIOS_INVOICE
            Else
                J_REMARKS = " SERVICE INVOICENO:" + " " + CSMIOS_INVOICE
            End If
            
            WARRANTY_J_AMOUNTTOPAY = 0
            WARRANTY_J_INVOICEAMT = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
            WARRANTY_J_BALANCE = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
            WARRANTY_J_AMOUNTPAID = 0
            If (Left(CSMIOS_INVOICE, 6) = "INT RO") Or COMPANY_CODE = "HSM" Then
                WARRANTY_J_INVOICEAMT = 0
                WARRANTY_J_BALANCE = 0
            Else
                WARRANTY_J_INVOICEAMT = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                WARRANTY_J_BALANCE = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
            End If
            SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                            " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                            ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & ",'" & J_REMARKS & "'," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
            gconDMIS.Execute SQL_STATEMENT
            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
            NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
            ':-)=================================
        End If
        '=========================================================================================================================================================
    End If
    ImportPurelyInternalNew = True
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & err.Number & vbCrLf & "Error Description :" & err.DESCRIPTION
    ImportPurelyInternalNew = False
    ShowVBError
End Function
Sub PURELYITERNAL()
If COMPANY_CODE <> "HCA" Then
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND ISNULL(ROTYPE,0) IN ('GJ','PMS') AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICECHG", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICECHG", "LABOR", "INTERNAL")))
                    End If
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    If INTERNAL_LABOR_COST > 0 Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICECHG", "LABOR", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICECHG", "LABOR", "INTERNAL")))
                        End If
                        J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                        J_CREDIT = 0
                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    
                        'INVENTORY
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                        J_DEBIT = 0
                        J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
            End If
            
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND ISNULL(ROTYPE,0) IN ('BP') AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABORBP", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABORBP", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICECHG", "LABORBP", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICECHG", "LABORBP", "INTERNAL")))
                    End If
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    If INTERNAL_LABOR_COST > 0 Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABORBP", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABORBP", "INTERNAL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICECHG", "LABORBP", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICECHG", "LABORBP", "INTERNAL")))
                        End If
                        J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                        J_CREDIT = 0
                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    
                        'INVENTORY
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                        J_DEBIT = 0
                        J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
            End If
Else
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    If INTERNAL_LABOR_COST > 0 Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                        J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                        J_CREDIT = 0
                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    
                        'INVENTORY
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                        J_DEBIT = 0
                        J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
            End If
End If
            
''SUBLET
If COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCE" Then
INTERNAL_LABOR_AMT = 0: INTERNAL_LABOR_COST = 0
 Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S')AND ROTYPE = 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                        INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Set rsUEA = New ADODB.Recordset
                            Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                            If Not rsUEA.EOF And Not rsUEA.BOF Then
                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                            Else
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                    If COMPANY_CODE = "HMH" Then
                    Else
                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                    ElseIf COMPANY_CODE = "HMH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR-GJ", "INTERNAL")))
                    ElseIf COMPANY_CODE = "HCE" Then
                            If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLETCHG", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLETCHG", "INTERNAL")))
                            End If
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    End If
                    J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    J_TAX = 0
                    J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    If COMPANY_CODE = "HPI" Then
                        J_CREDIT = 0
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                    Else
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    End If

                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Set rsUEA = New ADODB.Recordset
                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                        Else
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
        
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                        End If
                    End If
                
                If COMPANY_CODE = "DJM" Or COMPANY_CODE = "HMH" Then
                Else
                    'COST OF SALES
                If COMPANY_CODE = "HMH" Then
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR-GJ", "INTERNAL")))
                    End If
                ElseIf COMPANY_CODE = "HCE" Then
                        If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET", "INTERNAL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLETCHG", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLETCHG", "INTERNAL")))
                        End If
                Else
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                    End If
                End If
                    J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                    J_CREDIT = 0
                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                    Set rsUEA = New ADODB.Recordset
                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                    If Not rsUEA.EOF And Not rsUEA.BOF Then
                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                    Else
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DJM" Then
                Else
                    'INVENTORY
                    If COMPANY_CODE = "HCE" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                    Else
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    Set rsUEA = New ADODB.Recordset
                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                    If Not rsUEA.EOF And Not rsUEA.BOF Then
                        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                    Else
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
            End If
End If
'END SUBLET
            'PARTS
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' ) AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                If COMPANY_CODE = "HMH" Then
                Else
                If INTERNAL_PARTS_AMT > 0 Then
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       If COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS-GJ", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HCE" Then
                            If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTSCHG", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTSCHG", "INTERNAL")))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                        End If
                        J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    End If
                    If COMPANY_CODE = "HNE" Then
                    Else
                        'COST OF SALES
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HMH" Then
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS-GJ", "INTERNAL")))
                            End If
                        ElseIf COMPANY_CODE = "HCE" Then
                            If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTSCHG", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTSCHG", "INTERNAL")))
                            End If
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            End If
                        End If
                        J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                 If COMPANY_CODE <> "HNE" Then
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HCE" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                 End If
            End If

            'MATERIALS
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                        INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            If COMPANY_CODE = "HMH" Then
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                            Else
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                        End If
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop

                If INTERNAL_MATERIALS_AMT > 0 Then
                    If COMPANY_CODE = "HMH" Then
                    Else
                        If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HCA" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HCE" Then
                            If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "MATERIALS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "MATERIALS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "MATERIALSCHG", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "MATERIALSCHG", "INTERNAL")))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                        End If

                        J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                        End If
                        Else
                                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                         SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                        gconDMIS.Execute SQL_STATEMENT
                        End If
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    

                    If COMPANY_CODE = "HNE" Then
                    Else
                        'COST OF SALES
                        If COMPANY_CODE = "HMH" Then
                             If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                            Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS-GJ", "INTERNAL")))
                                End If
                        ElseIf COMPANY_CODE = "HCA" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                        ElseIf COMPANY_CODE = "HCE" Then
                            If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALS", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "MATERIALSCHG", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "MATERIALSCHG", "INTERNAL")))
                            End If
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                End If
                            Else
                                If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                End If
                            End If
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                         If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                        End If
                        Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If
                End If
                    
                    If COMPANY_CODE = "HNE" Then
                    Else
                    'INVENTORY
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HMH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM", "MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM", "MATERIALS")))
                    ElseIf COMPANY_CODE = "HCA" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    ElseIf COMPANY_CODE = "HCE" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM", "MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM", "MATERIALS")))
                    Else
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                        End If
                    End If
                    J_GROSS = 0: J_TAX = 0: J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                 End If
                End If
            End If

            
                'ACCESSORIES

             INTERNAL_ACCESSORIES_COST = 0
                Set rsINTERNAL_RO_DET = New ADODB.Recordset
                Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '4' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND ISNULL(ROTYPE,0) <> 'SR' AND REP_OR = '" & CSMIOS_REP_OR & "'")
                If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                    rsINTERNAL_RO_DET.MoveFirst
                    Do While Not rsINTERNAL_RO_DET.EOF
                        If (N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0) Then
                            INTERNAL_ACCESSORIES_AMT = INTERNAL_ACCESSORIES_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                            INTERNAL_ACCESSORIES_COST = INTERNAL_ACCESSORIES_COST + (N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                            If COMPANY_CODE = "HPI" Then
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                If COMPANY_CODE = "HMH" Then
                                J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                J_TAX = 0
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETCOST) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                Else
                                J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                J_TAX = 0
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                End If
                                J_DEBIT = Round(NumericVal(J_NET), 2)
                                J_CREDIT = 0
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                If COMPANY_CODE = "HCA" Then
                                        Set rsUEA = New ADODB.Recordset
                                        Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                            If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + " & J_DEBIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                            Else
                                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                    gconDMIS.Execute SQL_STATEMENT
                                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                            End If
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                        End If
                        rsINTERNAL_RO_DET.MoveNext
                    Loop

                    If INTERNAL_MATERIALS_AMT > 0 Then
                    If COMPANY_CODE = "HMH" Then
                    Else
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "CUSTOMER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "CUSTOMER")))
                            ElseIf COMPANY_CODE = "HCA" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                            ElseIf COMPANY_CODE = "HCE" Then
                                If CSMIOS_TERM = "" Or CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "INTERNAL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES", "INTERNAL")))
                                End If
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                            End If
                            J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            If CSMIOS_VAT_EXEMPT = True Then
                                J_TAX = 0
                            Else
                                J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            End If
                            J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            If COMPANY_CODE = "HCA" Then
                                    Set rsUEA = New ADODB.Recordset
                                    Set rsUEA = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        If Not rsUEA.EOF And Not rsUEA.BOF Then
                                                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET CREDIT = CREDIT + " & J_CREDIT & " WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE  = " & J_ACCT_CODE & "")
                                        Else
                                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                                gconDMIS.Execute SQL_STATEMENT
                                        End If
                        Else
                                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                         SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                        gconDMIS.Execute SQL_STATEMENT
                        End If
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                        

                        If COMPANY_CODE = "HNE" Then
                        Else
                            'COST OF SALES
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "ACCESSORIES", "CUSTOMER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "ACCESSORIES", "CUSTOMER")))
                            ElseIf COMPANY_CODE = "HCA" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            ElseIf COMPANY_CODE = "HCE" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "ACCESSORIES", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "ACCESSORIES", "INTERNAL")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    End If
                                End If
                            End If
                            J_GROSS = 0
                            J_TAX = 0
                            J_NET = 0
                            J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                            J_CREDIT = 0
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If
                        
                        'INVENTORY
                        If COMPANY_CODE = "HNE" Then
                        Else
                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                        ElseIf COMPANY_CODE = "HCE" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("AIN-PROCESS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("AIN-PROCESS")))
                        Else
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                            End If
                        End If
                        J_GROSS = 0: J_TAX = 0: J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(INTERNAL_ACCESSORIES_COST, 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                       End If
                    End If
                End If

End Sub
Sub DOTSERVICEPARTS()
If COMPANY_CODE = "HMR" Then
    ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
    J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
    J_CREDIT = Round(NumericVal(NumericVal(NumericVal(J_INVOICEAMT)) / 1.12 * 0.12), 2)
    ALL_CREDIT = ALL_CREDIT + J_CREDIT
    J_TAX = 0: J_GROSS = 0: J_NET = 0: J_DEBIT = 0
    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
    Call DOTAX_SERVICE
Else
    If Round(NumericVal(NumericVal(NumericVal(CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_LABOR) - NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)) / 1.12 * 0.12), 2) > 0 Then
        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("DOTSERVICE"))
        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("DOTSERVICE")))
        J_DEBIT = 0
        J_CREDIT = Round(NumericVal(NumericVal(NumericVal(CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_LABOR) - NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)) / 1.12 * 0.12), 2)
        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        ''INSERT AMIS_AP
        Set rsDOTCIS = New ADODB.Recordset
        Set rsDOTCIS = gconDMIS.Execute("SELECT * FROM AMIS_CHARTACCOUNT WHERE AcctCode = " & J_ACCT_CODE & " AND IS_SCHEDULE_ACCNT =1 AND TRANTYPE1 = 'DOTSERVICE'")
        If Not rsDOTCIS.EOF And Not rsDOTCIS.BOF Then
            xIDDOTAX = (gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE =" & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "").Fields(0).Value)
            xCNAME = (gconDMIS.Execute("select ACCOUNTNAME from ALL_ENTITY   WHERE ENTITYCODE = 'C' AND CODE = " & J_CUSTOMERCODE & "").Fields(0).Value)
            SQL_STATEMENT = "INSERT INTO AMIS_AP (Voucherno,Vendor_code,VENDOR_NAME,invoicedate,Amount2pay,AmountPaid,Balance,Acct_code,LastUpdated,INVOICENO,INVOICETYPE,JDATE,MODIFIED_DATE,REFCODE,ENTITYCODE,STATUS,JOURNAL_DET_ID) " & _
                            " values ('" & "SJ" + "-" + GetVoucherNo & "', " & J_CUSTOMERCODE & ", '" & xCNAME & "', " & J_INVOICEDATE & "," & J_CREDIT & ",'0'," & J_CREDIT & "," & J_ACCT_CODE & "," & J_JDATE & ", " & J_INVOICENO & ", " & J_INVOICETYPE & "," & J_JDATE & "," & J_JDATE & ",'" & "C" + CSMIOS_ACCT_NO & "','C','N'," & xIDDOTAX & ")"
            gconDMIS.Execute SQL_STATEMENT
        End If
        ''--------------------------------------
        ALL_CREDIT = ALL_CREDIT + J_CREDIT
        J_TAX = 0: J_GROSS = 0: J_NET = 0: J_DEBIT = 0: J_CREDIT = 0: xIDDOTAX = "": xCNAME = ""
    End If
    
    If Round(NumericVal(NumericVal(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - NumericVal(CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_ACCESSORIES_DISCOUNT)) / 1.12 * 0.12), 2) > 0 Then
        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("DOTPARTS"))
        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("DOTPARTS")))
        J_DEBIT = 0
        J_CREDIT = Round(NumericVal(NumericVal(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - NumericVal(CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_ACCESSORIES_DISCOUNT)) / 1.12 * 0.12), 2)
        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        ''INSERT AMIS_AP
        Set rsDOTCIS = New ADODB.Recordset
        Set rsDOTCIS = gconDMIS.Execute("SELECT * FROM AMIS_CHARTACCOUNT WHERE AcctCode = " & J_ACCT_CODE & " AND IS_SCHEDULE_ACCNT =1 AND TRANTYPE1 = 'DOTPARTS'")
        If Not rsDOTCIS.EOF And Not rsDOTCIS.BOF Then
            xIDDOTAX = (gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE =" & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "").Fields(0).Value)
            xCNAME = (gconDMIS.Execute("select ACCOUNTNAME from ALL_ENTITY   WHERE ENTITYCODE = 'C' AND CODE = " & J_CUSTOMERCODE & "").Fields(0).Value)
            SQL_STATEMENT = "INSERT INTO AMIS_AP (Voucherno,Vendor_code,VENDOR_NAME,invoicedate,Amount2pay,AmountPaid,Balance,Acct_code,LastUpdated,INVOICENO,INVOICETYPE,JDATE,MODIFIED_DATE,REFCODE,ENTITYCODE,STATUS,JOURNAL_DET_ID) " & _
                            " values ('" & "SJ" + "-" + GetVoucherNo & "', " & J_CUSTOMERCODE & ", '" & xCNAME & "', " & J_INVOICEDATE & "," & J_CREDIT & ",'0'," & J_CREDIT & "," & J_ACCT_CODE & "," & J_JDATE & ", " & J_INVOICENO & ", " & J_INVOICETYPE & "," & J_JDATE & "," & J_JDATE & ",'" & "C" + CSMIOS_ACCT_NO & "','C','N'," & xIDDOTAX & ")"
            gconDMIS.Execute SQL_STATEMENT
        End If
        ''--------------------------------------
        ALL_CREDIT = ALL_CREDIT + J_CREDIT
        J_TAX = 0: J_GROSS = 0: J_NET = 0: J_DEBIT = 0: J_CREDIT = 0: xIDDOTAX = "": xCNAME = ""
    End If
End If
End Sub
Sub DOTAX_SERVICE()
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xDUEDATE                                            As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xREFCUSTOMERCODE                                    As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xENTITYCODE                                         As String
    Dim xID                                                 As String
    Dim xBAL                                                As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0
Set DOTAX_SERVICE1 = New ADODB.Recordset
  Set DOTAX_SERVICE1 = gconDMIS.Execute("Select* from amis_chartaccount Where ACCTCODE = " & J_ACCT_CODE & " AND TRANTYPE1 = 'DEFERRED OUTPUT TAX' AND IS_SCHEDULE_ACCNT = 1  ")

    If Not DOTAX_SERVICE1.EOF And Not DOTAX_SERVICE1.BOF Then
                    xJType = "SJ"
                    xID = N2Str2IntZero(gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & J_ACCT_CODE & "").Fields(0).Value)
                    xENTITYCODE = N2Str2Null("C")
                    xREFCUSTOMERCODE = N2Str2Null(Null2String("C" & J_CUSTOMERCODE))
                    XCustomerCode = N2Str2Null(J_CUSTOMERCODE)
                    xCUST_NAME = N2Str2Null(gconDMIS.Execute("SELECT ACCTNAME FROM ALL_Customer_Table WHERE CUSCDE = " & J_CUSTOMERCODE & "").Fields(0).Value)
                    xInvoiceNo = N2Str2Null(J_INVOICENO)
                    xInvoiceType = N2Str2Null("SI")
                    xInvoicedate = N2Str2Null(J_JDATE)
                    xAMOUNT_TO_PAY = N2Str2Zero(J_CREDIT)
                    xAMOUNT_PAID = N2Str2Zero(J_CREDIT)
                    xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                    xACCT_CODE = N2Str2Null(J_ACCT_CODE)
                    xLAST_UPDATED = N2Str2Null(J_JDATE)
                    xDUEDATE = N2Str2Null(J_JDATE)
                    xJdate = N2Str2Null(J_JDATE)
                    xVOUCHERNO = N2Str2Null(Null2String(xJType) & "-" & Null2String(J_VOUCHERNO))
                    SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE,REFCODE,ENTITYCODE,journal_det_id) " & _
                                    "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & "," & xREFCUSTOMERCODE & "," & xENTITYCODE & "," & xID & ")"
                    gconDMIS.Execute SQL_STATEMENT
        End If
    End Sub
    Sub ARSERVICEPARTS()
    If Round(NumericVal(NumericVal(CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_LABOR) - NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)), 2) > 0 Then
        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
        J_DEBIT = Round(NumericVal(NumericVal(CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_LABOR) - NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)), 2)
        J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
        ALL_DEBIT = ALL_DEBIT + J_DEBIT
        ALL_CREDIT = ALL_CREDIT + J_CREDIT
        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        J_TAX = 0: J_GROSS = 0: J_NET = 0: J_DEBIT = 0: J_CREDIT = 0
    End If
    
    If Round(NumericVal(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - NumericVal(CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_ACCESSORIES_DISCOUNT)), 2) > 0 Then
        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
        J_DEBIT = Round(NumericVal(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - NumericVal(CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_ACCESSORIES_DISCOUNT)), 2)
        J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
        ALL_DEBIT = ALL_DEBIT + J_DEBIT
        ALL_CREDIT = ALL_CREDIT + J_CREDIT
        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        J_TAX = 0: J_GROSS = 0: J_NET = 0: J_DEBIT = 0: J_CREDIT = 0
    End If
End Sub
Sub rod()
If COMPANY_CODE = "HCE" Then
     Dim DES_FB As New ADODB.Recordset
     Dim FB_DES As Double
     Dim ACCT_ADJ As New ADODB.Recordset
     Dim ADJ_ACCT As String
     ADJ_ACCT = ""
     FB_DES = 0
     Set DES_FB = New ADODB.Recordset
     Set DES_FB = gconDMIS.Execute("SELECT SUM(ISNULL(DEBIT,0)) -  SUM(ISNULL(CREDIT,0)) AS DESCREPANCY FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & WARRANTY_VOUCHERNO & " AND JTYPE = 'SJ'")
     If Not DES_FB.EOF And Not DES_FB.BOF Then
     FB_DES = Round(N2Str2Zero(DES_FB!DESCREPANCY), 2)
     End If
     Set ACCT_ADJ = New ADODB.Recordset
     Set ACCT_ADJ = gconDMIS.Execute("SELECT TOP 1 ACCT_CODE AS ACCTCODE FROM AMIS_JOURNAL_DET A INNER JOIN AMIS_CHARTACCOUNT B ON A.ACCT_CODE = B.ACCTCODE WHERE A.VOUCHERNO = " & WARRANTY_VOUCHERNO & " AND A.JTYPE = 'SJ' AND B.TRANTYPE3 = 'DISCOUNT'")
     If Not ACCT_ADJ.EOF And Not ACCT_ADJ.BOF Then
     ADJ_ACCT = (ACCT_ADJ!AcctCode)
     End If
     If FB_DES = 0.01 Then
     gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT - 0.01 WHERE VOUCHERNO = " & WARRANTY_VOUCHERNO & " AND JTYPE = 'SJ' AND ACCT_CODE = '" & ADJ_ACCT & "'")
     ElseIf FB_DES = -0.01 Then
     gconDMIS.Execute ("UPDATE AMIS_JOURNAL_DET SET DEBIT = DEBIT + 0.01 WHERE VOUCHERNO = " & WARRANTY_VOUCHERNO & " AND JTYPE = 'SJ' AND ACCT_CODE = '" & ADJ_ACCT & "'")
     Else
     End If
Else
    Dim rsROD As New ADODB.Recordset
    Dim xROD As Double
    xROD = 0
    Set rsROD = New ADODB.Recordset
    Set rsROD = gconDMIS.Execute("SELECT (SUM(DEBIT) - SUM(CREDIT)) AS DIF FROM AMIS_JOURNAL_DET WHERE JType = " & J_JTYPE & " AND VoucherNo = " & J_VOUCHERNO & "")
    If Not rsROD.EOF And Not rsROD.BOF Then
        xROD = Round(rsROD!DIF, 2)
        If xROD = 0.01 Or xROD = 0.02 Then
            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("ROD"))
            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("ROD")))
            J_DEBIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
            J_CREDIT = xROD
            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        ElseIf xROD = -0.01 Or xROD = -0.02 Then
            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
            J_ACCT_CODE = N2Str2Null(ReturnAccountCode("ROD"))
            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("ROD")))
            J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
            J_DEBIT = (xROD * -1)
            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        End If
    End If
End If
End Sub
Function GetVoucherNoSJ() As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNoSJ = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNoSJ = "000001"
    End If
End Function

Function ReturnOutputTax2(InvType As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX' and TRANTYPE3 = 'SERVICE' and TRANTYPE2 = '" & InvType & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnOutputTax2 = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function
Sub CRJ_DEFFERED_TAX()
'---------------------DOTAX BY NORMAN
Dim xVOUCHERNO      As String
Dim xJType          As String
Dim xJdate          As String
Dim xVENDORCODE     As String
Dim xInvoiceNo      As String
Dim xInvoiceType    As String
Dim xAMOUNTPAID     As String
Dim xACCT_CODE      As String
Dim xPV_VOUCHERNO   As String
Dim xModidate       As String
Dim xInvoicedate    As String
Dim xENTITYCODE     As String
Dim xREFCODE        As String
Dim xSTATUS         As String
Dim xJournalDETID   As String


Set DOTAXDS = New ADODB.Recordset
Set DOTAXDS = gconDMIS.Execute("Select * from amis_ap ap inner join amis_chartaccount ca on ap.acct_code = ca.acctcode  Where ap.INVOICENO = '" & CMIS_DT_REFERENCE & "' AND ap.INVOICETYPE = '" & CMIS_DT_TRANTYPE & "' AND ap.VENDOR_CODE = " & CMIS_DT_CUSCDE & " and ca.IS_SCHEDULE_ACCNT = 1 AND TRANTYPE1 = 'DEFERRED OUTPUT TAX'  ")
If Not DOTAXDS.EOF And Not DOTAXDS.BOF Then
    xVOUCHERNO = N2Str2Null(J_VOUCHERNO)
    xJType = N2Str2Null(J_JTYPE)
    xJdate = N2Str2Null(J_JDATE)
    xVENDORCODE = N2Str2Null(Null2String(DOTAXDS!VENDOR_CODE))
    xInvoiceNo = N2Str2Null(Null2String(DOTAXDS!INVOICENO))
    xInvoiceType = N2Str2Null(Null2String(DOTAXDS!INVOICETYPE))
    If COMPANY_CODE = "DSSC" Then
        xAMOUNTPAID = J_DEBIT
    Else
        xAMOUNTPAID = N2Str2Null(Null2String(DOTAXDS!AMOUNT2PAY))
    End If
    xACCT_CODE = N2Str2Null(Null2String(DOTAXDS!ACCT_CODE))
    xPV_VOUCHERNO = N2Str2Null(Null2String(DOTAXDS!VOUCHERNO))
    xModidate = N2Str2Null(J_JDATE)
    xInvoicedate = N2Str2Null(J_JDATE)
    xENTITYCODE = N2Str2Null(Null2String(DOTAXDS!ENTITYCODE))
    xREFCODE = N2Str2Null(Null2String(DOTAXDS!REFCODE))
    xSTATUS = N2Str2Null("N")
    xJournalDETID = N2Str2IntZero(gconDMIS.Execute("SELECT ID FROM AMIS_JOURNAL_DET WHERE JTYPE = " & J_JTYPE & " AND VOUCHERNO = " & J_VOUCHERNO & " AND ACCT_CODE = " & xACCT_CODE & "").Fields(0).Value)

     SQL_STATEMENT = "INSERT INTO AMIS_DETAILS(VOUCHERNO,JTYPE,VENDORCODE,ACCT_CODE,PV_VOUCHERNO,INVOICENO,AMOUNTPAID,JDATE,ENTITYCODE,REFCODE,JOURNAL_DET_ID,STATUS,INVOICEDATE) " & _
                                       "VALUES(" & xVOUCHERNO & "," & xJType & "," & xVENDORCODE & "," & xACCT_CODE & "," & xPV_VOUCHERNO & "," & xInvoiceNo & "," & xAMOUNTPAID & "," & xJdate & "," & xENTITYCODE & "," & xREFCODE & "," & xJournalDETID & "," & xSTATUS & "," & xInvoicedate & ")"
    gconDMIS.Execute SQL_STATEMENT
End If
End Sub

Function bankaccountcode(XXX As String) As String
    Dim rsbankaccountcode As New ADODB.Recordset
    Set rsbankaccountcode = New ADODB.Recordset
    
    If COMPANY_CODE = "MGS" Or COMPANY_CODE = "DJM" Then
        Set rsbankaccountcode = gconDMIS.Execute("select * from ALL_BANKDEPOSITS where bankcode='" & XXX & "'")
    Else
        Set rsbankaccountcode = gconDMIS.Execute("select * from all_banks where bankcode='" & XXX & "'")
    End If
        
    If Not rsbankaccountcode.EOF And Not rsbankaccountcode.BOF Then
        bankaccountcode = (rsbankaccountcode!AcctCode)
    End If
End Function

Private Sub Grid1_DblClick()
    If Grid1.Rows = 0 Then Exit Sub
    If Grid1.ActiveCell.Col <> 3 Then Exit Sub
    
    Dim HDInvNo As String
    HDInvNo = Right(Grid1.ActiveCell.Text, 6)
    
    JOURNALTYPE = ""
    Unload frmAMISJournalEntry_CRJ
    JOURNALTYPE = "CRJ"
    Call frmAMISJournalEntry_CRJ.LOADJOURNAL("CRJ")
    frmAMISJournalEntry_CRJ.Show
    Call frmAMISJournalEntry_CRJ.StoreSearch3(HDInvNo, "")
End Sub
