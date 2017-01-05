VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DMIS 2.0"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   ForeColor       =   &H00000000&
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "About.frx":000C
   ScaleHeight     =   6180
   ScaleWidth      =   4335
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   0
      Picture         =   "About.frx":549D
      ScaleHeight     =   1395
      ScaleWidth      =   4290
      TabIndex        =   11
      Top             =   0
      Width           =   4290
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      Picture         =   "About.frx":A092
      ScaleHeight     =   5415
      ScaleWidth      =   4335
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton checkdup 
         Height          =   465
         Left            =   3000
         TabIndex        =   52
         Top             =   4860
         Width           =   1245
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   315
         Left            =   1410
         TabIndex        =   53
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EC87
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   8730
         TabIndex        =   51
         Top             =   2490
         Width           =   3945
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EDDB
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8730
         TabIndex        =   50
         Top             =   1140
         Width           =   3945
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EEE4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   8745
         TabIndex        =   49
         Top             =   90
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maui Hosmillo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   5340
         TabIndex        =   48
         Top             =   1500
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mabie Barja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   5340
         TabIndex        =   47
         Top             =   1200
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Joy Montes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   5340
         TabIndex        =   46
         Top             =   900
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hannah Brito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   5340
         TabIndex        =   45
         Top             =   600
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ariel Villarin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   5340
         TabIndex        =   44
         Top             =   330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yeng Paquejo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   5340
         TabIndex        =   43
         Top             =   60
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Samantha Yzabell"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   5340
         TabIndex        =   42
         Top             =   2430
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Laurence Jaucian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   5340
         TabIndex        =   41
         Top             =   2100
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ana Marie DJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   5340
         TabIndex        =   40
         Top             =   1800
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special Thanks to"
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
         Index           =   33
         Left            =   5340
         TabIndex        =   34
         Top             =   4530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "University of Nueva Caceres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   5340
         TabIndex        =   36
         Top             =   4800
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reah && Anjali"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   5340
         TabIndex        =   32
         Top             =   4200
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Miles, Gee and Anne"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   5340
         TabIndex        =   39
         Top             =   3900
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lourdes && Frisco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   5340
         TabIndex        =   38
         Top             =   3570
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fris Nino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   5340
         TabIndex        =   37
         Top             =   3300
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kimberly Joy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   5340
         TabIndex        =   35
         Top             =   3000
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gzeus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   5280
         TabIndex        =   33
         Top             =   2700
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caleb Motor Corporation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   5340
         TabIndex        =   31
         Top             =   5100
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   5340
         TabIndex        =   30
         Top             =   5400
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robert Mosqueda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3210
         TabIndex        =   29
         Top             =   2130
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jay Ortiz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3210
         TabIndex        =   28
         Top             =   1830
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Pesimo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3210
         TabIndex        =   27
         Top             =   1530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Butch Aruta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3210
         TabIndex        =   26
         Top             =   1230
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bernard Tolosa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3210
         TabIndex        =   25
         Top             =   930
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jonathan Alsum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3210
         TabIndex        =   24
         Top             =   630
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ashish Piya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3210
         TabIndex        =   23
         Top             =   330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kim Lim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3210
         TabIndex        =   22
         Top             =   60
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Soriano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3210
         TabIndex        =   21
         Top             =   4530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leizel Abainza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   3210
         TabIndex        =   20
         Top             =   4230
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Bino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3210
         TabIndex        =   19
         Top             =   3930
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leah Perreras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3210
         TabIndex        =   18
         Top             =   3630
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amy Antonio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3210
         TabIndex        =   17
         Top             =   3330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levy Cruz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   3210
         TabIndex        =   16
         Top             =   3030
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3210
         TabIndex        =   15
         Top             =   2730
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3210
         TabIndex        =   14
         Top             =   2430
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cel Ayap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3210
         TabIndex        =   13
         Top             =   5130
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hazel Barrameda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3210
         TabIndex        =   12
         Top             =   4830
         Width           =   3945
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Friends && Contributors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   130
         TabIndex        =   10
         Top             =   4350
         Width           =   3945
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "We would like to thank our"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   130
         TabIndex        =   9
         Top             =   4080
         Width           =   3945
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Automotive Business Solutions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   2820
         Width           =   3945
      End
      Begin MSForms.Label Label4 
         Height          =   525
         Left            =   130
         TabIndex        =   7
         Top             =   2250
         Width           =   3945
         ForeColor       =   16711680
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Dealership Management Information System"
         Size            =   "6959;926"
         FontName        =   "Arial Black"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFEA&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   5355
      Width           =   4335
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1740
         Top             =   210
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   405
         Left            =   3000
         TabIndex        =   5
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Credits"
         Height          =   405
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   15000
         X2              =   -150
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":EF8A
      Height          =   645
      Left            =   150
      TabIndex        =   2
      Top             =   3690
      Width           =   4095
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   3240
      Width           =   1485
      ForeColor       =   8421504
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "version 2.0.7"
      Size            =   "2619;926"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Left            =   150
      TabIndex        =   0
      Top             =   2760
      Width           =   1485
      ForeColor       =   16711680
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DMIS"
      Size            =   "2619;926"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim okey                                               As String
Dim cnt                                                As Integer

Private Sub checkdup_Click()
'CheckDupMaster
Call UploadQty_Details
'Call CheckUploadQty_Details
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Credits" Then
        Picture2.Visible = True
        Picture3.Visible = True
        cnt = 0
        Command1.Caption = "<<-- About DMIS 2.0"
        Command1.Width = 2265
    Else
        Call Reset
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim matibayako As ADODB.Recordset
Set matibayako = New ADODB.Recordset

    Dim rsCustomer                                     As ADODB.Recordset
    Dim k                                              As Integer
    Dim NewCtlCde                                      As String

Dim kawnter As Integer
Dim new_customer_code As String
Set matibayako = gconDMIS.Execute("Select * from ALL_Vendor order by nameofvendor asc")
If Not matibayako.EOF And Not matibayako.BOF Then
    matibayako.MoveFirst: kawnter = 0
    Do While Not matibayako.EOF
       kawnter = kawnter + 1
       Me.Caption = "record number: " & kawnter
       If IsNumeric(Left(Null2String(matibayako!lastname), 1)) = True Then
          new_customer_code = GetCustomerZCode(Null2String(matibayako!lastname))
       Else
          new_customer_code = GetCustomerCode(Null2String(matibayako!lastname))
       End If
       gconDMIS.Execute ("Update all_customer set cuscde = '" & new_customer_code & "' where cuscde ='" & matibayako!CUSCDE & "'")
       Screen.MousePointer = 11
        gconDMIS.Execute "delete from ALL_CusCtl"
        For k = 65 To 90
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!CODE, 2, 5)) + 1, "00000")
                gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
            Else
                gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
            End If
        Next
        Screen.MousePointer = 0
       matibayako.MoveNext
       DoEvents
    Loop
End If
End Sub

Private Sub Command4_Click()
Dim matibayako As ADODB.Recordset
Set matibayako = New ADODB.Recordset

    Dim rsCustomer                                     As ADODB.Recordset
    Dim k                                              As Integer
    Dim NewCtlCde                                      As String

Dim kawnter As Integer
Dim new_customer_code As String
Set matibayako = gconDMIS.Execute("Select * from ALL_Vendor order by nameofvendor asc")
If Not matibayako.EOF And Not matibayako.BOF Then
    matibayako.MoveFirst: kawnter = 0
    Do While Not matibayako.EOF
       kawnter = kawnter + 1
       Me.Caption = "record number: " & kawnter
       If IsNumeric(Left(Null2String(matibayako!nameofvendor), 1)) = True Then
          new_customer_code = GetVendorZCode(Null2String(matibayako!nameofvendor))
       Else
          new_customer_code = GetVendorCode(Null2String(matibayako!nameofvendor))
       End If
       gconDMIS.Execute ("Update all_vendor set code = '" & new_customer_code & "' where code ='" & matibayako!CODE & "'")
       Screen.MousePointer = 11
        gconDMIS.Execute "delete from ALL_venCtl"
        For k = 65 To 90
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "select Code from ALL_vendor where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!CODE, 2, 5)) + 1, "00000")
                gconDMIS.Execute "insert into ALL_venCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Vendor control character for " & Chr(k) & " -')"
            Else
                gconDMIS.Execute "insert into ALL_venCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "',' Vendor control character for " & Chr(k) & " -')"
            End If
        Next
        Screen.MousePointer = 0
       matibayako.MoveNext
       DoEvents
    Loop
End If
End Sub

Function GetCustomerCode(lastname As String) As String
    Dim Temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = Left(Trim(lastname), 1)
    Set Temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (Temprs.EOF Or Temprs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(Temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerCode = Left(lastname, 1) & "00001"
    End If
End Function

Function GetCustomerZCode(lastname As String) As String
    Dim Temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = "Z"
    Set Temprs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (Temprs.EOF Or Temprs.BOF) Then
        GetCustomerZCode = lAlpha & Format(Mid(Temprs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerZCode = lAlpha & "00001"
    End If
End Function

Function GetVendorCode(lastname As String) As String
    Dim Temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = Left(Trim(lastname), 1)
    Set Temprs = gconDMIS.Execute("Select CTLCDE From ALL_VENCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (Temprs.EOF Or Temprs.BOF) Then
        GetVendorCode = Left(lastname, 1) & Format(Mid(Temprs.Collect(0), 2, 5), "00000")
    Else
        GetVendorCode = Left(lastname, 1) & "00001"
    End If
End Function

Function GetVendorZCode(lastname As String) As String
    Dim Temprs                                         As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                         As String
    lAlpha = "Z"
    Set Temprs = gconDMIS.Execute("Select CTLCDE From ALL_VENCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (Temprs.EOF Or Temprs.BOF) Then
        GetVendorZCode = lAlpha & Format(Mid(Temprs.Collect(0), 2, 5), "00000")
    Else
        GetVendorZCode = lAlpha & "00001"
    End If
End Function

Private Sub Command5_Click()
Dim rsTable3 As ADODB.Recordset
Dim rsPartMas As ADODB.Recordset
Dim rshari As ADODB.Recordset

Dim HARI_DNP As Double
Dim HARI_SRP As Double

Set rsTable3 = New ADODB.Recordset
Set rsTable3 = gconDMIS.Execute("Select * from table3 order by partno asc")
If Not rsTable3.EOF And Not rsTable3.BOF Then
   rsTable3.MoveFirst
   Do While Not rsTable3.EOF
      Set rsPartMas = New ADODB.Recordset
      Set rsPartMas = gconDMIS.Execute("Select * from pmis_partmas where partno = " & N2Str2Null(rsTable3!partNo))
      If Not rsPartMas.EOF And Not rsPartMas.BOF Then
         Set rshari = New ADODB.Recordset
         Set rshari = gconDMIS.Execute("Select * from PMIS_DNPP where partnumber = " & N2Str2Null(rsTable3!partNo))
         If Not rshari.EOF And Not rshari.BOF Then
            HARI_DNP = N2Str2Zero(rshari!DNPP)
            HARI_SRP = N2Str2Zero(rshari!SRP)
            gconDMIS.Execute ("Update PMIS_Partmas set NON_HARI = 'N', mac = 0, dnp = " & HARI_DNP & ", srp = " & HARI_SRP & ", location = " & N2Str2Null(rsTable3!Location) & ", partdesc = " & N2Str2Null(rsTable3!PARTDESC) & " where partno = " & N2Str2Null(rsTable3!partNo))
         Else
            HARI_DNP = 0
            HARI_SRP = 0
            gconDMIS.Execute ("Update PMIS_Partmas set NON_HARI = 'Y', mac = 0, dnp = " & HARI_DNP & ", srp = " & HARI_SRP & ", location = " & N2Str2Null(rsTable3!Location) & ", partdesc = " & N2Str2Null(rsTable3!PARTDESC) & " where partno = " & N2Str2Null(rsTable3!partNo))
         End If
      Else
         gconDMIS.Execute ("Insert into PMIS_Partmas (NON_HARI,type, partno, mac, dnp, srp, [location], partdesc) values ('Y','P'," & _
                            N2Str2Null(rsTable3!partNo) & ",0," & HARI_DNP & ", " & HARI_SRP & ", " & N2Str2Null(rsTable3!Location) & "," & N2Str2Null(rsTable3!PARTDESC) & ")")
      End If
      rsTable3.MoveNext
      Me.Caption = rsTable3!partNo
      DoEvents
   Loop
End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Picture2.Visible = False
    Picture3.Visible = False
    InitCredits
    cnt = 16
    Label1.Caption = MODULENAME
    Label2.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Icon = frmMain.Icon
End Sub

Private Sub Timer1_Timer()
    Dim kim                                            As Integer
    If cnt < 15 Then cnt = cnt + 1
    If cnt = 15 And Label8.Top > 1620 Then
        If Label4.Top > 0 Then Label4.Top = Label4.Top - 30
        If Label5.Top > 0 Then Label5.Top = Label5.Top - 30
        If Label6.Top > 0 Then Label6.Top = Label6.Top - 30
        If Label7.Top > 0 Then Label7.Top = Label7.Top - 30
        For kim = 0 To 36
            If labCredits(kim).Top > 0 Then labCredits(kim).Top = labCredits(kim).Top - 30
        Next
        If Label8.Top > 0 Then Label8.Top = Label8.Top - 30
        If Label9.Top > 0 Then Label9.Top = Label9.Top - 30
        If Label10.Top > 0 Then Label10.Top = Label10.Top - 30


        'If Label4.Top < 1320 And Label4.Top > 800 Then Label4.ForeColor = &H808080 Else Label4.ForeColor = vbBlack
        'If Label5.Top < 1320 And Label5.Top > 800 Then Label5.ForeColor = &H808080 Else Label5.ForeColor = vbBlack
        'If Label6.Top < 1320 And Label6.Top > 800 Then Label6.ForeColor = &H808080 Else Label6.ForeColor = vbBlack
        'If Label7.Top < 1320 And Label7.Top > 800 Then Label7.ForeColor = &H808080 Else Label7.ForeColor = vbBlack
        'For kim = 0 To 36
        '    If labCredits(kim).Top > 0 Then labCredits(kim).Top = labCredits(kim).Top - 30
        '    If labCredits(kim).Top < 1320 And labCredits(kim).Top > 8 Then labCredits(kim).ForeColor = &H808080 Else labCredits(kim).ForeColor = vbBlack
        'Next
        'If Label8.Top < 1320 And Label8.Top > 800 Then Label8.ForeColor = &H808080 Else Label8.ForeColor = vbBlack
        'If Label9.Top < 1320 And Label9.Top > 800 Then Label9.ForeColor = &H808080 Else Label9.ForeColor = vbBlack
        'If Label10.Top < 1320 And Label10.Top > 800 Then Label10.ForeColor = &H808080 Else Label10.ForeColor = vbBlack
    End If
End Sub

Sub InitCredits()
    Dim kim                                            As Integer
    Label4.Top = 2250
    Label5.Top = 2850
    Label6.Top = 4080
    Label7.Top = 4350
    labCredits(0).Left = 130
    labCredits(0).Top = 5500
    Label8.Left = 130
    Label9.Left = 130
    Label10.Left = 130
    For kim = 1 To 36
        labCredits(kim).Left = 130
        labCredits(kim).Top = labCredits(kim - 1).Top + 250
    Next
    Label8.Top = 1620 + labCredits(36).Top + 2400
    Label9.Top = 2550 + labCredits(36).Top + 2400
    Label10.Top = 3870 + labCredits(36).Top + 2400
End Sub

Sub Reset()
    Command1.Caption = "Credits"
    Command1.Width = 1245
    Picture2.Visible = False
    Picture3.Visible = False
    cnt = 16: InitCredits
End Sub

Sub CheckDupMaster()
Dim rsmaterials As ADODB.Recordset
Dim rsParts As ADODB.Recordset
Dim rsStocks As ADODB.Recordset

Set rsmaterials = New ADODB.Recordset
Set rsmaterials = gconDMIS.Execute("Select * from table5 order by a asc")
If Not rsmaterials.EOF And Not rsmaterials.BOF Then
   rsmaterials.MoveFirst
   Screen.MousePointer = 11
   Do While Not rsmaterials.EOF
        Set rsStocks = New ADODB.Recordset
        Set rsStocks = gconDMIS.Execute("Select * from pmis_stockmas where stockno = " & N2Str2Null(rsmaterials!A))
        If Not rsStocks.EOF And Not rsStocks.BOF Then
           gconDMIS.Execute ("Update pmis_stockmas set type = 'M' where id = " & rsStocks!ID)
        Else
           MsgBox rsmaterials!A
        End If
        Me.Caption = rsmaterials!A
        DoEvents
      rsmaterials.MoveNext
   Loop
   Screen.MousePointer = 0
End If
End Sub

Sub UploadQty_Details()
Dim rsMats As ADODB.Recordset
Dim rsStocks As ADODB.Recordset

Dim leybel As Double

Dim partNo, partno_wspace As String

Set rsMats = New ADODB.Recordset
'Set rsMats = gconDMIS.Execute("Select * from lubs_mats_parts order by stockno asc")
Set rsMats = gconDMIS.Execute("Select * from lubs order by stockno asc")
If Not rsMats.EOF And Not rsMats.BOF Then
   rsMats.MoveFirst: leybel = 0
   Screen.MousePointer = 11
   Do While Not rsMats.EOF
      partNo = rsMats!stockno
      partno_wspace = Left(rsMats!stockno, 5) & " " & Right(rsMats!stockno, Len(rsMats!stockno) - 5)
      'MsgBox partNo & vbCrLf & partno_wspace
      Set rsStocks = New ADODB.Recordset
      'Set rsStocks = gconDMIS.Execute("Select * from pmis_stockmas where (stockno = " & N2Str2Null(rsMats!stockno) & ") or (stockno = '" & Left(Null2String(rsMats!stockno), 5) & " " & Right(Null2String(rsMats!stockno), Len(Null2String(rsMats!stockno)) - 5) & "')")
      Set rsStocks = gconDMIS.Execute("Select * from pmis_stockmas where stockno = " & N2Str2Null(rsMats!stockno))
      If Not rsStocks.EOF And Not rsStocks.BOF Then
         gconDMIS.Execute ("Update pmis_stockmas set onhand = " & N2Str2Zero(rsMats!ONHAND) & "," & _
                          " mac = " & N2Str2Zero(rsMats!Mac) & "," & _
                          " dnp = " & N2Str2Zero(rsMats!DNP) & "," & _
                          " srp = " & N2Str2Zero(rsMats!SRP) & "," & _
                          " lastm_mac = " & N2Str2Zero(rsMats!Mac) & "," & _
                          " lastm_oh = " & N2Str2Zero(rsMats!ONHAND) & _
                          " where stockno = " & N2Str2Null(rsStocks!stockno))
         gconDMIS.Execute ("update lubs_mats_parts set status = 'Uploaded' where stockno = '" & rsMats!stockno & "'")
      Else
         'gconDMIS.Execute ("update lubs_mats_parts set status = 'Not Loaded' where stockno = '" & rsMats!stockno & "'")
         gconDMIS.Execute ("Insert into PMIS_Partmas (NON_HARI,type, partno, mac, dnp, srp, [location], partdesc,onhand) values ('Y','M'," & _
                            N2Str2Null(rsMats!stockno) & "," & N2Str2Zero(rsMats!Mac) & "," & N2Str2Zero(rsMats!DNP) & ", " & N2Str2Zero(rsMats!SRP) & ", ''," & N2Str2Null(rsMats!stockdesc) & "," & rsMats!ONHAND & ")")
      End If
      Me.Caption = rsMats!stockno
      leybel = leybel + rsMats!ONHAND
      Label11.Caption = "onhand = " & leybel
      DoEvents
      rsMats.MoveNext
   Loop
   Screen.MousePointer = 0
End If
End Sub

Sub CheckUploadQty_Details()
Dim rsMats As ADODB.Recordset
Dim rsStocks As ADODB.Recordset

Dim leybel As Double

Dim partNo, partno_wspace As String

Set rsMats = New ADODB.Recordset
Set rsMats = gconDMIS.Execute("Select * from lubs_mats_parts order by stockno asc")
If Not rsMats.EOF And Not rsMats.BOF Then
   rsMats.MoveFirst: leybel = 0
   Screen.MousePointer = 11
   Do While Not rsMats.EOF
      partNo = rsMats!stockno
      partno_wspace = Left(rsMats!stockno, 5) & " " & Right(rsMats!stockno, Len(rsMats!stockno) - 5)
      'MsgBox partNo & vbCrLf & partno_wspace
      Set rsStocks = New ADODB.Recordset
      Set rsStocks = gconDMIS.Execute("Select * from pmis_stockmas where onhand <> " & rsMats!ONHAND & " and (stockno = " & N2Str2Null(rsMats!stockno) & ") or (stockno = '" & Left(Null2String(rsMats!stockno), 5) & " " & Right(Null2String(rsMats!stockno), Len(Null2String(rsMats!stockno)) - 5) & "')")
      If Not rsStocks.EOF And Not rsStocks.BOF Then
         'gconDMIS.Execute ("Update pmis_stockmas set onhand = " & N2Str2Zero(rsMats!ONHAND) & "," & _
                          " mac = " & N2Str2Zero(rsMats!Mac) & "," & _
                          " dnp = " & N2Str2Zero(rsMats!DNP) & "," & _
                          " srp = " & N2Str2Zero(rsMats!SRP) & "," & _
                          " lastm_mac = " & N2Str2Zero(rsMats!Mac) & "," & _
                          " lastm_oh = " & N2Str2Zero(rsMats!ONHAND) & _
                          " where (stockno = " & N2Str2Null(rsMats!stockno) & ") or (stockno = '" & Left(Null2String(rsMats!stockno), 5) & " " & Right(Null2String(rsMats!stockno), Len(Null2String(rsMats!stockno)) - 5) & "')")
         gconDMIS.Execute ("update lubs_mats_parts set status = 'Uploaded' where stockno = '" & rsMats!stockno & "'")
      Else
         'gconDMIS.Execute ("update lubs_mats_parts set status = 'Not Loaded' where stockno = '" & rsMats!stockno & "'")
         'gconDMIS.Execute ("Insert into PMIS_Partmas (NON_HARI,type, partno, mac, dnp, srp, [location], partdesc,onhand) values ('Y','P'," & _
                            N2Str2Null(rsMats!stockno) & "," & N2Str2Zero(rsMats!Mac) & "," & N2Str2Zero(rsMats!DNP) & ", " & N2Str2Zero(rsMats!SRP) & ", ''," & N2Str2Null(rsMats!stockdesc) & "," & rsMats!ONHAND & ")")
      End If
      Me.Caption = rsMats!stockno
      leybel = leybel + rsMats!ONHAND
      Label11.Caption = "onhand = " & leybel
      DoEvents
      rsMats.MoveNext
   Loop
   Screen.MousePointer = 0
End If
End Sub


