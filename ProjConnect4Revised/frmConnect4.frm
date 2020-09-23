VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConnect4 
   Caption         =   "C o n n e c t   4      --------------------      b y   P e t e   O a k e y"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   Icon            =   "frmConnect4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tlbrOptions 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1138
      ButtonWidth     =   4180
      ButtonHeight    =   979
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New Game"
            Key             =   ""
            Description     =   "Begin a new game"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   " Change Direction Of Play "
            Key             =   ""
            Description     =   "Play horizontally or vertically"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Show Priority Grid"
            Key             =   ""
            Description     =   "Show how the computer thinks"
            Object.Tag             =   ""
            Style           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Show Hint"
            Key             =   ""
            Description     =   "See what the computer would do"
            Object.Tag             =   ""
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit Game"
            Key             =   ""
            Description     =   "Bye bye"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSound 
      Height          =   645
      Left            =   105
      TabIndex        =   54
      Top             =   6615
      Width           =   2955
      Begin VB.CheckBox chkSound 
         Caption         =   "Sound On"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   55
         Top             =   210
         Width           =   2640
      End
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Hard"
      Height          =   330
      Index           =   3
      Left            =   210
      TabIndex        =   49
      Top             =   3150
      Width           =   1800
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Medium"
      Height          =   330
      Index           =   2
      Left            =   210
      TabIndex        =   48
      Top             =   2835
      Width           =   1800
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Easy"
      Height          =   330
      Index           =   1
      Left            =   210
      TabIndex        =   47
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Timer tmrDropPiece 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2310
      Top             =   2835
   End
   Begin VB.Frame fraPlayers 
      Height          =   645
      Left            =   105
      TabIndex        =   45
      Top             =   7350
      Width           =   2955
      Begin VB.CheckBox chkPlayers 
         Caption         =   "Player vs Computer"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   46
         Top             =   210
         Width           =   2640
      End
   End
   Begin VB.Frame fraPriority 
      Caption         =   "Priority Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   105
      TabIndex        =   2
      Top             =   3885
      Width           =   2745
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   41
         Left            =   2100
         TabIndex        =   44
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   40
         Left            =   2100
         TabIndex        =   43
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   39
         Left            =   2100
         TabIndex        =   42
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   38
         Left            =   2100
         TabIndex        =   41
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   37
         Left            =   2100
         TabIndex        =   40
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   36
         Left            =   2100
         TabIndex        =   39
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   35
         Left            =   1785
         TabIndex        =   38
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   34
         Left            =   1785
         TabIndex        =   37
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   33
         Left            =   1785
         TabIndex        =   36
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   32
         Left            =   1785
         TabIndex        =   35
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   31
         Left            =   1785
         TabIndex        =   34
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   30
         Left            =   1785
         TabIndex        =   33
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   29
         Left            =   1470
         TabIndex        =   32
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   28
         Left            =   1470
         TabIndex        =   31
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   27
         Left            =   1470
         TabIndex        =   30
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   26
         Left            =   1470
         TabIndex        =   29
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   25
         Left            =   1470
         TabIndex        =   28
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   24
         Left            =   1470
         TabIndex        =   27
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   23
         Left            =   1155
         TabIndex        =   26
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   22
         Left            =   1155
         TabIndex        =   25
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   21
         Left            =   1155
         TabIndex        =   24
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   20
         Left            =   1155
         TabIndex        =   23
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   19
         Left            =   1155
         TabIndex        =   22
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   18
         Left            =   1155
         TabIndex        =   21
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   17
         Left            =   840
         TabIndex        =   20
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   16
         Left            =   840
         TabIndex        =   19
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   15
         Left            =   840
         TabIndex        =   18
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   14
         Left            =   840
         TabIndex        =   17
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   13
         Left            =   840
         TabIndex        =   16
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   12
         Left            =   840
         TabIndex        =   15
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   11
         Left            =   525
         TabIndex        =   14
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   10
         Left            =   525
         TabIndex        =   13
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   9
         Left            =   525
         TabIndex        =   12
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   8
         Left            =   525
         TabIndex        =   11
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   7
         Left            =   525
         TabIndex        =   10
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   525
         TabIndex        =   9
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   210
         TabIndex        =   7
         Top             =   735
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label lblPriority 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   1995
         Width           =   330
      End
   End
   Begin VB.Frame fraBoard 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7365
      Left            =   3255
      TabIndex        =   0
      Top             =   735
      Width           =   8520
      Begin VB.Shape shHint 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1170
         Left            =   0
         Top             =   0
         Width           =   120
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   41
         Left            =   7140
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   40
         Left            =   7140
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   39
         Left            =   7140
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   38
         Left            =   7140
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   37
         Left            =   7140
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   36
         Left            =   7140
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   35
         Left            =   5985
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   34
         Left            =   5985
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   33
         Left            =   5985
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   32
         Left            =   5985
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   31
         Left            =   5985
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   30
         Left            =   5985
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   29
         Left            =   4830
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   28
         Left            =   4830
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   27
         Left            =   4830
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   26
         Left            =   4830
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   25
         Left            =   4830
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   24
         Left            =   4830
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   23
         Left            =   3675
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   22
         Left            =   3675
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   21
         Left            =   3675
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   20
         Left            =   3675
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   19
         Left            =   3675
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   18
         Left            =   3675
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   17
         Left            =   2520
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   16
         Left            =   2520
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   15
         Left            =   2520
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   14
         Left            =   2520
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   13
         Left            =   2520
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   12
         Left            =   2520
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   11
         Left            =   1365
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   10
         Left            =   1365
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   9
         Left            =   1365
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   8
         Left            =   1365
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   7
         Left            =   1365
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   6
         Left            =   1365
         Top             =   5985
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   5
         Left            =   210
         Top             =   210
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   4
         Left            =   210
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   3
         Left            =   210
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   2
         Left            =   210
         Top             =   3675
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   1
         Left            =   210
         Top             =   4830
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   1185
         Index           =   0
         Left            =   210
         Top             =   5985
         Width           =   1185
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   2100
      X2              =   2940
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   2100
      X2              =   2940
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   105
      X2              =   1995
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   105
      X2              =   1995
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Label lblP1ScoreTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   53
      Top             =   945
      Width           =   1905
   End
   Begin VB.Label lblP2ScoreTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   52
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label lblP1Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   51
      Top             =   945
      Width           =   855
   End
   Begin VB.Label lblP2Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   50
      Top             =   1365
      Width           =   855
   End
   Begin ComctlLib.ImageList imglstC4 
      Left            =   2205
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   79
      ImageHeight     =   79
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConnect4.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConnect4.frx":672E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmConnect4.frx":B192
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConnect4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ensure ALL variables are declared.

'Will hold the red piece, and the yellow piece.
Dim C(1) As StdPicture
Dim B As StdPicture

'Will hold 0 (red) or 1 (yellow).
Dim colour As Byte

'Will hold True if it's the computer's turn to drop next.
Dim computerTurnNext As Boolean

'Will hold the current position of the board that is being checked.
Dim pos As Integer

'Will hold the positions (indexes) in the img() control array (the board) of
'where the piece should drop from, and to.
Dim topOfDrop As Byte
Dim bottomOfDrop As Byte

'Each position in this array represents the respective position in the
'img() control array, and will hold a value which the AI will manipulate
'to determine what the best drop is.
Dim MovePriority(41) As Integer

'Will hold the furthest position in the respective direction, depending on
'the board position.  For example, if the board position (or index) is anything
'between 18 and 23 then pTop will hold 23.  If the board position is anything
'between 1 and 37 (step 6) then pLeft will hold 1, and so on.
Dim pTop As Byte     'Top.
Dim pBottom As Byte  'Bottom.
Dim pLeft As Byte    'Left.
Dim pRight As Byte   'Right.
Dim pTL As Byte      'Top left.
Dim pTR As Byte      'Top right.
Dim pBL As Byte      'Bottom left.
Dim pBR As Byte      'Bottom right.

'Will hold the number of pieces either side of the position being checked.
Dim row1 As Byte
Dim row2 As Byte

'Will prevent the AI from dropping once a game
'has finished.
Dim moveMade As Boolean

'Will hold number of drops made.
Dim dropsMade As Byte

Private Sub chkPlayers_Click()

If chkPlayers.Value = 0 Then
  lblP1ScoreTitle.Caption = "Player 1"
  lblP2ScoreTitle.Caption = "Player 2"
Else
  lblP1ScoreTitle.Caption = "Player"
  lblP2ScoreTitle.Caption = "Computer"
End If

End Sub

Private Sub Form_Load()

Set C(0) = imglstC4.ListImages(2).Picture
Set C(1) = imglstC4.ListImages(3).Picture
Set B = imglstC4.ListImages(1).Picture

Dim i As Integer
  For i = 0 To img.Count - 1
    img(i).Picture = B
    lblPriority(i).Caption = ""
    lblPriority(i).BackColor = &HE0E0E0
  Next i

If Val(lblP1Score.Caption) + Val(lblP2Score.Caption) = 0 Then
  chkPlayers.Value = 1
  chkSound.Value = 1
End If

Call PlaySound(Int(Rnd * 2) + 2)

colour = 0
dropsMade = 0

optDifficulty(3).Value = True

tlbrOptions.Buttons(4).Value = tbrUnpressed
tlbrOptions.Buttons(3).Value = tbrUnpressed

fraPriority.Visible = False

Call HideHint

If moveMade = False Then computerTurnNext = True

End Sub

Private Sub DefinePositions(ByVal pos As Integer)

'Top position.
pTop = pos - pos Mod 6 + 5
'Bottom position.
pBottom = pTop - 5
'Left position.
pLeft = pos Mod 6
'Right position.
pRight = pLeft + 36
'Top left position.
If pos < 36 And pBottom / 6 <= 5 - pLeft Then pTL = pos - ((pBottom / 6) * 5) Else pTL = pos - ((5 - pLeft) * 5)
'Bottom left position.
If pLeft <= (pos - pLeft) / 6 Then pBL = pos - (pLeft * 7) Else pBL = pos - (((pos - pLeft) / 6) * 7)
'Top right position.
If pLeft >= (pos - pLeft) / 6 Then pTR = pos + (35 - (7 * pLeft)) Else pTR = pos + ((6 - pBottom / 6) * 7)
'Bottom right position.
Dim i As Byte
 For i = 0 To 25 Step 5
  If pos > 35 - i Or (pos - (i / 5)) Mod 6 = 0 Then
   pBR = pos + i
   Exit For
  End If
 Next i

End Sub

Private Sub img_Click(Index As Integer)

'Increment variable.
dropsMade = dropsMade + 1

'Initialize variabes.
topOfDrop = Index - (Index Mod 6) + 5 'Top of chosen column.

'Do nothing except display an error message if the column chosen is full.
If img(topOfDrop).Picture <> B Then
 MsgBox "The column you have selected is full.  Please choose another one.", _
 vbOKOnly, "A Piece Cannot Be Dropped"
 Exit Sub
End If

'Initialize variabes.
For pos = Index - (Index Mod 6) To topOfDrop
 If img(pos).Picture = B Then
  bottomOfDrop = pos 'Bottom of chosen column.
  Exit For
 End If
Next pos

If colour = 0 Then lblPriority(bottomOfDrop).BackColor = vbRed Else _
lblPriority(bottomOfDrop).BackColor = vbYellow

'Drop a piece.
tmrDropPiece.Enabled = True

End Sub

Private Sub tlbrOptions_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button
  Case "New Game"
    Call Form_Load
  Case " Change Direction Of Play "
    Dim i As Integer
      If img(0).Top > img(1).Top Then
        For i = 0 To 41
          img(i).Move img(i).Left, (i Mod 6) * 1155 + 210
          lblPriority(i).Move lblPriority(i).Left, (i Mod 6) * 315 + 420
        Next i
      Else
        For i = 0 To 41
          img(i).Move img(i).Left, 5985 - (i Mod 6) * 1155
          lblPriority(i).Move lblPriority(i).Left, 1995 - (i Mod 6) * 315
        Next i
      End If
  Case "Show Priority Grid"
    If Button.Value = 1 Then fraPriority.Visible = True Else fraPriority.Visible = False
  Case "Show Hint"
    If Button.Value = 1 Then
      Call GetPos
    Else
      Call HideHint
    End If
  Case "Exit Game"
    Call CloseForm(Me)
End Select

End Sub

Private Sub HideHint()

With shHint
  .Visible = False
  .Top = 0
  .Left = 0
  .Width = 0
End With
      
End Sub

Private Sub GetPos()

Dim highNo As Integer
Dim checkNo As Integer
Dim imgIndex As Integer

For pos = 5 To 41 Step 6
  If img(pos).Picture <> B Then highNo = highNo + 1 Else checkNo = (pos - 5) / 6 + 1
Next pos

If highNo = 6 Then
  MsgBox "A piece can only be dropped in column " & checkNo _
  & ".  All other columns are full.", vbOKOnly, "User Stupidity Alert"
Else
  highNo = 0
  checkNo = 0
End If

For pos = 0 To 41
  MovePriority(pos) = 0
Next pos

Call AICheckForAWinOrBlock
Call AICheckForStupidMoves
Call UpdatePriorityGrid

  For pos = 0 To 41
     If lblPriority(pos).BackColor = &HE0E0E0 Then
        If lblPriority(pos).Caption <> "" Then checkNo = Val(lblPriority(pos).Caption)
        If checkNo > 79 And checkNo < 90 Then
          imgIndex = pos
          Exit For
        End If
        If Abs(checkNo) > highNo Then
          highNo = Abs(checkNo)
          imgIndex = pos
        End If
    End If
  Next pos

With shHint
  .Move img(imgIndex).Left, img(imgIndex).Top, img(imgIndex).Width, img(imgIndex).Height
  .Visible = True
End With

End Sub

Private Sub Win()

moveMade = True

Dim message As String

If colour = 0 Then
  PlaySound (Int(Rnd * 3) + 4)
  message = "Well done Player 1, you won this game, the scores have been updated."
  lblP1Score.Caption = Val(lblP1Score.Caption) + 1
ElseIf chkPlayers.Value = 1 Then
  PlaySound (Int(Rnd * 5) + 7)
  message = "Hahahahahaha I won this game, the scores have been updated."
  lblP2Score.Caption = Val(lblP2Score.Caption) + 1
Else
   PlaySound (Int(Rnd * 3) + 4)
  message = "Well done player 2, you won this game, the scores have been updated."
  lblP2Score.Caption = Val(lblP2Score.Caption) + 1
End If

MsgBox message, vbOKOnly, "A Winning Line Has Been Detected."

computerTurnNext = False

Call Form_Load

End Sub
Private Sub CheckForWin()

'Declare local variables.
Dim i As Integer
Dim j As Integer

'Initialize variables.
row1 = 0
row2 = 0

'Check all vertical lines.
For i = 0 To 36 Step 6
 For j = i To i + 5
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
  If row1 > 3 Then Call Win
 Next j
 row1 = 0
Next i

'Check all Horizontal lines.
For i = 0 To 5
 For j = i To i + 36 Step 6
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
  If row1 > 3 Then Call Win
 Next j
 row1 = 0
Next i

'Check the left to right diagonal lines.
For i = 3 To 5
 For j = i To i * 6 Step 5
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
   If row1 >= 4 Then
    Call Win
    End If
 Next j
 row1 = 0
Next i

For i = 11 To 23 Step 6
 For j = i To (((i + 1) / 6) - 1) + 35 Step 5
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
   If row1 >= 4 Then
    Call Win
    End If
 Next j
 row1 = 0
Next i

'Diagonal right to left.
For i = 6 To 18 Step 6
 For j = i To Abs(i / 6 - 42) Step 7
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
   If row1 >= 4 Then
    Call Win
    End If
 Next j
 row1 = 0
Next i

For i = 0 To 2
 For j = i To Abs((i - 7) * 5 - i) Step 7
  If img(j).Picture = C(colour) Then row1 = row1 + 1 Else row1 = 0
   If row1 >= 4 Then
    Call Win
    End If
 Next j
 row1 = 0
Next i

'Change colour of piece to drop.
If moveMade = False Then
  If colour = 0 Then colour = 1 Else colour = 0
Else
  moveMade = False
End If

If chkPlayers.Value = 1 Then
  If computerTurnNext = True Then Call AI Else computerTurnNext = True
End If

End Sub

Private Sub tmrDropPiece_Timer()

'Declare local variable.
Dim i As Integer

'Drop a piece.
If topOfDrop >= bottomOfDrop Then
  img(topOfDrop).Picture = C(colour)
  If topOfDrop Mod 6 <> 5 Then img(topOfDrop + 1).Picture = B
  If topOfDrop <> 0 Then topOfDrop = topOfDrop - 1 Else GoTo EndDrop
Else
'End the drop by disabling the Timer, and then check for a win.
EndDrop:
  tmrDropPiece.Enabled = False
  
  Call CheckForWin

End If

End Sub

Private Sub RandomDrop()

Do
  pos = Int(Rnd * 41)
  If img(pos).Picture = B Then Exit Do
Loop

Call img_Click(pos)
    
End Sub

Private Sub AI()

'Declare local variable.
Dim i As Integer

'The computer's having it's turn, so set this variable to False.
computerTurnNext = False

'Reset the MovePriority() array.
For i = 0 To 41
  MovePriority(i) = 1
Next i

'Get the difficulty level, and store it in i.
For i = 1 To 3
  If optDifficulty(i).Value = True Then Exit For
Next i

Select Case i
  'Easy. (Random drop)
  Case 1
    Call RandomDrop
  'Medium. (Looks for a win or block)
  Case 2
    Call AICheckForAWinOrBlock
  'Hard. (Looks for a win or a block, and won't play stupid)
  Case 3
    Call AICheckForAWinOrBlock
    Call AICheckForStupidMoves
End Select

Call UpdatePriorityGrid
Call PlayBestMove

End Sub

Private Sub AICheckForStupidMoves()

For pos = 0 To 41
  If pos Mod 6 <> 5 Then
    If pos Mod 6 = 0 Then GoTo PositionOK
    If img(pos - 1).Picture <> B Then
PositionOK:
      If img(pos).Picture = B Then
        pos = pos + 1
        Call DefinePositions(pos)
        Call StupidMoveChecks(0)
        Call StupidMoveChecks(1)
        pos = pos - 1
      End If
    End If
  End If
Next pos

End Sub

Private Sub UpdatePriorityGrid()

For pos = 0 To 41
  If img(pos).Picture = B Then
    If pos Mod 6 = 0 Then
      lblPriority(pos).Caption = MovePriority(pos)
    ElseIf img(pos - 1).Picture <> B Then
      lblPriority(pos).Caption = MovePriority(pos)
    Else
      lblPriority(pos).Caption = ""
    End If
  Else
    lblPriority(pos).Caption = ""
  End If
Next pos

End Sub

Private Sub StupidMoveChecks(whatColour As Integer)

Dim i As Integer

'Horizontal.
If pos <> pLeft Then
 For i = pLeft To pos - 6 Step 6
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pRight Then
 For i = pos + 6 To pRight Step 6
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If MovePriority(pos - 1) < 70 Then MovePriority(pos - 1) = 0 - Abs(whatColour - 1)
End If
row1 = 0
row2 = 0

'Diagonal top left to bottom right.
If pos <> pTL Then
 For i = pTL To pos - 5 Step 5
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pBR Then
 For i = pos + 5 To pBR Step 5
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If MovePriority(pos - 1) < 70 Then MovePriority(pos - 1) = 0 - Abs(whatColour - 1)
End If
row1 = 0
row2 = 0

'Diagonal bottom left to top right.
If pos <> pBL Then
 For i = pBL To pos - 7 Step 7
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pTR Then
 For i = pos + 7 To pTR Step 7
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If MovePriority(pos - 1) < 70 Then MovePriority(pos - 1) = 0 - Abs(whatColour - 1)
End If
row1 = 0
row2 = 0

End Sub

Private Sub PlayBestMove()

Randomize
Dim i As Integer
Dim randomChoice As Integer
Dim chosenPos As Integer
pos = 0
For i = 0 To 41
  If Val(lblPriority(i).Caption) > pos Then
    pos = Val(lblPriority(i).Caption)
    chosenPos = i
  ElseIf Val(lblPriority(i).Caption) = pos Then
    randomChoice = Int(Rnd * 2) ' randomChoice = 0 or 1.
    If randomChoice = 0 Then chosenPos = i
  End If
Next i

If img(chosenPos - (chosenPos Mod 6) + 5).Picture <> B Then
  Call RandomDrop
  Exit Sub
End If

If dropsMade = 1 Then
  chosenPos = Int(Rnd * 7) * 6
ElseIf dropsMade = 3 Then
  For i = 12 To 24 Step 6
    If img(i).Picture = B Then
      chosenPos = i
      Exit For
    End If
  Next i
End If

Call img_Click(chosenPos)

End Sub

Private Sub AICheckForAWinOrBlock()

Dim whatColour As Integer

For whatColour = 0 To 1
  'Loop through all positions on the board, and check for a win or block
  '(depending on whatColour) for each possible, valid drop.
  For pos = 0 To 41
   If pos Mod 6 = 0 Then GoTo PositionOK  '1. If bottom row.
   If img(pos - 1).Picture <> B Then      '2. Or if the position below
PositionOK:                               '   is not blank.
    If img(pos).Picture = B Then          '3. And the position is blank.
     'Initialize variables.
     Call DefinePositions(pos)            '4. Then it's a potential drop so
     'Check for a possible win.               set the variables using the
     Call AIWin(Abs(whatColour - 1))                        '   position's index, and work through
    End If                                 '  the AIWin routine.
   End If
  Next pos
Next whatColour

End Sub

Private Sub AIWin(whatColour As Byte)

'Declare local variable.
Dim i As Byte

'Set row1 and row2 to zero.
row1 = 0
row2 = 0

'Vertical.
If pos Mod 6 <> 0 Then
 For i = pBottom To pos - 1
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
 If row1 = 3 Then
  If whatColour = 1 Then
   MovePriority(pos) = 99
  Else
   MovePriority(pos) = 89
  End If
 ElseIf MovePriority(pos) < 80 Then
  MovePriority(pos) = MovePriority(pos) + whatColour + 1 + row1
 End If
End If
row1 = 0

'Horizontal.
If pos <> pLeft Then
 For i = pLeft To pos - 6 Step 6
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pRight Then
 For i = pos + 6 To pRight Step 6
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If whatColour = 1 Then
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 98
  End If
 Else
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 88
  End If
 End If
ElseIf MovePriority(pos) < 80 Then
 MovePriority(pos) = MovePriority(pos) + whatColour + 1 + row1 + row2
End If
row1 = 0
row2 = 0

'Diagonal top left to bottom right.
If pos <> pTL Then
 For i = pTL To pos - 5 Step 5
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pBR Then
 For i = pos + 5 To pBR Step 5
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If whatColour = 1 Then
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 97
  End If
 Else
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 87
  End If
 End If
ElseIf MovePriority(pos) < 80 Then
 MovePriority(pos) = MovePriority(pos) + whatColour + 1 + row1 + row2
End If
row1 = 0
row2 = 0

'Diagonal bottom left to top right.
If pos <> pBL Then
 For i = pBL To pos - 7 Step 7
  If img(i).Picture = C(whatColour) Then row1 = row1 + 1 Else row1 = 0
 Next i
End If
If pos <> pTR Then
 For i = pos + 7 To pTR Step 7
  If img(i).Picture = C(whatColour) Then row2 = row2 + 1 Else Exit For
 Next i
End If
If row1 + row2 > 2 Then
 If whatColour = 1 Then
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 97
  End If
 Else
  If MovePriority(pos) < 90 Then
   MovePriority(pos) = 87
  End If
 End If
ElseIf MovePriority(pos) < 80 Then
 MovePriority(pos) = MovePriority(pos) + whatColour + 1 + row1 + row2
End If
row1 = 0
row2 = 0

End Sub

