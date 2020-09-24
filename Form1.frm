VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   12630
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   14925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12630
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   1560
      Top             =   6720
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   4200
      TabIndex        =   42
      ToolTipText     =   "Right click any where on the window to get the game menu."
      Top             =   8160
      Width           =   5055
      Begin VB.HScrollBar HS 
         Height          =   255
         Left            =   2640
         Max             =   50
         Min             =   5
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   480
         Value           =   10
         Width           =   1815
      End
      Begin VB.HScrollBar HS2 
         Height          =   255
         Left            =   840
         Max             =   3
         Min             =   1
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1560
         Value           =   3
         Width           =   1815
      End
      Begin VB.PictureBox T 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   0
         Left            =   960
         MousePointer    =   99  'Custom
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   43
         ToolTipText     =   "Right click any where on the window to get the game menu."
         Top             =   120
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   48
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   6720
   End
   Begin VB.PictureBox MS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   1560
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   41
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox MS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0614
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox MS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":091E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox MS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   1080
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0C28
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   7200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox MS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0F32
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3333
      Left            =   0
      Picture         =   "Form1.frx":123C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   35
      Top             =   2040
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   0
      Left            =   4320
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3030
      Left            =   0
      Picture         =   "Form1.frx":1653
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   303
      Left            =   0
      Picture         =   "Form1.frx":1A3C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3300
      Left            =   960
      Picture         =   "Form1.frx":1E40
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   330
      Left            =   960
      Picture         =   "Form1.frx":224B
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3003
      Left            =   960
      Picture         =   "Form1.frx":2655
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   29
      Top             =   3000
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   33
      Left            =   960
      Picture         =   "Form1.frx":2A54
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3033
      Left            =   2040
      Picture         =   "Form1.frx":2E52
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3303
      Left            =   2040
      Picture         =   "Form1.frx":324D
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   333
      Left            =   2040
      Picture         =   "Form1.frx":3667
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3330
      Left            =   2040
      Picture         =   "Form1.frx":3A7F
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   30
      Left            =   3120
      Picture         =   "Form1.frx":3E83
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3
      Left            =   3120
      Picture         =   "Form1.frx":426F
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   300
      Left            =   3120
      Picture         =   "Form1.frx":4658
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   3000
      Left            =   3120
      Picture         =   "Form1.frx":4A4B
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   20
      Left            =   4200
      Picture         =   "Form1.frx":4E36
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   2
      Left            =   4200
      Picture         =   "Form1.frx":7110
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   200
      Left            =   4200
      Picture         =   "Form1.frx":93EA
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   2000
      Left            =   4200
      Picture         =   "Form1.frx":B6C4
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   10
      Left            =   5160
      Picture         =   "Form1.frx":D99E
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   15
      Top             =   4320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1
      Left            =   5160
      Picture         =   "Form1.frx":FC78
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   100
      Left            =   5160
      Picture         =   "Form1.frx":11F52
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1000
      Left            =   5160
      Picture         =   "Form1.frx":1422C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   101
      Left            =   5280
      Picture         =   "Form1.frx":16506
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1011
      Left            =   6120
      Picture         =   "Form1.frx":168DF
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1101
      Left            =   6120
      Picture         =   "Form1.frx":16C98
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   111
      Left            =   6120
      Picture         =   "Form1.frx":17082
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   11
      Left            =   7080
      Picture         =   "Form1.frx":1746D
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1001
      Left            =   7080
      Picture         =   "Form1.frx":1782C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   110
      Left            =   7080
      Picture         =   "Form1.frx":17BEB
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1010
      Left            =   5280
      Picture         =   "Form1.frx":17FBA
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1111
      Left            =   4320
      Picture         =   "Form1.frx":18358
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1110
      Left            =   6120
      Picture         =   "Form1.frx":18738
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox S 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Index           =   1100
      Left            =   7080
      Picture         =   "Form1.frx":18AFF
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   -100
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press ESC to exit the game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4440
      TabIndex        =   49
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please, wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3000
      TabIndex        =   36
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnNew 
         Caption         =   "New Round"
      End
      Begin VB.Menu mnReset 
         Caption         =   "Reset Round"
      End
      Begin VB.Menu mnSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnSol 
         Caption         =   "Solve"
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Dim BoardSize  As Integer
Const ClickMargin = 0.35
Const Pi = 3.14159265359
Const MaxGameSize = 18
' **********  Difficulty
'Const DiffBulk = 100
Dim Dicision As Long, Margine As Long, DifficultyMargin As Integer
Dim DiffLimits() As Single
'*****************


Dim Board() As String, Flag As Boolean, ReadyToClick As Boolean, Size1, Size2
Dim Temp() As String, Stack() As Boolean, GameLevel As String, ElecLoc As Integer
Dim Game() As String, IsGameLoaded As Boolean, RequiredTicks As Integer, SolveOrder() As Integer
Dim LampStatus() As Integer, Timing As Integer, LastTile As Integer, IsESCPressed As Boolean
Dim SolutionDisplayed As Boolean, IsSolutionWorking As Boolean, UnloadForm As Boolean

'Const EasyLimit = 0.345
'Const HardLimit = 0.1
'
'Const Level = 18




Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINT
    X As Long
    Y As Long
End Type

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)

Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)

Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)

Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Sub Bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta!)

' bmp_rotate(pic1, pic2, theta)
' Rotate the image in a picture box.
'   pic1 is the picture box with the bitmap to rotate
'   pic2 is the picture box to receive the rotated bitmap
'   theta is the angle of rotation

Dim c1x As Integer, c1y As Integer
Dim c2x As Integer, c2y As Integer
Dim a As Single
Dim p1x As Integer, p1y As Integer
Dim p2x As Integer, p2y As Integer
Dim n As Integer, r   As Integer

c1x = pic1.ScaleWidth \ 2
c1y = pic1.ScaleHeight \ 2
c2x = pic2.ScaleWidth \ 2
c2y = pic2.ScaleHeight \ 2

If c2x < c2y Then n = c2y Else n = c2x
n = n - 1
pic1hDC = pic1.hdc
pic2hDC = pic2.hdc

For p2x = 0 To n
     For p2y = 0 To n
        If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
        r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
        p1x = r * Cos(a + theta!)
        p1y = r * Sin(a + theta!)
        c0& = GetPixel(pic1hDC, c1x + p1x, c1y + p1y)
        c1& = GetPixel(pic1hDC, c1x - p1x, c1y - p1y)
        c2& = GetPixel(pic1hDC, c1x + p1y, c1y - p1x)
        c3& = GetPixel(pic1hDC, c1x - p1y, c1y + p1x)
        If c0& <> -1 Then xret& = SetPixel(pic2hDC, c2x + p2x, c2y + p2y, c0&)
        If c1& <> -1 Then xret& = SetPixel(pic2hDC, c2x - p2x, c2y - p2y, c1&)
        If c2& <> -1 Then xret& = SetPixel(pic2hDC, c2x + p2y, c2y - p2x, c2&)
        If c3& <> -1 Then xret& = SetPixel(pic2hDC, c2x - p2y, c2y + p2x, c3&)
     Next
    ' t% = DoEvents()
Next


End Sub


Public Sub ConstructBoard()

ReDim LampStatus(1 To 2, 1 To 1)

'inintiate the board

For X% = 1 To BoardSize%
  For Y% = 1 To BoardSize%
     Board(X, Y) = "0000"
  Next Y
Next X





'************  Locating a random location on the board to be a start point ************

Randomize
X = Int(Rnd * BoardSize + 1)
'X = 4
Randomize
Y = Int(Rnd * BoardSize + 1)
'Y = 4




ExitFromTile X * 100 + Y







End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' To move the window

Select Case KeyCode

Case 37
   If Form1.left > Form1.ScaleWidth / -2 Then Form1.left = Form1.left - 100
Case 38
   If Form1.top > Form1.ScaleHeight / -2 Then Form1.top = Form1.top - 100
Case 39
   If Form1.left < Screen.Width - Form1.ScaleWidth / 2 Then Form1.left = Form1.left + 100
Case 40
   If Form1.top < Screen.Height - Form1.ScaleHeight / 2 Then Form1.top = Form1.top + 100
Case 27
  If IsSolutionWorking Or Not IsGameLoaded Then IsESCPressed = True

End Select


'


End Sub

'
Private Sub Form_Load()



LastTile = 0
If Not UnloadForm Then mnAbout_Click

' to get the largest size of the window corresponding the current screen setting

Form1.Height = 2 * Screen.Height
Form1.Width = 2 * Screen.Width

Size1 = Screen.Width
Size2 = Screen.Height

' to get the max size fot the game

If Form1.Height < Form1.Width Then
   HS.Max = (Form1.Height \ T(0).Height) - 2
Else
   HS.Max = (Form1.Width \ T(0).Width) - 2
End If

If HS.Max > MaxGameSize Then HS.Max = MaxGameSize


'' the intial size will be Max/2 or default of the scroll
'If HS.Min <= Int(HS.Max / 2) + 1 Then
'   HS.Value = Int(HS.Max / 2) + 1
'Else
'   HS.Value = HS.Min
'End If

Randomize
HS.Value = HS.Min + Int(Rnd * (HS.Max - HS.Min + 1))


HS2.Value = 1 + Int(Rnd * 3)
LoadDiffLimits


'initiate a new game
NewGame






End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 And IsGameLoaded Then PopupMenu mnMenu



End Sub


Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And IsGameLoaded Then PopupMenu mnMenu


End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

T(LastTile).BorderStyle = 0


End Sub

Private Sub HS_Change()

'BoardSize = HS.Value
Label3.Caption = "Size is (" & HS.Value & " x " & HS.Value & ")"

End Sub

Private Sub HS_Scroll()

HS_Change


End Sub


Private Sub HS2_Change()




Select Case HS2.Value

Case 3
    DifficultyMargin = 2
    GameLevel = "many"
Case 2
    DifficultyMargin = 25
    GameLevel = "average"
Case 1
    DifficultyMargin = 50
    GameLevel = "low"
End Select





Label5.Caption = "No. of Lamps is " & GameLevel





End Sub

Private Sub HS2_Scroll()

HS2_Change


End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 And IsGameLoaded Then PopupMenu mnMenu


End Sub


Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 And IsGameLoaded Then PopupMenu mnMenu


End Sub


Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


T(LastTile).BorderStyle = 0


End Sub


Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And IsGameLoaded Then PopupMenu mnMenu


End Sub


Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And IsGameLoaded Then PopupMenu mnMenu


End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

T(LastTile).BorderStyle = 0


End Sub


Private Sub mnAbout_Click()

'If IsGameLoaded Then
T(LastTile).BorderStyle = 0

MsgBox "ELECTRO Maze is a very exciting game. You should re-construct the electric network to turn on all the lamps " + _
       "in minimum number of mouse ticks. You will notice that all the tiles on the board are rotatable. The idea of this " + _
       "game is based on a flash game called ConFuseBox. You can find it on http://www.ybbored.com ." + vbCrLf + vbCrLf + _
       "The electric network is built randomly to be hard to see the same network twice. To play the game, just move " + _
       "the mouse on a tile to get it focused. When you tick it, it will be rotated according to the direction of the appearing " + _
       "arrow when the mouse is nearby any corncer of the tile. When the mouse is on the middle region of the tile, a question " + _
       "mark appears and no rotation can be done." + vbCrLf + vbCrLf + _
       "If required, press any arrow key to move the window in the corresponding direction. Also, it should be noted that " + _
       "any change in the level or the size of the game will affect the next new round." + vbCrLf + vbCrLf + _
       "This game is designed by Wael Sayed. To send comments or suggestions, use the address wael_eng@hotmail.com.", vbOKOnly + vbInformation, "About ELECTRO Maze..."



End Sub

Private Sub mnExit_Click()

If IsGameLoaded Then T(LastTile).BorderStyle = 0

If MsgBox("Do you want to exit the game ?   ", vbYesNo + vbQuestion, "Exit Game...") = 6 Then End




End Sub

Private Sub mnNew_Click()



T(LastTile).BorderStyle = 0


If MsgBox("This will start a new round and the current one will be removed . " _
+ vbCrLf + "Do you want to proceed ?   ", vbYesNo + vbQuestion, "New Round...") = 6 Then
   
   UnloadTiles
   NewGame
End If



End Sub

Private Sub mnReset_Click()

If SolutionDisplayed Then Exit Sub



T(LastTile).BorderStyle = 0
If MsgBox("This will reset this round  . " _
+ vbCrLf + "Do you want to proceed ?   ", vbYesNo + vbQuestion, "Reset Round...") = 6 Then
      
     Frame1.Visible = False
      
      For X = 1 To BoardSize
        For Y = 1 To BoardSize%
           Board(X, Y) = Game(X, Y)
           T(X * 100 + Y).Picture = S(Int(Board(X, Y))).Picture
        Next Y
      Next X
  
      Frame1.Visible = True

End If



End Sub

Private Sub mnSol_Click()


If SolutionDisplayed Then Exit Sub

T(LastTile).BorderStyle = 0

If MsgBox("If the solution is displayed, you have to start a new round  . " _
+ vbCrLf + "Do you want to proceed ?   ", vbYesNo + vbQuestion, "Solution...") = 6 Then
    
    
    SolutionDisplayed = True
    IsSolutionWorking = True
    
    T(LastTile).BorderStyle = 0
    mnReset.Enabled = False

    
 '********************* Mouse Lock ***********************
    Dim client As RECT
    Dim upperleft As POINT

    'Get information about our wndow

    GetClientRect Me.hWnd, client
    upperleft.X = client.left
    upperleft.Y = client.top

    'Convert window coî–dinates to screen coî–dinates

    ClientToScreen Me.hWnd, upperleft

    'move our rectangle

    OffsetRect client, upperleft.X, upperleft.Y

    'limit the cursor movement

    client.top = client.bottom
    ClipCursor client


WWW: SS = ShowCursor(0)
    If SS >= 0 Then GoTo WWW
    
 '********************* Mouse Lock ***********************
    
    'IsESCPressed = False
    
    ReDim SolveOrder(1 To BoardSize ^ 2)
'ttt:
    For X% = 1 To BoardSize
       For Y% = 1 To BoardSize%
           SolveOrder((X - 1) * BoardSize + Y) = 100 * X + Y
       Next Y%
    Next X%


   For X% = 1 To (BoardSize ^ 2) / 2
      Randomize
      a% = Int(Rnd * (BoardSize ^ 2) + 1)
      Randomize
      b% = Int(Rnd * (BoardSize ^ 2) + 1)
      c% = SolveOrder(a)
      SolveOrder(a) = SolveOrder(b)
      SolveOrder(b) = c
   Next X


   For z% = 1 To (BoardSize ^ 2)
          X = SolveOrder(z) \ 100
          Y = SolveOrder(z) Mod 100
          ReadyToClick = True
          Randomize
          Flag = Int(Rnd * 2)
          
ddd:    DoEvents
          If PressESC Then
             For X% = 1 To BoardSize
               For Y% = 1 To BoardSize%
                  Board(X, Y) = Temp(X, Y)
                  T(X * 100 + Y) = S(Board(X, Y))
               Next Y%
             Next X%
             GoTo rrr
          End If
          
          
          If IsLamp(X%, Y%) Then
              If Temp(X%, Y%) <> Board(X%, Y%) Then
                 If Len(CStr(Int(Temp(X%, Y%)))) = Len(CStr(Int(Board(X%, Y%)))) Then
                    ToggleSingleLamp X * 100 + Y, 0
                 Else
                    T_MouseDown X% * 100 + Y%, 1, 0, 0, 0 '2, T(0).Width * (1 - 1 ^ TempX) - 2
                    GoTo ddd
                 End If
              
              End If
          Else
              If Temp(X%, Y%) <> Board(X%, Y%) Then
                T_MouseDown X% * 100 + Y%, 1, 0, 0, 0 '2, (T(0).Width - 2) * (1 - 1 ^ TempX)
                GoTo ddd
              End If
          End If

    Next z
            
rrr:
            
      DoEvents

     Sleep 700
     For X = 1 To 2
        ResetLamps
        DoEvents
        Sleep 100
        CheckElec
        ToggleLamps
        DoEvents
        Sleep 200
      Next X
     
          
          
     'Releases the cursor limits

    ClipCursor ByVal 0&
    
    HS.Visible = True
    HS2.Visible = True
    Label3.Visible = True
    Label5.Visible = True
    mnNew.Enabled = True

    
zzz: SS = ShowCursor(1)
    If SS < 0 Then GoTo zzz
    IsSolutionWorking = False
          
          
          'T(x * 100 + y) = S(Int(Temp(x, y)))

   ' SolutionDisplayed = True





End If

End Sub

Private Sub T_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And IsGameLoaded Then
  T(LastTile).BorderStyle = 0
  PopupMenu mnMenu
  Exit Sub
End If


If Not ReadyToClick Then Exit Sub
'If IsSolutionWorking Then Exit Sub


If Button = 1 Then

    ResetLamps
    
        'If IsLamp(Index \ 100, Index Mod 100) = 2 Then Stop: ToggleSingleLamp Index, 0  'To turn off the lamp as soon as it is ticked

    If Not IsSolutionWorking Then
        TempTile$ = Board(Index \ 100, Index Mod 100)
        Board(Index \ 100, Index Mod 100) = "0000"
        CheckElec
        Board(Index \ 100, Index Mod 100) = TempTile$
        ToggleLamps
        DoEvents
    End If
    
    If IsLamp(Index \ 100, Index Mod 100) = 2 Then ToggleSingleLamp Index, 0  'To turn off the lamp as soon as it is ticked

    If Not Flag Then
      Board(Index \ 100, Index Mod 100) = Rot_AntiClKWise((Board(Index \ 100, Index Mod 100)))
      Dirx% = 1
    Else
      Board(Index \ 100, Index Mod 100) = Rot_ClKWise((Board(Index \ 100, Index Mod 100)))
      Dirx% = -1
    End If
    
    CheckElec
ggg:
    Rot_Pic CInt(Board(Index \ 100, Index Mod 100)), Index, Dirx%
    ToggleLamps

    If SolutionDisplayed Then Exit Sub
    Label1.Caption = Label1.Caption + 1
    T(LastTile).BorderStyle = 0
    
    If Label1.Caption / RequiredTicks > 0.75 Then mnReset.Enabled = True
    
    
    If CountLitLamps Then
'        If MsgBox("You finished this round in " + CStr(Label1.Caption) + " ticks while the minimum required number of ticks is " + CStr(RequiredTicks) + "  . " + vbCrLf + _
'                  "Do you want to start a new one ?   " + vbCrLf + _
'                  "If you press No, you will Exit the game .  ", vbYesNo + vbQuestion, "New round...") = 6 Then
        
        If Label1.Caption = RequiredTicks Then
             msg$ = "Congratulations, you accoplished the round without any extra ticks . "
        Else
             msg$ = "You finished this round in " + CStr(Label1.Caption) + " ticks while the minimum ticks required is " + CStr(RequiredTicks) + " . "
        End If
        
        ret% = MsgBox(msg + vbCrLf + _
                      "Do you want to modify the settings of the new round ?" + vbCrLf _
                  , vbQuestion + vbYesNo, "New round...")
        If ret = 6 Then
           SolutionDisplayed = True
           HS.Visible = True
           HS2.Visible = True
           Label3.Visible = True
           Label5.Visible = True
           mnNew.Enabled = True
        ElseIf ret = 7 Then
           UnloadTiles
           NewGame
'        Else
'           SolutionDisplayed = True
        End If
    End If
 
End If

End Sub



Public Function Rot_AntiClKWise(Inp As String) As String


Rot_AntiClKWise = CStr(Inp)

'If Len(Rot_AntiClKWise) < 4 Then
'  For X% = 1 To 4 - Len(Rot_AntiClKWise)
'     Rot_AntiClKWise = "0" & Rot_AntiClKWise
'  Next X
'
'End If
 
 
Rot_AntiClKWise = right(Rot_AntiClKWise, 3) & left(Rot_AntiClKWise, 1)


End Function
Public Function Rot_ClKWise(Inp As String) As String




Rot_ClKWise = CStr(Inp)

'If Len(Rot_ClKWise) < 4 Then
'  For X% = 1 To 4 - Len(Rot_ClKWise)
'     Rot_ClKWise = "0" & Rot_ClKWise
'  Next X
'
'End If
 
 
Rot_ClKWise = right(Rot_ClKWise, 1) & left(Rot_ClKWise, 3)


End Function



Public Function RightSide(Loc As Integer, Inp As Integer) As Boolean




RightSide = False

If Inp = 1 Then
  Board(Loc \ 100, Loc Mod 100) = left(Board(Loc \ 100, Loc Mod 100), 2) + CStr(Inp) + right(Board(Loc \ 100, Loc Mod 100), 1)
Else
    X% = Loc \ 100
    Y% = Loc Mod 100
    If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then RightSide = False: Exit Function
    RightSide = Mid(Board(Loc \ 100, Loc Mod 100), 3, 1)
End If

'T(Loc).Visible = True
'T(Loc).Picture = S(CInt(Board(Loc \ 100, Loc mod 100))).Picture


End Function


Public Function UpSide(Loc As Integer, Inp As Integer) As Boolean




UpSide = False

If Inp = 1 Then
    Board(Loc \ 100, Loc Mod 100) = left(Board(Loc \ 100, Loc Mod 100), 1) + CStr(Inp) + right(Board(Loc \ 100, Loc Mod 100), 2)
Else
    X% = Loc \ 100
    Y% = Loc Mod 100
    If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then UpSide = False: Exit Function
    UpSide = Mid(Board(Loc \ 100, Loc Mod 100), 2, 1)
End If


'T(Loc).Visible = True
'T(Loc).Picture = S(CInt(Board(Loc \ 100, Loc mod 100))).Picture


End Function

Public Function DownSide(Loc As Integer, Inp As Integer) As Boolean


DownSide = False

If Inp = 1 Then
    Board(Loc \ 100, Loc Mod 100) = left(Board(Loc \ 100, Loc Mod 100), 3) + CStr(Inp)
Else
    X% = Loc \ 100
    Y% = Loc Mod 100
    If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then DownSide = False: Exit Function
    DownSide = Mid(Board(Loc \ 100, Loc Mod 100), 4, 1)
End If





'T(Loc).Visible = True
'
'T(Loc).Picture = S(CInt(Board(Loc \ 100, Loc mod 100))).Picture

End Function





Public Function LeftSide(Loc As Integer, Inp As Integer) As Boolean






LeftSide = False

If Inp = 1 Then
    Board(Loc \ 100, Loc Mod 100) = CStr(Inp) + right(Board(Loc \ 100, Loc Mod 100), 3)
Else
    X% = Loc \ 100
    Y% = Loc Mod 100
    If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then LeftSide = False: Exit Function
    LeftSide = Mid(Board(Loc \ 100, Loc Mod 100), 1, 1)
End If


'T(Loc).Visible = True
'
'T(Loc).Picture = S(CInt(Board(Loc \ 100, Loc mod 100))).Picture


End Function

Public Sub PutLamp(Loc As Integer, Inp As Integer)

Board(Loc \ 100, Loc Mod 100) = ""
For X% = 1 To 4
   If X = Inp Then
       Board(Loc \ 100, Loc Mod 100) = Board(Loc \ 100, Loc Mod 100) & "1"
   Else
       Board(Loc \ 100, Loc Mod 100) = Board(Loc \ 100, Loc Mod 100) & "0"
   End If
Next X










'T(Loc).Picture = S(CInt(Board(Loc \ 100, Loc mod 100))).Picture

End Sub

Public Sub ExitFromTile(Loc As Integer)


'T(Loc).BorderStyle = 1

Randomize
Dicision = Int(Rnd * 50 + 100)
Randomize
Margine = Int(Rnd * Dicision / DifficultyMargin + 1)



X% = Loc \ 100
Y% = Loc Mod 100


Direction$ = GetDirections(Loc)



For z% = 1 To Len(Direction)
     Dirc$ = Mid(Direction, z, 1)
     Randomize
     SS% = Int(Rnd * (Dicision + Margine) + 1)

         If Dirc = 1 Then  'left
            If CInt(Board((X - 1), Y)) = 0 Then
                If SS < Dicision Then
                   LeftSide X * 100 + Y, 1
                   RightSide (X - 1) * 100 + Y, 1
                   ExitFromTile (X - 1) * 100 + Y
                ElseIf SS >= Dicision Then
                   LeftSide X * 100 + Y, 1
                   PutLamp (X - 1) * 100 + Y, 3
                End If
            End If
         ElseIf Dirc = 2 Then  'up
            If CInt(Board(X, Y - 1)) = 0 Then
                If SS < Dicision Then
                   UpSide X * 100 + Y, 1
                   DownSide X * 100 + Y - 1, 1
                   ExitFromTile X * 100 + Y - 1
                ElseIf SS >= Dicision Then
                   UpSide X * 100 + Y, 1
                   PutLamp X * 100 + Y - 1, 4
                End If
            End If
         ElseIf Dirc = 3 Then  'right
            If CInt(Board(X + 1, Y)) = 0 Then
                If SS < Dicision Then
                   RightSide X * 100 + Y, 1
                   LeftSide (X + 1) * 100 + Y, 1
                   ExitFromTile (X + 1) * 100 + Y
                ElseIf SS >= Dicision Then
                   RightSide X * 100 + Y, 1
                   PutLamp (X + 1) * 100 + Y, 1
                End If
            End If
         Else   'down
            If CInt(Board(X, Y + 1)) = 0 Then
                If SS < Dicision Then
                   DownSide X * 100 + Y, 1
                   UpSide X * 100 + Y + 1, 1
                   ExitFromTile X * 100 + Y + 1
                ElseIf SS >= Dicision Then
                   DownSide X * 100 + Y, 1
                   PutLamp X * 100 + Y + 1, 2
                End If
            End If
         End If



Next z




End Sub

Public Function IsLamp(X As Integer, Y As Integer) As Byte



If InStr(1, Board(X, Y), "3") > 0 Then
  Exit Function
End If



For z% = 1 To 4
  If Mid(Board(X, Y), z, 1) = "0" Then
    c% = c% + 1
  Else
    LampSt% = Mid(Board(X, Y), z, 1)
  End If
Next z

If c% = 3 Then IsLamp = LampSt%




End Function

Public Function IsElec(X As Integer, Y As Integer) As Boolean


'X% = Loc \ 100
'Y% = Loc mod 100


If InStr(1, Board(X, Y), "3") > 0 Then IsElec = True




End Function


Public Function IsBoardFull() As Boolean

'Check if the board is full of tiles


IsBoardFull = True

For X% = 1 To BoardSize
  For Y% = 1 To BoardSize
     If Board(X, Y) = "0000" Then IsBoardFull = False: Exit Function
  Next Y
Next X



fff:

'select randome location to put the ELEC source
Randomize
X% = Int(Rnd * BoardSize + 1)
Randomize
Y% = Int(Rnd * BoardSize + 1)
Board(X, Y) = CStr(CInt(Board(X, Y)) * 3)
FormatAdapt X, Y

ElecLoc = X * 100 + Y
  


End Function

Public Sub CheckElec()


'check if there is a path between every Lamp and the ELEC source



For X% = 1 To UBound(LampStatus, 2)
     'If X = 13 Then Stop
      TrackLamp LampStatus(1, X), X, LampDirection(LampStatus(1, X))
Next X




End Sub

Public Sub TrackLamp(Loc As Integer, LampNo As Integer, Direction As Integer)


'debug.Assert
'Debug.Print "Lamp:", Loc, LampNo, Direction


X% = Loc \ 100

Y% = Loc Mod 100

If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then Exit Sub


HomeDir% = ReverseDir(Direction)


If Direction = 1 Then
   If RightSide((X - 1) * 100 + Y, 0) Then TrackWire (X - 1) * 100 + Y, LampNo, HomeDir%
   
ElseIf Direction = 2 Then
   If DownSide((X) * 100 + Y - 1, 0) Then TrackWire (X) * 100 + Y - 1, LampNo, HomeDir%
   
ElseIf Direction = 3 Then
   If LeftSide((X + 1) * 100 + Y, 0) Then TrackWire (X + 1) * 100 + Y, LampNo, HomeDir%
   
ElseIf Direction = 4 Then
   If UpSide((X) * 100 + Y + 1, 0) Then TrackWire (X) * 100 + Y + 1, LampNo, HomeDir%
   
End If




End Sub
Public Sub TrackWire(Loc As Integer, LampNo As Integer, HomeDir As Integer)


'Debug.Print "  Wire:", Loc, LampNo, HomeDir






X% = Loc \ 100
Y% = Loc Mod 100

If Stack(X, Y) = 0 Then
  Stack(X, Y) = 1
Else
  'Debug.Print "loop", Loc
  'Stop
  GoTo eee
End If



If X < 1 Or X > BoardSize Or Y < 1 Or Y > BoardSize Then GoTo eee

If IsLamp(X, Y) Then GoTo eee
If IsElec(X, Y) Then LampStatus(2, LampNo) = 1: GoTo eee


Direction$ = GetWireDir(Loc)


For z% = 1 To Len(Direction$)
    Di% = CInt(Mid(Direction, z, 1))
    If Di = HomeDir Then GoTo sss
    If Di = 1 Then
       If RightSide((X - 1) * 100 + Y, 0) Then
         If IsElec(X - 1, Y) Then
            LampStatus(2, LampNo) = 1
            'Debug.Print "  Elec:", Loc
            GoTo eee
         Else
            TrackWire (X - 1) * 100 + Y, LampNo, ReverseDir(Di)
         End If
       End If
    ElseIf Di = 2 Then
       If DownSide(X * 100 + Y - 1, 0) Then
         If IsElec(X, Y - 1) Then
            LampStatus(2, LampNo) = 1
            'Debug.Print "  Elec:", Loc
            GoTo eee
         Else
            TrackWire X * 100 + Y - 1, LampNo, ReverseDir(Di)
         End If
       End If
    ElseIf Di = 3 Then
       If LeftSide((X + 1) * 100 + Y, 0) Then
         If IsElec(X + 1, Y) Then
            LampStatus(2, LampNo) = 1
            'Debug.Print "  Elec:", Loc
            GoTo eee
         Else
            TrackWire (X + 1) * 100 + Y, LampNo, ReverseDir(Di)
         End If
       End If
    ElseIf Di = 4 Then
       If UpSide(X * 100 + Y + 1, 0) Then
         If IsElec(X, Y + 1) Then
            LampStatus(2, LampNo) = 1
            'Debug.Print "  Elec:", Loc
            GoTo eee
         Else
            TrackWire X * 100 + Y + 1, LampNo, ReverseDir(Di)
         End If
       End If
    End If
sss:
    
Next z


eee:
Stack(X, Y) = 0

End Sub


Public Function GetDirections(Loc As Integer) As String


X% = Loc \ 100
Y% = Loc Mod 100


GetDirections = ""

If X% > 1 And X <= BoardSize Then  'left
   Randomize
   D% = Int(Rnd * 2 + 1)
   If D = 1 Then
     GetDirections = GetDirections & "1"
   Else
     GetDirections = "1" & GetDirections
   End If
End If

If X% < BoardSize And X >= 1 Then  'right
   Randomize
   D% = Int(Rnd * 2 + 1)
   If D = 1 Then
     GetDirections = GetDirections & "3"
   Else
     GetDirections = "3" & GetDirections
   End If
End If


If Y% < BoardSize And Y >= 1 Then 'down
   Randomize
   D% = Int(Rnd * 2 + 1)
   If D = 1 Then
     GetDirections = GetDirections & "4"
   Else
     GetDirections = "4" & GetDirections
   End If
End If


If Y% > 1 And Y <= BoardSize Then  'up
   Randomize
   D% = Int(Rnd * 2 + 1)
   If D = 1 Then
     GetDirections = GetDirections & "2"
   Else
     GetDirections = "2" & GetDirections
   End If
End If




End Function

Public Function ReverseDir(Direction As Integer) As Integer


If Direction = 1 Then
  ReverseDir = 3
ElseIf Direction = 2 Then
  ReverseDir = 4
ElseIf Direction = 3 Then
  ReverseDir = 1
ElseIf Direction = 4 Then
  ReverseDir = 2
End If




End Function

Public Sub CollectLampLoc()

'collect the LAMPS location and initiate their status to OFF


For X% = 1 To BoardSize
  For Y% = 1 To BoardSize
     If IsLamp(X, Y) Then
        UboundX% = UBound(LampStatus, 2)
        LampStatus(1, UboundX%) = X * 100 + Y
        LampStatus(2, UboundX%) = 0
        ReDim Preserve LampStatus(1 To 2, 1 To UboundX% + 1)

     End If
  Next Y
Next X



ReDim Preserve LampStatus(1 To 2, 1 To UBound(LampStatus, 2) - 1)


End Sub

Public Function GetWireDir(Loc As Integer) As String



X% = Loc \ 100

Y% = Loc Mod 100

GetWireDir = ""


For z% = 1 To 4
   If Mid(Board(X, Y), z, 1) = "1" Then GetWireDir = GetWireDir + CStr(z)
Next z








End Function

Public Sub ToggleLamps()


For X% = 1 To UBound(LampStatus, 2)
  'If X = 13 Then Stop
  ToggleSingleLamp LampStatus(1, X), LampStatus(2, X)
Next X




End Sub

Public Sub ToggleSingleLamp(Loc As Integer, Status As Integer)


X% = Loc \ 100
Y% = Loc Mod 100

If Status = 0 Then
      If InStr(1, Board(X, Y), "2") Then Board(X, Y) = CStr(CInt(Board(X, Y)) / 2)
Else
     If InStr(1, Board(X, Y), "1") Then Board(X, Y) = CStr(CInt(Board(X, Y)) * 2)
End If
  
FormatAdapt X, Y


If IsGameLoaded Then T(Loc).Picture = S(Board(Loc \ 100, Loc Mod 100)).Picture ': T(Loc).Refresh




End Sub

Public Sub ResetLamps()

For X% = 1 To UBound(LampStatus, 2)
 LampStatus(2, X) = 0
 ToggleSingleLamp LampStatus(1, X), 0
Next X




End Sub

Public Sub Scramble()


For X% = 1 To BoardSize
  For Y% = 1 To BoardSize
     
     If Board(X, Y) <> "1111" And Board(X, Y) <> "3333" Then   'no need to rotate the crossed tiles
        Randomize
        z% = Int(Rnd * 4)
        RequiredTicks = RequiredTicks + z
        
        If z = 3 Then
            RequiredTicks = RequiredTicks - 2 ' 3 rotations equal to 1 in the reverse direction
        ElseIf z = 2 Then
            If (Board(X, Y) = "1010" Or Board(X, Y) = "0101" Or Board(X, Y) = "3030" Or Board(X, Y) = "0303") Then _
                 RequiredTicks = RequiredTicks - 2  ' 2 rotations for the ROD tiles is nothing
        End If
        
        For Tt% = 1 To z
          Board(X, Y) = Rot_ClKWise((Board(X, Y)))
        Next Tt%
     
      '  Debug.Print xxx, Board(x, y), z, RequiredTicks

     End If
     
  Next Y
Next X


End Sub

Public Sub NewGame()


Screen.MousePointer = 11
Timer2.Enabled = False
IsESCPressed = False
IsGameLoaded = False
RequiredTicks = 0
Frame1.Visible = False


BoardSize = HS.Value
'HS2.Value = 1


 mnNew.Enabled = False
 mnReset.Enabled = False

Form1.Width = T(0).Width * BoardSize
Refresh
Form1.Height = T(0).Height * (BoardSize + 1)
Refresh
Form1.left = (Screen.Width - Form1.Width) / 2
Refresh
Form1.top = (Screen.Height - Form1.Height) / 2
Refresh
Frame1.top = 0
Frame1.left = 0
Frame1.Width = Form1.Width
Frame1.Height = Form1.Height
Label4.top = (Form1.Height - Label4.Height) / 2
Label4.left = (Form1.Width - Label4.Width) / 2
Label2.left = (Form1.Width - Label2.Width) / 2
Label2.top = 0 ' (Label4.Height + Label4.top)

Label1.top = T(0).Height * BoardSize
Label1.left = (Form1.Width - Label1.Width) / 2
HS.top = Form1.Height - HS.Height * 1.1
HS2.top = Form1.Height - HS.Height * 1.1
HS.left = (Form1.Width - HS.Width)
HS2.left = 0
HS.Visible = False
HS2.Visible = False

Label3.Caption = "Size is (" & BoardSize & " x " & BoardSize & ")"
Label3.top = HS.top - Label3.Height
Label3.left = HS.left + (HS.Width - Label3.Width) / 2
Label3.Visible = False

'Label5.Caption = "Level is " & GameLevel
Label5.top = HS2.top - Label5.Height
Label5.left = HS2.left + (HS2.Width - Label5.Width) / 2
Label5.Visible = False



Form1.Show

ReDim Board(1 To BoardSize, 1 To BoardSize) As String
ReDim Temp(1 To BoardSize, 1 To BoardSize) As String
ReDim Stack(1 To BoardSize, 1 To BoardSize) As Boolean
ReDim Game(1 To BoardSize, 1 To BoardSize) As String
Dim DiffRatio As Single

'Dirx$ = App.Path '************
'If Right(Dirx$, 1) <> "\" Then Dirx$ = Dirx$ + "\" '************



'Open Dirx$ + "file.csv" For Output As 1 '************
'Print #1, CStr(BoardSize ^ 2) '************

'Counter% = 0  '************
'Stop
'Debug.Print "*****************************"

'lows! = 1 '************
bbb:
DoEvents
PressESC

ConstructBoard

If Not IsBoardFull Then GoTo bbb
  
CollectLampLoc


'Counter = Counter + 1 '************

If UBound(LampStatus, 2) = 1 Then GoTo bbb


DiffRatio = UBound(LampStatus, 2) / (BoardSize ^ 2)

'If DiffRatio > Highs! Then Highs = DiffRatio '************
'If DiffRatio < lows! Then lows = DiffRatio '************
'Debug.Print lows, DiffRatio, Highs '************


'If HS2.Value = 1 Then
'  If DiffRatio < EasyLimit Then
'    GoTo bbb
'  End If
'ElseIf HS2.Value = 2 Then
'  If Not (DiffRatio <= EasyLimit And DiffRatio >= HardLimit) Then
'      GoTo bbb
'  End If
'ElseIf HS2.Value = 3 Then
'  If DiffRatio > HardLimit Then
'      GoTo bbb
'  End If
'
'End If

'HS2.Value = 3
If HS2.Value = 3 Then
  If DiffRatio < DiffLimits(BoardSize, 2) Then
    GoTo bbb
  End If
ElseIf HS2.Value = 2 Then
 ' If Not (DiffRatio <= DiffLimits(BoardSize, 2) And DiffRatio >= (DiffLimits(BoardSize, 1)) + (DiffLimits(BoardSize, 2) - (DiffLimits(BoardSize, 1)) / 4)) Then
  If Not (DiffRatio <= DiffLimits(BoardSize, 2) And DiffRatio >= DiffLimits(BoardSize, 1)) Then
      
      GoTo bbb
  End If
ElseIf HS2.Value = 1 Then
  If DiffRatio > DiffLimits(BoardSize, 1) Then
      GoTo bbb
  End If

End If





'Print #1, CStr(UBound(LampStatus, 2)) ' + "," + CStr(BoardSize ^ 2) '************
  
'If Counter < 10000 Then GoTo bbb Else Close #1  ': Stop '************



CheckElec
ToggleLamps

' to save the solution in a temp array

For X = 1 To BoardSize
  For Y = 1 To BoardSize%
      Temp(X, Y) = Board(X, Y)
  Next Y
Next X

Scramble
ResetLamps




CheckElec
ToggleLamps

For X = 1 To BoardSize
  For Y = 1 To BoardSize%
      Load T(X * 100 + Y)
      T(X * 100 + Y).left = T(X * 100 + Y).Width * (X - 1)
      T(X * 100 + Y).top = T(X * 100 + Y).Height * (Y - 1)
      T(X * 100 + Y).Picture = S(CInt(Board(X, Y))).Picture
      T(X * 100 + Y).Visible = True
      Game(X, Y) = Board(X, Y)
  Next Y
Next X

IsGameLoaded = True

' Load the tiles and locate them



Frame1.Visible = True

Label1.Caption = 0

Screen.MousePointer = 0
Timer2.Enabled = True

'MsgBox UBound(LampStatus, 2)

End Sub

Public Function CountLitLamps() As Boolean

'check if ALL Lamps are ON

CountLitLamps = True


For X% = 1 To UBound(LampStatus, 2)
    If LampStatus(2, X) = 0 Then CountLitLamps = False: Exit Function
Next X



End Function

Private Sub T_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


' Change the mouse icon

'If IsSolutionWorking Then Exit Sub


ReadyToClick = False

If X > 0 And X < T(Index).ScaleWidth * ClickMargin And Y > 0 And Y < T(Index).ScaleHeight * ClickMargin Then 'up-left  anti
    T(Index).MouseIcon = MS(3).Picture
    Flag = False
    ReadyToClick = True
ElseIf X > T(Index).ScaleWidth * (1 - ClickMargin) And X < T(Index).ScaleWidth And Y > 0 And Y < T(Index).ScaleHeight * ClickMargin Then 'up-right clock
    T(Index).MouseIcon = MS(0).Picture
    Flag = True
    ReadyToClick = True
ElseIf X > 0 And X < T(Index).ScaleWidth * ClickMargin And Y > T(Index).ScaleHeight * (1 - ClickMargin) And Y < T(Index).ScaleHeight Then 'down-left clock
    T(Index).MouseIcon = MS(2).Picture
    Flag = True
    ReadyToClick = True
ElseIf X > T(Index).ScaleWidth * (1 - ClickMargin) And X < T(Index).ScaleWidth And Y > T(Index).ScaleHeight * (1 - ClickMargin) And Y < T(Index).ScaleHeight Then ' down-right anti
    T(Index).MouseIcon = MS(1).Picture
    Flag = False
    ReadyToClick = True
Else
     T(Index).MouseIcon = MS(4).Picture
End If




' Give the current tile a focus


For z% = 1 To BoardSize
    For Tc% = 1 To BoardSize
      If z% * 100 + Tc% = Index Then GoTo fff
       T(z% * 100 + Tc%).BorderStyle = 0
fff:
    Next Tc%
Next z%

T(Index).BorderStyle = 1

LastTile% = Index



End Sub



Public Function LampDirection(Loc As Integer) As Integer


LampDirection = InStr(1, Board(Loc \ 100, Loc Mod 100), "1")


End Function

Public Sub UnloadTiles()


   SolutionDisplayed = False
  ' Frame1.Visible = False
   
   z% = Sqr(T.Count - 2)
   For X% = 1 To z
     For Y% = 1 To z
       Unload T(X * 100 + Y)
     Next Y
   Next X
   
   Refresh



End Sub

Private Sub Timer1_Timer()



If Not IsGameLoaded Then Exit Sub


If Size1 > Screen.Width Or Size2 > Screen.Height Then
   T(LastTile).BorderStyle = 0
   MsgBox "The game will be restarted because the screen dimensions have been changed  .  ", vbCritical + vbOKOnly, "Display Error..."
   UnloadTiles
   UnloadForm = True
   Form_Load
End If



End Sub



Public Sub FormatAdapt(X As Integer, Y As Integer)

  'adapt the format of the tile



WWW:
Board(X, Y) = "0" + Board(X, Y)
If Len(Board(X, Y)) > 4 Then
  Board(X, Y) = right(Board(X, Y), 4)
Else
  GoTo WWW
End If



End Sub

Private Sub Timer2_Timer()


Timing = Timing + 1

If Timing = 10 Then Timing = 0

If Timing < 1 Then
  txt$ = T(0).ToolTipText
Else
  txt$ = ""
End If


For X% = 1 To BoardSize
  For Y% = 1 To BoardSize
    T(X * 100 + Y).ToolTipText = txt
  Next Y
Next X





End Sub



Public Sub Rot_Pic(S_index As Integer, T_Loc As Integer, Dir As Integer)


steps% = 10

If IsSolutionWorking Then steps% = 5: delay% = 3
Ang = Dir * Pi / 2 / steps%

T(0).Picture = T(T_Loc).Picture
T(T_Loc).Picture = Nothing


For X = 1 To steps%
    Bmp_rotate T(0), T(T_Loc), Ang * X
    Sleep delay%
    T(T_Loc).Refresh
Next X


If IsSolutionWorking Then GoTo fff
Bmp_rotate T(0), T(T_Loc), Ang * (X + 0)
T(T_Loc).Refresh
Bmp_rotate T(0), T(T_Loc), Ang * (X - 1)
T(T_Loc).Refresh
Bmp_rotate T(0), T(T_Loc), Ang * (X - 2)
T(T_Loc).Refresh
Bmp_rotate T(0), T(T_Loc), Ang * (X - 1)
T(T_Loc).Refresh

fff:
T(T_Loc).Picture = S(S_index).Picture



End Sub

Public Sub LoadDiffLimits()



ReDim DiffLimits(HS.Min To MaxGameSize, 1 To 2)



'      HardLimit                    EasyLimit
DiffLimits(HS.Min, 1) = 0.11: DiffLimits(HS.Min, 2) = 0.47
DiffLimits(HS.Min + 1, 1) = 0.08: DiffLimits(HS.Min + 1, 2) = 0.42
DiffLimits(HS.Min + 2, 1) = 0.081: DiffLimits(HS.Min + 2, 2) = 0.44
DiffLimits(HS.Min + 3, 1) = 0.09: DiffLimits(HS.Min + 3, 2) = 0.42
DiffLimits(HS.Min + 4, 1) = 0.086: DiffLimits(HS.Min + 4, 2) = 0.39
DiffLimits(HS.Min + 5, 1) = 0.089: DiffLimits(HS.Min + 5, 2) = 0.385
DiffLimits(HS.Min + 6, 1) = 0.09: DiffLimits(HS.Min + 6, 2) = 0.396
DiffLimits(HS.Min + 7, 1) = 0.09: DiffLimits(HS.Min + 7, 2) = 0.38
DiffLimits(HS.Min + 8, 1) = 0.094: DiffLimits(HS.Min + 8, 2) = 0.375
DiffLimits(HS.Min + 9, 1) = 0.096: DiffLimits(HS.Min + 9, 2) = 0.367
DiffLimits(HS.Min + 10, 1) = 0.1: DiffLimits(HS.Min + 10, 2) = 0.351
DiffLimits(HS.Min + 11, 1) = 0.097: DiffLimits(HS.Min + 11, 2) = 0.359
DiffLimits(HS.Min + 12, 1) = 0.1: DiffLimits(HS.Min + 12, 2) = 0.356
DiffLimits(HS.Min + 13, 1) = 0.1: DiffLimits(HS.Min + 13, 2) = 0.345

End Sub

Public Function PressESC() As Boolean



If IsESCPressed Then
  IsESCPressed = False
  PressESC = True
Else
  Exit Function
End If




If IsSolutionWorking Then Exit Function

mnExit_Click
IsESCPressed = False
PressESC = False




'T(LastTile).BorderStyle = 0
'If MsgBox("You will exit the game  . " _
'+ vbCrLf + "Do you want to proceed ?   ", vbYesNo + vbQuestion, "Reset Round...") = 6 Then
'   Exit Function
'Else
'   PressESC = False
'End If



End Function
