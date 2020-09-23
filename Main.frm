VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "City Game Map Maker"
   ClientHeight    =   5685
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8220
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7620
      Top             =   3240
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   5640
      TabIndex        =   61
      Top             =   2580
      Width           =   1875
      Begin VB.PictureBox picSeason 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1320
         Picture         =   "Main.frx":030A
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   75
         ToolTipText     =   "Land"
         Top             =   1200
         Width           =   300
      End
      Begin VB.PictureBox picSeason 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   960
         Picture         =   "Main.frx":07FC
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   74
         ToolTipText     =   "Land"
         Top             =   1200
         Width           =   300
      End
      Begin VB.PictureBox picSeason 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   600
         Picture         =   "Main.frx":0CEE
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   73
         ToolTipText     =   "Land"
         Top             =   1200
         Width           =   300
      End
      Begin VB.PictureBox picSeason 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":11E0
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   72
         ToolTipText     =   "Land"
         Top             =   1200
         Width           =   330
      End
      Begin VB.PictureBox Items 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":16D2
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   69
         ToolTipText     =   "Land"
         Top             =   480
         Width           =   330
      End
      Begin VB.PictureBox Items 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   600
         Picture         =   "Main.frx":1BC4
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   68
         ToolTipText     =   "Water"
         Top             =   480
         Width           =   300
      End
      Begin VB.PictureBox Items 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   960
         Picture         =   "Main.frx":20B6
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   67
         ToolTipText     =   "Trees"
         Top             =   480
         Width           =   300
      End
      Begin VB.PictureBox Items 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1320
         Picture         =   "Main.frx":25A8
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   66
         ToolTipText     =   "Demolish"
         Top             =   480
         Width           =   300
      End
      Begin VB.PictureBox ItemSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":2A9A
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   65
         ToolTipText     =   "Land"
         Top             =   840
         Width           =   330
      End
      Begin VB.PictureBox ItemSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   600
         Picture         =   "Main.frx":2F8C
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   64
         ToolTipText     =   "Land"
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox ItemSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   960
         Picture         =   "Main.frx":347E
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   63
         ToolTipText     =   "Land"
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox ItemSize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   1320
         Picture         =   "Main.frx":3970
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   62
         ToolTipText     =   "Land"
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lblSelected 
         BackStyle       =   0  'Transparent
         Caption         =   "Land"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.PictureBox BufferM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7380
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   60
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox BufferS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   7380
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   59
      Top             =   5580
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8100
      Picture         =   "Main.frx":3E62
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8460
      Picture         =   "Main.frx":4354
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8100
      Picture         =   "Main.frx":4846
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8460
      Picture         =   "Main.frx":4D38
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8100
      Picture         =   "Main.frx":522A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8460
      Picture         =   "Main.frx":571C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8100
      Picture         =   "Main.frx":5C0E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8460
      Picture         =   "Main.frx":6100
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   8460
      Picture         =   "Main.frx":65F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   8100
      Picture         =   "Main.frx":6AE4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   8460
      Picture         =   "Main.frx":6FD6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   8100
      Picture         =   "Main.frx":74C8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   8460
      Picture         =   "Main.frx":79BA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   8100
      Picture         =   "Main.frx":7EAC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   8460
      Picture         =   "Main.frx":839E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   8100
      Picture         =   "Main.frx":8890
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   8100
      Picture         =   "Main.frx":8D82
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   8460
      Picture         =   "Main.frx":9274
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   8100
      Picture         =   "Main.frx":9766
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3780
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   8460
      Picture         =   "Main.frx":9C58
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3780
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   8100
      Picture         =   "Main.frx":A14A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   8460
      Picture         =   "Main.frx":A63C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4140
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   8100
      Picture         =   "Main.frx":AB2E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   8460
      Picture         =   "Main.frx":B020
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4500
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   8460
      Picture         =   "Main.frx":B512
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   8100
      Picture         =   "Main.frx":BA04
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   8460
      Picture         =   "Main.frx":BEF6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   8100
      Picture         =   "Main.frx":C3E8
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5580
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   8460
      Picture         =   "Main.frx":C8DA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   8100
      Picture         =   "Main.frx":CDCC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicmTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   8460
      Picture         =   "Main.frx":D2BE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicTree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   8100
      Picture         =   "Main.frx":D7B0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox BufferMap 
      AutoRedraw      =   -1  'True
      Height          =   4395
      Left            =   9300
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox picMM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   5640
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   25
      Top             =   60
      Width           =   2355
   End
   Begin VB.VScrollBar VScroll 
      Height          =   5175
      Left            =   5280
      Min             =   1
      TabIndex        =   24
      Top             =   120
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   60
      Min             =   1
      TabIndex        =   23
      Top             =   5340
      Value           =   1
      Width           =   5175
   End
   Begin VB.PictureBox MainPic 
      AutoRedraw      =   -1  'True
      Height          =   5160
      Left            =   60
      ScaleHeight     =   340
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   22
      Top             =   120
      Width           =   5160
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8760
      Picture         =   "Main.frx":DCA2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8760
      Picture         =   "Main.frx":E194
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8760
      Picture         =   "Main.frx":E686
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSummer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8760
      Picture         =   "Main.frx":EB78
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8760
      Picture         =   "Main.frx":F06A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8760
      Picture         =   "Main.frx":F55C
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8760
      Picture         =   "Main.frx":FA4E
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicAutumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8760
      Picture         =   "Main.frx":FF40
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8760
      Picture         =   "Main.frx":10432
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8760
      Picture         =   "Main.frx":10924
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8760
      Picture         =   "Main.frx":10E16
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   11
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicWinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8760
      Picture         =   "Main.frx":11308
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8760
      Picture         =   "Main.frx":117FA
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8760
      Picture         =   "Main.frx":11CEC
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8760
      Picture         =   "Main.frx":121DE
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicSpring 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8760
      Picture         =   "Main.frx":126D0
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   8400
      Picture         =   "Main.frx":12BC2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   8400
      Picture         =   "Main.frx":130B4
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8400
      Picture         =   "Main.frx":135A6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   8400
      Picture         =   "Main.frx":13A98
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   8400
      Picture         =   "Main.frx":13F8A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox BufferG 
      AutoRedraw      =   -1  'True
      Height          =   435
      Left            =   7500
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   1395
      Left            =   5700
      TabIndex        =   71
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New Map"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "Open Map"
      End
      Begin VB.Menu div3 
         Caption         =   "-"
      End
      Begin VB.Menu mSave 
         Caption         =   "Save Map"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save Map As"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu mImport 
         Caption         =   "Import Map"
         Shortcut        =   ^O
      End
      Begin VB.Menu div2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Restart
End Sub
Sub Restart()
    Hoyde = 50
    Bredde = 50
    WBredde = 16
    WHoyde = 16
    WstartX = 1
    WstartY = 1
    

    
    ReDim Boarddata(1 To Bredde, 1 To Hoyde)
    
    VScroll.Max = Hoyde - WHoyde
    HScroll.Max = Bredde - WBredde
    VScroll.Value = 1
    HScroll.Value = 1
    PaintGround
    
    picMM.Cls
    PaintMap 1, picMM
End Sub


Private Sub Form_Load()
    BufferG.Width = MainPic.Width
    BufferG.Height = MainPic.Height
    BufferM.Width = MainPic.Width
    BufferM.Height = MainPic.Height
    BufferS.Width = MainPic.Width
    BufferS.Height = MainPic.Height
    ISize = 1
    MapName = "New Map"
    Restart
    FirstTime = True
    
    Text = "Map Maker verson 1.0" & vbNewLine
    Text = Text & "for" & vbNewLine
    Text = Text & "City Game 2001"
    lblInfo.Caption = Text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Dirty Then
        Svar = MsgBox("The map is not saved. Continue?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then
            Cancel = 1
        Else
            End
        End If
    End If
End Sub

Private Sub HScroll_Change()
    WstartX = HScroll.Value
    picMM.Cls
    BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
    picMM.Line (WstartX - 2, WstartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMM.Refresh
    PaintGround
End Sub

Private Sub Items_Click(Index As Integer)
    For a = 0 To 3
    Items.Item(a).BorderStyle = 0
    Next a
    Items.Item(Index).BorderStyle = 1
    lblSelected.Caption = Items.Item(Index).ToolTipText
    SelItem = Index
    
End Sub

Private Sub ItemSize_Click(Index As Integer)
    For a = 0 To 3
    ItemSize.Item(a).BorderStyle = 0
    Next a
    ItemSize.Item(Index).BorderStyle = 1
    ISize = Index + 1
End Sub

Private Sub lblCOOR_Click()

End Sub

Private Sub mainPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitX As Integer
Dim HitY As Integer
    'HØYRE KNAPP
    If Button = 2 Then
        HitX = GetXY(X)
        HitY = GetXY(Y)
        
        If HitX > WBredde / 2 Then
            HitX = HitX
            
            WstartX = WstartX + HitX - (WBredde / 2)
        Else
            WstartX = WstartX - ((WBredde / 2) - HitX)
        End If
        
        If HitY > WHoyde / 2 Then
            HitY = HitY
            
            WstartY = WstartY + HitY - (WBredde / 2)
        Else
            WstartY = WstartY - ((WHoyde / 2) - HitY)
        End If
        
        If WstartX <= 0 Then WstartX = 1
        If WstartX >= Bredde - WBredde Then WstartX = Bredde - WBredde
        If WstartY <= 0 Then WstartY = 1
        If WstartY >= Hoyde - WHoyde Then WstartY = Hoyde - WHoyde
        
        HScroll.Value = WstartX
        VScroll.Value = WstartY
        
    End If
    
    If Button = 1 Then
        Dirty = True
        HitX = GetXY(X) + WstartX
        HitY = GetXY(Y) + WstartY
        DoTheShit HitX, HitY
        PaintGround
        PaintMapSmall 1, Form1.picMM
        OldX = HitX
        OldY = HitY
    End If
End Sub


Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim glY As Integer
Dim glX As Integer
        
    If Button = 1 Then
        glX = GetXY(X) + WstartX
        glY = GetXY(Y) + WstartY
        
        If glX = OldX And glY = OldY Then
            OldX = glX
            OldY = glY
            Exit Sub
        End If
        DoTheShit glX, glY
        PaintGround
        PaintMapSmall 1, Form1.picMM
        OldX = glX
        OldY = glY
    End If
End Sub



Private Sub mExit_Click()
    End
End Sub

Private Sub mImport_Click()
    Import.Show , Me
End Sub

Private Sub mNew_Click()
    frmNew.Show , Me
End Sub

Private Sub mOpen_Click()
    frmLoad.Show , Me
End Sub

Private Sub mSave_Click()
    If FirstTime = True Then
        If Dir(App.path & "\maps\" & MapName & ".map") <> "" Then
            Svar = MsgBox("Game already exsist. Overwrite?", vbOKCancel + vbInformation, GameTitle)
            If Not Svar = vbOK Then Exit Sub
        End If
    End If
    FirstTime = False
    
    If Dir(App.path & "\maps\" & MapName & ".map") = "" Then
        frmSave.Show , Me
    Else
        SaveMap App.path & "\maps\" & MapName & ".map"
        Dirty = False
    End If
End Sub

Private Sub mSaveAs_Click()
    frmSave.Show , Me
End Sub

Private Sub picMM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picMM.Cls
        BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
        picMM.Line (X - (WBredde / 2) - 2, Y - (WHoyde / 2) - 2)-Step(WBredde, WHoyde), vbWhite, B
        picMM.Refresh
        WstartX = X - (WBredde / 2)
        WstartY = Y - (WHoyde / 2)
        
        If WstartX <= 0 Then WstartX = 1
        If WstartX >= Bredde - WBredde Then WstartX = Bredde - WBredde
        If WstartY <= 0 Then WstartY = 1
        If WstartY >= Hoyde - WHoyde Then WstartY = Hoyde - WHoyde
        
        HScroll.Value = WstartX
        VScroll.Value = WstartY
    End If
End Sub

Private Sub picMM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMM.Cls
    BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
    picMM.Line (WstartX - 2, WstartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMM.Refresh
End Sub

Private Sub picSeason_Click(Index As Integer)
    For a = 0 To 3
    picSeason.Item(a).BorderStyle = 0
    Next a
    picSeason.Item(Index).BorderStyle = 1
    SetSeason Index + 1
    PaintGround
End Sub

Private Sub Timer1_Timer()
    Form1.Caption = AppTitle & " - " & MapName
End Sub

Private Sub VScroll_Change()
    WstartY = VScroll.Value
    picMM.Cls
    BitBlt picMM.hDC, 0, 0, Bredde * 4, Hoyde * 4, BufferMap.hDC, 0, 0, SRCCOPY
    picMM.Line (WstartX - 2, WstartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMM.Refresh
    PaintGround
End Sub



