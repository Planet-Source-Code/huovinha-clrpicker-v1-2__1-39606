VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "clrPicker"
   ClientHeight    =   2790
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "main.frx":038A
   ScaleHeight     =   2790
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3060
      MaxLength       =   3
      TabIndex        =   49
      Text            =   "255"
      Top             =   2340
      Width           =   920
   End
   Begin VB.TextBox G 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2100
      MaxLength       =   3
      TabIndex        =   48
      Text            =   "255"
      Top             =   2340
      Width           =   970
   End
   Begin VB.TextBox R 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   47
      Text            =   "255"
      Top             =   2340
      Width           =   920
   End
   Begin VB.TextBox HEXvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   46
      Text            =   "FFFFFF"
      Top             =   2340
      Width           =   2775
      Visible         =   0   'False
   End
   Begin VB.TextBox LongValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   45
      Text            =   "16777215"
      Top             =   2340
      Width           =   2775
      Visible         =   0   'False
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   3720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   41
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   3720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   40
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   3720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   39
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   3720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   3720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   37
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   36
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   35
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   34
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   33
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3360
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   32
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   29
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   24
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2280
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   495
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   1245
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1560
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1620
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1240
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   500
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox CurColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   2040
      Width           =   945
   End
   Begin VB.PictureBox Picker 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   120
      Picture         =   "main.frx":50B9
      ScaleHeight     =   1725
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
   Begin VB.Label ViewAs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Long"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   44
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label ViewAs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HEX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   43
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label ViewAs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[RGB]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   42
      Top             =   2040
      Width           =   855
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy to clipboard"
         Begin VB.Menu mnuRGB 
            Caption         =   "RGB"
         End
         Begin VB.Menu mnuHex 
            Caption         =   "HEX"
         End
         Begin VB.Menu mnuLong 
            Caption         =   "Long value"
         End
      End
      Begin VB.Menu mnuSeprator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function GetHEXValue()

On Error Resume Next

Dim HEXr As String, HEXg As String, HEXb As String

    HEXr = Hex$(R.Text)
    If Len(HEXr) = 1 Then HEXr = "0" & HEXr
    
    HEXg = Hex$(G.Text)
    If Len(HEXg) = 1 Then HEXg = "0" & HEXg
    
    HEXb = Hex$(B.Text)
    If Len(HEXb) = 1 Then HEXb = "0" & HEXb
    
    HEXvalue.Text = HEXr & HEXg & HEXb

End Function

Public Function GetLongValue()

On Error Resume Next

    LongValue.Text = CurColor.BackColor

End Function




Public Function GetRGBValue()

On Error Resume Next

    Dim ColorR As String, ColorG As String, ColorB As String

    ColorR = CurColor.BackColor And 255
    ColorG = (CurColor.BackColor And 65280) / 256
    ColorB = (CurColor.BackColor And 16711680) / 65535
    
    R.Text = ColorR
    G.Text = ColorG
    B.Text = ColorB

End Function

Private Sub B_Change()

On Error Resume Next

    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    
    GetLongValue
    GetHEXValue

End Sub


Private Sub Form_Load()
Me.Caption = "clrPicker v" & App.Major & "." & App.Minor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub G_Change()

On Error Resume Next

    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    
    GetLongValue
    GetHEXValue

End Sub

Private Sub HEXvalue_Change()

On Error Resume Next

    GetLongValue
    GetRGBValue

End Sub

Private Sub LongValue_Change()

On Error Resume Next

    CurColor.BackColor = LongValue.Text

    GetRGBValue
    GetHEXValue

End Sub

Private Sub MiniPicker_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    CurColor.BackColor = MiniPicker(Index).BackColor
    GetLongValue
    GetRGBValue
    GetHEXValue
    
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHex_Click()

On Error Resume Next

    Clipboard.Clear
    Clipboard.SetText HEXvalue.Text
   
End Sub

Private Sub mnuLong_Click()

On Error Resume Next

    Clipboard.Clear
    Clipboard.SetText LongValue.Text
    
End Sub


Private Sub mnuRGB_Click()

On Error Resume Next

    Clipboard.Clear
    Clipboard.SetText R.Text & ", " & G.Text & ", " & B.Text
   
End Sub

Private Sub Picker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    CurColor.BackColor = Picker.Point(X, Y)
    GetLongValue
    GetRGBValue
    GetHEXValue
    
End Sub

Private Sub Picker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

If Button = 1 Or 2 Then
    CurColor.BackColor = Picker.Point(X, Y)
    GetLongValue
    GetRGBValue
    GetHEXValue
Else: End If

End Sub


Private Sub R_Change()

On Error Resume Next

    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    
    GetLongValue
    GetHEXValue

End Sub

Private Sub ViewAs_Click(Index As Integer)

On Error Resume Next

    Select Case (Index)
    
        Case 0
            R.Visible = True
            G.Visible = True
            B.Visible = True
            HEXvalue.Visible = False
            LongValue.Visible = False
            ViewAs(0).Caption = "[RGB]"
            ViewAs(1).Caption = "HEX"
            ViewAs(2).Caption = "Long"
        
        Case 1
            R.Visible = False
            G.Visible = False
            B.Visible = False
            HEXvalue.Visible = True
            LongValue.Visible = False
            ViewAs(0).Caption = "RGB"
            ViewAs(1).Caption = "[HEX]"
            ViewAs(2).Caption = "Long"
        
        Case 2
            R.Visible = False
            G.Visible = False
            B.Visible = False
            HEXvalue.Visible = False
            LongValue.Visible = True
            ViewAs(0).Caption = "RGB"
            ViewAs(1).Caption = "HEX"
            ViewAs(2).Caption = "[Long]"
    
    End Select

End Sub


Private Sub ViewAs_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    Select Case (Index)
    
        Case 0
            ViewAs(0).ForeColor = 8421504
        
        Case 1
            ViewAs(1).ForeColor = 8421504
        
        Case 2
            ViewAs(2).ForeColor = 8421504
    
    End Select

End Sub


Private Sub ViewAs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

    Select Case (Index)
    
        Case 0
            ViewAs(0).ForeColor = 0
        
        Case 1
            ViewAs(1).ForeColor = 0
        
        Case 2
            ViewAs(2).ForeColor = 0
    
    End Select

End Sub


