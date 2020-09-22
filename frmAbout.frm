VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Tunnel"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplayShip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   2
      Left            =   1440
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   10
      Top             =   1560
      Width           =   630
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox picDisplayShip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   1
      Left            =   2640
      Picture         =   "frmAbout.frx":184C
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   630
   End
   Begin VB.PictureBox picDisplayShip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Index           =   0
      Left            =   240
      Picture         =   "frmAbout.frx":1CBE
      ScaleHeight     =   630
      ScaleWidth      =   630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lblBody 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":2130
      Height          =   1455
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label lblBody 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "09 / 08 / 00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblBody 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblBody 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Anton Venema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblBody 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tunnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "Tunnel v2.0"
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

    'Hide form when OK is clicked
    frmAbout.Hide

End Sub
