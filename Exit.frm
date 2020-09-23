VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Thank You"
   ClientHeight    =   1905
   ClientLeft      =   3510
   ClientTop       =   6225
   ClientWidth     =   4680
   Icon            =   "Exit.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1080
      Picture         =   "Exit.frx":0742
      ScaleHeight     =   975
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You for using LR Software"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload frmAbout1
Unload frmAbout
Unload frmSplash
Unload Me
End Sub
