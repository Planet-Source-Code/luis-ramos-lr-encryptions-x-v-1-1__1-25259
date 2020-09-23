VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   Caption         =   "LR Encryptions V.1.1"
   ClientHeight    =   2925
   ClientLeft      =   4065
   ClientTop       =   3750
   ClientWidth     =   3495
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   Picture         =   "Main.frx":0442
   ScaleHeight     =   2925
   ScaleWidth      =   3495
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4C6C
            Key             =   "lrecrypt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":4F88
            Key             =   "lrfileecrypt"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1535
      ButtonWidth     =   2725
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LR Text Encryptions"
            ImageKey        =   "lrecrypt"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LR File Ecryption"
            ImageKey        =   "lrfileecrypt"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   0
      Left            =   720
      Picture         =   "Main.frx":52A4
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   1575
      Left            =   0
      Shape           =   2  'Oval
      Top             =   960
      Width           =   2775
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      X1              =   2400
      X2              =   3360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000FF&
      X1              =   3360
      X2              =   3120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      X1              =   3120
      X2              =   3120
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      X1              =   3360
      X2              =   3120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      X1              =   3000
      X2              =   3000
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   3000
      X2              =   2760
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   2760
      X2              =   2760
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   2400
      X2              =   2640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   2400
      X2              =   2640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2520
      X2              =   2520
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   2040
      X2              =   2280
      Y1              =   2040
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   2040
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   2040
      X2              =   2280
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   2040
      X2              =   2040
      Y1              =   1920
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   360
      X2              =   360
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   1
      Left            =   1920
      Picture         =   "Main.frx":9ACE
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   2
      Left            =   0
      Picture         =   "Main.frx":E2F8
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   3
      Left            =   1800
      Picture         =   "Main.frx":12B22
      Top             =   360
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   4
      Left            =   1560
      Picture         =   "Main.frx":1734C
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   1170
      Index           =   5
      Left            =   360
      Picture         =   "Main.frx":1BB76
      Top             =   960
      Width           =   1170
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Begin VB.Menu mnulrstring 
            Caption         =   "LR Encryptions"
         End
         Begin VB.Menu mnulrfile 
            Caption         =   "LR File Ecryption QL"
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hide
Form2.Show
End Sub

Private Sub mnuexit_Click()
Hide
Form2.Show
End Sub

Private Sub mnulrfile_Click()
Form3.Show
End Sub

Private Sub mnulrstring_Click()
Form1.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    Form1.Show
ElseIf Button.Index = 3 Then
    Form3.Show
End If
End Sub
