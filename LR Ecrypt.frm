VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "LR Encryptions"
   ClientHeight    =   4215
   ClientLeft      =   3885
   ClientTop       =   2655
   ClientWidth     =   3780
   Icon            =   "LR Ecrypt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "LR Ecrypt.frx":030A
   ScaleHeight     =   4215
   ScaleWidth      =   3780
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   1111
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Encrypt"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Decrypt"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageKey        =   "about"
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404040&
      Caption         =   "LR Ecrypt HLV.1.0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      Caption         =   "LR Ecrypt XP V.2.0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Encryption:"
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox vText 
      BackColor       =   &H8000000C&
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":0614
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":076E
            Key             =   "about"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":0A92
            Key             =   "encrypt"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":0DB6
            Key             =   "decrypt"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR Ecrypt.frx":152E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Index           =   2
      Begin VB.Menu mnuencrypt 
         Caption         =   "&Encrypt"
      End
      Begin VB.Menu mnudecrypt 
         Caption         =   "&Decrypt"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuinstructions 
         Caption         =   "&Instructions"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()
Hide
Form2.Show
End Sub

Private Sub Command4_Click()
vText = ""
End Sub



Private Sub Command1_Click()
vText = ""
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuClear_Click()
vText = ""
End Sub

Private Sub mnudecrypt_Click()
vText = DeCode(vText)
End Sub

Private Sub mnuencrypt_Click()
vText = Encode(vText)
End Sub

Private Sub mnuexit_Click()
Hide
Form4.Show
End Sub

Private Sub mnuinstructions_Click()
frmAbout.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Option1.Value = True Then
    If Button.Index = 1 Then
        vText = Encode(vText)
    ElseIf Button.Index = 3 Then
        vText = DeCode(vText)
    End If
End If
If Option2.Value = True Then
    If Button.Index = 1 Then
        vText = Encodehl(vText)
    ElseIf Button.Index = 3 Then
        vText = Encodehl(vText)
    End If
End If
If Button.Index = 5 Then
    Hide
    Form4.Show
ElseIf Button.Index = 13 Then
    frmAbout.Show
End If

    
End Sub
