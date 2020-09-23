VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   Caption         =   "LR File Ecrypt QL"
   ClientHeight    =   7395
   ClientLeft      =   1335
   ClientTop       =   1185
   ClientWidth     =   9015
   Icon            =   "LR File Ecrypt QL.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7395
   ScaleWidth      =   9015
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lstKey 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Drag and Drop the file to used as the key"
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Browse"
      Top             =   2520
      Width           =   285
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   3735
   End
   Begin VB.PictureBox picSmall 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5640
      Picture         =   "LR File Ecrypt QL.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR File Ecrypt QL.frx":0454
            Key             =   "Crypto"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstToolbar 
      Left            =   6600
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR File Ecrypt QL.frx":08A6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LR File Ecrypt QL.frx":0A00
            Key             =   "Start"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbDragDrop 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7140
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "0 Total KB"
            TextSave        =   "0 Total KB"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modified"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCrypto 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1005
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Exit   "
            Object.ToolTipText     =   "Exit the Program"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "   Start   "
            Object.ToolTipText     =   "Start Encryption\Decryption"
            ImageKey        =   "Start"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.OptionButton optAction 
         Caption         =   "Decrypt"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton optAction 
         Caption         =   "Encrypt"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      Picture         =   "LR File Ecrypt QL.frx":0B5A
      ScaleHeight     =   855
      ScaleWidth      =   2895
      TabIndex        =   8
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblKey 
      BackStyle       =   0  'Transparent
      Caption         =   "Key:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks, Brad Martinez, modSH's Programer"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Path to write Encrypted file to:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnustart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuinstructions 
         Caption         =   "Instuctions"
      End
      Begin VB.Menu mnuhow 
         Caption         =   "How it Works?"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()

    Dim Path As String
    Path = ShowDialog(lblPath.Caption)

    If Path <> "" Then
        txtPath.Text = Path
    End If

End Sub

Private Sub Form_Load()

    ImageList1.ImageHeight = 16
    ImageList1.ImageWidth = 16

    optAction(0).Value = True

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub

    If Me.Width < 3855 Then
        Me.Width = 3855
    End If
    If Me.Height < 6075 Then
        Me.Height = 6075
    End If

    lstFiles.Move 0, 2800, ScaleWidth, ScaleHeight - stbDragDrop.Height - 2400
    
    lstKey.Width = ScaleWidth
    lstKey.ColumnHeaders(1).Width = ScaleWidth - 200
    
    txtPath.Width = ScaleWidth - 400
    ProgressBar1.Width = ScaleWidth
    cmdBrowse.Left = ScaleWidth - 300
    
    stbDragDrop.Panels(1).Width = ScaleWidth - 300
    
End Sub

Private Sub lstFiles_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim l As MSComctlLib.ListItem
    Dim i As Integer
    Dim SelFiles As New Collection

    If KeyCode = 46 Then
        
        For i = lstFiles.ListItems.Count To 1 Step -1
            If lstFiles.ListItems(i).Selected Then
                lstFiles.ListItems.Remove i
            End If
        Next i

        stbDragDrop.Panels(1).Text = GetTotalFileSize & " Total KB"
    End If

End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        PopupMenu mnuPopup
    End If

End Sub

Private Sub lstFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Call ListView_OLEDragDrop(Data, lstFiles)

End Sub

Public Sub GetIcon(Path As String)
    On Local Error GoTo errHandler

    Dim hImgSmall As Long
    Dim fName As String
    Dim fnFilter As String
    Dim l As Long
    
    fName = Path

    hImgSmall = SHGetFileInfo(fName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    
    picSmall.Picture = LoadPicture()

    l = ImageList_Draw(hImgSmall, shinfo.iIcon, picSmall.hdc, 0, 0, ILD_TRANSPARENT)

    Exit Sub
errHandler:

    MsgBox Err.Description, vbCritical, "LR File Ecrypt QL"
    picSmall.Picture = LoadPicture()

End Sub

Private Sub lstFiles_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    Effect = 1

End Sub

Public Function FileExists(FileName As String) As Boolean

    Dim l As MSComctlLib.ListImage
    
    For Each l In ImageList1.ListImages
        If l.Key = FileName Then
            FileExists = True
            Exit Function
        End If
    Next l

End Function

Public Function GetFile(FullPath As String) As String
    On Error Resume Next

    Dim s As String
    Dim Delimiter As Integer
    Dim i As Integer
    
    s = FullPath

    For i = Len(s) To 0 Step -1
        If Mid(s, i, 1) = "\" Then
            Delimiter = Len(s) - i
            Exit For
        End If
    Next i

    s = Right(s, Delimiter)
    
    GetFile = s

End Function

Public Sub ListView_OLEDragDrop(Data As MSComctlLib.DataObject, ListView As MSComctlLib.ListView)
    On Error GoTo errHandler

    Dim i As Integer
    Dim j As Integer
    Dim FileCount As Integer
    Dim FileName As String
    Dim FileLength As Long
    Dim m As New clsMousePointer
    m.SetCursor
    
    FileCount = Data.Files.Count

    For i = 1 To FileCount
        
        FileLength = FileLen(Data.Files(i))
        
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            MsgBox Data.Files(i) & "\ is not a file. Only Files can be added.", vbExclamation
        ElseIf FileLength = 0 Then
            MsgBox Data.Files(i) & " does not contain any data and cannot be used as a key or included in the Archive.", vbExclamation
        Else
        
            FileName = GetFile(Data.Files(i))
        
            If GetTotalFileSize <> -1 Then
                stbDragDrop.Panels(1).Text = GetTotalFileSize & " Total Bytes"
            Else
                MsgBox "The max size of the Archive has been reached. No more files can be added.", vbInformation
                
                lstFiles.ListItems.Remove (lstFiles.ListItems.Count)
                
                Exit Sub
            End If
            
            j = j + 1
            
            If ListView.ListItems.Count > 255 Then
                MsgBox "Only 255 Files can be inserted.", vbInformation, "LR File Ecrypt QL"
                Exit Sub
            End If
            
            Call GetIcon(Data.Files(i))
        
            If Not FileExists(FileName) Then
                ImageList1.ListImages.Add , FileName, picSmall.Image
            End If
            
            ListView.ListItems.Add , , Data.Files(i), , FileName
            ListView.ListItems(ListView.ListItems.Count).ListSubItems.Add , , FileLength & "Bytes"
            ListView.ListItems(ListView.ListItems.Count).ListSubItems.Add , , FileDateTime(Data.Files(i))
            
        End If
    Next i

    Exit Sub
    
errHandler:
    MsgBox "An error was encountered while trying to add " & Data.Files(i) & "." & _
        " Make sure the no other users or applications are accessing the file.", vbExclamation, "LR File Ecrypt QL"

End Sub

Private Sub lstKey_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Data.Files.Count > 1 Then
        MsgBox "Only one file can be used as the key.", vbInformation
        lblKey.Caption = "Key: "
    Else
        lblKey.Caption = "Key: " & FileLen(Data.Files(1)) & " Bytes"
        lstKey.ListItems.Clear
        Call ListView_OLEDragDrop(Data, lstKey)
    End If

End Sub

Public Function GetTotalFileSize() As Long
    On Error GoTo errHandler

    Dim TotalFileLength As Long
    Dim l As MSComctlLib.ListItem
    
    For Each l In lstFiles.ListItems
        TotalFileLength = TotalFileLength + Left(l.ListSubItems(1).Text, Len(l.ListSubItems(1).Text) - 5)
    Next l

    GetTotalFileSize = TotalFileLength

    Exit Function
    
errHandler:
    GetTotalFileSize = -1

End Function


Private Function ShowDialog(Caption As String) As String

  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  
  With BI
    .hOwner = Me.hWnd
    
    nFolder = GetFolderValue(13)
    
    ' Fill the item id list with the pointer of the selected folder item, rtns 0 on success
    ' ==================================================
    ' If this function fails because the selected folder doesn't exist,
    ' .pidlRoot will be uninitialized & will equal 0 (CSIDL_DESKTOP)
    ' and the root will be the Desktop.
    ' DO NOT specify the CSIDL_ constants for .pidlRoot !!!!
    ' The SHBrowseForFolder() call below will generate a fatal exception
    ' (GPF) if the folder indicated by the CSIDL_ constant does not exist!!
    ' ==================================================
    If SHGetSpecialFolderLocation(ByVal Me.hWnd, ByVal nFolder, IDL) = NOERROR Then
      .pidlRoot = IDL.mkid.cb
    End If
    
    ' Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(MAX_PATH, 0)
    
    ' Set the dialog's banner text
    .lpszTitle = Caption
    
    ' Set the type of folders to display & return
    ' -play with these option constants to see what can be returned
    .ulFlags = GetReturnType()
    
  End With
  
  ' Show the Browse dialog
  pIdl = SHBrowseForFolder(BI)
  
  ' If the dialog was cancelled...
  If pIdl = 0 Then Exit Function
    
  ' Fill sPath w/ the selected path from the id list
  ' (will rtn False if the id list can't be converted)
  sPath = String$(MAX_PATH, 0)
  SHGetPathFromIDList ByVal pIdl, ByVal sPath

  ' Display the path and the name of the selected folder
  ShowDialog = Left(sPath, InStr(sPath, vbNullChar) - 1)
  'MsgBox Left$(BI.pszDisplayName, _
                             InStr(BI.pszDisplayName, vbNullChar) - 1)
  
  
  ' Frees the memory SHBrowseForFolder()
  ' allocated for the pointer to the item id list
  CoTaskMemFree pIdl
  
End Function

Private Function GetFolderValue(wIdx As Integer) As Long
' Returns the value of the system folder constant specified by wIdx
' See BrowsDlg.bas for the system folder nFolder values
    
    ' The Desktop
    If wIdx < 2 Then
      GetFolderValue = 0
    
    ' Programs Folder --> Start Menu Folder
    ElseIf wIdx < 12 Then
      GetFolderValue = wIdx
    
    ' Desktop Folder --> ShellNew Folder
    Else   ' wIdx >= 12
      GetFolderValue = wIdx + 4
    End If

End Function

Private Function GetReturnType() As Long
  Dim dwRtn As Long
  dwRtn = dwRtn Or BIF_RETURNONLYFSDIRS
  GetReturnType = dwRtn
End Function

Private Sub mnuClear_Click()

    lstFiles.ListItems.Clear

End Sub

Private Sub mnuDelete_Click()

    Call lstFiles_KeyDown(46, 1)

End Sub

Private Sub mnuabout_Click()
frmAbout1.Show
End Sub

Private Sub mnuexit_Click()
Hide
Form4.Show
End Sub

Private Sub mnustart_Click()
    Dim Path As String
    Dim SourcePath As String


If lstKey.ListItems.Count <> 1 Then
            MsgBox "Missing Key. Please drag a file to use as the key into the box.", vbExclamation, "LR File Ecrypt QL"
            Exit Sub
        End If
        If txtPath.Text = "" Then
            MsgBox "The Path specified as the File Destination does not exist, please enter a correct Path.", vbExclamation, "LR File Ecrypt SQ"
            txtPath.SetFocus
            Exit Sub
        End If
        If Dir(txtPath.Text, vbDirectory) = "" Then
            MsgBox "The Path specified as the File Destination does not exist, please enter a correct Path.", vbExclamation, "LR File Ecrypt SQ"
            txtPath.Text = ""
            txtPath.SetFocus
            Exit Sub
        End If
        
        If optAction(0) Then
            If lstFiles.ListItems.Count = 0 Then
                MsgBox "No files were selected to Encrypt. Please make sure there is at least one file in the list.", vbExclamation, "LR File Ecrypt SQ"
                Exit Sub
            End If
        ElseIf optAction(1) Then
            If lstFiles.ListItems.Count <> 1 Then
                MsgBox "Only one *.cpt file can be Decrypted at once. Please make sure there is only one valid Archive in the list.", vbExclamation, "LR File Ecrypt SQ"
                Exit Sub
            End If
            If lstFiles.ListItems.Count = 1 Then
                If UCase(Right(lstFiles.ListItems(1).Text, 3)) <> "CPT" Then
                    MsgBox "The file you are attempting to Decrypt is not a valid *.cpt file. Please make sure there is only one valid Archive in the list.", vbExclamation, "LR File Ecrypt SQ"
                    Exit Sub
                End If
            End If
        End If

        
        
        If optAction(0) Then
            'Encrypt
        
            If Right(txtPath.Text, 1) <> "\" Then
                Path = txtPath.Text & "\" & "ARCHIVE_" & Format(Date, "MM-DD-YYYY") & " " & Format(Time, "h-mm-ss am/pm") & ".cpt"
            Else
                Path = txtPath.Text & "ARCHIVE_" & Format(Date, "MM-DD-YYYY") & " " & Format(Time, "h-mm-ss am/pm") & ".cpt"
            End If
        
            If Not CreateArchive(Path) Then
                MsgBox "The files were unable to be Encrypted.", vbExclamation, "LR File Ecrypt SQ"
            End If
             
            
        ElseIf optAction(1) Then
            'Decrypt
        
            If Right(txtPath.Text, 1) <> "\" Then
                Path = txtPath.Text & "\"
            End If
            
            SourcePath = lstFiles.ListItems(1).Text
        
            If Not DecodeArchive(SourcePath, Path) Then
                MsgBox "The Archive was not able to be Decrypted.", vbExclamation, "LR File Ecrypt SQ"
            End If
                
        End If

    
End Sub

Private Sub optAction_Click(Index As Integer)

    If Index = 0 Then
        'Encrypt
        lblPath.Caption = "Path to write Encrypted file to:"
        lblFiles.Caption = "Files to Encrypt:"
    
    ElseIf Index = 1 Then
        'Decrypt
        lblPath.Caption = "Path to write Decrypted files to:"
        lblFiles.Caption = "File to Decrypt:"
    
    End If

End Sub

Public Function CreateArchive(Path As String) As Boolean
    On Error GoTo errHandler

    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim Status As String
    Dim FileCount As Integer
    Dim TotalLength As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim KeyPath As String
    Dim FileNames As String
    Dim FileLengths() As Long
    Dim m As New clsMousePointer
    m.SetCursor
    
    Status = stbDragDrop.Panels(1).Text
    
    
    TotalLength = GetTotalFileSize
    FileCount = lstFiles.ListItems.Count
    
    If TotalLength = -1 Then
        MsgBox "The Archive is too large. Some files need to be removed.", vbExclamation, "LR File Ecrypt QL"
        CreateArchive = False
        Exit Function
    End If
    
    KeyPath = lstKey.ListItems(1).Text
    ReDim FileLengths(FileCount)
    
    
    For i = 1 To FileCount
        FileNames = FileNames & GetFile(lstFiles.ListItems(i).Text) & "|"
    Next i
    
    For i = 1 To FileCount
        FileLengths(i - 1) = Left(lstFiles.ListItems(i).ListSubItems(1), Len(lstFiles.ListItems(i).ListSubItems(1)) - 5)
        FileNames = FileNames & FileLengths(i - 1) & "|"
    Next i
    
    ReDim b(Len(FileNames) + 1)
    b(0) = CByte(FileCount)
    For i = 1 To Len(FileNames)
        b(i) = Asc(Mid(FileNames, i, 1))
    Next i
   
   
   
    l = UBound(b)
    ReDim Preserve b(UBound(b) + TotalLength)
    
    For i = 1 To FileCount
        ReDim b2(FileLengths(i - 1))
        
        stbDragDrop.Panels(1).Text = "Adding: " & lstFiles.ListItems(i).Text
        
        Open lstFiles.ListItems(i).Text For Binary Access Read As #1
            Get #1, , b2()
        Close #1

        For j = 0 To UBound(b2) - 1
            b(l) = b2(j)
            l = l + 1
        Next j
        
    Next i
    
    ReDim Preserve b(UBound(b) + 2)
    
    b(UBound(b) - 2) = Asc("C")
    b(UBound(b) - 1) = Asc("P")
    b(UBound(b)) = Asc("T")
    
    
    
    stbDragDrop.Panels(1).Text = "Encrypting..."
    
    i = 0
    ReDim b2(FileLen(KeyPath))
    Open KeyPath For Binary Access Read As #1
        Get #1, , b2()
    Close #1
    For l = 0 To UBound(b)
        If b2(i) = 0 Then
            b2(i) = 33
        End If
        
        'Encrypt
        b(l) = CByte(Encr(CInt(b(l)), CInt(b2(i))))
        
        If i = UBound(b2) Then
            i = 0
        Else
            i = i + 1
        End If
        ProgressBar1.Visible = True
        ProgressBar1.Value = l
        ProgressBar1.Max = UBound(b)
        
    Next l
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
    
    
    Open Path For Binary Access Write As #1
        Put #1, , b
    Close #1
    
    MsgBox "Files Archived at " & Path, vbInformation, "LR File Ecrypt QL"
    
    stbDragDrop.Panels(1).Text = Status
    
    CreateArchive = True
    
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical, "LR File Ecrypt QL"
    
    stbDragDrop.Panels(1).Text = Status
    CreateArchive = False
    
End Function

Public Function DecodeArchive(SourcePath As String, DestinationPath As String) As Boolean
    On Error GoTo errHandler

    Dim j As Integer
    Dim l As Long
    Dim i As Long
    Dim Trailer As String
    Dim NextReadPosition As Long
    Dim FileCount As Integer
    Dim FileLength As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim KeyPath As String
    Dim FileNames() As String
    Dim FileLengths() As Long
    Dim Status As String
    Dim m As New clsMousePointer
    m.SetCursor

    'Used | as delimiter, asc = 124
    
    'Set the status
    Status = stbDragDrop.Panels(1).Text

    FileLength = FileLen(SourcePath)
    ReDim b(FileLength)
    
    Open SourcePath For Binary Access Read As #1
        Get #1, , b()
    Close #1

    'This is where the decryption happens
    KeyPath = lstKey.ListItems(1).Text

    ReDim b2(FileLen(KeyPath))
    Open KeyPath For Binary Access Read As #1
        Get #1, , b2()
    Close #1
    
    'Set the status
    stbDragDrop.Panels(1).Text = "Decrypting..."
    
    For l = 0 To UBound(b)
        'Exclude 0 values
        If b2(i) = 0 Then
            b2(i) = 33
        End If
        
        'Encrypt
        b(l) = CByte(Decr(CInt(b(l)), CInt(b2(i))))
        
        If i = UBound(b2) Then
            i = 0
        Else
            i = i + 1
        End If
        
    Next l
    i = 0
    l = 0
    

    'Checks if valid *.CPT file
    Trailer = Chr(b(FileLength - 3))
    Trailer = Trailer & Chr(b(FileLength - 2))
    Trailer = Trailer & Chr(b(FileLength - 1))

    If Trailer <> "CPT" Then
        MsgBox "The target Archive is not a valid *.cpt file, or the incorrect Key is being used.", vbExclamation, "Crypto"
        DecodeArchive = False
        Exit Function
    End If


    'Get the file count
    FileCount = b(0) - 1
    ReDim FileNames(FileCount)
    ReDim FileLengths(FileCount)
    
    
    'Fill the FileName and FileSize arrays
    For i = 1 To FileLength
        If b(i) <> 124 Then
           FileNames(j) = FileNames(j) & Chr(b(i))
        Else
            If j = FileCount Then
                NextReadPosition = i + 1
                j = 0
                Exit For
            Else
                j = j + 1
            End If
        End If
    Next i
    
    For i = NextReadPosition To FileLength
        If b(i) <> 124 Then
           FileLengths(j) = FileLengths(j) & Chr(b(i))
        Else
            If j = FileCount Then
                NextReadPosition = i + 1
                j = 0
                Exit For
            Else
                j = j + 1
            End If
        End If
    Next i

    'Extract the Files
    For j = 0 To FileCount
        ReDim b2(FileLengths(j) - 1)
    
        For i = NextReadPosition To NextReadPosition + FileLengths(j) - 1
            b2(l) = b(i)
            l = l + 1
        Next i
        l = 0

        'Set the status
        stbDragDrop.Panels(1).Text = "Extracting: " & FileNames(j)

        'Delete if file already exists with same name
        If Dir(DestinationPath & FileNames(j), vbDirectory) <> "" Then
            Kill (DestinationPath & FileNames(j))
        End If

        Open DestinationPath & FileNames(j) For Binary Access Write As #1
            Put #1, , b2
        Close #1
        NextReadPosition = NextReadPosition + FileLengths(j)
    Next j
    
    'Set the status
    stbDragDrop.Panels(1).Text = Status
    
    MsgBox FileCount + 1 & " Files Dercypted to " & Left(DestinationPath, Len(DestinationPath) - 1) & ".", vbInformation, "Crypto"
    
    DecodeArchive = True
    
    Exit Function
    
errHandler:

    MsgBox Err.Description, vbCritical, "LR File Ecrypt QL"
    DecodeArchive = False
    stbDragDrop.Panels(1).Text = Status
    
    
End Function




Private Sub tlbCrypto_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim Path As String
    Dim SourcePath As String

    If Button.Index = 1 Then
        Hide
        Form4.Show
        
    
    ElseIf Button.Index = 3 Then
    
        
        If lstKey.ListItems.Count <> 1 Then
            MsgBox "Missing Key. Please drag a file to use as the key into the box.", vbExclamation, "LR File Ecrypt QL"
            Exit Sub
        End If
        If txtPath.Text = "" Then
            MsgBox "The Path specified as the File Destination does not exist, please enter a correct Path.", vbExclamation, "LR File Ecrypt SQ"
            txtPath.SetFocus
            Exit Sub
        End If
        If Dir(txtPath.Text, vbDirectory) = "" Then
            MsgBox "The Path specified as the File Destination does not exist, please enter a correct Path.", vbExclamation, "LR File Ecrypt SQ"
            txtPath.Text = ""
            txtPath.SetFocus
            Exit Sub
        End If
        
        If optAction(0) Then
            If lstFiles.ListItems.Count = 0 Then
                MsgBox "No files were selected to Encrypt. Please make sure there is at least one file in the list.", vbExclamation, "LR File Ecrypt SQ"
                Exit Sub
            End If
        ElseIf optAction(1) Then
            If lstFiles.ListItems.Count <> 1 Then
                MsgBox "Only one *.cpt file can be Decrypted at once. Please make sure there is only one valid Archive in the list.", vbExclamation, "LR File Ecrypt SQ"
                Exit Sub
            End If
            If lstFiles.ListItems.Count = 1 Then
                If UCase(Right(lstFiles.ListItems(1).Text, 3)) <> "CPT" Then
                    MsgBox "The file you are attempting to Decrypt is not a valid *.cpt file. Please make sure there is only one valid Archive in the list.", vbExclamation, "LR File Ecrypt SQ"
                    Exit Sub
                End If
            End If
        End If

        
        
        If optAction(0) Then
            'Encrypt
        
            If Right(txtPath.Text, 1) <> "\" Then
                Path = txtPath.Text & "\" & "ARCHIVE_" & Format(Date, "MM-DD-YYYY") & " " & Format(Time, "h-mm-ss am/pm") & ".cpt"
            Else
                Path = txtPath.Text & "ARCHIVE_" & Format(Date, "MM-DD-YYYY") & " " & Format(Time, "h-mm-ss am/pm") & ".cpt"
            End If
        
            If Not CreateArchive(Path) Then
                MsgBox "The files were unable to be Encrypted.", vbExclamation, "LR File Ecrypt SQ"
            End If
             
            
        ElseIf optAction(1) Then
            'Decrypt
        
            If Right(txtPath.Text, 1) <> "\" Then
                Path = txtPath.Text & "\"
            End If
            
            SourcePath = lstFiles.ListItems(1).Text
        
            If Not DecodeArchive(SourcePath, Path) Then
                MsgBox "The Archive was not able to be Decrypted.", vbExclamation, "LR File Ecrypt SQ"
            End If
                
        End If

    End If
            
    

End Sub

Public Function Encr(OrigVal As Integer, KeyVal As Integer) As Integer

    If KeyVal = 0 Then
        Encr = OrigVal
    Else
        Encr = OrigVal - KeyVal
        If Encr = -255 Then
            Encr = 255
        ElseIf Encr < 0 Then
            Encr = 256 + Encr
        End If
    End If
    
End Function

Public Function Decr(EncrVal As Integer, KeyVal As Integer) As Integer
    
    If KeyVal = 0 Then
        Decr = EncrVal
    Else
        Decr = EncrVal + KeyVal
        If Decr = 510 Then
            Decr = 0
        ElseIf Decr > 255 Then
            Decr = Decr - 256
        End If
    End If
    
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub



