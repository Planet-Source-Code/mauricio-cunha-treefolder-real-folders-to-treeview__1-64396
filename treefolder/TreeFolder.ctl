VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TreeFolder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   PropertyPages   =   "TreeFolder.ctx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   5655
   ToolboxBitmap   =   "TreeFolder.ctx":0035
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5040
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393217
      Indentation     =   34
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imgFolder"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList imgFolder 
      Left            =   4080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "TreeFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PATH As Long = 260
Private Const LVM_FIRST = &H1000

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const TEXT_RESOURCE_WORKSPACE = 4162
Private Const TEXT_RESOURCE_MYCOMPUTER = 9216
Private Const TEXT_RESOURCE_CONTROLPANEL = 4161

Private Const TEXT_RESOURCE_COL_NAME = 8976
Private Const TEXT_RESOURCE_COL_SIZE = 8978
Private Const TEXT_RESOURCE_COL_TYPE = 8979
Private Const TEXT_RESOURCE_COL_MODIFIED = 8980
Private Const TEXT_RESOURCE_COL_CREATED = 8996

Private Const ICON_RESOURCE_WORKSPACE = 34
Private Const ICON_RESOURCE_MYCOMPUTER = 16
Private Const ICON_RESOURCE_MYDOCUMENTS = 20
Private Const ICON_RESOURCE_NETWOORK = 17
Private Const ICON_RESOURCE_CONTROLPANEL = 35

Private Const FILE_SYSTEM_OBJECT_DRIVE_CDROM = 4
Private Const FILE_SYSTEM_OBJECT_DRIVE_FIXED = 2
Private Const FILE_SYSTEM_OBJECT_DRIVE_RAMDISK = 5
Private Const FILE_SYSTEM_OBJECT_DRIVE_REMOTE = 3
Private Const FILE_SYSTEM_OBJECT_DRIVE_REMOVABLE = 1
Private Const FILE_SYSTEM_OBJECT_DRIVE_UNKNOWN = 0

Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_ALIAS = 1024
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_ARCHIVE = 32
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_COMPRESSED = 2048
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_DIRECTORY = 16
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_HIDDEN = 2
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_NORMAL = 0
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_READONLY = 1
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_SYSTEM = 4
Private Const FILE_SYSTEM_OBJECT_FILEATTRIBUTE_VOLUME = 8

Private Const FILE_SYSTEM_OBJECT_SPECIALFOLDERS_SYSTEM = 1
Private Const FILE_SYSTEM_OBJECT_SPECIALFOLDERS_TEMPORARY = 2
Private Const FILE_SYSTEM_OBJECT_SPECIALFOLDERS_WINDOWS = 0

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type tSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Type TFILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
   
Private Type TWIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As TFILETIME
  ftLastAccessTime As TFILETIME
  ftLastWriteTime As TFILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Type TRANDYS_OWN_DRIVE_INFO
  DrvSectors As Long
  DrvBytesPerSector As Long
  DrvFreeClusters As Long
  DrvTotalClusters As Long
  DrvSpaceFree As Long
  DrvSpaceUsed As Long
  DrvSpaceTotal As Long
End Type

Public Enum EBorderStyle
 EBSNone = 0
 EBSFixedSingle = 1
End Enum

Public Enum ERootFolder
  ERFWorkSpace = 0
  ERFMyComputer = 1
  ERFControlPanel = 2
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As TWIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As TWIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferL As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferL As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As tSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pIdl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Const GWL_STYLE As Long = (-16)
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_WINDOWTEXT As Long = 8

Private Const TVI_ROOT    As Long = &HFFFF0000
Private Const TVI_FIRST   As Long = &HFFFF0001
Private Const TVI_LAST    As Long = &HFFFF0002
Private Const TVI_SORT    As Long = &HFFFF0003

Private Const TVIF_STATE As Long = &H8

'treeview styles
Private Const TVS_HASLINES As Long = 2
Private Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Private Const TVIS_BOLD   As Long = &H10

Private Const TV_FIRST As Long = &H1100
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_GETITEM As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM As Long = (TV_FIRST + 13)
Private Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Private Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Private Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Private Const TVGN_ROOT                 As Long = &H0
Private Const TVGN_NEXT                 As Long = &H1
Private Const TVGN_PREVIOUS             As Long = &H2
Private Const TVGN_PARENT               As Long = &H3
Private Const TVGN_CHILD                As Long = &H4
Private Const TVGN_FIRSTVISIBLE         As Long = &H5
Private Const TVGN_NEXTVISIBLE          As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE      As Long = &H7
Private Const TVGN_DROPHILITE           As Long = &H8
Private Const TVGN_CARET                As Long = &H9

Private Type TV_ITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type



Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32" _
                         (ByVal clr As OLE_COLOR, _
                          ByVal hPal As Long, _
                          dwRGB As Long) As Long


Private FSO As Object
Private FileInfo As tSHFILEINFO
Private NbFile As Long
Private FileFSToOpen As String
Private StringToFind As String
Private ProgressCancel As Boolean
Private TypeView
Private SizeOn As Boolean
Private OldX
Private InitialFormWith
Private DriveError As Boolean

Private mShell32 As String
Private mShowFolders As Boolean
Private mBackColor As OLE_COLOR
Private mHideFavorites As Boolean
Private mHideMyDocuments As Boolean
Private mHideNetwork As Boolean
Private mHideControlPanel As Boolean

Private NodeX As Node
Private NodeY As Node
Private Node0 As Node
Private Node1 As Node
Private mPath As String

Public Event AfterRename(Cancel As Integer, NewName As String)
Public Event BeforeRename(Cancel As Integer)
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event DriveCheck(DriveName As String, Value As Boolean)
Public Event DriveClick(DriveName As String)
Public Event DriveNotReady(DriveName As String)
Public Event FolderCheck(FolderName As String, Value As Boolean)
Public Event FolderClick(FolderName As String)
Public Event FolderDblClick(FolderName As String)
Public Event FolderExpand(FolderName As String)
Public Event FolderCollapse(FolderName As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Public Event Progress(Percent As Integer)
Public Event RootClick(RootType As ERootFolder)
Public Event RootDblClick(RootType As ERootFolder)
Public Event RootCheck(RootType As ERootFolder, Value As Boolean)
Public Event RootExpand(RootType As ERootFolder)
Public Event RootCollapse(RootType As ERootFolder)

Private Function CalcPercent(iValue As Integer, iMax As Integer) As Integer
  If iMax <= 0 Then
   CalcPercent = 0
  ElseIf iValue > iMax Then
   CalcPercent = 100
  Else
   CalcPercent = (iValue / iMax * 100)
  End If
End Function

Private Function GetSpecialPath(pFolder As Long) As String
  Dim sPath As String
  Dim IDL As Long
  Dim strPath As String
  Dim lngPos As Long
      
      If SHGetSpecialFolderLocation(0, pFolder, IDL) = 0 Then
          sPath = String(255, 0)
          SHGetPathFromIDListA IDL, sPath
          lngPos = InStr(sPath, Chr(0))
          If lngPos > 0 Then strPath = Left(sPath, lngPos - 1)
        GetSpecialPath = sPath
      End If
End Function

Private Function GetResourceStringFromFile(sModule As String, idString As Long) As String
  Dim hModule As Long
  Dim nChars As Long
  Dim Buffer As String * 260
  
     hModule = LoadLibrary(sModule)
     If hModule Then
        nChars = LoadString(hModule, idString, Buffer, 260)
        If nChars Then GetResourceStringFromFile = Left(Buffer, nChars)
        FreeLibrary hModule
     End If
End Function

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer, Optional IsLibrary As Boolean, Optional IconIDInLibrary As Long = -1, Optional IsShell32 As Boolean = True) As Long
  Dim SmallIcon As Long
  Dim NewImage As ListImage
  Dim IconIndex As Integer
  Dim ImageCheck As ListImage
  Dim IconCount As Long
  Dim LibLargeIcon As Long
  Dim LibSmallIcon As Long
      
      For Each ImageCheck In AddtoImageList.ListImages
       If LCase(ImageCheck.Key) = LCase(FileName) Then
        ExtractIcon = ImageCheck.Index
        Exit Function
       End If
      Next
      
      If IsLibrary = False Then
       GoSub Extract_Normal_Icon
      Else
       GoSub Extract_Library_Icon
      End If
      Exit Function
      
      
Extract_Normal_Icon:
      If PixelsXY = 16 Then
          SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
      Else
          SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
      End If
      
      If SmallIcon <> 0 Then
        With PictureBox
          .Height = 15 * PixelsXY
          .Width = 15 * PixelsXY
          .ScaleHeight = 15 * PixelsXY
          .ScaleWidth = 15 * PixelsXY
          .Picture = LoadPicture("")
          .AutoRedraw = True
          .Refresh
          SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
        End With
        
        IconIndex = AddtoImageList.ListImages.Count + 1
        Set NewImage = AddtoImageList.ListImages.Add(IconIndex, FileName, PictureBox.Image)
        ExtractIcon = IconIndex
      End If
  Return
  
  
Extract_Library_Icon:
   Dim FileToWorK As String
   
   If IsShell32 = True Then FileToWorK = mShell32 Else FileToWorK = FileName
   
   IconCount = ExtractIconEx(FileToWorK, -1, 0, 0, 0)
   If IconCount < 0 Then
    ExtractIcon = 0
    Exit Function
   End If
   
   
   Call ExtractIconEx(FileToWorK, IconIDInLibrary, LibLargeIcon, LibSmallIcon, 1)
   With PictureBox
      .Height = 15 * PixelsXY
      .Width = 15 * PixelsXY
      .ScaleHeight = 15 * PixelsXY
      .ScaleWidth = 15 * PixelsXY
      .Picture = LoadPicture("")
      .AutoRedraw = True
      Call DrawIconEx(.hdc, 0, 0, IIf(PixelsXY = 16, LibSmallIcon, LibLargeIcon), PixelsXY, PixelsXY, 0, 0, 3)
      .Refresh
   End With
   
   IconIndex = AddtoImageList.ListImages.Count + 1
   Set NewImage = AddtoImageList.ListImages.Add(IconIndex, FileName, PictureBox.Image)
   ExtractIcon = IconIndex
  Return
End Function

Private Function FormatFileSize(lFileSize As Double) As String
  Select Case lFileSize
      Case 0 To 1023
       FormatFileSize = Format(lFileSize, "##0") & " Bytes"
      Case 1024 To 1048575
       FormatFileSize = Format(lFileSize / 1024#, "#,##0") & " KB"
      Case 1024# ^ 2 To 1073741823
       FormatFileSize = Format(lFileSize / (1024# ^ 2), "#,##0.00") & " MB"
      Case Is > 1073741823#
      FormatFileSize = Format(lFileSize / (1024# ^ 3), "#,###,##0.00") & " GB"
  End Select
End Function

Public Property Get SelectedFolder() As String
  SelectedFolder = IIf(Not (TreeView1.SelectedItem Is Nothing), TreeView1.SelectedItem.Key, Empty)
End Property

Private Sub LoadTreeView(zPath As String)
  On Error GoTo err1
  
  Dim PrevDir As String
  Dim TempFile As String
  Dim I As Integer
  Dim R As Long
  Dim D As Object
  Dim DriveName As String
  Dim WinPath As String
  
  If TreeView1.Nodes.Count = 0 Then
    WinPath = FSO.GetSpecialFolder(FILE_SYSTEM_OBJECT_SPECIALFOLDERS_WINDOWS)
      
    PrevDir = "Root_WorkSpace"
    Set Node0 = TreeView1.Nodes.Add(, , PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_WORKSPACE))
    Node0.Image = ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_WORKSPACE)
      
    TempFile = FixPath(WinPath) & "explorer.exe"
    PrevDir = "Root_MyComputer"
    Set Node1 = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, GetResourceStringFromFile(mShell32, TEXT_RESOURCE_MYCOMPUTER))
    Node1.Image = ExtractIcon(TempFile, imgFolder, picBuffer, ICON_RESOURCE_MYCOMPUTER)
    
    For Each D In FSO.Drives
     DriveName = ""
     Select Case D.DriveType
      Case FILE_SYSTEM_OBJECT_DRIVE_REMOVABLE
       DriveName = GetResourceStringFromFile(mShell32, 9220)
      Case FILE_SYSTEM_OBJECT_DRIVE_FIXED
       DriveName = D.VolumeName
       If DriveName = "" Then DriveName = GetResourceStringFromFile(mShell32, 9397)
      Case FILE_SYSTEM_OBJECT_DRIVE_CDROM
       If D.IsReady = True Then DriveName = D.VolumeName
      Case FILE_SYSTEM_OBJECT_DRIVE_REMOTE
       DriveName = D.ShareName
     End Select
     
     Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, FixPath(D.Path), DriveName & " (" & D.Path & ")", ExtractIcon(FixPath(D.Path), imgFolder, picBuffer, 16))
         NodeX.Tag = 0
     Set NodeX = TreeView1.Nodes.Add(FixPath(D.Path), tvwChild, FixPath(D.Path) & "_empty_child")
    Next
    
    Node0.Expanded = True
    Node1.Expanded = True
  
    If mHideFavorites = False Then
      PrevDir = GetSpecialPath(&H6)
      Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16))
          NodeX.Tag = 0
      Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
    End If
    
    If mHideMyDocuments = False Then
      PrevDir = GetSpecialPath(&H5)
      Set NodeX = TreeView1.Nodes.Add(Node1.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_MYDOCUMENTS))
          NodeX.Tag = 0
      Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
    End If
    
    If mHideNetwork = False Then
      PrevDir = GetSpecialPath(&H13)
      Set NodeX = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, FSO.GetBaseName(PrevDir), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_NETWOORK))
          NodeX.Tag = 0
      Set NodeX = TreeView1.Nodes.Add(PrevDir, tvwChild, PrevDir & "_empty_child")
    End If
    
    If mHideControlPanel = False Then
      PrevDir = "Root_ControlPanel"
      Set NodeX = TreeView1.Nodes.Add(Node0.Key, tvwChild, PrevDir, GetResourceStringFromFile(FixPath(FSO.GetSpecialFolder(FILE_SYSTEM_OBJECT_SPECIALFOLDERS_SYSTEM)) & "shell32.dll", TEXT_RESOURCE_CONTROLPANEL), ExtractIcon(PrevDir, imgFolder, picBuffer, 16, True, ICON_RESOURCE_CONTROLPANEL))
          NodeX.Tag = 0
    End If
    
    GetFolders zPath
  End If
  Exit Sub
  
  
err1:
  If Err.Number > 0 Then
   Err.Raise Err.Number, Ambient.DisplayName & "_LoadTreeView", Err.Description
   Exit Sub
  End If
End Sub

Private Function StripNull(ByVal WhatStr As String) As String
  Dim pos As Integer
  pos = InStr(WhatStr, Chr(0))
  If pos > 0 Then StripNull = Left(WhatStr, pos - 1) Else StripNull = WhatStr
End Function

Private Sub Refresh()
 LoadTreeView FixPath(mPath)
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
 RaiseEvent AfterRename(Cancel, NewString)
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
 RaiseEvent BeforeRename(Cancel)
End Sub

Private Sub TreeView1_Click()
 RaiseEvent Click
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Image = 1 Or Node.Image = 2 Then Exit Sub
  If InStr(1, Node.Key, "Root_") > 0 Then
   If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootCollapse(ERFMyComputer)
   If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootCollapse(ERFWorkSpace)
   If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootCollapse(ERFControlPanel)
  Else
    RaiseEvent FolderCollapse(Node.Key)
  End If
  
  Node.Tag = 0
  While Node.Children > 1
    TreeView1.Nodes.Remove (Node.Child.Index)
  Wend
End Sub

Private Sub TreeView1_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
If Node.Tag = 0 Then
    TreeView1.Nodes.Remove (Node.Child.Index)
    Call GetFolders(Node.Key, Node.Checked)
    If InStr(1, Node.Key, "Root_") > 0 Then
     If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootExpand(ERFMyComputer)
     If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootExpand(ERFWorkSpace)
     If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootExpand(ERFControlPanel)
    Else
      RaiseEvent FolderExpand(Node.Key)
    RaiseEvent Change
    End If
    
    If DriveError = False Then
     Node.Tag = 1
    Else
     Node.Tag = 0
     DriveError = False
     Node.Parent.Expanded = False
    End If
End If
End Sub

Private Sub GetFolders(DirName As String, Optional bChecked As Boolean = False)
  Dim DirText As String
  Dim DirPath As String
  Dim I As Integer
  Dim J As Integer
  Dim L As Integer
  Dim B As Boolean
  Dim COL As New Collection
  
     On Error GoTo Errorhandle
     
     LockWindowUpdate TreeView1.hwnd
     Dir1.Path = DirName
     Dir1.Refresh
     
     For I = 0 To Dir1.ListCount - 1
          COL.Add Dir1.List(I)
     Next I
     
     For I = 1 To COL.Count
           RaiseEvent Progress(CalcPercent(I, COL.Count))
           DirText = ""
           L = Len(COL.Item(I))
           
           For J = L To 0 Step -1
            If Mid(COL.Item(I), J, 1) = "\" Then Exit For
           Next J
           
           DirText = Right(COL.Item(I), L - J)
           DirPath = FixPath(DirName) & DirText
           
           If DirName = GetSpecialPath(&H13) Then
            Set NodeY = TreeView1.Nodes.Add(DirName, tvwChild, COL.Item(I), DirText, ExtractIcon(DirPath, imgFolder, picBuffer, 16, True, 85))
            NodeY.Checked = bChecked
           Else
            Set NodeY = TreeView1.Nodes.Add(DirName, tvwChild, COL.Item(I), DirText, ExtractIcon(DirPath, imgFolder, picBuffer, 16))
            NodeY.Checked = bChecked
           End If
           Dir1.Path = COL.Item(I)
           Dir1.Refresh
           If Dir1.ListCount > 0 Then
               NodeY.Tag = 0
               NodeY.Checked = bChecked
               Set NodeY = TreeView1.Nodes.Add(COL.Item(I), tvwChild, COL.Item(I) & "\a")
           End If
     Next I
     LockWindowUpdate 0
     RaiseEvent Progress(0)
  Exit Sub
  
Errorhandle:
      RaiseEvent DriveNotReady(DirName)
      RaiseEvent Progress(0)
      DriveError = True
      Resume Next
End Sub

Private Sub SetFolder(PathFolder As String)
  On Error Resume Next
  Dim H As Integer
  Dim T As Integer
  Dim StrPT As String
  Dim StrOM As String
  
  H = 3
MainProcess:
  Do
   StrPT = TreeView1.Nodes(H).Key
    If InStr(1, StrPT, ":") > 0 Then
      StrOM = LCase(Mid(StrPT, InStr(1, StrPT, ":") - 1, 2) & Mid(StrPT, InStr(1, StrPT, ":") + 1))
      If Left(LCase(PathFolder), Len(StrOM)) = LCase(StrOM) Then
        TreeView1.Nodes(H).Expanded = True
        TreeView1.Nodes(H).Selected = True
        If TreeView1.Nodes(H).Children > 0 Then H = TreeView1.Nodes(H).Child.Index Else Exit Do
        GoTo MainProcess
        Exit Do
      End If
    If H = TreeView1.Nodes(H).LastSibling.Index Then Exit Do
    H = TreeView1.Nodes(H).Next.Index
   End If
  Loop
  TreeView1.Nodes(H).EnsureVisible
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
  If InStr(1, Node.Key, "Root_") > 0 Then
    If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootCheck(ERFMyComputer, Node.Checked)
    If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootCheck(ERFWorkSpace, Node.Checked)
    If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootCheck(ERFControlPanel, Node.Checked)
  Else
    If Len(Node.Key) = 3 Then RaiseEvent DriveCheck(Node.Key, Node.Checked)
    RaiseEvent FolderCheck(Node.Key, Node.Checked)
  End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  If InStr(1, Node.Key, "Root_") > 0 Then
   If InStr(1, LCase(Node.Key), "computer") Then RaiseEvent RootClick(ERFMyComputer)
   If InStr(1, LCase(Node.Key), "workspace") Then RaiseEvent RootClick(ERFWorkSpace)
   If InStr(1, LCase(Node.Key), "controlpanel") Then RaiseEvent RootClick(ERFControlPanel)
  Else
    If Len(Node.Key) = 3 Then RaiseEvent DriveClick(Node.Key)
    RaiseEvent FolderClick(Node.Key)
    RaiseEvent Change
    mPath = Node.Key
  End If
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
 ShellAbout UserControl.hwnd, "Explorer Controls by Mauricio Cunha", "Control to show files and folder like Windows Explorer." & vbCrLf & "Developed by Mauricio Cunha mcunha98@terra.com.br", Empty
End Sub

Private Sub TreeView1_OLECompleteDrag(Effect As Long)
 RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub TreeView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, state As Integer)
  RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, state)
End Sub

Private Sub TreeView1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
  RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub TreeView1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
  RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub TreeView1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Initialize()
 Set FSO = CreateObject("Scripting.FileSystemObject")
 mShell32 = FixPath(FSO.GetSpecialFolder(FILE_SYSTEM_OBJECT_SPECIALFOLDERS_SYSTEM)) & "shell32.dll"
 Call Refresh
End Sub

Private Sub UserControl_InitProperties()
 mBackColor = TranslateColor(vbWindowBackground)
 Path = CurDir
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
 BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
 TreeView1.CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
 Set TreeView1.Font = PropBag.ReadProperty("Font", Ambient.Font)
 TreeView1.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
 TreeView1.HideSelection = PropBag.ReadProperty("HideSelection", True)
 TreeView1.HotTracking = PropBag.ReadProperty("HotTracking", False)
 TreeView1.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
 Set TreeView1.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 TreeView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 TreeView1.OleDragMode = PropBag.ReadProperty("OleDragMode", 0)
 TreeView1.OleDropMode = PropBag.ReadProperty("OleDropMode", 0)
 TreeView1.SingleSel = PropBag.ReadProperty("SingleSel", False)
 HideControlPanel = PropBag.ReadProperty("HideControlPanel", False)
 HideFavorites = PropBag.ReadProperty("HideFavorites", False)
 HideMyDocuments = PropBag.ReadProperty("HideMyDocuments", False)
 HideNetwork = PropBag.ReadProperty("HideNetwork", False)
 Path = PropBag.ReadProperty("Path", CurDir)
 Style = PropBag.ReadProperty("Style", TreeView1.Style)
 LineStyle = PropBag.ReadProperty("LineStyle", TreeView1.LineStyle)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("BackColor", mBackColor, vbWindowBackground)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
 Call PropBag.WriteProperty("CheckBoxes", TreeView1.CheckBoxes, False)
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
 Call PropBag.WriteProperty("Font", TreeView1.Font, Ambient.Font)
 Call PropBag.WriteProperty("FullRowSelect", TreeView1.FullRowSelect, False)
 Call PropBag.WriteProperty("HideSelection", TreeView1.HideSelection, True)
 Call PropBag.WriteProperty("HideControlPanel", mHideControlPanel, False)
 Call PropBag.WriteProperty("HideFavorites", mHideFavorites, False)
 Call PropBag.WriteProperty("HideMyDocuments", mHideMyDocuments, False)
 Call PropBag.WriteProperty("HideNetwork", mHideNetwork, False)
 Call PropBag.WriteProperty("HotTracking", TreeView1.HotTracking, False)
 Call PropBag.WriteProperty("LabelEdit", TreeView1.LabelEdit, 0)
 Call PropBag.WriteProperty("MouseIcon", TreeView1.MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", TreeView1.MousePointer, 0)
 Call PropBag.WriteProperty("OleDragMode", TreeView1.OleDragMode, 0)
 Call PropBag.WriteProperty("OleDropMode", TreeView1.OleDropMode, 0)
 Call PropBag.WriteProperty("SingleSel", TreeView1.SingleSel, False)
 Call PropBag.WriteProperty("Style", TreeView1.Style, TreeStyleConstants.tvwTreelinesPlusMinusPictureText)
 Call PropBag.WriteProperty("LineStyle", TreeView1.LineStyle, TreeLineStyleConstants.tvwRootLines)
 Call PropBag.WriteProperty("ShowFolders", mShowFolders, True)
 Call PropBag.WriteProperty("Path", mPath, CurDir)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
  TreeView1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Function FixPath(sPath As String) As String
 FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
End Function

Public Property Get BorderStyle() As EBorderStyle
 BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal NewValue As EBorderStyle)
 UserControl.BorderStyle = NewValue
 PropertyChanged "BorderStyle"
End Property

Public Property Get Path() As String
 If Not (TreeView1.SelectedItem Is Nothing) Then
  Path = TreeView1.SelectedItem.Key
 Else
  Path = mPath
 End If
End Property
Public Property Let Path(ByVal NewValue As String)
 mPath = NewValue
 Call Refresh
 SetFolder NewValue
 PropertyChanged "Path"
 RaiseEvent Change
End Property

Public Property Get Font() As StdFont
 Set Font = TreeView1.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
 Set TreeView1.Font = NewValue
 PropertyChanged "Font"
End Property

Public Property Get FullRowSelect() As Boolean
 FullRowSelect = TreeView1.FullRowSelect
End Property
Public Property Let FullRowSelect(ByVal NewValue As Boolean)
 TreeView1.FullRowSelect = NewValue
 PropertyChanged "FullRowSelect"
End Property

Public Property Get CheckBoxes() As Boolean
 CheckBoxes = TreeView1.CheckBoxes
End Property
Public Property Let CheckBoxes(ByVal NewValue As Boolean)
 TreeView1.CheckBoxes = NewValue
 PropertyChanged "CheckBoxes"
 TreeView1.Refresh
End Property

Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 TreeView1.Enabled = NewValue
 PropertyChanged "Enabled"
End Property

Public Property Get HideSelection() As Boolean
 HideSelection = TreeView1.HideSelection
End Property
Public Property Let HideSelection(ByVal NewValue As Boolean)
 TreeView1.HideSelection = NewValue
 PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
 HotTracking = TreeView1.HotTracking
End Property
Public Property Let HotTracking(ByVal NewValue As Boolean)
 TreeView1.HotTracking = NewValue
 PropertyChanged "HotTracking"
End Property

Public Property Get LabelEdit() As LabelEditConstants
 LabelEdit = TreeView1.LabelEdit
End Property
Public Property Let LabelEdit(ByVal NewValue As LabelEditConstants)
 TreeView1.LabelEdit = NewValue
 PropertyChanged "LabelEdit"
End Property

Public Sub Clear()
 TreeView1.Nodes.Clear
End Sub

Public Property Get FolderChecked(Index As Long) As Boolean
 If ValidIndex(Index) = False Then Exit Property
 FolderChecked = TreeView1.Nodes(Index).Checked
End Property
Public Property Let FolderChecked(Index As Long, Value As Boolean)
 If ValidIndex(Index) = False Then Exit Property
 TreeView1.Nodes(Index).Checked = Value
End Property

Private Function ValidIndex(V As Long) As Boolean
 If TreeView1.Nodes.Count = 0 Then
  ValidIndex = False
  Exit Function
 ElseIf V > TreeView1.Nodes.Count Then
  ValidIndex = False
  Exit Function
 ElseIf V < 0 Then
  ValidIndex = False
  Exit Function
 Else
  ValidIndex = True
 End If
End Function

Public Property Get LineStyle() As TreeLineStyleConstants
  LineStyle = TreeView1.LineStyle
End Property
Public Property Let LineStyle(ByVal NewValue As TreeLineStyleConstants)
  TreeView1.LineStyle = NewValue
  PropertyChanged "LineStyle"
End Property

Public Property Get Style() As TreeStyleConstants
  Style = TreeView1.Style
End Property
Public Property Let Style(ByVal NewValue As TreeStyleConstants)
  TreeView1.Style = NewValue
  PropertyChanged "Style"
End Property

Public Property Get SelectedCount() As Long
  Dim SC As Long
  Dim I As Long
  
  SC = 0
  If TreeView1.Nodes.Count >= 1 Then
   For I = 1 To TreeView1.Nodes.Count
    If TreeView1.CheckBoxes = False Then
      If TreeView1.Nodes(I).Selected = True Then SC = SC + 1
    Else
      If TreeView1.Nodes(I).Checked = True Then SC = SC + 1
    End If
   Next I
  End If
  
  SelectedCount = SC
End Property

Public Property Get FolderCount() As Long
  If Not (TreeView1.SelectedItem Is Nothing) Then
   FolderCount = TreeView1.Nodes.Count
  End If
End Property

Public Property Get SubFolderCount() As Long
  SubFolderCount = TreeView1.SelectedItem.Children
End Property

Public Property Get MousePointer() As MSComctlLib.MousePointerConstants
 MousePointer = TreeView1.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MSComctlLib.MousePointerConstants)
 TreeView1.MousePointer = NewValue
 PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = TreeView1.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewValue As StdPicture)
 Set TreeView1.MouseIcon = NewValue
 PropertyChanged "MouseIcon"
End Property

Public Property Get OleDragMode() As MSComctlLib.OLEDragConstants
  OleDragMode = TreeView1.OleDragMode
End Property
Public Property Let OleDragMode(ByVal NewValue As MSComctlLib.OLEDragConstants)
  TreeView1.OleDragMode = NewValue
  PropertyChanged "OleDragMode"
End Property

Public Property Get OleDropMode() As MSComctlLib.OLEDropConstants
  OleDropMode = TreeView1.OleDropMode
End Property
Public Property Let OleDropMode(ByVal NewValue As MSComctlLib.OLEDropConstants)
  TreeView1.OleDropMode = NewValue
  PropertyChanged "OleDropMode"
End Property

Public Property Get SingleSel() As Boolean
  SingleSel = TreeView1.SingleSel
End Property
Public Property Let SingleSel(ByVal NewValue As Boolean)
  TreeView1.SingleSel = NewValue
  PropertyChanged "SingleSel"
End Property

Public Property Get HideFavorites() As Boolean
  HideFavorites = mHideFavorites
End Property
Public Property Let HideFavorites(ByVal NewValue As Boolean)
  mHideFavorites = NewValue
  PropertyChanged "HideFavorites"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideMyDocuments() As Boolean
  HideMyDocuments = mHideMyDocuments
End Property
Public Property Let HideMyDocuments(ByVal NewValue As Boolean)
  mHideMyDocuments = NewValue
  PropertyChanged "HideMyDocuments"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideNetwork() As Boolean
  HideNetwork = mHideNetwork
End Property
Public Property Let HideNetwork(ByVal NewValue As Boolean)
  mHideNetwork = NewValue
  PropertyChanged "HideNetwork"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get HideControlPanel() As Boolean
  HideControlPanel = mHideControlPanel
End Property
Public Property Let HideControlPanel(ByVal NewValue As Boolean)
  mHideControlPanel = NewValue
  PropertyChanged "HideControlPanel"
  TreeView1.Nodes.Clear
  Call Refresh
End Property

Public Property Get FolderNodes() As Nodes
  Set FolderNodes = TreeView1.Nodes
End Property

Public Property Get BackColor() As OLE_COLOR
  Dim NewValue As Long
  NewValue = SendMessage(TreeView1.hwnd, TVM_GETBKCOLOR, 0, ByVal 0)
  If NewValue = -1 Then NewValue = GetSysColor(COLOR_WINDOW)
  BackColor = NewValue
End Property
Public Property Let BackColor(NewValue As OLE_COLOR)
   Dim Style As Long
   mBackColor = TranslateColor(NewValue)
   Call SendMessage(TreeView1.hwnd, TVM_SETBKCOLOR, 0, ByVal mBackColor)
   Style = GetWindowLong(TreeView1.hwnd, GWL_STYLE)
   If Style And TVS_HASLINES Then
      Call SetWindowLong(TreeView1.hwnd, GWL_STYLE, Style Xor TVS_HASLINES)
      Call SetWindowLong(TreeView1.hwnd, GWL_STYLE, Style)
   End If
End Property

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
  If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
    TranslateColor = &HFFFF
  End If
End Function
