VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TreeFolder by Mauricio Cunha (mcunha98)"
   ClientHeight    =   4515
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAutoCheck 
      Caption         =   "Automatic check sub folders"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CheckBox chkCheckBoxes 
      Caption         =   "CheckBoxes"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ListBox lstHideRoot 
      Height          =   1185
      ItemData        =   "FrmTest.frx":0000
      Left            =   3840
      List            =   "FrmTest.frx":0010
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ComboBox cmbStyle 
      Height          =   315
      ItemData        =   "FrmTest.frx":0045
      Left            =   4680
      List            =   "FrmTest.frx":0061
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox cmbLineStyle 
      Height          =   315
      ItemData        =   "FrmTest.frx":00F8
      Left            =   4680
      List            =   "FrmTest.frx":0102
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin TreeFolderLib.TreeFolder TreeFolder1 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmTest.frx":011C
      OleDragMode     =   1
      OleDropMode     =   1
      ShowFolders     =   0   'False
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Hide Custom Roots:"
      Height          =   195
      Index           =   3
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Line Style:"
      Height          =   195
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Style:"
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   390
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Current path:"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   6615
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkCheckBoxes_Click()
  TreeFolder1.CheckBoxes = IIf(chkCheckBoxes.Value = 1, True, False)
  chkAutoCheck.Enabled = IIf(chkCheckBoxes.Value = 1, True, False)
End Sub

Private Sub chkEnabled_Click()
  TreeFolder1.Enabled = IIf(chkEnabled.Value = 1, True, False)
End Sub

Private Sub cmbLineStyle_Click()
  TreeFolder1.LineStyle = cmbLineStyle.ListIndex
End Sub

Private Sub cmbStyle_Click()
  TreeFolder1.Style = cmbStyle.ListIndex
End Sub

Private Sub Form_Load()
  Label1.Caption = TreeFolder1.Path
  
  cmbLineStyle.ListIndex = TreeFolder1.LineStyle
  cmbStyle.ListIndex = TreeFolder1.Style
  
  lstHideRoot.Selected(0) = TreeFolder1.HideControlPanel
  lstHideRoot.Selected(1) = TreeFolder1.HideFavorites
  lstHideRoot.Selected(2) = TreeFolder1.HideMyDocuments
  lstHideRoot.Selected(3) = TreeFolder1.HideNetwork
  
  chkCheckBoxes.Value = IIf(TreeFolder1.CheckBoxes = True, 1, 0)
  chkEnabled.Value = IIf(TreeFolder1.Enabled = True, 1, 0)
End Sub

Private Sub lstHideRoot_Click()
  TreeFolder1.HideControlPanel = lstHideRoot.Selected(0)
  TreeFolder1.HideFavorites = lstHideRoot.Selected(1)
  TreeFolder1.HideMyDocuments = lstHideRoot.Selected(2)
  TreeFolder1.HideNetwork = lstHideRoot.Selected(3)
End Sub

Private Sub TreeFolder1_Change()
  Label1.Caption = TreeFolder1.Path
End Sub

Private Sub TreeFolder1_DriveCheck(DriveName As String, Value As Boolean)
  If chkAutoCheck.Value = 1 Then
    Dim N As Node
    For Each N In TreeFolder1.FolderNodes
      If Left(N.Key, Len(DriveName)) = DriveName Then
        N.Checked = Value
      End If
    Next
  End If
End Sub

Private Sub TreeFolder1_FolderCheck(FolderName As String, Value As Boolean)
  If chkAutoCheck.Value = 1 Then
    Dim N As Node
    For Each N In TreeFolder1.FolderNodes
      If Left(N.Key, Len(FolderName)) = FolderName Then
        N.Checked = Value
      End If
    Next
  End If
End Sub
