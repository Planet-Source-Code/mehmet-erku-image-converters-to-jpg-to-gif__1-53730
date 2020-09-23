VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formsrc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ýmage Converter"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11040
   Icon            =   "File Searcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   563
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   736
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Text            =   "bmp"
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   2400
      ScaleHeight     =   4095
      ScaleWidth      =   6735
      TabIndex        =   18
      Top             =   8040
      Visible         =   0   'False
      Width           =   6735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   8100
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   735
      Left            =   4680
      TabIndex        =   12
      Top             =   600
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "..To BitMap"
         Height          =   375
         Left            =   4920
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..To Gif"
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Convert ... To Jpeg"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Text            =   "50"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   " Quality %"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   20
      Left            =   4560
      Top             =   0
   End
   Begin VB.ListBox lstResult 
      Height          =   1425
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   10695
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "."
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   6615
   End
   Begin MSComctlLib.ListView lvwHD 
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8880
      TabIndex        =   22
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Pattern"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulunanlar / Files Found"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblCurpath 
      AutoSize        =   -1  'True
      Caption         =   "Current path"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Current path"
      Top             =   7680
      Width           =   870
   End
   Begin VB.Label lblFilesFound 
      AutoSize        =   -1  'True
      Caption         =   "Files Found"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Files Found"
      Top             =   7320
      Width           =   810
   End
   Begin VB.Label lblFilesSearched 
      AutoSize        =   -1  'True
      Caption         =   "Files Searched"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Files Searched"
      Top             =   6960
      Width           =   1050
   End
   Begin VB.Label lblFilter 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "formsrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by mehmet erkuþ
'12.05.2004
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const MAX_PATH = 1024
Private Const CSIDL_TEMPLATES = &H15 'ShellNew folder
Private WithEvents cGif As GIF
Attribute cGif.VB_VarHelpID = -1

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) _
    As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As _
    Long, pidl As ItemIDList) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As _
    String) As Long

Private lpIDList As Long
Private sBuffer As String
Private szTitle As String
Private tBrowseInfo As BrowseInfo
Private RecVal


Private Sub cmdSearch_Click()
    Abort = False
    
  Dim varible
  Dim lWidth
  Dim a
  Dim b
  Dim c
  Dim n
  Dim d As String
    
    lWidth = lvwHD.Width / 7
    lvwHD.ColumnHeaders.Clear
    lvwHD.ColumnHeaders.Add , , "File [Click item for View]", lWidth
    lvwHD.ColumnHeaders.Add , , "Pattern", lWidth
    lvwHD.ColumnHeaders.Add , , "Length", lWidth
    lvwHD.ColumnHeaders.Add , , "File Date-Time", lWidth
    lvwHD.ColumnHeaders.Add , , "Converted Size", lWidth
    lvwHD.ColumnHeaders.Add , , " Quality", lWidth
    lvwHD.ColumnHeaders.Add , , "Created File", lWidth
    
    lvwHD.View = lvwReport
    lvwHD.ListItems.Clear
    
    Call FileSearch(lstResult, txtDir, txtFilter)
    DoEvents
    Abort = True
    
End Sub

Private Sub cmdStop_Click()
    Abort = True
End Sub

Sub Command1_Click()
    'On Error GoTo Exit_Proc
    
    szTitle = "Select Folder"
    cmdSearch.Enabled = False
    
    ' Get folder from user
    szTitle = "Select Folder"
    
    ' Get folder from user
    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        cmdSearch.Enabled = True
        
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then
            sBuffer = sBuffer & "\"
        End If
        
        txtDir.Text = sBuffer
    Else
        
    End If
    
    'Exit_Proc:
    ' Exit Sub
    
End Sub

Private Sub Command2_Click()
    Abort = False
    
  Dim i As Integer
  Dim filebmp
  Dim filejpg As String
  Dim si As String
  Dim c As New cDIBSection
  Dim qual1
    
    lstResult.Clear
    ProgressBar1.Max = lvwHD.ListItems.count
    
    For i = 1 To lvwHD.ListItems.count
        
        filebmp = lvwHD.ListItems(i).Text 'Left(dlg.FileName, Len(dlg.FileName) - 4)
        Picture2.Picture = LoadPicture(filebmp)
        
        lstResult.AddItem filebmp
        lvwHD.ListItems(i).Selected = True
        lvwHD.ListItems(i).EnsureVisible '= True
        lvwHD.ListItems(i).Checked = True
        DoEvents
        
        filejpg = Left(lvwHD.ListItems(i).Text, Len(lvwHD.ListItems(i).Text) - 4) & ".jpg"
        'Saving the file as Bitmap
        'SavePicture picMain.Picture, FileName & ".bmp"
        'Change the bitmap file to the jpg file with the Bmp2Jpg.dll
        '100 is the Compress Quality
        If Dir(filejpg, vbReadOnly) <> "" Then
            SetAttr filejpg, vbNormal
            
        End If
        
        si = filejpg 'fileToSave
        c.CreateFromPicture Picture2.Picture
        qual1 = Combo1
        SaveJPG c, si, qual1
        ProgressBar1.Value = i
        Label5.Caption = Int((100 / ProgressBar1.Max) * i) & " % "
        DoEvents
        lvwHD.SelectedItem.SubItems(4) = FileLen(filejpg)
        lvwHD.SelectedItem.SubItems(5) = Combo1
        lvwHD.SelectedItem.SubItems(6) = filejpg
        
        'Deleting the bitmap file
        'Kill FileName & ".bmp"
    Next i
    
        DoEvents
    Abort = True

End Sub

Private Sub Command3_Click()
    Abort = False
    
  Dim i As Integer
  Dim filebmp
  Dim filejpg As String
  Dim si As String
  Dim c As New cDIBSection
  Dim qual1
    
    lstResult.Clear
    ProgressBar1.Max = lvwHD.ListItems.count
    
    For i = 1 To lvwHD.ListItems.count
        
        filebmp = lvwHD.ListItems(i).Text 'Left(dlg.FileName, Len(dlg.FileName) - 4)
        Picture2.Picture = LoadPicture(filebmp)
        
        lstResult.AddItem filebmp
        lvwHD.ListItems(i).Selected = True
        lvwHD.ListItems(i).EnsureVisible
        lvwHD.ListItems(i).Checked = True
        DoEvents
        
        filejpg = Left(lvwHD.ListItems(i).Text, Len(lvwHD.ListItems(i).Text) - 4) & ".gif"
        If Dir(filejpg, vbReadOnly) <> "" Then
            SetAttr filejpg, vbNormal
            
        End If
        
        'Saving the file as Bitmap
        'SavePicture picMain.Picture, FileName & ".bmp"
        'Change the bitmap file to the jpg file with the Bmp2Jpg.dll
        '100 is the Compress Quality
        si = filejpg 'fileToSave
        c.CreateFromPicture Picture2.Picture
        qual1 = Combo1
        
        Set cGif = New GIF
        cGif.SaveGIF Picture2.Picture, filejpg, Picture2.hDc, 1, Picture2.Point(0, 0)
        
        SaveJPG c, si, qual1
        ProgressBar1.Value = i
        Label5.Caption = Int((100 / ProgressBar1.Max) * i) & " % "
        DoEvents
        lvwHD.SelectedItem.SubItems(4) = FileLen(filejpg)
        lvwHD.SelectedItem.SubItems(5) = Combo1
        lvwHD.SelectedItem.SubItems(6) = filejpg
        
        If Abort = True Then
            Exit For
            Exit Sub
        End If
        
        'Deleting the bitmap file
        'Kill FileName & ".bmp"
    Next i
        DoEvents
    Abort = True

End Sub

Private Sub Command4_Click()
    Abort = False
    
  Dim i As Integer
  Dim filebmp
  Dim filejpg As String
  Dim si As String
  Dim c As New cDIBSection
  Dim qual1
    
    lstResult.Clear
    ProgressBar1.Max = lvwHD.ListItems.count
    
    For i = 1 To lvwHD.ListItems.count
        
        filebmp = lvwHD.ListItems(i).Text 'Left(dlg.FileName, Len(dlg.FileName) - 4)
        Picture2.Picture = LoadPicture(filebmp)
        
        lstResult.AddItem filebmp
        lvwHD.ListItems(i).Selected = True
        lvwHD.ListItems(i).EnsureVisible
        lvwHD.ListItems(i).Checked = True
        DoEvents
        
        filejpg = Left(lvwHD.ListItems(i).Text, Len(lvwHD.ListItems(i).Text) - 4) & ".bmp"
        'Saving the file as Bitmap
        If Dir(filejpg, vbReadOnly) <> "" Then
            SetAttr filejpg, vbNormal
            
        End If
        
        SavePicture Picture2.Picture, filejpg
        ProgressBar1.Value = i
        Label5.Caption = Int((100 / ProgressBar1.Max) * i) & " % "
        DoEvents
        lvwHD.SelectedItem.SubItems(4) = FileLen(filejpg)
        lvwHD.SelectedItem.SubItems(5) = Combo1
        lvwHD.SelectedItem.SubItems(6) = filejpg
        
        'Deleting the bitmap file
        'Kill FileName & ".bmp"
        If Abort = True Then
            Exit For
            Exit Sub
        End If
    Next i
       DoEvents
    Abort = True
 
End Sub

Private Sub Command5_Click()
    MsgBox "This app modified by mehmet erkuþ with psc code ...."
End Sub

Private Sub Form_Load()
  Dim varible
  Dim lWidth
  Dim a
  Dim b
  Dim c
  Dim n
  Dim d As String
    lvwHD.Visible = True
    lWidth = lvwHD.Width / 7
    lvwHD.ColumnHeaders.Clear
    lvwHD.ColumnHeaders.Add , , "Dosya Adý", lWidth
    lvwHD.ColumnHeaders.Add , , "Pattern", lWidth
    lvwHD.ColumnHeaders.Add , , "Toplam Boyutu", lWidth
    lvwHD.ColumnHeaders.Add , , "Eriþim", lWidth
    lvwHD.ColumnHeaders.Add , , "Kalite", lWidth
    lvwHD.ColumnHeaders.Add , , "Converted Size", lWidth
    
    lvwHD.ColumnHeaders.Add , , "Updated File", lWidth
    
    lvwHD.View = lvwReport
    lvwHD.ListItems.Clear
    For n = 1 To 100
        Combo1.AddItem n
    Next n
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label2_Click()
    lstResult.Clear
    
End Sub

Private Sub lstResult_DblClick()
    StartDoc lstResult.Text
End Sub

Private Sub lvwHD_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo errprx
    If (Abort = True) And Dir(lvwHD.SelectedItem.Text, vbNormal) <> "" Then
        frmZoom.Show , Me
        frmZoom.OpenImage (lvwHD.SelectedItem.Text)
        
    End If
errprx:
    Exit Sub
    
End Sub

Private Sub tmrUpdate_Timer()
    lblFilesFound = "Bulunan Dosya: " & FilesFound
    lblFilesSearched = "Bakýlan Dosya : " & FileSearchCount
    lblCurpath = CurrentName
End Sub
