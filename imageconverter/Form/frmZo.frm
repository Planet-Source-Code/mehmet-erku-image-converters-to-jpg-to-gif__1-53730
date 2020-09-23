VERSION 5.00
Begin VB.Form frmZoom 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9255
   Icon            =   "frmZo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   3660
      Picture         =   "frmZo.frx":0D2A
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2205
   End
   Begin VB.Image Image2 
      Height          =   870
      Left            =   840
      Picture         =   "frmZo.frx":327D
      Top             =   480
      Width           =   6825
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As _
    Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As _
    String) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor _
    As Long) As Long
Private Const MAX_PATH As Long = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Type PanState
    x As Long
    y As Long
End Type
Private PanSet As PanState

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type


Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim incre As Double
    incre = 0.01
    If KeyAscii = 45 Then
        Image1.Visible = False
        Image1.Height = Image1.Height - (Image1.Height * incre)
        Image1.Width = Image1.Width - (Image1.Width * incre)
        Image1.Visible = True
    End If
    If KeyAscii = 43 Then
        Image1.Visible = False
        Image1.Height = Image1.Height + (Image1.Height * incre)
        Image1.Width = Image1.Width + (Image1.Width * incre)
        Image1.Visible = True
    End If
    SetImageInCenter
End Sub

Private Function OpenDirectoryTV(ohwnd As Long, Optional odtvTitle As String) As String
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
    szTitle = odtvTitle
    With tBrowseInfo
        .hwndOwner = ohwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        OpenDirectoryTV = sBuffer
    End If
End Function

Private Sub Form_Resize()
    On Error GoTo errh
    If Me.Windowstate = vbMinimized Then Exit Sub
    SetImageInCenter
errh:
End Sub

Private Sub SetImageInCenter()
    Image2.Left = (Me.Width - Image2.Width) / 2
    Image1.Move (ScaleWidth - Image1.Width) / 2, (ScaleHeight - Image1.Height) / 2
End Sub

Public Sub OpenImage(imagePath As String)
    'On Error GoTo karakterat
    Me.Caption = imagePath
    Image1.Stretch = False
    Image1.Picture = LoadPicture(imagePath)
    Image1.Stretch = True
    
    SetImageInCenter
    'karakterat:
    'Unload Me
    'Exit Sub
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And Shift = 0 Then
        PanSet.x = x
        PanSet.y = y
        MousePointer = vbSizePointer
    End If
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim nTop As Integer
  Dim nLeft As Integer
    MousePointer = vbSizePointer
    
    On Local Error Resume Next
    
    If Button = vbLeftButton And Shift = 0 Then
        
        With Image1
            nTop = -(.Top + (y - PanSet.y))
            nLeft = -(.Left + (x - PanSet.x))
        End With
        
        Image1.Move -nLeft, -nTop
        
    End If
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MousePointer = 0
    
End Sub
