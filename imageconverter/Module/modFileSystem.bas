Attribute VB_Name = "modFileSystem"
Option Explicit

Private Const SW_SHOWMAXIMIZED = 3
Private Const ArrGrow As Long = 5000
Private Const MaxLong As Long = 2147483647

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1

Enum eFileAttribute
ATTR_READONLY = &H1
ATTR_HIDDEN = &H2
ATTR_SYSTEM = &H4
ATTR_DIRECTORY = &H10
ATTR_ARCHIVE = &H20
ATTR_NORMAL = &H80
ATTR_TEMPORARY = &H100
End Enum

Enum eSortMethods
SortNot = 0
SortByNames = 1
End Enum

Enum eSizeConstants
BIPerB = 8
BPERKB = 1024
KBPerMB = 1024
MBPerGB = 1024
GBPerTB = 1024
TBPerPT = 1024
End Enum

Type tFile
    Name As String
    Path As String
    FullName As String
    CreationDate As String
    AccessDate As String
    WriteDate As String
    Size As Currency
    Attr As VbFileAttribute
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved As Long
    dwReserved1 As Long
    FileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp _
    As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As _
    Long) As Long

'File Stuff
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
    As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData _
    As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) _
    As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'Memory stuff
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length _
    As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, _
    ByVal Fill As Byte)

Public FileSearchCount As Long
Public FilesFound As Long
Public RecurseAmmount As Long
Public CurrentName As String
Public Abort As Boolean
Private CURWFD As WIN32_FIND_DATA
Type SHItemID
    cb      As Long
    abID    As Byte
End Type

Type ItemIDList
    mkid    As SHItemID
End Type
Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type


Private Sub Compress_RLE(ByteArray() As Byte)
  Dim OutStream() As Byte
  Dim X As Long
  Dim RLE_Count As Long
  Dim OutPos As Long
  Dim FileLong As Long
  Dim Char As Long
  Dim OldChar As Long
    
    ReDim OutStream(LBound(ByteArray) To UBound(ByteArray) * 1.33) 'Worst case
    FileLong = UBound(ByteArray)
    OldChar = -1
    
    For X = LBound(ByteArray) To UBound(ByteArray)
        Char = ByteArray(X)
        
        If Char = OldChar Then
            RLE_Count = RLE_Count + 1
            
            If RLE_Count < 4 Then
                OutStream(OutPos) = Char
                OutPos = OutPos + 1
            End If
            If RLE_Count = 258 Then
                OutStream(OutPos) = RLE_Count - 3
                OutPos = OutPos + 1
                RLE_Count = 0
                OldChar = -1
            End If
        Else
            If RLE_Count > 2 Then
                OutStream(OutPos) = RLE_Count - 3
                OutPos = OutPos + 1
            End If
            
            OutStream(OutPos) = Char
            OutPos = OutPos + 1
            RLE_Count = 1
            OldChar = Char
        End If
    Next
    
    If RLE_Count > 2 Then
        OutStream(OutPos) = RLE_Count - 3
        OutPos = OutPos + 1
    End If
    
    ReDim ByteArray(OutPos + 3)
    CopyMemory ByteArray(OutPos), FileLong, 4
    CopyMemory ByteArray(0), OutStream(0), OutPos
End Sub

Private Sub DeCompress_RLE(ByteArray() As Byte)
  Dim OutStream() As Byte
  Dim FileLong As Long
  Dim X As Long
  Dim Char As Long
  Dim OldChar As Long
  Dim RLE_Count As Long
  Dim OutPos As Long
  Dim RRun1 As Boolean
  Dim RRun2 As Boolean
    CopyMemory FileLong, ByteArray(UBound(ByteArray) - 3), 4
    ReDim OutStream(LBound(ByteArray) To FileLong)
    OldChar = -1
    
    For X = LBound(ByteArray) To UBound(ByteArray) - 4
        If RRun1 Then
            If RRun2 Then
                RLE_Count = ByteArray(X)
                If RLE_Count Then
                    FillMemory OutStream(OutPos), RLE_Count, Char
                    OutPos = OutPos + RLE_Count
                End If
                RRun1 = False
                RRun2 = False
                OldChar = -1
            Else
                Char = ByteArray(X)
                
                OutStream(OutPos) = Char
                OutPos = OutPos + 1
                
                If Char = OldChar Then RRun2 = True Else RRun1 = False: OldChar = Char
            End If
        Else
            Char = ByteArray(X)
            OutStream(OutPos) = Char
            OutPos = OutPos + 1
            
            If Char = OldChar Then RRun1 = True Else OldChar = Char
        End If
    Next
    
    ReDim ByteArray(0 To OutPos - 1)
    CopyMemory ByteArray(0), OutStream(0), OutPos
End Sub

Function Base64Enc(s As String) As String
  Static Enc() As Byte
  Dim b() As Byte
  Dim  Out() As Byte
  Dim  i As Long
  Dim  j As Long
  Dim  l As Long
    
    If (Not Val(Not Enc)) = 0 Then 'Null-Ptr = not initialized
        Enc = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
    End If
    
    l = Len(s): b = StrConv(s, vbFromUnicode)
    ReDim Preserve b(0 To (UBound(b) \ 3) * 3 + 2)
    ReDim Preserve Out(0 To (UBound(b) \ 3) * 4 + 3)
    For i = 0 To UBound(b) - 1 Step 3
        Out(j) = Enc(b(i) \ 4): j = j + 1
        Out(j) = Enc((b(i + 1) \ 16) Or (b(i) And 3) * 16): j = j + 1
        Out(j) = Enc((b(i + 2) \ 64) Or (b(i + 1) And 15) * 4): j = j + 1
        Out(j) = Enc(b(i + 2) And 63): j = j + 1
    Next
    For i = 1 To i - l
        Out(UBound(Out) - i + 1) = 61
    Next
    Base64Enc = StrConv(Out, vbUnicode)
End Function

Sub NONAME_Encode(data() As Byte, Key() As Byte)
  Const RANDOMSIZE As Long = 2047
  Dim RandomTable(RANDOMSIZE) As Byte
  Dim i As Long
  Dim KeyPos As Long
  Dim  RandomSeed As Long
  Dim LKey As Long
  Dim  UKey As Long
  Dim  LData As Long
  Dim  UData As Long
  Dim TotalAdd As Long
  Dim  CurKey As Long
    
    LKey = LBound(Key)
    UKey = UBound(Key)
    LData = LBound(data)
    UData = UBound(data)
    
    For i = LKey To UKey
        RandomSeed = RandomSeed + Key(i)
    Next
    Randomize RandomSeed
    
    For i = LBound(RandomTable) To UBound(RandomTable)
        RandomTable(i) = Int(Rnd * 256)
    Next
    
    KeyPos = LKey
    For i = LData To UData
        CurKey = Key(KeyPos)
        data(i) = data(i) Xor CurKey Xor RandomTable(TotalAdd) Xor (TotalAdd And 255)
        TotalAdd = ((RandomTable(CurKey) + TotalAdd) Xor CurKey) And RANDOMSIZE
        If KeyPos >= UKey Then KeyPos = LKey Else KeyPos = KeyPos + 1
    Next
End Sub

Sub NONAME_Encrypt(data() As Byte, Key() As Byte)
    Call NONAME_Encode(data, Key)
End Sub

Sub NONAME_Decrypt(data() As Byte, Key() As Byte)
    Call NONAME_Encode(data, Key)
End Sub

Function FileGetFirst(Path As String, data As tFile) As Long
    FileGetFirst = FindFirstFile(Path & "*", CURWFD)
    DataToFile Path, CURWFD, data
End Function

Function FileGetNext(Path As String, hSearch As Long, data As tFile) As Long
    FileGetNext = FindNextFile(hSearch, CURWFD)
    DataToFile Path, CURWFD, data
End Function

Sub DataToFile(Path As String, WFD As WIN32_FIND_DATA, data As tFile)
    With data
        'Strings need to be converted
        .Name = StripNulls(WFD.FileName)
        .Path = Path
        .Size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        .Attr = 0
        If WFD.dwFileAttributes And ATTR_ARCHIVE Then .Attr = .Attr Or vbArchive
        If WFD.dwFileAttributes And ATTR_DIRECTORY Then .Attr = .Attr Or vbDirectory
        If WFD.dwFileAttributes And ATTR_HIDDEN Then .Attr = .Attr Or vbHidden
        If WFD.dwFileAttributes And ATTR_NORMAL Then .Attr = .Attr Or vbNormal
        If WFD.dwFileAttributes And ATTR_READONLY Then .Attr = .Attr Or vbReadOnly
        If WFD.dwFileAttributes And ATTR_SYSTEM Then .Attr = .Attr Or vbSystem
    End With
End Sub

Private Function StripNulls(Str As String) As String
  Dim Pos As Long
    Pos = InStr(1, Str, vbNullChar)
    If Pos Then
        StripNulls = Left$(Str, Pos - 1)
    Else
        StripNulls = Str
    End If
End Function

Function OpenBinaryFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenBinaryFile = FreeFile
    
    If bWrite Then
        Open FilePath For Binary Access Write As #OpenBinaryFile
    Else
        Open FilePath For Binary Access Read As #OpenBinaryFile
    End If
End Function

Function OpenRandomFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenRandomFile = FreeFile
    
    If bWrite Then
        Open FilePath For Random Access Write As #OpenRandomFile
    Else
        Open FilePath For Random Access Read As #OpenRandomFile
    End If
End Function

Function OpenTextFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenTextFile = FreeFile
    
    If bWrite Then
        Open FilePath For Output As #OpenTextFile
    Else
        Open FilePath For Input As #OpenTextFile
    End If
End Function

Sub CloseFile(FileNumber As Integer)
    Close #FileNumber
End Sub

Function StartDoc(DocName As String) As Long
    StartDoc = ShellExecute(0, "Open", DocName, vbNullString, vbNullString, 1)
End Function

Function BrowseWebPage(PageName As String) As Long
    BrowseWebPage = ShellExecute(0, "Open", PageName, vbNullString, vbNullString, vbNullString)
End Function

Function Execute(FileName As String, Optional Windowstate As Long = vbMinimizedFocus) As Boolean
    On Error GoTo Handler
    Call Shell(FileName, Windowstate)
    Execute = True
    Handler:
End Function

Function SafeDelete(FilePath As String) As Long
  Dim FileNum As Integer
  Dim CurNum As Long
    'Resize the byte array
    
    For CurNum = 0 To 5
        'Generate a random byte array
        FileNum = OpenBinaryFile(FilePath, True)
        Do While EOF(FileNum) = False
            Put #FileNum, , CByte(Int(Rnd * 256))
        Loop
        CloseFile FileNum
    Next CurNum
    
    Kill (FilePath)
End Function

Sub SaveBytes(FilePath As String, Bytes() As Byte)
    FileClear FilePath
    
  Dim FileNum As Integer
    FileNum = FreeFile
    Open FilePath For Binary Access Write As #FileNum
    Put #FileNum, , Bytes()
    Close #FileNum
End Sub

Sub OpenBytes(FilePath As String, Bytes() As Byte)
  Dim FileNum As Integer
    FileNum = FreeFile
    
    Open FilePath For Binary Access Read As #FileNum
    ReDim Bytes(0 To LOF(FileNum) - 1)
    Get #FileNum, , Bytes()
    Close #FileNum
End Sub

Function LoadTextFile(FilePath As String) As String
  Dim FileNum As Integer
    FileNum = FreeFile
    
    Open FilePath For Input As #FileNum
    LoadTextFile = Input(LOF(FileNum), #FileNum)
    Close #FileNum
End Function

Sub SaveTextFile(TheString As String, FilePath As String)
  Dim FileNum As Integer
    FileNum = FreeFile
    
    Open FilePath For Output As #FileNum
    Print #FileNum, TheString
    Close #FileNum
End Sub
'Folder

Function MakeFolder(FolderName As String) As Boolean
    On Error GoTo Handler
    Call MkDir(FolderName)
    MakeFolder = True
    Handler:
End Function

Function DeleteFolder(FolderName As String) As Boolean
    On Error GoTo Handler
    Call RmDir(FolderName)
    DeleteFolder = True
    Handler:
End Function

Function FolderExists(FolderName As String) As Boolean
    On Error GoTo Handler
    If Dir$(FolderName, vbDirectory) <> vbNullString Then FolderExists = True
    Handler:
End Function
'File

Function FileMove(Src As String, Dest As String) As Boolean
    On Error GoTo Handler
    FileCopy Src, Dest
    FileMove = True
    Handler:
End Function

Function FileMake(FileName As String) As Boolean
    On Error GoTo Handler
  Dim FileNum As Integer
    FileNum = OpenBinaryFile(FileName, True)
    CloseFile FileNum
    FileMake = True
    Handler:
End Function

Function FileExists(FileName As String) As Boolean
    On Error GoTo Handler
    If Dir$(FileName) <> vbNullString Then FileExists = True
    Handler:
End Function

Function FileClear(FilePath As String)
    On Error GoTo Handler
    
  Dim FileNum As Integer
    FileNum = OpenTextFile(FilePath, True)
    Print #FileNum, vbNullString
    CloseFile FileNum
    
    FileClear = True
    Handler:
End Function

Function GetDirectoryFolders(Directory As String) As String()
  Dim TheName As String
  Dim count As Long
  Dim Names() As String
    ReDim Names(ArrGrow)
    
    TheName = Dir$(Directory, vbDirectory)
    
    Do While TheName <> vbNullString
        If TheName <> "." And TheName <> ".." Then
            If (GetAttr(Directory & TheName) And vbDirectory) Then
                If count > UBound(Names) Then ReDim Preserve Names(count + ArrGrow)
                Names(count) = TheName
                count = count + 1
            End If
        End If
        TheName = Dir$
    Loop
    ReDim Preserve Names(0 To count - 1)
    
    GetDirectoryFolders = Names
End Function

Function GetDirectoryFiles(Directory As String) As String()
  Dim TheName As String
  Dim count As Long
  Dim Names() As String
    ReDim Names(ArrGrow)
    
    TheName = Dir$(Directory, vbNormal)
    
    Do While TheName <> vbNullString
        If TheName <> "." And TheName <> ".." Then
            If count > UBound(Names) Then ReDim Preserve Names(count + ArrGrow)
            Names(count) = TheName
            count = count + 1
        End If
        TheName = Dir$
    Loop
    ReDim Preserve Names(0 To count - 1)
    
    GetDirectoryFiles = Names
End Function

Function GetDirectoryFoldersAndFiles(Directory As String) As tFile()
  Dim count As Long
  Dim Files() As tFile
    ReDim Files(0 To ArrGrow)
    FileSearchCount = 0
    
    AddFoldersAndFiles Directory, count, Files
    ReDim Preserve Files(count - 1)
    GetDirectoryFoldersAndFiles = Files
End Function

Function AddFile(Files() As tFile, count As Long, Name As String, Path As String)
    If count > UBound(Files) Then ReDim Files(0 To count + ArrGrow)
    With Files(count)
        .Name = Name
        .Path = Path
        .Attr = GetAttr(.Path & .Name)
        If (.Attr And vbDirectory) = 0 Then .Size = FileLen(.Path & .Name)
    End With
    count = count + 1
End Function

Sub AddFilesToListBox(TheListBox As ListBox, Files() As tFile)
    TheListBox.Clear
    
    LockWindowUpdate TheListBox.hwnd
  Dim i As Long
    For i = LBound(Files) To UBound(Files)
        Call TheListBox.AddItem(Files(i).Path & Files(i).Name)
        SetAttr Files(i).Path & Files(i).Name, vbNormal
        
    Next
    LockWindowUpdate 0
End Sub

Function FolderFind(Directory As String, Optional Filter As String, Optional eSortMethods As eSortMethods = _
        SortNot, Optional MinSize As Long = 0, Optional MaxSize As Long = MaxLong) As tFile()
  Dim i As Long
  Dim Files() As tFile
  Dim count As Long
  Dim Count2 As Long
  Dim PrevCount As Long
  Dim StartCount As Long
  Dim Added As Boolean
  Dim FilteredFiles() As tFile
    ReDim Files(ArrGrow)
    ReDim FilteredFiles(ArrGrow)
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    Added = True
    FileSearchCount = 0
    Call AddFoldersAndFiles(Directory, count, Files)
    
    If count Then
        Do
            StartCount = count
            For i = PrevCount To count - 1
                If Files(i).Attr And vbDirectory Then
                    AddFoldersAndFiles Files(i).Path & Files(i).Name, count, Files
                End If
            Next
            
            If PrevCount = count Then Exit Do
            PrevCount = StartCount
        Loop
    End If
    
    For i = LBound(Files) To count - 1
        With Files(i)
            If InStr(1, .Name, Filter, vbTextCompare) <> 0 Then
                If .Size >= MinSize And .Size <= MaxSize Then
                    If Count2 > UBound(FilteredFiles) Then ReDim Preserve FilteredFiles(0 To Count2 + ArrGrow)
                    
                    FilteredFiles(Count2) = Files(i)
                    Count2 = Count2 + 1
                End If
            End If
        End With
    Next
    
    If Count2 Then ReDim Preserve FilteredFiles(Count2 - 1)
    
    Select Case eSortMethods
    Case SortByNames
        Call FileSortName(FilteredFiles, LBound(FilteredFiles), UBound(FilteredFiles), -1)
    End Select
    
    If Count2 Then FolderFind = FilteredFiles
End Function

Function AddFoldersAndFiles(Directory As String, count As Long, Files() As tFile) As Long
  Dim File As tFile
  Dim hSearch As Long
    
    hSearch = FindFirstFile(Directory & "*", CURWFD) 'FileGetFirst(Directory, File)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function
    
    Do
        If File.Name <> "." And File.Name <> ".." And Len(File.Name) <> 0 Then
            DoEvents    'Translate messages
            If count > UBound(Files) Then ReDim Preserve Files(count + ArrGrow)
            With Files(count)
                .Path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .Path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.Path & File.Name
                End If
                
                count = count + 1
            End With
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File)
    
    FindClose hSearch
End Function

Function GetRecurseFolders(ByVal Directory As String, count As Long, Files() As tFile) As Long
  Dim File As tFile
  Dim StartCount As Long
  Dim  i As Long
  Dim  hSearch As Long
    StartCount = count
    
    hSearch = FindFirstFile(Directory & "*", CURWFD)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function
    
    Do
        If File.Name <> "." And File.Name <> ".." And File.Name <> vbNullString Then
            DoEvents    'Translate messages
            If count > UBound(Files) Then ReDim Preserve Files(count + ArrGrow)
            With Files(count)
                .Path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .Path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.Path & File.Name
                End If
            End With
            
            count = count + 1
            
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File)
    
    For i = StartCount To count - 1
        If Files(i).Attr And vbDirectory Then Call GetRecurseFolders(Files(i).FullName, count, Files)
    Next
    FindClose hSearch
End Function

Function GetRecurseFoldersListBox(TheListBox As ListBox, ByVal Directory As String, Filter As String, count _
        As Long, Files() As tFile) As Long
  Dim File As tFile
  Dim  StartCount As Long
  Dim  i As Long
  Dim  hSearch As Long
    StartCount = count
  Dim varible
    hSearch = FindFirstFile(Directory & "*", CURWFD)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function
    
    Do
        If File.Name <> "." And File.Name <> ".." And File.Name <> vbNullString Then
            DoEvents    'Translate messages
            If count > UBound(Files) Then ReDim Preserve Files(count + ArrGrow)
            With Files(count)
                .Path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .Path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.Path & File.Name
                End If
            End With
            
            count = count + 1
            
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File) <> 0 And (Abort = False)
    FindClose hSearch
    
    For i = StartCount To count - 1
        If InStrRev(Files(i).Name, Filter, , vbTextCompare) Then
            TheListBox.AddItem Files(i).FullName
            FilesFound = FilesFound + 1
            
            'buraya ekle
            If Files(i).Attr <> vbDirectory And InStrRev(Right(Files(i).Name, 3), formsrc.Text1, , vbTextCompare) _
                Then
                
                Set varible = formsrc.lvwHD.ListItems.Add(, , Files(i).FullName)
                varible.SubItems(1) = Right(Files(i).FullName, 3)
                varible.SubItems(2) = Files(i).Size
                varible.SubItems(3) = FileDateTime(Files(i).FullName)
                varible.SubItems(4) = "0"
                varible.SubItems(5) = formsrc.Combo1
            End If
        End If
        If Files(i).Attr And vbDirectory Then GetRecurseFoldersListBox TheListBox, Files(i).FullName, Filter, count, _
            Files
    Next
End Function

Function FileSearch(ListBox As ListBox, Directory As String, Filter As String) As tFile()
  Dim Files() As tFile
  Dim count As Long
    Call SearchStart(Files)
    ListBox.Clear
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    Call GetRecurseFoldersListBox(ListBox, Directory, Filter, count, Files)
    ReDim Preserve Files(0 To count - 1)
    FileSearch = Files
End Function

Private Sub SearchStart(Files() As tFile)
    ReDim Files(ArrGrow)
    Abort = False
    FileSearchCount = 0
    FilesFound = 0
End Sub

Function GetFoldersAndFiles(ByVal Directory As String) As tFile()
  Dim Files() As tFile
  Dim count As Long
    ReDim Files(0)
    FileSearchCount = 0
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    Call GetRecurseFolders(Directory, count, Files)
    ReDim Preserve Files(0 To count - 1)
    GetFoldersAndFiles = Files
End Function

Private Sub FileSortName(Arr() As tFile, lLbound As Long, lUbound As Long, Direction As Long)
    If lUbound <= lLbound Then Exit Sub
    
  Static Buffer As tFile
  Dim Compare As String
  Dim CurHigh As Long
  Dim CurLow As Long
    
    CurLow = lLbound
    CurHigh = lUbound
    Compare = Arr((lLbound + lUbound) \ 2).FullName
    
    Do While CurLow <= CurHigh
        Do While StrComp(Arr(CurLow).FullName, Compare, vbTextCompare) = Direction And CurLow <> lUbound: CurLow _
            = CurLow + 1: Loop
            Do While StrComp(Compare, Arr(CurHigh).FullName, vbTextCompare) = Direction And CurHigh <> lLbound: _
                CurHigh = CurHigh - 1: Loop
                
                If CurLow <= CurHigh Then
                    Buffer = Arr(CurLow)
                    Arr(CurLow) = Arr(CurHigh)
                    Arr(CurHigh) = Buffer
                    CurLow = CurLow + 1
                    CurHigh = CurHigh - 1
                End If
            Loop
            
            If lLbound < CurHigh Then FileSortName Arr(), lLbound, CurHigh, Direction
            If CurLow < lUbound Then FileSortName Arr(), CurLow, lUbound, Direction
        End Sub
