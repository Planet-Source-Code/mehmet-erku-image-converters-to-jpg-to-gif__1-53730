Attribute VB_Name = "mIntelJPEGLibrary"
Option Explicit
' ==================================================================================
' Requires:    cDIBSectionmod.cls
'              ijl15.dll (Intel)
' An interface to Intel's IJL (Intel JPG Library) for use in VB.
' ==================================================================================

Private Enum IJLERR
    IJL_OK = 0
End Enum

Private Enum IJLIOTYPE
    ''// Write an entire JFIF bit stream.
    IJL_JFILE_WRITEWHOLEIMAGE = 8&
End Enum

Type JPEG_CORE_PROPERTIES_VB
    UseJPEGPROPERTIES As Long                      '// default = 0
    '// DIB specific I/O data specifiers.
    DIBBytes As Long ';                  '// default = NULL 4
    DIBWidth As Long ';                  '// default = 0 8
    DIBHeight As Long ';                 '// default = 0 12
    DIBPadBytes As Long ';               '// default = 0 16
    DIBChannels As Long ';               '// default = 3 20
    DIBColor As Long ';                  '// default = IJL_BGR 24
    DIBSubsampling As Long  ';            '// default = IJL_NONE 28
    '// JPEG specific I/O data specifiers.
    JPGFile As Long 'LPTSTR              JPGFile;                32   '// default = NULL
    JPGBytes As Long ';                  '// default = NULL 36
    JPGSizeBytes As Long ';              '// default = 0 40
    JPGWidth As Long ';                  '// default = 0 44
    JPGHeight As Long ';                 '// default = 0 48
    JPGChannels As Long ';               '// default = 3
    JPGColor As Long           ';                  '// default = IJL_YCBCR
    JPGSubsampling As Long  ';            '// default = IJL_411
    JPGThumbWidth As Long ' ;             '// default = 0
    JPGThumbHeight As Long ';            '// default = 0
    '// JPEG conversion properties.
    cconversion_reqd As Long ';          '// default = TRUE
    upsampling_reqd As Long ';           '// default = TRUE
    jquality As Long ';                  '// default = 75.  90 is my preferred quality setting.
    '// Low-level properties - 20,000 bytes.  If the whole structure
    ' is written out then VB fails with an obscure error message
    ' "Too Many Local Variables" !
    ' These all default if they are not otherwise specified so there
    ' is no trouble.
    jprops(0 To 19999) As Byte
End Type

Private Declare Function ijlInit Lib "ijl15.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl15.dll" (jcprops As Any) As Long
Private Declare Function ijlWrite Lib "ijl15.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( lpvDest As Any, lpvSource As Any, ByVal cbCopy _
    As Long)


Public Function SaveJPG( ByRef cDib As cDIBSection, ByVal sFile As String, Optional ByVal lQuality As Long _
        = 90 ) As Boolean
    
  Dim tJ As JPEG_CORE_PROPERTIES_VB
  Dim bFile() As Byte
  Dim lptr As Long
  Dim lR As Long
    
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
        ' Set up the DIB information:
        ' Store DIBWidth:
        tJ.DIBWidth = cDib.Width
        ' Store DIBHeight:
        tJ.DIBHeight = -cDib.Height
        ' Store DIBBytes (pointer to uncompressed JPG data):
        tJ.DIBBytes = cDib.DIBSectionBitsPtr
        ' Very important: tell IJL how many bytes extra there
        ' are on each DIB scan line to pad to 32 bit boundaries:
        tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
        
        ' Set up the JPEG information:
        
        ' Store JPGFile:
        bFile = StrConv(sFile, vbFromUnicode)
        ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
        bFile(UBound(bFile)) = 0
        lptr = VarPtr(bFile(0))
        CopyMemory tJ.JPGFile, lptr, 4
        ' Store JPGWidth:
        tJ.JPGWidth = cDib.Width
        ' .. & JPGHeight member values:
        tJ.JPGHeight = cDib.Height
        ' Set the quality/compression to save:
        tJ.jquality = lQuality
        ' Write the image:
        lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
        If lR = IJL_OK Then
            SaveJPG = True
        Else
            ' Throw error
            MsgBox "Failed to save to JPG", vbExclamation
        End If
        ' Ensure we have freed memory:
        ijlFree tJ
    Else
        ' Throw error:
        MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
    End If
End Function
