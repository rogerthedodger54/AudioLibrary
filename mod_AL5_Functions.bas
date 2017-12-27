Attribute VB_Name = "mod_AL5_Functions"
Option Explicit

Private Const ID3v2_TPE2 As String = "TPE2"   'Band / orchestra / accompaniment

Private Declare Function Has_ID3v1Tag _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String) As Boolean
Private Declare Function Has_ID3v2Tag _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String) As Boolean
Private Declare Function GetTagInfo_ID3v1 _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName) As String()
Private Declare Function GetTagInfo_ID3v2 _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName) As String()
Private Declare Function FindFirstFrameOffset _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String, ByVal lngOffset As Long) As Long
Private Declare Function ID3v2Checksum _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String, ByVal lnglength As Long) As Double
Private Declare Function ID3v1Checksum _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String) As Double
Private Declare Function SetGUID _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal strfilename As String) As String
Private Declare Function GetAudioStreamInfo _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal mp3FileName As String) As StreamInfo
'Public Declare Function LogFile _
        Lib " (m:\VB Projects\Audio Projects\create dll\MP3Tools.dll)" _
        (ByVal strOut As String)
Private Declare Function MP3Checksum _
        Lib "m:\VB Projects\Audio Projects\create dll\MP3Tools.dll" _
        (ByVal lngAudioStart As Long, ByVal lngAudioEnd As Long, ByVal stru_filename As String) As String

Private Declare Function CoCreateGuid _
        Lib "OLE32.DLL" _
        (pGuid As GUID) As Long

Private Type StreamInfo
    mp3DurationSec As Double
    mp3Frames As Long
    mpgAudioVersion As String
    mpgLayerDesc As String
    mp3BitsPerSecond As Long
    mp3SamplesPerSecond As Long
    mp3ChannelsPerSample As String
    mp3Emphasis As String
    mp3CRC As Boolean
    mp3Copyrighted As Boolean
    mp3Original As Boolean
    mp3Private As Boolean
    mp3ModeExt As String
    PreviousOffset As Long
    PreviousFrameLength As Long
End Type

Private AudioStreamInfo As StreamInfo

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Declare Sub CopyMemory _
                    Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    Destination As Any, _
                    Source As Any, _
                    ByVal Length As Long)
Private strError As String
Public strID3V1Tags As String
Public strID3V1Headers As String
Public strID3V2Header As String
Public strID3V2Headers As String
Public strMP3UID As String
Public strMP3UIDHeaders As String
Public strMP3Info As String
Public strMP3InfoHeaders As String
Public strChecksums As String
Public strChecksumHeaders As String

Public Function FileExists(ByVal strDest As String) As Boolean

    Dim intLen As Integer

    If strDest <> vbNullString Then
        On Error Resume Next
        intLen = Len(Dir$(strDest))
        On Error GoTo PROC_ERR
        FileExists = (Not Err And intLen > 0)
    Else
        FileExists = False
    End If

PROC_EXIT:
Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
           "FileExists"
    Err.Clear
    FileExists = False
    Resume PROC_EXIT

End Function

Public Function ReadFile( _
  ByVal intfile As Integer, _
  ByRef abytbuffer() As Byte, _
  ByVal lngnumberofbytes As Long) _
  As Long

  ' Comments  : This function attempts to read the specified number of
  '             bytes from the file.
  ' Parameters: intFile - The file to read from
  '             abytBuffer - The buffer to read the bytes into
  '             lngNumberOfBytes - The number of bytes to read
  ' Returns   : The actual number of bytes read.
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  Dim lngLen As Long
  Dim lngActualBytesRead As Long
  Dim lngStart As Long
    
  ' Get the starting position of the next read
  lngStart = Loc(intfile) + 1
  ' Get the length of the file
  lngLen = LOF(intfile)

  ' Check to see if there is more data to read from the file
  If lngStart < lngLen Then
    ' Check to see if we are attempting to read past the end of the file
    If lngStart + lngnumberofbytes < lngLen Then
      lngActualBytesRead = lngnumberofbytes
    Else
      ' If we are attempting to read more data than is left in the file,
      ' calculate
      ' how much data we should read
      lngActualBytesRead = lngLen - (lngStart - 1)
    End If
    
    ' Create the buffer to hold the data
    ReDim abytbuffer(lngActualBytesRead - 1) As Byte
    ' Do the read
    Get intfile, lngStart, abytbuffer
  Else
    ' If we attempted to read past the end of file, return zero bytes read
    lngActualBytesRead = 0
  End If
  
  ' Return the number of bytes read
  ReadFile = lngActualBytesRead
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReadFile"
  Resume PROC_EXIT
  
End Function


Public Function DirExists(strDir As String) As Boolean
  On Error GoTo PROC_ERR

  DirExists = Len(Dir$(strDir & "\.", vbDirectory)) > 0
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "DirExists"
  Resume PROC_EXIT
  
End Function

Public Function RecurseFolderList(ByVal FolderName As String, ByVal strType As String) As Boolean

    On Error Resume Next
    Dim fso, f, fc, fj, f1
    Dim strfile As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Err.Number > 0 Then
        RecurseFolderList = False
        Exit Function
    End If
    
    
    If fso.FolderExists(FolderName) Then
    
        Set f = fso.GetFolder(FolderName)
        Set fc = f.SubFolders
        Set fj = f.Files
        
        'For each subfolder in the Folder
        On Error Resume Next
        For Each f1 In fc
            DoEvents
            If Err = 0 Then
                'Do something with the Folder Name
                'Debug.Print f1
                'Then recurse this function with the sub-folder to get any'
                ' sub-folders
                Call RecurseFolderList(f1, strType)
            Else
                Err.Clear
            End If
        Next
        
        'For each folder check for any files
        For Each f1 In fj
            DoEvents
            If Err = 0 Then
                strfile = f1
                'Form1.lblFile.Caption = "Searching " & vbCrLf & BeforeLast(strFile, "\")
                If AfterLast(strfile, ".") = strType Then
                    DoEvents
                    Form1.DirSelect.Path = BeforeLast(strfile, "\")
                    'Debug.Print strFile
                    Call FileInfo(strfile)
                    Close
                    'Stop
                Else
                    'Stop
                End If
             'Debug.Print f1
            Else
                Err.Clear
            End If
        Next
        
        Set f = Nothing
        Set fc = Nothing
        Set fj = Nothing
        Set f1 = Nothing
        
    Else
        DoEvents
        RecurseFolderList = False
    End If
    
    Set fso = Nothing

End Function

Public Function AfterFirst(strIn As String, strFirst As String) As String
       '---------------------------------------------------------------------------------------
       ' Procedure : AfterFirst
       ' Author    : Roger Pearce
       ' Date      : 04/03/2009
       ' Purpose   :
       '---------------------------------------------------------------------------------------
       '
    AfterFirst = Right(strIn, Len(strIn) - InStr(1, strIn, strFirst) - (Len(strFirst) - 1))

End Function

Public Function AfterLast(sFrom As String, strAfterLast As String) As String
       '---------------------------------------------------------------------------------------
       ' Procedure : AfterLast
       ' Author    : Roger Pearce
       ' Date      : 04/03/2009
       ' Purpose   : Extract text from last match of strAfterLast
       '---------------------------------------------------------------------------------------
       '



    If InStr(1, sFrom, strAfterLast) Then
        AfterLast = Right(sFrom, Len(sFrom) - InStrRev(sFrom, strAfterLast) - (Len(strAfterLast) - 1))
    Else
        AfterLast = ""
    End If

End Function

Public Function BeforeFirst(strIn As String, strFirst As String) As String
       '---------------------------------------------------------------------------------------
       ' Procedure : BeforeFirst
       ' Author    : Roger Pearce
       ' Date      : 04/03/2009
       ' Purpose   : Extract text before strFirst
       '---------------------------------------------------------------------------------------
       '


    BeforeFirst = Left(strIn, InStr(1, strIn, strFirst) - 1)

End Function

Public Function BeforeLast(strIn As String, strBeforeLast As String) As String
       '---------------------------------------------------------------------------------------
       ' Procedure : BeforeLast
       ' Author    : User
       ' Date      : 06/03/2009
       ' Purpose   : Returns left part of string up to delimit string
       '---------------------------------------------------------------------------------------
       '


    If InStr(1, strIn, strBeforeLast) Then
        BeforeLast = Left(strIn, InStrRev(strIn, strBeforeLast))
    Else
        BeforeLast = ""
    End If


End Function

Public Function FileInfo(strfile As String)
    ' Process each file
    
    Dim blnFileExists As Boolean
    Dim blnResult As Boolean
    Dim strID3v1Result() As String
    Dim strid3v2result() As String
    Dim strMP3Result(14) As String
    Dim intCount As Integer
    Dim lnglength As Long
    Dim strUniCodefile As String
    Dim strGUID As String
    Dim lngSOA As Long
    Dim lngEOA As Long
    Dim lngFileLength As Long
    Dim lngAudioLength As Long
    Dim strTemp As String
    Dim ifile As Byte
    
    'On Error GoTo FileInfo_error
    blnResult = FileExists(strfile)
    If Not blnResult Then Exit Function
    lngFileLength = FileLen(strfile)
    lngEOA = lngFileLength
    If FileLen(strfile) = 0 Then Exit Function
    strUniCodefile = StrConv(strfile, vbUnicode)
    strGUID = StrConv(SetGUID(StrConv(strUniCodefile, vbUnicode)), vbFromUnicode)
    strTemp = strfile & vbTab & strGUID
    Call TextOut(strMP3UID, strTemp, 1)
    ifile = FreeFile
    On Error Resume Next
    Open strfile For Input Access Read As #ifile
        If Err <> 0 Then
            MsgBox "FileInfo file open error " & Err
            Err.Clear
        End If
    Close #ifile
    ' ID3v1 info
    blnResult = Has_ID3v1Tag(strUniCodefile)
    '   If ID3v1 tag present then get info
    If blnResult Then
        lngEOA = lngEOA - 128
        strID3v1Result = GetTagInfo_ID3v1(strUniCodefile)
        strID3v1Result(8) = strfile
        For intCount = 0 To 8
            If intCount > 0 Then
                strTemp = strTemp & vbTab & strID3v1Result(intCount)
            Else
                strTemp = strID3v1Result(0)
            End If
        Next intCount
        Call TextOut(strID3V1Tags, strTemp, 1)
        'For intCount = 0 To 8
            'Debug.Print strID3v1Result(intCount)
        'Next intCount
    End If
    strTemp = ""
    ' ID3v2 info
    blnResult = Has_ID3v2Tag(strUniCodefile)
    If blnResult Then
        ' Get length of ID3 Tags
        lnglength = FindFirstFrameOffset(strUniCodefile, 1) - 1
        strid3v2result() = GetTagInfo_ID3v2(strUniCodefile)
        strid3v2result(100) = Hex(ID3v2Checksum(strUniCodefile, lnglength))
        If strid3v2result(101) <> "" Then LogFile (strid3v2result(101))
        For intCount = 0 To 100
            If intCount > 0 Then
                strTemp = strTemp & vbTab & strid3v2result(intCount)
            Else
                strTemp = strid3v2result(0)
            End If
        Next intCount
        Call TextOut(strID3V2Header, strTemp, 1)
        'For intCount = 0 To 100
            'Debug.Print intCount, strID3v2Result(intCount)
        'Next intCount
    End If
    strTemp = ""
    ' Get length of audio
    lngSOA = lnglength + 1
    lngAudioLength = lngEOA - lngSOA
    'MsgBox "Audio Frames Info" & vbCrLf & Hex(lngFileLength) & vbCrLf & _
        Hex(lngAudioLength) & vbCrLf & _
        Hex(lngEOA) & vbCrLf & Hex(lngSOA)
    ' Get mp3 info
    AudioStreamInfo = GetAudioStreamInfo(strUniCodefile)
    ' Prepare output
    strMP3Result(0) = strfile
    strMP3Result(1) = CStr(lngFileLength)
    strMP3Result(2) = CStr(lnglength)
    With AudioStreamInfo
        strMP3Result(3) = .mp3DurationSec
        strMP3Result(4) = .mp3Frames
        strMP3Result(5) = .mpgAudioVersion
        strMP3Result(6) = .mpgLayerDesc
        strMP3Result(7) = .mp3BitsPerSecond
        strMP3Result(8) = .mp3SamplesPerSecond
        strMP3Result(9) = .mp3ChannelsPerSample
        strMP3Result(10) = .mp3Emphasis
        strMP3Result(11) = .mp3CRC
        strMP3Result(12) = .mp3Copyrighted
        strMP3Result(13) = .mp3Original
        strMP3Result(14) = .mp3Private
    End With
    For intCount = 0 To 14
        If intCount > 0 Then
            strTemp = strTemp & vbTab & strMP3Result(intCount)
        Else
            strTemp = strMP3Result(0)
        End If
    Next intCount
    Call TextOut(strMP3Info, strTemp, 1)
    strTemp = MP3Checksum(lngSOA, AudioStreamInfo.PreviousOffset, strUniCodefile)
    'Stop
    'MsgBox strFile
    ' The end of audio for checksums is taken from after the frame header.
    strTemp = strfile & vbTab & strid3v2result(100) & "Z" & strID3v1Result(7) & _
        "Z" & Hex(lngFileLength - lngSOA - 127) & "Z" & Hex(lngSOA - 1) & _
        "Z" & strTemp & "Z" & Hex(AudioStreamInfo.PreviousOffset + 5)
    Call TextOut(strChecksums, strTemp, 1)
    'Stop
'Exit Function

'FileInfo_error:
'MsgBox Error
'Stop
End Function

Public Function LogFile(ByVal strOut As String)
    '---------------------------------------------------------------------------------------
    ' Procedure : LogFile
    ' Author    : Roger Pearce
    ' Date      : 12/09/2013
    ' Purpose   : Output error message to AudioLibrarian 5 errors.txt file
    '---------------------------------------------------------------------------------------
    '
    Dim lfile As Byte
    Dim strfile As String
    
    strfile = App.Path & "\AL5 errors.txt"
    If Len(strOut) > 1 Then
        lfile = FreeFile
        On Error GoTo File_error
        Open strfile For Append As #lfile
            Print #lfile, strOut
        Close #lfile
    Else
        Stop
    End If
Exit Function

File_error:
MsgBox Error
End Function

Public Sub TextOut(strfile As String, strdata As String, bytType As Byte)
    '---------------------------------------------------------------------------------------
    ' Procedure : TextOut
    ' Author    : Roger Pearce
    ' Date      : 16-Jul-2010
    ' Purpose   : Output to text file
    ' Inputs    : Information to output to file
    ' Assumes   :
    ' Returns   : Nothing
    ' Effects   :
    '---------------------------------------------------------------------------------------
    '

    Dim bytFile As Byte
    Dim strPath As String
    
    On Error GoTo TextOut_Error

    If Len(strfile) < 1 Then
        Stop
    End If
    bytFile = FreeFile
    Select Case bytType
        Case 0
            Open strfile For Output Access Write As #bytFile
        Case 1
            Open strfile For Append Access Write As #bytFile
    End Select
    Print #bytFile, strdata
    Close #bytFile

    
Exit Sub

TextOut_Error:
    Select Case Err
        Case 76
            Err.Clear
            strPath = BeforeLast(strfile, "\")
            On Error Resume Next
            MkDir strPath
            If Err <> 0 Then
                Stop
            End If
                Resume
        Case Else
            strError = Date & " " & Time & vbTab & "Error" & Err.Number & " (" & _
                Err.Description & ") at line " & Erl & _
                " in procedure TextOut of Module modFunctions. File being processed " & _
                strfile
            'Call LogError(strError, DebugLevel)
            'Call LogError(strError, 0)
            Err.Clear
    End Select


End Sub

Public Function Read_GetTagInfo_ID3v2(ByVal mp3FileName$) As Boolean
Dim songtitle$
Dim songartist$
Dim songalbum$
Dim songyear As Long
Dim songcomment$
Dim songtracknumber As Long
Dim songgenre As String
Dim songComposer$
Dim songOriginalArtist$
Dim songCopyright$
Dim songURL$
Dim songEncodedBy$

songtitle$ = ""
songartist$ = ""
songalbum$ = ""
songyear = 0
songcomment$ = ""
songtracknumber = 0
songgenre = ""

songComposer$ = ""
songOriginalArtist$ = ""
songCopyright$ = ""
songURL$ = ""
songEncodedBy$ = ""

Dim fNum As Long, fOffset As Long, fLen As Long

On Error GoTo FileError
fNum = FreeFile()
Open mp3FileName$ For Binary As #fNum
On Error GoTo 0

fLen = LOF(fNum)
fOffset = 1

Dim tempFrameOffset As Long, tempTagLen As Long
Dim tempYear$, tempTrack$, tempGenre$

Dim tempStr$:  tempStr$ = String$(3, 0)
Get #fNum, fOffset, tempStr$
If tempStr$ <> "ID3" Then
  GoTo FileError
End If

ReadTagHeader mp3FileName$, fOffset, tempTagLen, 0, 0, False, False, False, tempFrameOffset, fNum
fOffset = tempFrameOffset

Dim frameLen As Long, frameType$
Dim id3v2bytes() As Byte

Do
  If fOffset > tempTagLen Then Exit Do

  ReadTagFrame mp3FileName$, fOffset, frameLen, frameType$, False, fNum
  
  If frameType$ = ID3v2_TPE2 Then
    ReDim id3v2bytes(0 To frameLen - 1)
    Get #fNum, fOffset, id3v2bytes()
    ID3v2_GetTPE2info id3v2bytes(), songtitle$
  End If
  
  fOffset = fOffset + frameLen
Loop

Close #fNum
Read_GetTagInfo_ID3v2 = True

Exit Function


FileError:
 Close #fNum
 Read_GetTagInfo_ID3v2 = False
 songtitle$ = ""
 songartist$ = ""
 songalbum$ = ""
 songyear = 0
 songcomment$ = ""
 songtracknumber = 0
 songgenre = ""


End Function
Private Function ReadTagHeader(ByVal mp3FileName$, ByVal tagOffset As Long, ByRef tagLength As Long, ByRef tagVersion As Single, ByRef tagExtendedHeaderLength As Long, ByRef tagIsUnsync As Boolean, ByRef tagHasFooter As Boolean, ByRef tagIsExperimental As Boolean, ByRef tagFrameOffset As Long, Optional ByVal openFileNumber As Long = -1) As Boolean

Dim fNum As Long, bArray(0 To 13) As Byte
If openFileNumber = -1 Then
  On Error GoTo FileError
  fNum = FreeFile()
  Open mp3FileName$ For Binary As #fNum
  On Error GoTo 0
Else
  fNum = openFileNumber
End If

Get #fNum, tagOffset, bArray()
If openFileNumber = -1 Then Close #fNum

Dim tempHasExtended As Boolean

tagLength = ID3v2_4BytesToLONG_SynchSafe(bArray(6), bArray(7), bArray(8), bArray(9)) + 10

tagVersion = bArray(3) + (bArray(4) / 10)
tagIsUnsync = MP3_GetBit(bArray(5), 7)
tempHasExtended = MP3_GetBit(bArray(5), 6)
tagIsExperimental = MP3_GetBit(bArray(5), 5)
tagHasFooter = MP3_GetBit(bArray(5), 4)

If tempHasExtended = False Then
  tagExtendedHeaderLength = 0
  tagFrameOffset = tagOffset + 10
Else
  tagExtendedHeaderLength = ID3v2_4BytesToLONG_SynchSafe(bArray(10), bArray(11), bArray(12), bArray(13))
  tagFrameOffset = tagOffset + 20
End If

If tagHasFooter <> False Then tagLength = tagLength + 10

ReadTagHeader = True

Exit Function

FileError:
 If openFileNumber = -1 Then Close #fNum
 ReadTagHeader = False
 tagLength = 0
 tagVersion = 0
 tagExtendedHeaderLength = 0
 tagIsUnsync = False
 tagIsExperimental = False
 tagHasFooter = False
 
   
End Function

Private Function ReadTagFrame(ByVal mp3FileName$, ByVal tagFrameOffset As Long, ByRef tagFrameLength As Long, ByRef tagFrameType As String, ByRef tagFrameIsCompressed As Boolean, Optional ByVal openFileNumber As Long = -1) As Boolean

Dim fNum As Long, bArray(0 To 9) As Byte
If openFileNumber = -1 Then
  On Error GoTo FileError
  fNum = FreeFile()
  Open mp3FileName$ For Binary As #fNum
  On Error GoTo 0
Else
  fNum = openFileNumber
End If

Get #fNum, tagFrameOffset, bArray()
If openFileNumber = -1 Then Close #fNum

Dim tempStr$: tempStr$ = String$(4, 0)

Dim i As Long
For i = 0 To 3
  Mid$(tempStr$, i + 1, 1) = Chr$(bArray(i))
Next i
tagFrameType$ = tempStr$

tagFrameLength = ID3v2_4BytesToLONG_SynchSafe(bArray(4), bArray(5), bArray(6), bArray(7)) + 10


ReadTagFrame = True

Exit Function

FileError:
 If openFileNumber = -1 Then Close #fNum
 ReadTagFrame = False
 tagFrameLength = 0
 tagFrameType = ""
 tagFrameIsCompressed = False
 
End Function

Private Function GetTextFrameInfo(ByRef id3v2FrameBytes() As Byte, Optional ByVal textStartPos As Long = 11) As String


Dim tempFrameLength As Long
tempFrameLength = ID3v2_4BytesToLONG_SynchSafe(id3v2FrameBytes(4), id3v2FrameBytes(5), id3v2FrameBytes(6), id3v2FrameBytes(7)) + 10

Dim tempTextFlags As Long
tempTextFlags = id3v2FrameBytes(10)

Dim tempStr$, outStr$, i As Long, tempBytes() As Byte

If tempTextFlags = 0 Then
  For i = textStartPos To tempFrameLength - 1
    tempStr$ = tempStr$ & Chr$(id3v2FrameBytes(i))
  Next i

ElseIf tempTextFlags = 1 Or tempTextFlags = 2 Then
  ReDim tempBytes(0 To tempFrameLength - textStartPos)
  For i = textStartPos To tempFrameLength - 1
    tempBytes(i - textStartPos) = id3v2FrameBytes(i)
  Next i
  If tempFrameLength - 1 >= textStartPos Then
    tempStr$ = ID3v2_ReadUnicode_UTF16(tempBytes(), (tempTextFlags = 2))
  End If

ElseIf tempTextFlags = 3 Then
  ReDim tempBytes(0 To tempFrameLength - textStartPos)
  For i = textStartPos To tempFrameLength - 1
    tempBytes(i - textStartPos) = id3v2FrameBytes(i)
  Next i
  tempStr$ = ID3v2_ReadUnicode_UTF8(tempBytes())
  
End If


outStr$ = ""
For i = Len(tempStr$) To 1 Step -1
  If Asc(Mid$(tempStr$, i, 1)) <> 0 Then
    outStr$ = Mid$(tempStr$, 1, i)
    Exit For
  End If
Next i

GetTextFrameInfo = outStr$

End Function

Private Sub ID3v2_GetTPE2info(ByRef id3v2FrameBytes() As Byte, ByRef songtitle$)

songtitle$ = "TPE2" & vbTab & GetTextFrameInfo(id3v2FrameBytes())

End Sub
Private Sub ID3v2_4BytesFromLONG_SynchSafe(ByVal longValue As Long, ByRef ByteHigh As Long, ByRef ByteMidH As Long, ByRef ByteMidL As Long, ByRef ByteLow As Long)

ByteLow = longValue Mod 128
longValue = (longValue - ByteLow) \ 128

ByteMidL = longValue Mod 128
longValue = (longValue - ByteMidL) \ 128

ByteMidH = longValue Mod 128
longValue = (longValue - ByteMidH) \ 128

ByteHigh = longValue Mod 128
longValue = (longValue - ByteHigh) \ 128

End Sub

Private Function ID3v2_4BytesToLONG_SynchSafe(ByVal ByteHigh As Byte, ByVal ByteMidH As Byte, ByVal ByteMidL As Byte, ByVal ByteLow As Byte) As Long

Dim longHigh As Long, longMidH As Long, longMidL As Long, longLow As Long
longHigh = ByteHigh
longMidH = ByteMidH
longMidL = ByteMidL
longLow = ByteLow

Dim outVal As Long
outVal = (longHigh * 2097152) + (longMidH * 16384) + (longMidL * 128) + longLow

ID3v2_4BytesToLONG_SynchSafe = outVal

End Function
Private Function MP3_GetBit(ByVal aNumber As Long, ByVal bitNumber As Long) As Boolean

Dim expVal As Long
expVal = 2 ^ bitNumber

If (aNumber And expVal) = expVal Then
  MP3_GetBit = True
Else
  MP3_GetBit = False
End If

End Function

Private Function ID3v2_ReadUnicode_UTF16(ByRef strBytes() As Byte, Optional ByVal BigEndian As Boolean = True) As String

Dim cVal1 As Long, cVal2 As Long

cVal1 = strBytes(0)
cVal2 = strBytes(1)

Dim tStart As Long

If cVal1 = &HFF And cVal2 = &HFE Then
  BigEndian = True
  tStart = 2
ElseIf cVal1 = &HFE And cVal2 = &HFF Then
  BigEndian = False
  tStart = 2
Else
  tStart = 0
End If

Dim i As Long, byteCount As Long
Dim outStr$

For i = tStart To UBound(strBytes())
  If Not byteCount Mod 2 = 1 Then
    If BigEndian <> False Then
      cVal1 = strBytes(i)
    Else
      cVal1 = strBytes(i)
      cVal1 = cVal1 * 256
    End If
    outStr$ = outStr$ & ChrW$(cVal1)

  End If
  
  byteCount = byteCount + 1
Next i

ID3v2_ReadUnicode_UTF16 = outStr$

End Function
Private Function ID3v2_ReadUnicode_UTF8(ByRef strBytes() As Byte) As String

Dim cVal1 As Long, cVal2 As Long

Dim i As Long, j As Long, byteCount As Long
Dim charLen As Long
Dim outStr$


For i = 0 To UBound(strBytes())
  If strBytes(i) < 128 Then
    outStr$ = outStr$ & Chr$(strBytes(i))
    byteCount = byteCount + 1
  Else
    charLen = 0
    cVal2 = 0
    For j = 7 To 2 Step -1
      If MP3_GetBit(strBytes(i), j) <> False Then charLen = charLen + 1 Else Exit For
    Next j
    
    For j = (charLen - 1) To 0 Step -1
      If j > 0 Then
        cVal1 = (strBytes(i + j) And 127)
        cVal2 = cVal2 + (cVal1 * 64)
      Else
        cVal1 = (strBytes(i + j) And ((2 ^ (7 - charLen)) - 1))
        cVal2 = cVal2 + (cVal1 * 64)
      End If
    Next j
    outStr$ = outStr$ & ChrW$(cVal2)
    i = i + charLen - 1
  End If
  
Next i

ID3v2_ReadUnicode_UTF8 = outStr$

End Function

Private Sub ID3v2_GetCOMMinfo(ByRef id3v2FrameBytes() As Byte, ByRef songcomment$)

songcomment$ = ID3v2_GetTextFrameInfo(id3v2FrameBytes(), 15)

End Sub
