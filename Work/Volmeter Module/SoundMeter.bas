Attribute VB_Name = "SoundMeter"
Dim rc As Long                      ' return code
Dim ok As Boolean                   ' boolean return code
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' handle to volume memory
Dim audbytearray As AUDINPUTARRAY
Dim audByteHigh As AUDINPUTARRAY
Dim posval As Integer
Dim tempval As Integer
Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type
Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type
Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type
Type AUDINPUTARRAY
    bytes(5000) As Byte
End Type
Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long


Private i As Integer, j As Integer, msg As String * 200, hWaveIn As Long
Private Const NUM_BUFFERS = 2
Private format As WAVEFORMAT, hmem(NUM_BUFFERS) As Long, inHdr(NUM_BUFFERS) As WAVEHDR
Public BUFFER_SIZE
Private Const DEVICEID = 0
Private fRecording As Boolean
Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub
Public Function StartInput() As Boolean
On Error GoTo err
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg
        End If
    Next
    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
    Exit Function
err:
    StartInput = False
End Function
Public Function StopInput() As Integer
    On Error GoTo err
    
    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
    GlobalFree volHmem
    StopInput = 0
    Exit Function
err:
    StopInput = 1
End Function
Public Function getVolume(pbuff As Long) As Integer
Dim n As Integer
   On Error Resume Next
   
         Do While Not inHdr(0).dwFlags And WHDR_DONE
         ' perhaps I ought to put a time limit on this bit!
         Loop
            iValue.Caption = CStr(0)
            iValue.Refresh
            CopyStructFromPtr audbytearray, inHdr(0).lpData, inHdr(0).dwBufferLength
            rc = waveInAddBuffer(hWaveIn, inHdr(0), Len(inHdr(0)))

    tempval = 0
    posval = 0
    For n = 0 To BUFFER_SIZE - 1
        posval = audbytearray.bytes(n) - 128
        If posval < 0 Then posval = 0 - posval
        If posval > tempval Then tempval = posval
    Next n
        getVolume = tempval
        pbuff = inHdr(0).lpData
End Function


