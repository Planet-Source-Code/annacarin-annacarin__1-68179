VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picWave 
      Height          =   2040
      Left            =   480
      ScaleHeight     =   1980
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OpenWav..."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This tiny project is all about showing the math behind reading different standard wave-formats.
'I have taken it from a bigger project of mine and stripped off all the graphic goodies.
'That's why it's ugly but simple.
'The main work here is the routine to read 24bit waves. I could not find any help on the web so I created my own.
'Please use it as you like and please change the code to be even faster. I submit some very short example audio files.
'Best regards AnnaCarin.

Option Explicit
Option Compare Text
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'================================

Private Type WAVEFMT
    signature As String * 4     ' must contain 'RIFF'
    RIFFsize As Long            ' size of file (in bytes) minus 8
    type As String * 4          ' must contain 'WAVE'
    fmtchunk As String * 4      ' must contain 'fmt ' (including blankspace)
    fmtsize As Long             ' size of format chunk, must be 16
    FORMAT As Integer           ' normally 1 (PCM), 3 = floating point decimaltal.
    channels As Integer         ' number of channels, 1=mono, 2=stereo
    samplerate As Long          ' sampling frequency: 44100, 48000 or 96000
    average_bps As Long         ' average bytes per second; samplerate * channels
    align As Integer            ' 1=byte aligned, 2=word aligned
    bitspersample As Integer    ' should be 16 or 24 or 32.
    datchunk As String * 4      ' must contain 'data'
    samples As Long             ' number of samples measured in NUMBER OF BYTES.
End Type

Private Type ThreeByte
    'Integer As Integer
    Byte1 As Byte
    IntM256 As Integer
    'Byte2 As Byte
    'Byte3 As Byte
End Type

Private Type POINT
    X As Long
    Y As Long
End Type
 
'----------------------------------------------------------
' 882 comes from:
'   44100 = 1 second
'   441   = 1/100 second
'   882   = 441 * 2 (2 = stereo)
'----------------------------------------------------------
'Private Type arrUdtWAVEBLOCK
 '   arriWavinfo(1 To 882) As Integer 'Plats för 882 sample-värden (0,01 s) i en dylik som nedan definieras som arriCandidatesPerCol.
'End Type
 
Private Type SCROLLER
    Min As Long
    Max As Long
    value As Long
    topval As Long  'Position represented by top line of wavform
End Type
 
Dim bteFreeFile As Byte
Dim arrlWavMin() As Long, arrlWavMinR() As Long, arrLWavMax() As Long, arrLWavMaxR() As Long
Dim arriCandidatesPerCol() As Integer, arrlCandidatesPerCol() As Long, arrSngCandidatesPerCol() As Single
Dim arrUdvColCandidates24bit() As ThreeByte, LastPt As POINT, VScroll As SCROLLER
Dim ipicBoxWidth As Integer
Const TWO_BYTES_per_INTEGER As Byte = 2, THREE_BYTESin24BIT As Byte = 3, FOUR_BYTESinLONG As Byte = 4
Const WAVE_FORMAT_PCM_INTEGERS_IS_ONE As Byte = 1, WAVE_FORMAT_IEEE_FLOAT_IS_THREE As Byte = 3
Const WAVE_FORMAT_EXTENSIBLE_IS_HFFFE As Long = 65534, Three_Normal_Sizes_Of_FMTCHUNK As String = "16;18;40"
'Const Chunk size: 16 or 18 or 40

Private Sub Command1_Click()
Dim sFile As String
With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "File Audio(*.WAV|*.WAV|All(*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
'On Error GoTo Fehler
 
 Call ValidateFile(sFile)
 'Call AddFiles(sFile)

End Sub

Public Sub ValidateFile(sPath)
 Dim wavh As WAVEFMT, ctr As Long, lNrealSamplesPerChannel As Long 'Variabelkluster alltså userdefined type.
 Dim sngDuration As Single, sHeaderData As String
 Static sPreserveOrigPath As String
'----- Does the wavefile exist? -----------
If sPath = "" Then Exit Sub
If Dir(sPath, vbNormal) = "" Then Exit Sub 'Bail if the file doesn't exist on the drive.
Select Case Right(sPath, 4)
 Case ".wav"
 'continue down i this routine.
End Select
'--- Examine the files header chunk! ------------
bteFreeFile = FreeFile
Open sPath For Binary Access Read As #bteFreeFile
Get #bteFreeFile, , wavh
'-----------------------------------------
'  Only allow wavh.bitspersample = 16, 24, 32!!!
'-----------------------------------------
If wavh.signature <> "RIFF" Then MsgBox "Sorry, this file lacks the fourCC label 'RIFF' and instead it says " & wavh.signature: Exit Sub
If wavh.type <> "WAVE" Then MsgBox "Sorry, this file lacks the fourCC label 'WAVE' and instead it says " & wavh.type: Exit Sub
If wavh.fmtchunk <> "fmt " Then MsgBox "Sorry, this file lacks the fourCC label 'fmt ' and instead it says " & wavh.fmtchunk: Exit Sub
If InStr(Three_Normal_Sizes_Of_FMTCHUNK, CStr(wavh.fmtsize)) = 0 Then
 MsgBox "This file has fmtChunkSize = " & wavh.fmtsize & " and should normally be = 16 or 18 or 40! I'll check if SizeChunk matches LOF else I'll bail!"
    If wavh.RIFFsize + 8 <> LOF(bteFreeFile) Then
        MsgBox "!wavh.RIFFsize + 8 = " & wavh.RIFFsize + 8 & " and LOF(bteFreeFile) = " & LOF(bteFreeFile) & "! I will bail!"
        Exit Sub
    Else
    MsgBox "MATCH I'll keep on going. !wavh.RIFFsize + 8 = " & wavh.RIFFsize + 8 & " and LOF(bteFreeFile) = " & LOF(bteFreeFile) & "!"
    End If
 End If
If wavh.FORMAT <> WAVE_FORMAT_PCM_INTEGERS_IS_ONE And wavh.FORMAT <> WAVE_FORMAT_IEEE_FLOAT_IS_THREE And wavh.FORMAT <> WAVE_FORMAT_EXTENSIBLE_IS_HFFFE Then MsgBox "Sorry, this file has format = " & wavh.FORMAT & " and PCMWave should be = 1, IEEEfloat should be 3 and WaveExtensible should be 65534! I'll Bail!": Exit Sub
If wavh.channels <> 1 And wavh.channels <> 2 Then MsgBox "Sorry, this file has nChannels = " & wavh.channels & " and this programme only allows = 1 OR 2!": Exit Sub
If wavh.bitspersample <> 16 And wavh.bitspersample <> 24 And wavh.bitspersample <> 32 Then MsgBox "Sorry, this file has bitspersample = " & wavh.bitspersample & " and this programme only allows bitspersample = 16 AND 24 AND 32!": Exit Sub
If wavh.samplerate <> 44100 And wavh.samplerate <> 48000 And wavh.samplerate <> 96000 Then MsgBox "Sorry, this file has samplerate = " & wavh.samplerate & "/second  and this programme only allows samplerate = 41000 AND 48000 AND 96000!": Exit Sub
If wavh.datchunk <> "data" Then
For ctr = 0 To (LOF(bteFreeFile) - Loc(bteFreeFile)) '100 'Sometimes there are lots of padding zeros before the data chunk so lets go chase the datachunk!
 sHeaderData = Input(1, #bteFreeFile) 'Get #bteFreeFile, 1, wavh
 If sHeaderData = "d" Then
  sHeaderData = Input(3, #1)
  If sHeaderData = "ata" Then 'wavh.samples are located after FOURCC 'data'.
   wavh.datchunk = "data": Get #bteFreeFile, , wavh.samples: Exit For
  End If
 End If
Next ctr
End If
If Loc(bteFreeFile) > (LOF(bteFreeFile) - 2) Then MsgBox "Sorry, this file lacks the fourCC datchunk 'data' and instead i says " & wavh.datchunk: Exit Sub
'=========================================
'The file has passed the validation!


'--- Create my own useful lNrealSamplesPerChannel! ------ Praktisk för att flaggan wavh.samples använder enheten bytes medan i verkligheten både integers (CD-ljud) och Long kan användas.
Select Case wavh.bitspersample
 Case 16 'Monofiler har 2 bytes per sampling medan stereo har 4.
  If wavh.channels = 1 Then lNrealSamplesPerChannel = wavh.samples \ TWO_BYTES_per_INTEGER Else lNrealSamplesPerChannel = wavh.samples \ (TWO_BYTES_per_INTEGER * 2)
 Case 24 'Monofiler har 3 bytes per sampling medan stereo har 6.
  If wavh.channels = 1 Then lNrealSamplesPerChannel = wavh.samples \ THREE_BYTESin24BIT Else lNrealSamplesPerChannel = wavh.samples \ (THREE_BYTESin24BIT * 2)
 Case 32 'Monofiler har 4 bytes per sampling medan stereo har 8.
  If wavh.channels = 1 Then lNrealSamplesPerChannel = wavh.samples \ FOUR_BYTESinLONG Else lNrealSamplesPerChannel = wavh.samples \ (FOUR_BYTESinLONG * 2)
End Select

sngDuration = lNrealSamplesPerChannel / wavh.samplerate


Call ExtractExtremeValues(sngDuration, lNrealSamplesPerChannel, wavh.samples, wavh.channels, wavh.bitspersample, wavh.FORMAT) 'Varje kolumn representerar ett större område så därför är respektive områdes extremvärde mest representativt.
Call DrawWaveData(wavh.channels)

sPath = sPreserveOrigPath 'Since mplayer changes the path into "audiodump.wav"

End Sub
Private Sub ExtractExtremeValues(sngDuration, lNrealSamplesPerChannel, nSampleBytes As Long, nChannels As Integer, nBitsPerSample As Integer, FORMAT As Integer)
Dim iMaxBreddPicBox As Integer, lNsamplesMultiplNchannels As Long

'---The width limit for picBoxes in VB is 16383 pixels!
iMaxBreddPicBox = 245745 / Screen.TwipsPerPixelX 'Antal pixlar = 16383.


'==== Determine horisontal resolution! ===================|
' This project uses a soecial standard resolution of 0,01s per column.|
' That makes room for 163,83 seconds of musik in the picbox!          |
' Files loonger than that will be zoomed to fit the maxwidth of the picBox.|
'==============================================================|

If sngDuration > 163.83 Then ipicBoxWidth = iMaxBreddPicBox Else ipicBoxWidth = (sngDuration * 100) + 1
    '--- Trivial: The music is over 2'43'' and will fill the whole picbox after zooming! ---
    Select Case nBitsPerSample '16 (integer) eller 24 eller 32 (Long).
        Case 16
            lNsamplesMultiplNchannels = nSampleBytes \ TWO_BYTES_per_INTEGER
            ReDim arriCandidatesPerCol(1 To lNsamplesMultiplNchannels \ ipicBoxWidth) 'Division med 2 beror på att sSamples visar bytes - ej integers.
            'lColumnCount = number of small pices read from the drive - should equal the horizontal number of pixels of pixBox.
            picWave.ToolTipText = "ZoomFactor MUST CONCERN stereo/mono. ADJUST! = " & UBound(arriCandidatesPerCol)
        Case 24
            lNsamplesMultiplNchannels = nSampleBytes \ THREE_BYTESin24BIT
            ReDim arrUdvColCandidates24bit(1 To lNsamplesMultiplNchannels \ ipicBoxWidth)
            ReDim arrlCandidatesPerCol(1 To UBound(arrUdvColCandidates24bit))
            'lColumnCount = number of small pices read from the drive - should equal the horizontal number of pixels of pixBox.
            picWave.ToolTipText = "ZoomFactor MUST CONCERN stereo/mono. ADJUST! = " & UBound(arrUdvColCandidates24bit)
        
        Case 32
            lNsamplesMultiplNchannels = nSampleBytes \ FOUR_BYTESinLONG
            Select Case FORMAT
             Case WAVE_FORMAT_PCM_INTEGERS_IS_ONE
              ReDim arrlCandidatesPerCol(1 To lNsamplesMultiplNchannels \ ipicBoxWidth)
             Case WAVE_FORMAT_IEEE_FLOAT_IS_THREE
              ReDim arrSngCandidatesPerCol(1 To lNsamplesMultiplNchannels \ ipicBoxWidth)
            End Select
            'lColumnCount = the number of small piceces being read from the drive - should be asa many as the number of horizontal pixels of pixBox.
            'picWave.ToolTipText = "ZoomFactor MUST COMPENSATE for stereo/mono PLEASE ADJUST = " & UBound(arrlCandidatesPerCol): imgWave(1).ToolTipText = picWave.ToolTipText
        End Select
    

ReDim arrlWavMin(ipicBoxWidth - 1), arrlWavMinR(ipicBoxWidth - 1), arrLWavMax(ipicBoxWidth - 1), arrLWavMaxR(ipicBoxWidth - 1)
'There's only room in picWave for 2,73 minuter (2'43'' or 28,8 MB) CDsound with a column every 0,001 second.
If ipicBoxWidth * Screen.TwipsPerPixelX > 245745 Then MsgBox "Programmet försöker göra plats för fler kolumner är fler än vad picBox kan rymma. Måste göra en zoomad version!"
picWave.Width = ipicBoxWidth * Screen.TwipsPerPixelX  'Make horizontal room for all graphics.
Me.Caption = picWave.Width

Screen.MousePointer = vbHourglass


'================= Följande rutin väljer ut extremvärdena! =============
Select Case nBitsPerSample
 Case 16
  Call FillarrLWavMax16(nChannels)
 Case 24
 'Call FillarrLWavMax24Slow(nChannels)
 Call FillarrLWavMax24Speedy(nChannels)
 Case 32
    Select Case FORMAT
        Case WAVE_FORMAT_PCM_INTEGERS_IS_ONE
          Call FillarrLWavMax32Integers(nChannels)
        Case WAVE_FORMAT_IEEE_FLOAT_IS_THREE
          Call FillarrLWavMax32Float(nChannels)
    End Select
End Select


Close #bteFreeFile

VScroll.Min = 0
VScroll.Max = ipicBoxWidth
VScroll.value = 0
        
Screen.MousePointer = vbDefault

End Sub
Private Sub FillarrLWavMax16(nChannels)
Dim lColCount As Long, ctr As Long, lMin As Long, lMinR As Long, lMax As Long, lMaxR As Long

'Do
For lColCount = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1
    Get #bteFreeFile, , arriCandidatesPerCol
    'Find min/max to get a better view of the wave
    lMin = 0: lMinR = 0: lMax = 0: lMaxR = 0
    'Should look at all values but 'Step 32' speeds it up a bit. A nice C++ function to find the max/min would be handy here!
    If nChannels = 2 Then
         '----------- StereoFil. -----------
         For ctr = 1 To UBound(arriCandidatesPerCol) - 1 Step 32 ' Step 2 tittar alltså varenda värde - varannan vänster/höger 'Väljer ut extremSAMPLEvärdena lColCount ett tidsblock på 100 ms.
            If arriCandidatesPerCol(ctr) < lMin Then lMin = arriCandidatesPerCol(ctr) 'Left Channel.
            If arriCandidatesPerCol(ctr) > lMax Then lMax = arriCandidatesPerCol(ctr) 'Left Channel.
            If arriCandidatesPerCol(ctr + 1) < lMinR Then lMinR = arriCandidatesPerCol(ctr + 1) 'Right Channel.
            If arriCandidatesPerCol(ctr + 1) > lMaxR Then lMaxR = arriCandidatesPerCol(ctr + 1) 'Right Channel.
         Next ctr
        Else
         '------------ MonoFil. ------------
         For ctr = 1 To UBound(arriCandidatesPerCol) 'Step 32 'Väljer ut extremSAMPLEvärdena lColCount ett tidsblock på 100 ms.
            'If arriCandidatesPerCol.arriWavinfo(ctr) > 0 Then MsgBox arriCandidatesPerCol.arriWavinfo(ctr)
            If arriCandidatesPerCol(ctr) < lMin Then lMin = arriCandidatesPerCol(ctr)
            If arriCandidatesPerCol(ctr) > lMax Then lMax = arriCandidatesPerCol(ctr)
         Next ctr
    End If
    'lColCount = lColCount + 1
    If nChannels = 2 Then
    'Stereofil
     arrlWavMin(lColCount) = lMin \ 1024 + 32  'Left. Long values become +/-64 then
     arrlWavMinR(lColCount) = lMinR \ 1024 + 96  'Right. Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 1024 + 32 'Left. add 64 to make into co-ordinates
     arrLWavMaxR(lColCount) = lMaxR \ 1024 + 96  'Right. add 64 to make into co-ordinates
     'Debug.Print lMax; lMaxR
    Else
    'monoFil. *** Max utslag på picWave = 127 och minutslag = 0.
     arrlWavMin(lColCount) = lMin \ 512 + 64  'Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 512 + 64  'add 64 to make into co-ordinates
     End If
    Next lColCount
'Loop Until EOF(1)

End Sub

Private Sub FillarrLWavMax24Slow(nChannels)
Dim lColCount As Long, ctr As Long, lMin As Long, lMinR As Long, lMax As Long, lMaxR As Long
Dim lByte1 As Long, lIntM256 As Long 'lByte2 As Long, lByte3 As Long
'SIGNED 24 BYTE INTEGERS INTE FLOAT!!!
For lColCount = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1
   Get #bteFreeFile, , arrUdvColCandidates24bit
    For ctr = 1 To UBound(arrlCandidatesPerCol) 'Mystiskt. VB vägrar summera nummer från UDT.
     lByte1 = arrUdvColCandidates24bit(ctr).Byte1: lIntM256 = arrUdvColCandidates24bit(ctr).IntM256: lIntM256 = lIntM256 * 256 'lByte2 = arrUdvColCandidates24bit(ctr).Byte2: lByte3 = arrUdvColCandidates24bit(ctr).Byte3
     If lIntM256 < 0 Then
      arrlCandidatesPerCol(ctr) = lIntM256 + (-255 + lByte1)
     Else
      arrlCandidatesPerCol(ctr) = lIntM256 + lByte1 'lByte1 + lByte2 * 256 + lByte3 * 65536 'arrUdvColCandidates24bit(ctr).Byte1 + arrUdvColCandidates24bit(ctr).Byte2 + arrUdvColCandidates24bit(ctr).Byte3 'arrUdvColCandidates24bit(ctr).Integer + arrUdvColCandidates24bit(ctr).Byte * 65536
     End If
    Next ctr
    'Find min/max to get a better view of the wave
    lMin = 0: lMinR = 0: lMax = 0: lMaxR = 0
    'Should look at all values but 'Step 32' speeds it up a bit. A nice C++ function to find the max/min would be handy here!
    If nChannels = 2 Then
         '----------- StereoFil. -----------
         For ctr = 1 To UBound(arrlCandidatesPerCol) - 1 Step 32 ' Step 3 tittar alltså varenda värde - varannan vänster/höger 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr + 1) < lMinR Then lMinR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
            If arrlCandidatesPerCol(ctr + 1) > lMaxR Then lMaxR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
         Next ctr
        Else
         '------------ MonoFil. ------------
         For ctr = 1 To UBound(arrlCandidatesPerCol) 'Step 32 'Picking the extremSAMPLEvalues in timeframes of 100 ms.
            'If arrlCandidatesPerCol.arriWavinfo(ctr) > 0 Then MsgBox arrlCandidatesPerCol.arriWavinfo(ctr)
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr)
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr)
         Next ctr
    End If
    'lColCount = lColCount + 1
    If nChannels = 2 Then
    'Stereofile \67108864 = Hex 4000 000! A PURE MULTIPLICATION FROM 16-bitversion.
     arrlWavMin(lColCount) = lMin \ 266304 + 32  'Left. Long values become +/-64 then
     arrlWavMinR(lColCount) = lMinR \ 266304 + 96  'Right. Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 266304 + 32 'Left. add 64 to make into co-ordinates
     arrLWavMaxR(lColCount) = lMaxR \ 266304 + 96  'Right. add 64 to make into co-ordinates
     'Debug.Print lMax; lMaxR
    Else
    'monoFile. 33554432 = Hex 2000 000! A PURE MULTIPLICATION FROM 16-bitversion.
    '*** Max deviation in picWave should be = 127 och minimum = 0.
     arrlWavMin(lColCount) = (lMin \ 133152) + 64 'Long values become +/-64 then
     arrLWavMax(lColCount) = (lMax \ 133152) + 64 'add 64 to make into co-ordinates
     End If
    Next lColCount
'Loop Until EOF(1)
'JA! filen slutar med LISTB00INFOISFTS bla bla00
End Sub
Private Sub FillarrLWavMax24Speedy(nChannels)
Dim lColCount As Long, ctr As Long, lMin As Long, lMinR As Long, lMax As Long, lMaxR As Long
Dim l24bit As Long, i As Long, bteTest As Byte, lLocInFile As Long, iTotBytesPcol  As Integer
'SIGNED 24 BYTE INTEGERS INTE FLOAT!!!
iTotBytesPcol = UBound(arrlCandidatesPerCol) * THREE_BYTESin24BIT

         '========= Processing one column at a time! ===========
lLocInFile = Loc(bteFreeFile)
Seek bteFreeFile, (lLocInFile + 2)  '+2 forwards ONE position. I DON'T UNDERSTAND WHY IT IS NEEDED!!!
lLocInFile = Loc(bteFreeFile)
'========= Looping from column number zero to ipicBoxWidth - 1 ====================
For lColCount = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1
   i = 1: lLocInFile = Loc(bteFreeFile)
   '===== Filling the matrix arrlCandidatesPerCol ==================================
   '===== Fetching a huge block at a time me 32bitLongs because it's optimizing the drives mecanical performance. ===============
   For ctr = lLocInFile To lLocInFile + iTotBytesPcol - 1 Step 3 'Forcing the reading from disc to be carried aut with a distance of 3 bytes! Step 3!
    Get #bteFreeFile, ctr, arrlCandidatesPerCol(i)
    'Get #bteFreeFile, ctr, bteTest 'Fetching a huge block at a time me 32bitLongs because it's optimizing the drives mecanical performance.
    
    i = i + 1
   Next ctr
   '======= Convert 32bitars Long to 24bitarsShortLong! =======
   For ctr = 1 To UBound(arrlCandidatesPerCol)
     If arrlCandidatesPerCol(ctr) And &H800000 Then 'Shift 8 steps left.
      arrlCandidatesPerCol(ctr) = (arrlCandidatesPerCol(ctr) And &H7FFFFF) * &H100& Or &H80000000
    Else
      arrlCandidatesPerCol(ctr) = (arrlCandidatesPerCol(ctr) And &H7FFFFF) * &H100&
    End If
    arrlCandidatesPerCol(ctr) = (arrlCandidatesPerCol(ctr) And &HFFFFFF00) \ &H100& 'Shift 8 steps right and preserve parity.
   Next ctr
   '==============================================================
    'Find min/max to get a better view of the wave
    lMin = 0: lMinR = 0: lMax = 0: lMaxR = 0
    'Should look at all values but 'Step 32' speeds it up a bit. A nice C++ function to find the max/min would be handy here!
    If nChannels = 2 Then
         '----------- StereoFil. -----------
         For ctr = 1 To UBound(arrlCandidatesPerCol) - 1 Step 32 ' Step 3 tittar alltså varenda värde - varannan vänster/höger 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr + 1) < lMinR Then lMinR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
            If arrlCandidatesPerCol(ctr + 1) > lMaxR Then lMaxR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
         Next ctr
        Else
         '------------ MonoFil. ------------
         For ctr = 1 To UBound(arrlCandidatesPerCol) 'Step 32 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            'If arrlCandidatesPerCol.arriWavinfo(ctr) > 0 Then MsgBox arrlCandidatesPerCol.arriWavinfo(ctr)
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr)
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr)
         Next ctr
    End If
    'lColCount = lColCount + 1
    If nChannels = 2 Then
    'Stereofil \67108864 = Hex 4000 000! A PURE MULTIPLICATION FROM 16-bitversion.
     arrlWavMin(lColCount) = lMin \ 266304 + 32  'Left. Long values become +/-64 then
     arrlWavMinR(lColCount) = lMinR \ 266304 + 96  'Right. Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 266304 + 32 'Left. add 64 to make into co-ordinates
     arrLWavMaxR(lColCount) = lMaxR \ 266304 + 96  'Right. add 64 to make into co-ordinates
     'Debug.Print lMax; lMaxR
    Else
    'A 24bitsigned WaveFile has the ambitus of plusminus &H 7F FF FF & = plusminus 8.388.607!
    'monoFil. 33554432 = Hex 2000 000! eller 133152 A PURE MULTIPLICATION FROM 16-bitversion
    '*** Max deviation in picWave should be = 127 and minimum = 0.
     arrlWavMin(lColCount) = (lMin \ 133152) + 64 'Long values become +/-64 then
     arrLWavMax(lColCount) = (lMax \ 133152) + 64 'add 64 to make into co-ordinates
    End If
Next lColCount
'Loop Until EOF(1)
'JA! filen slutar med LISTB00INFOISFTS bla bla00
End Sub


Private Sub FillarrLWavMax32Integers(nChannels) 'This routine gets fast because of LongString being native for windows.
Dim lColCount As Long, ctr As Long, lMin As Long, lMinR As Long, lMax As Long, lMaxR As Long
Dim iTotBytesPcol As Integer, lLocInFile As Long
'SIGNED 32BYTE INTEGERS!!! Hösta biten visar pariteten+-. Max positive deviation = 7F FF FF FF. Max negative deviation FF FF FF FF.
iTotBytesPcol = UBound(arrlCandidatesPerCol) * FOUR_BYTESinLONG

         '========= Bearbetar en kolumn i taget! ===========
lLocInFile = Loc(bteFreeFile)
'Seek bteFreeFile, (lLocInFile + 2)  '+2 forwards ONE position. I DON'T UNDERSTAND WHY IT IS NEEDED!!!
'lLocInFile = Loc(bteFreeFile)

For lColCount = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1
    Get #bteFreeFile, , arrlCandidatesPerCol
    'Find min/max to get a better view of the wave
    lMin = 0: lMinR = 0: lMax = 0: lMaxR = 0
    'Should look at all values but 'Step 32' speeds it up a bit. A nice C++ function to find the max/min would be handy here!
    If nChannels = 2 Then
         '----------- StereoFil. -----------
         For ctr = 1 To UBound(arrlCandidatesPerCol) - 1 Step 2 ' Step 2 is thus examining every single complete value - alternating left/right 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr + 1) < lMinR Then lMinR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
            If arrlCandidatesPerCol(ctr + 1) > lMaxR Then lMaxR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
         Next ctr
        Else
         '------------ MonoFil. ------------
         For ctr = 1 To UBound(arrlCandidatesPerCol) 'Step 32 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            'If arrlCandidatesPerCol.arriWavinfo(ctr) > 0 Then MsgBox arrlCandidatesPerCol.arriWavinfo(ctr)
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr)
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr)
         Next ctr
    End If
    'lColCount = lColCount + 1
    If nChannels = 2 Then
    'Stereofil \67108864 = Hex 4000 000! A PURE MULTIPLICATION FROM 16-bitversion.
     arrlWavMin(lColCount) = lMin \ 67108864 + 32  'Left. Long values become +/-64 then
     arrlWavMinR(lColCount) = lMinR \ 67108864 + 96  'Right. Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 67108864 + 32 'Left. add 64 to make into co-ordinates
     arrLWavMaxR(lColCount) = lMaxR \ 67108864 + 96  'Right. add 64 to make into co-ordinates
     'Debug.Print lMax; lMaxR
    Else
    'monoFil. 33554432 = Hex 2000 000! A PURE MULTIPLICATION FROM 16-bitversion.
     arrlWavMin(lColCount) = lMin \ 33554432 + 64 'Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 33554432 + 64 'add 64 to make into co-ordinates
     End If
    Next lColCount
'Loop Until EOF(1)

End Sub
Private Sub FillarrLWavMax32Float(nChannels) 'FLOAT AUDIO ÄR TAL MELLAN 1 & MINUS 1!
Dim lColCount As Long, ctr As Long, lMin As Long, lMinR As Long, lMax As Long, lMaxR As Long
Dim iTotBytesPcol As Integer, lLocInFile As Long
Const PIC_BOX_HALF_HEIGHT_64 As Long = 64
'SIGNED 32BYTE FLOATS = decimaltal mellan -1 och +1!!!
iTotBytesPcol = UBound(arrSngCandidatesPerCol) * FOUR_BYTESinLONG

         '========= Processing one column at a time! ===========
'lLocInFile = Loc(bteFreeFile)
'Seek bteFreeFile, (lLocInFile + 2)  '+2 forwards ONE position. I DON'T UNDERSTAND WHY IT IS NEEDED!!!
lLocInFile = Loc(bteFreeFile)

For lColCount = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1
    Get #bteFreeFile, , arrSngCandidatesPerCol '4 byte float.
    'Find min/max to get a better view of the wave
    lMin = 0: lMinR = 0: lMax = 0: lMaxR = 0
    'Should look at all values but 'Step 32' speeds it up a bit. A nice C++ function to find the max/min would be handy here!
    If nChannels = 2 Then
         '----------- StereoFil. -----------
         For ctr = 1 To UBound(arrlCandidatesPerCol) - 1 Step 32 ' Step 2 tittar alltså varenda värde - varannan vänster/höger 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            If arrlCandidatesPerCol(ctr) < lMin Then lMin = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr) > lMax Then lMax = arrlCandidatesPerCol(ctr) 'Left Channel.
            If arrlCandidatesPerCol(ctr + 1) < lMinR Then lMinR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
            If arrlCandidatesPerCol(ctr + 1) > lMaxR Then lMaxR = arrlCandidatesPerCol(ctr + 1) 'Right Channel.
         Next ctr
        Else
         '------------ MonoFil. ------------
         For ctr = 1 To UBound(arrSngCandidatesPerCol) 'Step 32 'Väljer ut extremSAMPLEvärdena i ett tidsblock på 100 ms.
            'If arrlCandidatesPerCol.arriWavinfo(ctr) > 0 Then MsgBox arrlCandidatesPerCol.arriWavinfo(ctr)
            If arrSngCandidatesPerCol(ctr) * PIC_BOX_HALF_HEIGHT_64 < lMin Then lMin = arrSngCandidatesPerCol(ctr) * PIC_BOX_HALF_HEIGHT_64
            If arrSngCandidatesPerCol(ctr) * PIC_BOX_HALF_HEIGHT_64 > lMax Then lMax = arrSngCandidatesPerCol(ctr) * PIC_BOX_HALF_HEIGHT_64
         Next ctr
    End If
    'lColCount = lColCount + 1
    If nChannels = 2 Then
    'Stereofil \67108864 = Hex 4000 000! A PURE MULTIPLICATION FROM 16-bitversion.
     arrlWavMin(lColCount) = lMin \ 67108864 + 32  'Left. Long values become +/-64 then
     arrlWavMinR(lColCount) = lMinR \ 67108864 + 96  'Right. Long values become +/-64 then
     arrLWavMax(lColCount) = lMax \ 67108864 + 32 'Left. add 64 to make into co-ordinates
     arrLWavMaxR(lColCount) = lMaxR \ 67108864 + 96  'Right. add 64 to make into co-ordinates
     'Debug.Print lMax; lMaxR
    Else
    'monoFil. Lmin är som ytterst = -64. PicBoxHöjden = 127.
     arrlWavMin(lColCount) = lMin + 64 'Long values become +/-64 then
     arrLWavMax(lColCount) = lMax + 64 'add 64 to make into co-ordinates
     End If
    Next lColCount
'Loop Until EOF(1)

End Sub


'-------------------------------------------------
' Main drawing routine
'-------------------------------------------------
' MoveToEx and LineTo are GDI functions, many
' times faster than their VB equivalents and
' just as easy to use'
'-------------------------------------------------
Sub DrawWaveData(nChannels As Integer)
    Dim ctr As Long, vstart As Long, hline As Integer

    picWave.Cls
    picWave.ForeColor = vbGreen
    
    'Draw each line
 If nChannels = 2 Then
     'StereoFil.
    For ctr = 0 To (picWave.Width \ Screen.TwipsPerPixelX) - 1 '1600 'IIf(lColumnCount < 400, lColumnCount, 400)
        MoveToEx picWave.hDC, ctr, arrlWavMin(ctr), LastPt 'Left.
        LineTo picWave.hDC, ctr, arrLWavMax(ctr) 'Left.
        MoveToEx picWave.hDC, ctr, arrlWavMinR(ctr), LastPt 'Right.
        LineTo picWave.hDC, ctr, arrLWavMaxR(ctr) 'Right.
    Next ctr
    
    MoveToEx picWave.hDC, 0, 96, LastPt 'Draw left channel center line
    LineTo picWave.hDC, picWave.ScaleWidth, 96
    MoveToEx picWave.hDC, 0, 34, LastPt 'Draw right channel center line
    LineTo picWave.hDC, picWave.ScaleWidth, 34
    
 Else
     'MonoFil.
    For ctr = 0 To ipicBoxWidth - 1 '(picWave.Width \ Screen.TwipsPerPixelX) - 1 '1600 'IIf(lColumnCount < 400, lColumnCount, 400)
        MoveToEx picWave.hDC, ctr, arrlWavMin(ctr), LastPt
        LineTo picWave.hDC, ctr, arrLWavMax(ctr)
    Next ctr
    
    MoveToEx picWave.hDC, 0, 64, LastPt 'Draw center line
    LineTo picWave.hDC, picWave.ScaleWidth, 64
 End If
End Sub




