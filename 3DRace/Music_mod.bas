Attribute VB_Name = "Music_mod"
Option Explicit
Private dx As New DirectX7
Private perf As DirectMusicPerformance
Private perf2 As DirectMusicPerformance
Private seg As DirectMusicSegment
Private segstate As DirectMusicSegmentState
Private loader As DirectMusicLoader
Private GetStartTime As Long
Private Offset As Long
Private mtTime As Long
Private mtLength As Double
Private dTempo As Double
Private timesig As DMUS_TIMESIGNATURE
Private IsPlayingCheck As Boolean
Private msg As String
Private time As Double
Private fIsPaused As Boolean
Private ISITPAUSED As Boolean
Private Total_Time As Double
Private Current_Time As Double
Private Percent_Time As Double

Private InitDM As Boolean

Public MusicVolume As Integer

Public MusicFileName(1) As String

Sub EndMusic()
    Call Music_mod.StopMusic
    Set perf = Nothing
    Set perf2 = Nothing
    Set loader = Nothing
End Sub

Sub Initialize_Music()
On Error GoTo MusOut:

    MusicFileName(0) = "Music.mid"

    Set loader = dx.DirectMusicLoaderCreate()
    Set perf2 = dx.DirectMusicPerformanceCreate()
    Call perf2.Init(Nothing, 0)
    perf2.SetPort -1, 80
    Call perf2.GetMasterAutoDownload
    Set perf = dx.DirectMusicPerformanceCreate()
    Call perf.Init(Nothing, 0)
    perf.SetPort -1, 80
    Call perf.SetMasterAutoDownload(True)
    perf.SetMasterVolume (50 * 42 - 3000)
    InitDM = True
    
    Exit Sub
MusOut:
    InitDM = False
End Sub

Function IsMusicAtEnd() As Boolean
    If InitDM = False Then Exit Function
    If perf.IsPlaying(seg, segstate) = False Then
       IsMusicAtEnd = True
    Else
        IsMusicAtEnd = False
    End If
End Function

Sub PlayMusic()
If InitDM = False Then Exit Sub
  If seg Is Nothing Then
        Exit Sub
    End If
    If fIsPaused Then
        Offset = mtTime - GetStartTime + Offset + 1
        Call seg.SetStartPoint(Offset)
        Set segstate = perf.PlaySegment(seg, 0, 0)
    Else
        Offset = 0
        If perf.IsPlaying(seg, segstate) = True Then
            Call perf.Stop(seg, segstate, 0, 0)
        End If
        seg.SetStartPoint (0)
        Set segstate = perf.PlaySegment(seg, 0, 0)
        Exit Sub
    End If
    fIsPaused = False
End Sub
Sub StopMusic()
    If InitDM = False Then Exit Sub
    If segstate Is Nothing Then
    Exit Sub
    End If
    Call perf.Stop(seg, segstate, 0, 0)
    time = 0
End Sub
Sub SetMusic(Volume As Integer)
    If InitDM = False Then Exit Sub
    perf.SetMasterVolume (Volume * 42 - 3000)
End Sub
Sub LoopMusic()
If InitDM = False Then Exit Sub
If perf.IsPlaying(seg, segstate) = False Then
    seg.SetStartPoint (0)
    Set segstate = perf.PlaySegment(seg, 0, 0)
End If
End Sub
Sub Load_Music(FileNumber As Byte)
    Dim FileName As String
    If InitDM = False Then Exit Sub
    
    FileName = App.path & "\Audio\" & MusicFileName(FileNumber)
    Dim Minutes As Integer
    Dim a As Integer
    Dim length As Integer
    Dim length2 As Integer
    
    On Error GoTo LocalErrors
    
    If Not seg Is Nothing And Not segstate Is Nothing Then
        If perf.IsPlaying(seg, segstate) = True Then
            Call perf.Stop(seg, segstate, 0, 0)
        ElseIf ISITPAUSED = True Then
            Call perf.Stop(seg, segstate, 0, 0)
        End If
    End If
    
    Set loader = Nothing
    Set loader = dx.DirectMusicLoaderCreate
    Set seg = loader.LoadSegment(FileName)
    length = Len(FileName)
    length2 = length
    Dim path As String
    Do While path <> "\"
    path = Mid(FileName, length, 1)
    length = length - 1
    Loop
    Dim SearchDir As String
    SearchDir = left(FileName, length)
    loader.SetSearchDirectory (left(FileName, length + 1))
    perf2.SetMasterAutoDownload True
    
    mtTime = perf2.GetMusicTime()
    Call perf2.PlaySegment(seg, 0, mtTime + 2000)
    
    dTempo = perf2.GetTempo(mtTime + 2000, 0)
    mtLength = (((seg.GetLength() / 768) * 60) / dTempo)
    Total_Time = mtLength
    Call perf2.Stop(seg, Nothing, 0, 0)
    seg.SetStandardMidiFile
            
    Exit Sub
LocalErrors:
    If Not seg Is Nothing Then
        Call perf2.Stop(seg, Nothing, 0, 0)
    End If
    FileName = vbNullString
End Sub

