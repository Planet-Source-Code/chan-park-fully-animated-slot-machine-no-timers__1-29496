Attribute VB_Name = "mdlTimeFunctions"
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Returns the number of loops the running computer can execute in a second
'If bDoevents is true, then test will use Doevents in loop.
'lTestLoops is the number of loops the test will use. Increase for greater accuracy.
'lTestLoops should be atleast 100000 if bDoevents = false.
'lNoTries is the number of tests. Increase for greater accuracy.
'Returns 0 if lTestLoops is too small
'Note: This function can be used to make for-loop-delays run uniformly on all computers
Public Function LoopsPerSecond(ByVal bDoevents As Boolean, ByVal lTestLoops As Long, ByVal lNoTries As Long) As Long
Dim lOrig As Long
Dim lNew As Long

Dim lNoLoops As Long
Dim lTotalTime As Long

Dim i As Long
Dim j As Long

If bDoevents = False Then
    For i = 1 To lNoTries
        lOrig = GetTickCount()
        For j = 1 To lTestLoops
        Next
        lNew = GetTickCount()
        lTotalTime = lTotalTime + (lNew - lOrig)
    Next
Else
    For i = 1 To lNoTries
        lOrig = GetTickCount()
        For j = 1 To lTestLoops
            DoEvents
        Next
        lNew = GetTickCount()
        lTotalTime = lTotalTime + (lNew - lOrig)
    Next
End If

lNoLoops = lTestLoops * lNoTries

If lTotalTime = 0 Then
    LoopsPerSecond = 0
Else
    LoopsPerSecond = Int(lNoLoops / lTotalTime * 1000)
End If
End Function
