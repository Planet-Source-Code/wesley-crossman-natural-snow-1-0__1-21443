Attribute VB_Name = "ModSnow"
'used to get data for setting high priority
Declare Function GetCurrentProcess Lib "kernel32" () As Long
'used to apply a high priority
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
'used to check for a slow computer
Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)

'delay to return power to Windows
Public SnowDel%
'sets if the mouse can be used to place
'objects for snow to interact with
Public AllowMouse As Boolean
'controls if the snow is created automatically
Public CreateNewSnow As Boolean
'number of "snow move" cycles to undergo in each move call
Public NumInCycle%
'refresh snow every once every RefreshFreq times
Public RefreshFreq%
'controls if game is paused
Public Paused As Boolean
'if the snow falls from top or frame
Public GeneratorMode As Boolean
'snow color settings
Public RActive As Boolean, GActive As Boolean, BActive As Boolean
Public RMult!, GMult!, BMult!
'the program's current priority
Public HighPriority As Boolean
'snow's palette
Public SnowPalette&(1 To 255)


'this variable is overridden if NoFric or AbsFric are true
Public FrictionLevel As Integer

'*** the following variables should never both be at the same time ***

'no movement if particle in front
Public NoFric As Boolean
'always move if particle in front (if possible)
Public AbsFric As Boolean

Sub SetCustPalette()
For a = 1 To 255
 r! = IIf(RActive, a * RMult, 0)
 r = IIf(r > 255, 255, r)
 r = IIf(r < 0, 0, r)
 g! = IIf(GActive, a * GMult, 0)
 g = IIf(g > 255, 255, g)
 g = IIf(g < 0, 0, g)
 b! = IIf(BActive, a * BMult, 0)
 b = IIf(b > 255, 255, b)
 b = IIf(b < 0, 0, b)
 SnowPalette(a) = RGB(r, g, b)
Next
End Sub
