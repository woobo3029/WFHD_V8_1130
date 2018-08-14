Attribute VB_Name = "Module3"
'Attribute VB_Name = "Module1"
Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
    End Type
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function SetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Integer, ByVal dwProcessAffinityMask As Long) As Boolean
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


