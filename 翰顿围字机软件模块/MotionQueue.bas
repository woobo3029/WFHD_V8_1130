Attribute VB_Name = "DeviceMotionQueue"
Option Explicit

Private Const MaxQueueCount = 20000

Public Type MotionEventType
    Task_Name As String * 20
    API_Name As String * 20
    d1 As Double
    d2 As Double
    d3 As Double
    d4 As Double
    d5 As Double
    d6 As Double
    d7 As Double
    d8 As Double
    action As ActionType
End Type

Public Type MotionQueueType
    Count As Long
    InPos As Long
    OutPos As Long
    Locked As Boolean
    MotionEvent() As MotionEventType
End Type

Public MotionQueue As MotionQueueType

Sub InitMotionQueue()
    ReDim MotionQueue.MotionEvent(MaxQueueCount)
    MotionQueue.Count = 0
    MotionQueue.InPos = 0
    MotionQueue.OutPos = 0
    MotionQueue.Locked = False
End Sub

Sub ResetMotionQueuePos()
    MotionQueue.Count = MotionQueue.InPos + 1
    MotionQueue.InPos = 0
    MotionQueue.OutPos = 0
    MotionQueue.Locked = False
End Sub

Sub UninitMotionQueue()
    Erase MotionQueue.MotionEvent
    MotionQueue.Count = 0
    MotionQueue.InPos = 0
    MotionQueue.OutPos = 0
    MotionQueue.Locked = True
End Sub

Function MotionQueueCount() As Integer
    MotionQueueCount = MotionQueue.Count
End Function

Sub LockMotionQueue()
    MotionQueue.Locked = True
End Sub

Sub UnlockMotionQueue()
    MotionQueue.Locked = False
End Sub

Function IsMotionQueueLocked() As Boolean
    IsMotionQueueLocked = MotionQueue.Locked
End Function

Function PutMotionQueue(MotionEvent As MotionEventType) As Boolean
    Dim I As Byte
    Dim t As Variant
    
    'Debug.Print ">>> Put"
    With MotionQueue
        If .Count >= MaxQueueCount Then
            PutMotionQueue = False
            Debug.Print "<<< Error: Count >= MaxQueueCount"
            Exit Function
        End If
            
        t = Timer
        Do While .Locked And TimeDiff(Timer, t) < 5
            DoEvents
        Loop
        If .Locked Then
            PutMotionQueue = False
            Debug.Print "<<< Error: Locked & Time out"
            Exit Function
        End If
        
        .Locked = True
        
        .MotionEvent(.InPos) = MotionEvent

        .InPos = (.InPos + 1) Mod MaxQueueCount
        .Locked = False
        
        .Count = .Count + 1
        PutMotionQueue = True
        'Debug.Print "<<< Ok "; .Count
    End With
End Function

Function GetMotionQueue(MotionEvent As MotionEventType) As Boolean
    Dim I As Byte
    Dim t As Variant
    
    With MotionQueue
        If .Count <= 0 Then
            GetMotionQueue = False
            'Debug.Print "<<< Error: Count <=0"
            Exit Function
        End If
        'Debug.Print ">>> Get"
           
        t = Timer
        Do While .Locked And TimeDiff(Timer, t) < 5
            DoEvents
        Loop
        If .Locked Then
            GetMotionQueue = False
            'Debug.Print "<<< Error: Locked & Time out"
            Exit Function
        End If
        
        .Locked = True
        
        MotionEvent = .MotionEvent(.OutPos)
        
        .OutPos = (.OutPos + 1) Mod MaxQueueCount
        .Locked = False
        
        'Debug.Print "<<< Ok "; .Count
        .Count = .Count - 1
        GetMotionQueue = True
    End With
End Function

