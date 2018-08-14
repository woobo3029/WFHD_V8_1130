Attribute VB_Name = "DMCCOM"
' Windows interface to Galil Motion Controllers

' For Visual Basic 4.0 and Higher Only!

' All functions return an error code. 0 is function completed successfully.
' Any error code < 0 is a local error (see the error codes belong).
' Any error code > 0 is an Win32 API error.
' These are documented in the Win32 Programming Reference.

' Constant values

' Controller models
Public Const DMC400 = "DMC-400"
Public Const DMC600 = "DMC-600"
Public Const DMC700 = "DMC-700"
Public Const DMC1000 = "DMC-1000"
Public Const DMC1200 = "DMC-1200"
Public Const DMC1410 = "DMC-1410"
Public Const DMC1411 = "DMC-1411"
Public Const DMC1412 = "DMC-1412"
Public Const DMC1417 = "DMC-1417"
Public Const DMC1500 = "DMC-1500"
Public Const DMC1600 = "DMC-1600"
Public Const DMC1700 = "DMC-1700"
Public Const DMC1800 = "DMC-1800"
Public Const DMC1802 = "DMC-1802"
Public Const DMC2000 = "DMC-2000"
Global Const DMC2100 = "DMC-2100"
Global Const DMC90064 = "IOC-90064"
 
' Error Codes
Public Const DMCNOERROR = 0
Public Const DMCERROR_TIMEOUT = -1
Public Const DMCERROR_COMMAND = -2
Public Const DMCERROR_CONTROLLER = -3
Public Const DMCERROR_FILE = -4
Public Const DMCERROR_DRIVER = -5
Public Const DMCERROR_HANDLE = -6
Public Const DMCERROR_HMODULE = -7
Public Const DMCERROR_MEMORY = -8
Public Const DMCERROR_BUFFERFULL = -9
Public Const DMCERROR_RESPONSEDATA = -10
Public Const DMCERROR_DMA = -11
Public Const DMCERROR_ARGUMENT = -12
Public Const DMCERROR_DATARECORD = -13
Public Const DMCERROR_DOWNLOAD = -14
Public Const DMCERROR_FIRMWARE = -15
Public Const DMCERROR_CONVERSION = -16
Public Const DMCERROR_RESOURCE = -17
Public Const DMCERROR_REGISTRY = -18
Public Const DMCERROR_BUSY = -19
Public Const DMCERROR_DEVICE_DISCONNECTED = -20

' IMPORTANT: Constant values for data record item offsets can change between
' firmware revisions. Use the QZ command or the function DMCGetDataRecordRevision
' to determine what revision of data record access you have.

' The DMCGetDataRecordByItemId function retrieves a data record item by unique
' Id while the DMCGetDataRecord function retrieves a data record item by offset.
' While data record item offsets can change with the firmware revision, the data
' record item Ids always remain the same.

' Constant values for data record data types
Public Const DRTypeUnknown = 0
Public Const DRTypeCHAR = 1
Public Const DRTypeUCHAR = 2
Public Const DRTypeSHORT = 3
Public Const DRTypeUSHORT = 4
Public Const DRTypeLONG = 5
Public Const DRTypeULONG = 6

' Constant values for data record item Ids to be used with the function
' DMCGetDataRecordByItemId
Public Const DRIdSampleNumber = 1
Public Const DRIdGeneralInput0 = 2
Public Const DRIdGeneralInput1 = 3
Public Const DRIdGeneralInput2 = 4
Public Const DRIdGeneralInput3 = 5
Public Const DRIdGeneralInput4 = 6
Public Const DRIdGeneralInput5 = 7
Public Const DRIdGeneralInput6 = 8
Public Const DRIdGeneralInput7 = 9
Public Const DRIdGeneralInput8 = 10
Public Const DRIdGeneralInput9 = 11
Public Const DRIdGeneralOutput0 = 12
Public Const DRIdGeneralOutput1 = 13
Public Const DRIdGeneralOutput2 = 14
Public Const DRIdGeneralOutput3 = 15
Public Const DRIdGeneralOutput4 = 16
Public Const DRIdGeneralOutput5 = 17
Public Const DRIdGeneralOutput6 = 18
Public Const DRIdGeneralOutput7 = 19
Public Const DRIdGeneralOutput8 = 20
Public Const DRIdGeneralOutput9 = 21
Public Const DRIdErrorCode = 22
Public Const DRIdGeneralStatus = 23
Public Const DRIdSegmentCountS = 24
Public Const DRIdCoordinatedMoveStatusS = 25
Public Const DRIdCoordinatedMoveDistanceS = 26
Public Const DRIdSegmentCountT = 27
Public Const DRIdCoordinatedMoveStatusT = 28
Public Const DRIdCoordinatedMoveDistanceT = 29
Public Const DRIdAnalogInput1 = 30
Public Const DRIdAnalogInput2 = 31
Public Const DRIdAnalogInput3 = 32
Public Const DRIdAnalogInput4 = 33
Public Const DRIdAnalogInput5 = 34
Public Const DRIdAnalogInput6 = 35
Public Const DRIdAnalogInput7 = 36
Public Const DRIdAnalogInput8 = 37
Public Const DRIdAxisStatus = 38
Public Const DRIdAxisSwitches = 39
Public Const DRIdAxisStopCode = 40
Public Const DRIdAxisReferencePosition = 41
Public Const DRIdAxisMotorPosition = 42
Public Const DRIdAxisPositionError = 43
Public Const DRIdAxisAuxillaryPosition = 44
Public Const DRIdAxisVelocity = 45
Public Const DRIdAxisTorque = 46

' Constant values for axis Ids to be used with the function
' DMCGetDataRecordByItemId
Public Const DRIdAxis1 = 1
Public Const DRIdAxis2 = 2
Public Const DRIdAxis3 = 3
Public Const DRIdAxis4 = 4
Public Const DRIdAxis5 = 5
Public Const DRIdAxis6 = 6
Public Const DRIdAxis7 = 7
Public Const DRIdAxis8 = 8

' Data record offsets

' Rev 1 constants
'    QZ command returns <#axes>,12,6,26

' Rev 2 constants
'    QZ command returns <#axes>,26,6,26
'    This rev added items to the general section for extended I/O.

' Rev 3 constants
'    QZ command returns <#axes>,24,16,26
'    This rev added items to the general section for the coordinated motion T axis.

' Rev 4 constants
'    QZ command returns <#axes>,24,16,28
'    This rev added 1 item to the axis section for analog inputs.
'    Note: each axis will now include the current value for 1 analog input.
'    X axis - analog 1, Y axis - analog 2, and so on. You must have an 8 axis
'    controller to get data for all 8 analog inputs.

' Rev 5 constants
'    QZ command returns 0,8,0,0
'    This rev added to accomodate the IOC-90064.
'    Note: this card's data record is much smaller compared
'          to previous revisions.  The sample number, error
'          code, general status, and 8 general inputs and 8 general
'          outputs are supported.

' Rev 1 General data item offsets
Public Const REV1GenOffSampleNumber = 0
Public Const REV1GenOffGeneralInput1 = 2
Public Const REV1GenOffGeneralInput2 = 3
Public Const REV1GenOffGeneralInput3 = 4
Public Const REV1GenOffSpare = 5
Public Const REV1GenOffGeneralOutput1 = 6
Public Const REV1GenOffGeneralOutput2 = 7
Public Const REV1GenOffErrorCode = 8
Public Const REV1GenOffGeneralStatus = 9
Public Const REV1GenOffSegmentCount = 10
Public Const REV1GenOffCoordinatedMoveStatus = 12
Public Const REV1GenOffCoordinatedMoveDistance = 14
Public Const REV1GenOffAxis1 = 18
Public Const REV1GenOffAxis2 = 44
Public Const REV1GenOffAxis3 = 70
Public Const REV1GenOffAxis4 = 96
Public Const REV1GenOffAxis5 = 122
Public Const REV1GenOffAxis6 = 148
Public Const REV1GenOffAxis7 = 174
Public Const REV1GenOffAxis8 = 200
Public Const REV1GenOffEnd = 226

' Rev 1 axis data item offsets
Public Const REV1AxisOffNoAxis = 0
Public Const REV1AxisOffAxisStatus = 0
Public Const REV1AxisOffAxisSwitches = 2
Public Const REV1AxisOffAxisStopCode = 3
Public Const REV1AxisOffAxisReferencePosition = 4
Public Const REV1AxisOffAxisMotorPosition = 8
Public Const REV1AxisOffAxisPositionError = 12
Public Const REV1AxisOffAxisAuxillaryPosition = 16
Public Const REV1AxisOffAxisVelocity = 20
Public Const REV1AxisOffAxisTorque = 24
Public Const REV1AxisOffEnd = 26

' Rev 2 General data item offsets
Public Const REV2GenOffSampleNumber = 0
Public Const REV2GenOffGeneralInput0 = 2
Public Const REV2GenOffGeneralInput1 = 3
Public Const REV2GenOffGeneralInput2 = 4
Public Const REV2GenOffGeneralInput3 = 5
Public Const REV2GenOffGeneralInput4 = 6
Public Const REV2GenOffGeneralInput5 = 7
Public Const REV2GenOffGeneralInput6 = 8
Public Const REV2GenOffGeneralInput7 = 9
Public Const REV2GenOffGeneralInput8 = 10
Public Const REV2GenOffGeneralInput9 = 11
Public Const REV2GenOffGeneralOutput0 = 12
Public Const REV2GenOffGeneralOutput1 = 13
Public Const REV2GenOffGeneralOutput2 = 14
Public Const REV2GenOffGeneralOutput3 = 15
Public Const REV2GenOffGeneralOutput4 = 16
Public Const REV2GenOffGeneralOutput5 = 17
Public Const REV2GenOffGeneralOutput6 = 18
Public Const REV2GenOffGeneralOutput7 = 19
Public Const REV2GenOffGeneralOutput8 = 20
Public Const REV2GenOffGeneralOutput9 = 21
Public Const REV2GenOffErrorCode = 22
Public Const REV2GenOffGeneralStatus = 23
Public Const REV2GenOffSegmentCount = 24
Public Const REV2GenOffCoordinatedMoveStatus = 26
Public Const REV2GenOffCoordinatedMoveDistance = 28
Public Const REV2GenOffAxis1 = 32
Public Const REV2GenOffAxis2 = 58
Public Const REV2GenOffAxis3 = 84
Public Const REV2GenOffAxis4 = 110
Public Const REV2GenOffAxis5 = 136
Public Const REV2GenOffAxis6 = 162
Public Const REV2GenOffAxis7 = 188
Public Const REV2GenOffAxis8 = 214
Public Const REV2GenOffEnd = 240

' Rev 2 axis data item offsets
Public Const REV2AxisOffNoAxis = 0
Public Const REV2AxisOffAxisStatus = 0
Public Const REV2AxisOffAxisSwitches = 2
Public Const REV2AxisOffAxisStopCode = 3
Public Const REV2AxisOffAxisReferencePosition = 4
Public Const REV2AxisOffAxisMotorPosition = 8
Public Const REV2AxisOffAxisPositionError = 12
Public Const REV2AxisOffAxisAuxillaryPosition = 16
Public Const REV2AxisOffAxisVelocity = 20
Public Const REV2AxisOffAxisTorque = 24
Public Const REV2AxisOffEnd = 26

' Rev 3 General data item offsets
Public Const REV3GenOffSampleNumber = 0
Public Const REV3GenOffGeneralInput0 = 2
Public Const REV3GenOffGeneralInput1 = 3
Public Const REV3GenOffGeneralInput2 = 4
Public Const REV3GenOffGeneralInput3 = 5
Public Const REV3GenOffGeneralInput4 = 6
Public Const REV3GenOffGeneralInput5 = 7
Public Const REV3GenOffGeneralInput6 = 8
Public Const REV3GenOffGeneralInput7 = 9
Public Const REV3GenOffGeneralInput8 = 10
Public Const REV3GenOffGeneralInput9 = 11
Public Const REV3GenOffGeneralOutput0 = 12
Public Const REV3GenOffGeneralOutput1 = 13
Public Const REV3GenOffGeneralOutput2 = 14
Public Const REV3GenOffGeneralOutput3 = 15
Public Const REV3GenOffGeneralOutput4 = 16
Public Const REV3GenOffGeneralOutput5 = 17
Public Const REV3GenOffGeneralOutput6 = 18
Public Const REV3GenOffGeneralOutput7 = 19
Public Const REV3GenOffGeneralOutput8 = 20
Public Const REV3GenOffGeneralOutput9 = 21
Public Const REV3GenOffErrorCode = 22
Public Const REV3GenOffGeneralStatus = 23
Public Const REV3GenOffSegmentCountS = 24
Public Const REV3GenOffCoordinatedMoveStatusS = 26
Public Const REV3GenOffCoordinatedMoveDistanceS = 28
Public Const REV3GenOffSegmentCountT = 32
Public Const REV3GenOffCoordinatedMoveStatusT = 34
Public Const REV3GenOffCoordinatedMoveDistanceT = 36
Public Const REV3GenOffAxis1 = 40
Public Const REV3GenOffAxis2 = 66
Public Const REV3GenOffAxis3 = 92
Public Const REV3GenOffAxis4 = 118
Public Const REV3GenOffAxis5 = 144
Public Const REV3GenOffAxis6 = 170
Public Const REV3GenOffAxis7 = 196
Public Const REV3GenOffAxis8 = 222
Public Const REV3GenOffEnd = 248

' Rev 3 axis data item offsets
Public Const REV3AxisOffNoAxis = 0
Public Const REV3AxisOffAxisStatus = 0
Public Const REV3AxisOffAxisSwitches = 2
Public Const REV3AxisOffAxisStopCode = 3
Public Const REV3AxisOffAxisReferencePosition = 4
Public Const REV3AxisOffAxisMotorPosition = 8
Public Const REV3AxisOffAxisPositionError = 12
Public Const REV3AxisOffAxisAuxillaryPosition = 16
Public Const REV3AxisOffAxisVelocity = 20
Public Const REV3AxisOffAxisTorque = 24
Public Const REV3AxisOffEnd = 26

' Rev 4 General data item offsets
Public Const REV4GenOffSampleNumber = 0
Public Const REV4GenOffGeneralInput0 = 2
Public Const REV4GenOffGeneralInput1 = 3
Public Const REV4GenOffGeneralInput2 = 4
Public Const REV4GenOffGeneralInput3 = 5
Public Const REV4GenOffGeneralInput4 = 6
Public Const REV4GenOffGeneralInput5 = 7
Public Const REV4GenOffGeneralInput6 = 8
Public Const REV4GenOffGeneralInput7 = 9
Public Const REV4GenOffGeneralInput8 = 10
Public Const REV4GenOffGeneralInput9 = 11
Public Const REV4GenOffGeneralOutput0 = 12
Public Const REV4GenOffGeneralOutput1 = 13
Public Const REV4GenOffGeneralOutput2 = 14
Public Const REV4GenOffGeneralOutput3 = 15
Public Const REV4GenOffGeneralOutput4 = 16
Public Const REV4GenOffGeneralOutput5 = 17
Public Const REV4GenOffGeneralOutput6 = 18
Public Const REV4GenOffGeneralOutput7 = 19
Public Const REV4GenOffGeneralOutput8 = 20
Public Const REV4GenOffGeneralOutput9 = 21
Public Const REV4GenOffErrorCode = 22
Public Const REV4GenOffGeneralStatus = 23
Public Const REV4GenOffSegmentCountS = 24
Public Const REV4GenOffCoordinatedMoveStatusS = 26
Public Const REV4GenOffCoordinatedMoveDistanceS = 28
Public Const REV4GenOffSegmentCountT = 32
Public Const REV4GenOffCoordinatedMoveStatusT = 34
Public Const REV4GenOffCoordinatedMoveDistanceT = 36
Public Const REV4GenOffAxis1 = 40
Public Const REV4GenOffAxis2 = 68
Public Const REV4GenOffAxis3 = 96
Public Const REV4GenOffAxis4 = 124
Public Const REV4GenOffAxis5 = 152
Public Const REV4GenOffAxis6 = 180
Public Const REV4GenOffAxis7 = 208
Public Const REV4GenOffAxis8 = 236
Public Const REV4GenOffEnd = 264

' Rev 4 axis data item offsets
Public Const REV4AxisOffNoAxis = 0
Public Const REV4AxisOffAxisStatus = 0
Public Const REV4AxisOffAxisSwitches = 2
Public Const REV4AxisOffAxisStopCode = 3
Public Const REV4AxisOffAxisReferencePosition = 4
Public Const REV4AxisOffAxisMotorPosition = 8
Public Const REV4AxisOffAxisPositionError = 12
Public Const REV4AxisOffAxisAuxillaryPosition = 16
Public Const REV4AxisOffAxisVelocity = 20
Public Const REV4AxisOffAxisTorque = 24
Public Const REV4AxisOffAnalogInput = 26
Public Const REV4AxisOffEnd = 28

' Rev 5 General data item offsets
Public Const DRREV5GenOffSampleNumber = 0
Public Const DRREV5GenOffConfigByte = 2
Public Const DRREV5GenOffGeneralIO0 = 3
Public Const DRREV5GenOffGeneralIO1 = 4
Public Const DRREV5GenOffGeneralIO2 = 5
Public Const DRREV5GenOffGeneralIO3 = 6
Public Const DRREV5GenOffGeneralIO4 = 7
Public Const DRREV5GenOffGeneralIO5 = 8
Public Const DRREV5GenOffGeneralIO6 = 9
Public Const DRREV5GenOffGeneralIO7 = 10
Public Const DRREV5GenOffErrorCode = 11
Public Const DRREV5GenOffGeneralStatus = 12

' ** The following constants are OBSOLETE **
' General offsets for firmware without coordinated motion T axis - data record revsion 2
Public Const DRGenOffsetsSampleNumber = 0
Public Const DRGenOffsetsGeneralInput0 = 2
Public Const DRGenOffsetsGeneralInput1 = 3
Public Const DRGenOffsetsGeneralInput2 = 4
Public Const DRGenOffsetsGeneralInput3 = 5
Public Const DRGenOffsetsGeneralInput4 = 6
Public Const DRGenOffsetsGeneralInput5 = 7
Public Const DRGenOffsetsGeneralInput6 = 8
Public Const DRGenOffsetsGeneralInput7 = 9
Public Const DRGenOffsetsGeneralInput8 = 10
Public Const DRGenOffsetsGeneralInput9 = 11
Public Const DRGenOffsetsGeneralOutput0 = 12
Public Const DRGenOffsetsGeneralOutput1 = 13
Public Const DRGenOffsetsGeneralOutput2 = 14
Public Const DRGenOffsetsGeneralOutput3 = 15
Public Const DRGenOffsetsGeneralOutput4 = 16
Public Const DRGenOffsetsGeneralOutput5 = 17
Public Const DRGenOffsetsGeneralOutput6 = 18
Public Const DRGenOffsetsGeneralOutput7 = 19
Public Const DRGenOffsetsGeneralOutput8 = 20
Public Const DRGenOffsetsGeneralOutput9 = 21
Public Const DRGenOffsetsErrorCode = 22
Public Const DRGenOffsetsGeneralStatus = 23
Public Const DRGenOffsetsSegmentCount = 24
Public Const DRGenOffsetsCoordinatedMoveStatus = 26
Public Const DRGenOffsetsCoordinatedMoveDistance = 28
Public Const DRGenOffsetsAxis1 = 32
Public Const DRGenOffsetsAxis2 = 58
Public Const DRGenOffsetsAxis3 = 84
Public Const DRGenOffsetsAxis4 = 110
Public Const DRGenOffsetsAxis5 = 136
Public Const DRGenOffsetsAxis6 = 162
Public Const DRGenOffsetsAxis7 = 188
Public Const DRGenOffsetsAxis8 = 214
Public Const DRGenOffsetsEnd = 240

' ** The following constants are OBSOLETE **
' General offsets for firmware with coordinated motion T axis - data record revsion 3
Public Const wTDRGenOffsetsSampleNumber = 0
Public Const wTDRGenOffsetsGeneralInput0 = 2
Public Const wTDRGenOffsetsGeneralInput1 = 3
Public Const wTDRGenOffsetsGeneralInput2 = 4
Public Const wTDRGenOffsetsGeneralInput3 = 5
Public Const wTDRGenOffsetsGeneralInput4 = 6
Public Const wTDRGenOffsetsGeneralInput5 = 7
Public Const wTDRGenOffsetsGeneralInput6 = 8
Public Const wTDRGenOffsetsGeneralInput7 = 9
Public Const wTDRGenOffsetsGeneralInput8 = 10
Public Const wTDRGenOffsetsGeneralInput9 = 11
Public Const wTDRGenOffsetsGeneralOutput0 = 12
Public Const wTDRGenOffsetsGeneralOutput1 = 13
Public Const wTDRGenOffsetsGeneralOutput2 = 14
Public Const wTDRGenOffsetsGeneralOutput3 = 15
Public Const wTDRGenOffsetsGeneralOutput4 = 16
Public Const wTDRGenOffsetsGeneralOutput5 = 17
Public Const wTDRGenOffsetsGeneralOutput6 = 18
Public Const wTDRGenOffsetsGeneralOutput7 = 19
Public Const wTDRGenOffsetsGeneralOutput8 = 20
Public Const wTDRGenOffsetsGeneralOutput9 = 21
Public Const wTDRGenOffsetsErrorCode = 22
Public Const wTDRGenOffsetsGeneralStatus = 23
Public Const wTDRGenOffsetsSegmentCountS = 24
Public Const wTDRGenOffsetsCoordinatedMoveStatusS = 26
Public Const wTDRGenOffsetsCoordinatedMoveDistanceS = 28
Public Const wTDRGenOffsetsSegmentCountT = 32
Public Const wTDRGenOffsetsCoordinatedMoveStatusT = 34
Public Const wTDRGenOffsetsCoordinatedMoveDistanceT = 36
Public Const wTDRGenOffsetsAxis1 = 40
Public Const wTDRGenOffsetsAxis2 = 66
Public Const wTDRGenOffsetsAxis3 = 92
Public Const wTDRGenOffsetsAxis4 = 118
Public Const wTDRGenOffsetsAxis5 = 144
Public Const wTDRGenOffsetsAxis6 = 170
Public Const wTDRGenOffsetsAxis7 = 196
Public Const wTDRGenOffsetsAxis8 = 222
Public Const wTDRGenOffsetsEnd = 248

' Constant values for data record axis data item offsets
' IMPORTANT - Values can change between revisions
Public Const DRAxisOffsetsNoAxis = 0
Public Const DRAxisOffsetsAxisStatus = 0
Public Const DRAxisOffsetsAxisSwitches = 2
Public Const DRAxisOffsetsAxisStopCode = 3
Public Const DRAxisOffsetsAxisReferencePosition = 4
Public Const DRAxisOffsetsAxisMotorPosition = 8
Public Const DRAxisOffsetsAxisPositionError = 12
Public Const DRAxisOffsetsAxisAuxillaryPosition = 16
Public Const DRAxisOffsetsAxisVelocity = 20
Public Const DRAxisOffsetsAxisTorque = 24
Public Const DRAxisOffsetsEnd = 26

' Constant values for GALILREGISTRY structure

' Controller Type
Public Const ControllerTypeISABus = 0
Public Const ControllerTypeSerial = 1
Public Const ControllerTypePCIBus = 2
Public Const ControllerTypeUSB = 3

' Device Drivers
Public Const DeviceDriverWinRT = 0
Public Const DeviceDriverGalil = 1

' Serial Handshake
Public Const SerialHandshakeHardware = 0
Public Const SerialHandshakeSoftware = 1

' Data Record Access
Public Const DataRecordAccessNone = 0
Public Const DataRecordAccessDMA = 1
Public Const DataRecordAccessFIFO = 2

' Ethernet Protocol
Global Const EthernetProtocolTCP = 0
Global Const EthernetProtocolUDP = 1

' Structures

' To add/change/delete registry information
Type GALILREGISTRY
        Model As String * 16
        DeviceNumber As Integer
        DeviceDriver As Integer
        Timeout As Long
        Delay As Long
        ControllerType As Integer
        CommPort As Integer
        CommSpeed As Long
        Handshake As Integer
        Address As Integer
        Interrupt As Integer
        DataRecordAccess As Integer
        DMAChannel As Integer
        DataRecordSize As Integer
        RefreshRate As Integer
        SerialNumber As Integer
        PNPHardwareKey As String * 64
End Type

' Function prototypes

#If Win32 Then
Public Declare Function DMCOpen Lib "dmc32.dll" (ByVal Controller As Integer, ByVal hWnd As Long, phDmc As Long) As Long
' Open communications with the Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' hWnd             The window handle to use for notifying the application
'                  program of an interrupt.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCOpen2 Lib "dmc32.dll" (ByVal Controller As Integer, ByVal ThreadID As Long, phDmc As Long) As Long
' Open communications with the Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' ThreadID         The thread id to use for notifying the application
'                  program of an interrupt.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCGetHandle Lib "dmc32.dll" (ByVal Controller As Integer, phDmc As Long) As Long
' Get the handle associated with a particular Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCClose Lib "dmc32.dll" (ByVal hDmc As Long) As Long
' Close communications with the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCCommand Lib "dmc32.dll" (ByVal hDmc As Long, ByVal CommandString As String, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Send a command to the Galil controller.
' NOTE: This function can only send commands or groups of commands up to
' 1024 bytes long.

' hDmc             Handle to the Galil controller.
' CommandString    The command to send to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCFastCommand Lib "dmc32.dll" (ByVal hDmc As Long, ByVal CommandString As String) As Long
' Send a command to the Galil controller without the overhead of waiting for a response. Use this function with
' caution as command errors will not be reported and the out-going FIFO or communciations buffer
' may fill up. This function is intended to be used in routines which provide data records for the Galil
' DL and QD commands which do not return a response. Other uses may be to send contour data.
' NOTE: This function can only send commands or groups of commands up to
' 1024 bytes long.

' hDmc             Handle to the Galil controller.
' CommandString    The command to send to the Galil controller.

Public Declare Function DMCGetUnsolicitedResponse Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Query the Galil controller for unsolicited responses. These are messages
' output from programs running in the background in the Galil controller.

' hDmc             Handle to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCWriteData Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long, BytesWritten As Long) As Long
' Low-level I/O routine to write data to the Galil controller. Data is written
' to the Galil controller only if it is "ready" to receive it. The function
' will attempt to write exactly cbBuffer characters to the controller.
' NOTE: For Win32 and WinRT driver the maximum number of bytes which can be written
' each time is 64. There are no restrictions with the Galil driver.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer to write the data from. Data does not need to be
'                  NULL terminated.
' BufferLength     Length of the data in the buffer.
' BytesWritten     Number of bytes written.

Public Declare Function DMCReadData Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long, BytesRead As Long) As Long
' Low-level I/O routine to read data from the Galil controller. The routine
' will read what ever is currently in the FIFO (bus controller) or
' communications port input queue (serial controller). The function will read
' up to cbBuffer characters from the controller. The data placed in the user
' buffer (pchBuffer) is NOT NULL terminated. The data returned is not guaranteed
' to be a complete response - you may have to call this function repeatedly to
' get a complete response.
' NOTE: For Win32 and WinRT driver the maximum number of bytes which can be read
' each time is 64. There are no restrictions with the Galil driver.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer to read the data into. Data will not be NULL
'                  terminated.
' BufferLength     Length of the buffer.
' BytesRead        Number of bytes read.

Public Declare Function DMCGetAdditionalResponseLen Lib "dmc32.dll" (ByVal hDmc As Long, ResponseLength As Long) As Long
' Query the Galil controller for the length of the additional response data. There will be more
' response data available if DMCCommand returned DMCERROR_BUFFERFULL.

' hDmc             Handle to the Galil controller.
' ResponseLength   Length of the additional response data.

Public Declare Function DMCGetAdditionalResponse Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Query the Galil controller for more response data. There will be more
' response data available if DMCCommand returned DMCERROR_BUFFERFULL.

' hDmc             Handle to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCError Lib "dmc32.dll" (ByVal hDmc As Long, ByVal ErrorCode As Long, ByVal Message As String, ByVal MessageLength As Long) As Long
' Retrieve the error message text from a DMCERROR_COMMAND error.

' hDmc             Handle to the Galil controller.
' ErrorCode        Error returned from API function.
' Message          Buffer to receive the error message text.
' MessageLength    Length of the buffer.

Public Declare Function DMCClear Lib "dmc32.dll" (ByVal hDmc As Long) As Long
' Clear the Galil controller FIFO.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCReset Lib "dmc32.dll" (ByVal hDmc As Long) As Long
' Reset the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCMasterReset Lib "dmc32.dll" (ByVal hDmc As Long) As Long
' Master reset the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCVersion Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Version As String, ByVal VersionLength As Long) As Long
' Get the version of the Galil controller.

' hDmc             Handle to the Galil controller.
' Version          Buffer to receive the version information.
' VersionLength    Length of the buffer.

Public Declare Function DMCDownloadFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal Label As String) As Long
' Download a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to download to the Galil controller.
' Label            Program label to download to. This argument is ignored if
'                  NULL.

Public Declare Function DMCDownloadFromBuffer Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal Label As String) As Long
' Download a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer of DMC commands to download to the Galil controller.
' Label            Program label to download to. This argument is ignored if
'                  NULL.

Public Declare Function DMCUploadFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Upload a file from the Galil controller.

' FileName         File name to upload from the Galil controller.

Public Declare Function DMCUploadToBuffer Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long) As Long
' Upload a file from the Galil controller.

' Buffer           Buffer of DMC commands to upload from the Galil controller.
' BufferLength     Length of the buffer.

Public Declare Function DMCSendFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Send a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to send to the Galil controller.

Public Declare Function DMCArrayDownload Lib "dmc32.dll" (ByVal hDmc As Long, ByVal ArrayName As String, ByVal FirstElement As Integer, ByVal LastElement As Integer, ByVal data As String, ByVal DataLength As Long, BytesWritten As Long) As Long
' Download an array to the Galil controller. The array must exist. Array data can be
' delimited by a comma or CR (0x0D) or CR/LF (0x0D0A).
' NOTE: The firmware on the controller must be recent enough to support the QD command.

' hDmc             Handle to the Galil controller.
' ArrayName        Array name to download to the Galil controller.
' FirstElement     First array element.
' LastElement      Last array element.
' Data             Buffer to write the array data from. Data does not need to be
'                  NULL terminated.
' DataLength       Length of the array data in the buffer.
' BytesWritten     Number of bytes written.

Public Declare Function DMCArrayUpload Lib "dmc32.dll" (ByVal hDmc As Long, ByVal ArrayName As String, ByVal FirstElement As Integer, ByVal LastElement As Integer, ByVal Buffer As String, ByVal BufferLength As Long, BytesRead As Long, ByVal Comma As Integer) As Long
' Upload an array from the Galil controller. The array must exist. Array data will be
' delimited by a comma or CR (0x0D) depending of the value of fComma.
' NOTE: The firmware on the controller must be recent enough to support the QU command.

' hDmc             Handle to the Galil controller.
' ArrayName        Array name to upload from the Galil controller.
' FirstElement     First array element.
' LastElement      Last array element.
' Buffer           Buffer to read the array data into. Array data will not be
'                  NULL terminated.
' BufferLength     Length of the buffer.
' BytesRead        Number of bytes read.

Public Declare Function DMCRefreshDataRecord Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Length As Long) As Long
' Refresh the data record used for fast polling.

' hDmc             Handle to the Galil controller.
' Length           Refresh size in bytes. Set to 0 unless you do not want a full-buffer
'                  refresh.

Public Declare Function DMCGetDataRecord Lib "dmc32.dll" (ByVal hDmc As Long, ByVal GeneralOffset As Integer, ByVal AxisInfoOffset As Integer, DataType As Integer, data As Long) As Long
' Get a data item from the data record used for fast polling. Gets one item from the
' data record by using offsets. To retrieve data record items by Id instead of offset,
' use the function DMCGetDataRecordByItemId.

' hDmc             Handle to the Galil controller.
' GeneralOffset    Data record offset for general data item.
' AxisInfoOffset   Additional data record offset for axis data item.
' DataType         Data type of the data item. If you are using the standard,
'                  pre-defined offsets, set this argument to zero before calling this
'                  function. The actual data type of the data item is returned on output.
' Data             Buffer to receive the data record data. Output only.

Public Declare Function DMCGetDataRecordByItemId Lib "dmc32.dll" (ByVal hDmc As Long, ByVal ItemId As Integer, ByVal AxisId As Integer, DataType As Integer, data As Long) As Long
' Get a data item from the data record used for fast polling. Gets one item from the
' data record by using Id. To retrieve data record items by offset instead of Id,
' use the function DMCGetDataRecord.

' hDmc             Handle to the Galil controller.
' ItemId           Data record item Id.
' AxisId           Axis Id used for axis data items.
' DataType         Data type of the data item. The data type of the
'                  data item is returned on output. Output Only.
' Data             Buffer to receive the data record data. Output only.

Public Declare Function DMCGetDataRecordRevision Lib "dmc32.dll" (ByVal hDmc As Long, Revision As Integer) As Long
' Get the revision of the data record structure used for fast polling.

' hDmc             Handle to the Galil controller.
' Revision         The revision of the data record structure is returned on
'                  output. Output Only.

Public Declare Function DMCDiagnosticsOn Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal AppendFile As Integer) As Long
' Turn on diagnostics.

' hDmc             Handle to the Galil controller.
' FileName         File name for the diagnostic file.
' AppendFile       True if the file will open for append, otherwise False.

Public Declare Function DMCDiagnosticsOff Lib "dmc32.dll" (ByVal hDmc As Long) As Long
' Turn off diagnostics.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCGetTimeout Lib "dmc32.dll" (ByVal hDmc As Long, Timeout As Long) As Long
' Get current timeout value.

' hDmc             Handle to the Galil controller.
' Timeout          Current timeout value in milliseconds.

Public Declare Function DMCSetTimeout Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Timeout As Long) As Long
' Set timeout value.

' hDmc             Handle to the Galil controller.
' Timeout          Timeout value in milliseconds.

Public Declare Function DMCGetDelay Lib "dmc32.dll" (ByVal hDmc As Long, Delay As Long) As Long
' Get current delay value.
' *** THIS FUNCTION IS OBSOLETE. DELAY IS NO LONGER USED ***

' hDmc             Handle to the Galil controller.
' Delay            Current delay value.

Public Declare Function DMCSetDelay Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Delay As Long) As Long
' Set delay value.
' *** THIS FUNCTION IS OBSOLETE. DELAY IS NO LONGER USED ***

' hDmc             Handle to the Galil controller.
' Delay            Delay value.

Public Declare Function DMCBinaryCommand Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Command As String, ByVal CommandLength As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Send a DMC command in binary format to the Galil controller.

' hDmc             Handle to the Galil controller.
' Command          The command to send to the Galil controller.
' CommandLength    The length of the command (binary commands are not null-terminated).
' Response         Buffer to receive the response data. If the buffer is too
'                  small to recieve all the response data from the controller,
'                  the error code DMCERROR_BUFFERFULL will be returned. The
'                  user may get additional response data by calling the
'                  function DMCGetAdditionalResponse. The length of the
'                  additonal response data may ascertained by call the
'                  function DMCGetAdditionalResponseLen. If the response
'                  data from the controller is too large for the internal
'                  additional response buffer, the error code
'                  DMCERROR_RESPONSEDATA will be returned. Output only.
' ResponseLength   Length of the buffer.

Public Declare Function DMCSendBinaryFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Send a file consisting of DMC commands in binary format to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to send to the Galil controller.

Public Declare Function DMCCommand_BinaryToAscii Lib "dmc32.dll" (ByVal hDmc As Long, ByVal BinCommand As String, ByVal BinCommandLength As Long, ByVal AscResult As String, ByVal AscResultLength As Long, AscResultReturnedLength As Long) As Long
' Convert a binary DMC command to an ascii DMC command.

' hDmc                     Handle to the Galil controller.
' BinCommand               Binary DMC command(s) to be converted.
' BinCommandLength         Length of DMC command(s).
' AscResult                Buffer to receive the translated DMC command.
' AscResultLength          Length of the buffer.
' AscResultReturnedLength  Length of the translated DMC command.

Public Declare Function DMCCommand_AsciiToBinary Lib "dmc32.dll" (ByVal hDmc As Long, ByVal AscCommand As String, ByVal AscCommandLength As Long, ByVal BinResult As String, ByVal BinaryResult As Long, BinResultReturnedLength As Long) As Long
' Convert an ascii DMC command to a binary DMC command.

' hDmc                     Handle to the Galil controller.
' AscCommand               Ascii DMC command(s) to be converted.
' AscCommandLength         Length of DMC command(s).
' BinResult                Buffer to receive the translated DMC command.
' BinResultLength          Length of the buffer.
' BinResultReturnedLength  Length of the translated DMC command.

Public Declare Function DMCFile_AsciiToBinary Lib "dmc32.dll" (ByVal hDmc As Long, ByVal InputFileName As String, ByVal OutputFileName As String) As Long
' Convert a file consisting of ascii commands to a file consisting of binary commands.

' hDmc              Handle to the Galil controller.
' InputFileName     File name for the input ascii file.
' OutputFileName    File name for the output binary file.

Public Declare Function DMCFile_BinaryToAscii Lib "dmc32.dll" (ByVal hDmc As Long, ByVal InputFileName As String, ByVal OutputFileName As String) As Long
' Convert a file consisting of binary commands to a file consisting of ascii commands.

' hDmc              Handle to the Galil controller.
' InputFileName     File name for the input binary file.
' OutputFileName    File name for the output ascii file.

Public Declare Function DMCReadSpecialConversionFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Read into memory a special BinaryToAscii/AsciiToBinary conversion table.

' hDmc              Handle to the Galil controller.
' FileName          File name for the special conversion file.

Public Declare Function DMCAddGalilRegistry Lib "dmc32.dll" (GALILREGISTRY As GALILREGISTRY, Controller As Integer) As Long
' Add a Galil controller to the Windows registry.

' galilregistry    Pointer to a GALILREGISTRY struct.
' Controller       Galil controller number assigned by the successful completion of this function.

Public Declare Function DMCModifyGalilRegistry Lib "dmc32.dll" (ByVal Controller As Integer, GALILREGISTRY As GALILREGISTRY) As Long
' Change a Galil controller in the Windows registry.

' Controller       Galil controller number.
' galilregistry    Pointer to a GALILREGISTRY struct.

Public Declare Function DMCDeleteGalilRegistry Lib "dmc32.dll" (ByVal Controller As Integer) As Long
' Delete a Galil controller in the Windows registry.

' Controller       Galil controller number. Use -1 to delete all Galil controllers.

Public Declare Function DMCGetGalilRegistryInfo Lib "dmc32.dll" (ByVal Controller As Integer, GALILREGISTRY As GALILREGISTRY) As Long
' Get Windows registry information for a given Galil controller.

' Controller       Galil controller number.
' galilregistry    Pointer to a GALILREGISTRY struct.

Public Declare Function DMCRegisterPnpControllers Lib "dmc32.dll" (count As Integer) As Long
' Update Windows registry for all Galil Plug-and-Play (PnP) controllers. This function
' will add new controllers to the registry or update existing ones.

' Count             Pointer to the number of Galil PnP controllers registered (and/or updated).

Public Declare Function DMCSelectController Lib "dmc32.dll" (ByVal hWnd As Long) As Integer
' Select a Galil motion controller from a list of registered controllers. Returns the
' selected controller number or -1 if no controller was selected.
' NOTE: This function invokes a dialog window.

' hwnd              The window handle of the calling application. If NULL, the
'                   window with the current input focus is used.

Public Declare Sub DMCEditRegistry Lib "dmc32.dll" (ByVal hWnd As Integer)
' Edit the Windows registry: add, change, or delete Galil motion controllers.
' NOTE: This function invokes a dialog window.

' hwnd              The window handle of the calling application. If NULL, the
'                   window with the current input focus is used.

Public Declare Function DMCWaitForMotionComplete Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Axes As String, ByVal DispatchMsgs As Integer) As Long
' Wait for motion complete by creating a thread to query the controller. The function returns
' when motion is complete.

' hDmc              Handle to the Galil controller.
' Axes              Which axes to wait for: X, Y, Z, W, E, F, G, H, or S for
'                   coordinated motion. To wait for more than one axis (other than
'                   coordinated motion), simply concatenate the axis letters in the string.
' DispatchMsgs      Set to TRUE if you want to get and dispatch Windows messages
'                   while waiting for motion complete. This flag is always TRUE for Win16.

Public Declare Function DMCDownloadFirmwareFile Lib "dmc32.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal DisplayDialog As Integer) As Long
' Update the controller's firmware. This function will open a binary firmware file and refresh
' the flash EEPROM of the controller.

' hDmc              Handle to the Galil controller.
' FileName          File name to download to the Galil controller.
' DisplayDialog     Display a progress dialog to the user.

Public Declare Function DMCReadRegister Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Offset As Integer, Status As Byte) As Long
' Read from a register (FIFO) of a bus controller.
' NOTE: This function is for Galil bus controllers and Win32 only.

' ** THIS FUNCTION IS FOR EXPERIENCED PROGRAMMERS ONLY **

' hDmc              Handle to the Galil controller.
' Offset            Register offset. 0 = mailbox, 1 = status.
' Status            Buffer to receive status register data.

Public Declare Function DMCWriteRegister Lib "dmc32.dll" (ByVal hDmc As Long, ByVal Offset As Integer, ByVal Status As Byte) As Long
' Write to a register (FIFO) of a bus controller.
' NOTE: This function is for Galil bus controllers and Win32 only.

' ** THIS FUNCTION IS FOR EXPERIENCED PROGRAMMERS ONLY **

' hDmc              Handle to the Galil controller.
' Offset            Register offset. 0 = mailbox, 1 = status.
' Status            Status register data.

#Else

Public Declare Function DMCOpen Lib "dmc16.dll" (ByVal Controller As Integer, ByVal hWnd As Integer, phDmc As Long) As Long
' Open communications with the Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' hWnd             The window handle to use for notifying the application
'                  program of an interrupt.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCOpen2 Lib "dmc16.dll" (ByVal Controller As Integer, ByVal ThreadID As Long, phDmc As Long) As Long
' Open communications with the Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' ThreadID         The thread id to use for notifying the application
'                  program of an interrupt.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCGetHandle Lib "dmc16.dll" (ByVal Controller As Integer, phDmc As Long) As Long
' Get the handle associated with a particular Galil controller.

' Controller       A number between 1 and 16. Up to 16 Galil controllers may be
'                  addressed per process.
' phDmc            Handle to the Galil controller to be use for all subsequent
'                  API calls.

Public Declare Function DMCClose Lib "dmc16.dll" (ByVal hDmc As Long) As Long
' Close communications with the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCCommand Lib "dmc16.dll" (ByVal hDmc As Long, ByVal CommandString As String, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Send a command to the Galil controller.
' NOTE: This function can only send commands or groups of commands up to
' 1024 bytes long.

' hDmc             Handle to the Galil controller.
' CommandString    The command to send to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCFastCommand Lib "dmc16.dll" (ByVal hDmc As Long, ByVal CommandString As String) As Long
' Send a command to the Galil controller without the overhead of waiting for a response. Use this function with
' caution as command errors will not be reported and the out-going FIFO or communciations buffer
' may fill up. This function is intended to be used in routines which provide data records for the Galil
' DL and QD commands which do not return a response. Other uses may be to send contour data.
' NOTE: This function can only send commands or groups of commands up to
' 1024 bytes long.

' hDmc             Handle to the Galil controller.
' CommandString    The command to send to the Galil controller.

Public Declare Function DMCGetUnsolicitedResponse Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Query the Galil controller for unsolicited responses. These are messages
' output from programs running in the background in the Galil controller.

' hDmc             Handle to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCWriteData Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long, BytesWritten As Long) As Long
' Low-level I/O routine to write data to the Galil controller. Data is written
' to the Galil controller only if it is "ready" to receive it. The function
' will attempt to write exactly cbBuffer characters to the controller.
' NOTE: For Win32 and WinRT driver the maximum number of bytes which can be written
' each time is 64. There are no restrictions with the Galil driver.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer to write the data from. Data does not need to be
'                  NULL terminated.
' BufferLength     Length of the data in the buffer.
' BytesWritten     Number of bytes written.

Public Declare Function DMCReadData Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long, BytesRead As Long) As Long
' Low-level I/O routine to read data from the Galil controller. The routine
' will read what ever is currently in the FIFO (bus controller) or
' communications port input queue (serial controller). The function will read
' up to cbBuffer characters from the controller. The data placed in the user
' buffer (pchBuffer) is NOT NULL terminated. The data returned is not guaranteed
' to be a complete response - you may have to call this function repeatedly to
' get a complete response.
' NOTE: For Win32 and WinRT driver the maximum number of bytes which can be read
' each time is 64. There are no restrictions with the Galil driver.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer to read the data into. Data will not be NULL
'                  terminated.
' BufferLength     Length of the buffer.
' BytesRead        Number of bytes read.

Public Declare Function DMCGetAdditionalResponseLen Lib "dmc16.dll" (ByVal hDmc As Long, ResponseLength As Long) As Long
' Query the Galil controller for the length of the additional response data. There will be more
' response data available if DMCCommand returned DMCERROR_BUFFERFULL.

' hDmc             Handle to the Galil controller.
' ResponseLength   Length of the additional response data.

Public Declare Function DMCGetAdditionalResponse Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Query the Galil controller for more response data. There will be more
' response data available if DMCCommand returned DMCERROR_BUFFERFULL.

' hDmc             Handle to the Galil controller.
' Response         Buffer to receive the response data.
' ResponseLength   Length of the buffer.

Public Declare Function DMCError Lib "dmc16.dll" (ByVal hDmc As Long, ByVal ErrorCode As Long, ByVal Message As String, ByVal MessageLength As Long) As Long
' Retrieve the error message text from a DMCERROR_COMMAND error.

' hDmc             Handle to the Galil controller.
' ErrorCode        Error returned from API function.
' Message          Buffer to receive the error message text.
' MessageLength    Length of the buffer.

Public Declare Function DMCClear Lib "dmc16.dll" (ByVal hDmc As Long) As Long
' Clear the Galil controller FIFO.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCReset Lib "dmc16.dll" (ByVal hDmc As Long) As Long
' Reset the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCMasterReset Lib "dmc16.dll" (ByVal hDmc As Long) As Long
' Master reset the Galil controller.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCVersion Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Version As String, ByVal VersionLength As Long) As Long
' Get the version of the Galil controller.

' hDmc             Handle to the Galil controller.
' Version          Buffer to receive the version information.
' VersionLength    Length of the buffer.

Public Declare Function DMCDownloadFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal Label As String) As Long
' Download a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to download to the Galil controller.
' Label            Program label to download to. This argument is ignored if
'                  NULL.

Public Declare Function DMCDownloadFromBuffer Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal Label As String) As Long
' Download a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' Buffer           Buffer of DMC commands to download to the Galil controller.
' Label            Program label to download to. This argument is ignored if
'                  NULL.

Public Declare Function DMCUploadFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Upload a file from the Galil controller.

' FileName         File name to upload from the Galil controller.

Public Declare Function DMCUploadToBuffer Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Buffer As String, ByVal BufferLength As Long) As Long
' Upload a file from the Galil controller.

' Buffer           Buffer of DMC commands to upload from the Galil controller.
' BufferLength     Length of the buffer.

Public Declare Function DMCSendFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Send a file to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to send to the Galil controller.

Public Declare Function DMCArrayDownload Lib "dmc16.dll" (ByVal hDmc As Long, ByVal ArrayName As String, ByVal FirstElement As Integer, ByVal LastElement As Integer, ByVal data As String, ByVal DataLength As Long, BytesWritten As Long) As Long
' Download an array to the Galil controller. The array must exist. Array data can be
' delimited by a comma or CR (0x0D) or CR/LF (0x0D0A).
' NOTE: The firmware on the controller must be recent enough to support the QD command.

' hDmc             Handle to the Galil controller.
' ArrayName        Array name to download to the Galil controller.
' FirstElement     First array element.
' LastElement      Last array element.
' Data             Buffer to write the array data from. Data does not need to be
'                  NULL terminated.
' DataLength       Length of the array data in the buffer.
' BytesWritten     Number of bytes written.

Public Declare Function DMCArrayUpload Lib "dmc16.dll" (ByVal hDmc As Long, ByVal ArrayName As String, ByVal FirstElement As Integer, ByVal LastElement As Integer, ByVal Buffer As String, ByVal BufferLength As Long, BytesRead As Long, ByVal Comma As Integer) As Long
' Upload an array from the Galil controller. The array must exist. Array data will be
' delimited by a comma or CR (0x0D) depending of the value of fComma.
' NOTE: The firmware on the controller must be recent enough to support the QU command.

' hDmc             Handle to the Galil controller.
' ArrayName        Array name to upload from the Galil controller.
' FirstElement     First array element.
' LastElement      Last array element.
' Buffer           Buffer to read the array data into. Array data will not be
'                  NULL terminated.
' BufferLength     Length of the buffer.
' BytesRead        Number of bytes read.

Public Declare Function DMCRefreshDataRecord Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Length As Long) As Long
' Refresh the data record used for fast polling.

' hDmc             Handle to the Galil controller.
' Length           Refresh size in bytes. Set to 0 unless you do not want a full-buffer
'                  refresh.

Public Declare Function DMCGetDataRecord Lib "dmc16.dll" (ByVal hDmc As Long, ByVal GeneralOffset As Integer, ByVal AxisInfoOffset As Integer, DataType As Integer, data As Long) As Long
' Get a data item from the data record used for fast polling.

' hDmc             Handle to the Galil controller.
' GeneralOffset    Data record offset for general data item.
' AxisInfoOffset   Additional data record offset for axis data item.
' DataType         Data type of the data item. If you are using the standard,
'                  pre-defined offsets, set this argument to zero before calling this
'                  function. The actual data type of the data item is returned on output.
' Data             Buffer to receive the data record data. Output only.

Public Declare Function DMCGetDataRecordByItemId Lib "dmc16.dll" (ByVal hDmc As Long, ByVal ItemId As Integer, ByVal AxisId As Integer, DataType As Integer, data As Long) As Long
' Get a data item from the data record used for fast polling. Gets one item from the
' data record by using Id. To retrieve data record items by offset instead of Id,
' use the function DMCGetDataRecord.

' hDmc             Handle to the Galil controller.
' ItemId           Data record item Id.
' AxisId           Axis Id used for axis data items.
' DataType         Data type of the data item. The data type of the
'                  data item is returned on output. Output Only.
' Data             Buffer to receive the data record data. Output only.

Public Declare Function DMCGetDataRecordRevision Lib "dmc16.dll" (ByVal hDmc As Long, Revision As Integer) As Long
' Get the revision of the data record structure used for fast polling.

' hDmc             Handle to the Galil controller.
' Revision         The revision of the data record structure is returned on
'                  output. Output Only.

Public Declare Function DMCDiagnosticsOn Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal AppendFile As Integer) As Long
' Turn on diagnostics.

' hDmc             Handle to the Galil controller.
' FileName         File name for the diagnostic file.
' AppendFile       True if the file will open for append, otherwise False.

Public Declare Function DMCDiagnosticsOff Lib "dmc16.dll" (ByVal hDmc As Long) As Long
' Turn off diagnostics.

' hDmc             Handle to the Galil controller.

Public Declare Function DMCGetTimeout Lib "dmc16.dll" (ByVal hDmc As Long, Timeout As Long) As Long
' Get current timeout value.

' hDmc             Handle to the Galil controller.
' Timeout          Current timeout value in milliseconds.

Public Declare Function DMCSetTimeout Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Timeout As Long) As Long
' Set timeout value.

' hDmc             Handle to the Galil controller.
' Timeout          Timeout value in milliseconds.

Public Declare Function DMCGetDelay Lib "dmc16.dll" (ByVal hDmc As Long, Delay As Long) As Long
' Get current delay value.
' *** THIS FUNCTION IS OBSOLETE. DELAY IS NO LONGER USED ***

' hDmc             Handle to the Galil controller.
' Delay            Current delay value.

Public Declare Function DMCSetDelay Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Delay As Long) As Long
' Set delay value.
' *** THIS FUNCTION IS OBSOLETE. DELAY IS NO LONGER USED ***

' hDmc             Handle to the Galil controller.
' Delay            Delay value.

Public Declare Function DMCBinaryCommand Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Command As String, ByVal CommandLength As Long, ByVal Response As String, ByVal ResponseLength As Long) As Long
' Send a DMC command in binary format to the Galil controller.

' hDmc             Handle to the Galil controller.
' Command          The command to send to the Galil controller.
' CommandLength    The length of the command (binary commands are not null-terminated).
' Response         Buffer to receive the response data. If the buffer is too
'                  small to recieve all the response data from the controller,
'                  the error code DMCERROR_BUFFERFULL will be returned. The
'                  user may get additional response data by calling the
'                  function DMCGetAdditionalResponse. The length of the
'                  additonal response data may ascertained by call the
'                  function DMCGetAdditionalResponseLen. If the response
'                  data from the controller is too large for the internal
'                  additional response buffer, the error code
'                  DMCERROR_RESPONSEDATA will be returned. Output only.
' ResponseLength   Length of the buffer.

Public Declare Function DMCSendBinaryFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Send a file consisting of DMC commands in binary format to the Galil controller.

' hDmc             Handle to the Galil controller.
' FileName         File name to send to the Galil controller.

Public Declare Function DMCCommand_BinaryToAscii Lib "dmc16.dll" (ByVal hDmc As Long, ByVal BinCommand As String, ByVal BinCommandLength As Long, ByVal AscResult As String, ByVal AscResultLength As Long, AscResultReturnedLength As Long) As Long
' Convert a binary DMC command to an ascii DMC command.

' hDmc                     Handle to the Galil controller.
' BinCommand               Binary DMC command(s) to be converted.
' BinCommandLength         Length of DMC command(s).
' AscResult                Buffer to receive the translated DMC command.
' AscResultLength          Length of the buffer.
' AscResultReturnedLength  Length of the translated DMC command.

Public Declare Function DMCCommand_AsciiToBinary Lib "dmc16.dll" (ByVal hDmc As Long, ByVal AscCommand As String, ByVal AscCommandLength As Long, ByVal BinResult As String, ByVal BinaryResult As Long, BinResultReturnedLength As Long) As Long
' Convert an ascii DMC command to a binary DMC command.

' hDmc                     Handle to the Galil controller.
' AscCommand               Ascii DMC command(s) to be converted.
' AscCommandLength         Length of DMC command(s).
' BinResult                Buffer to receive the translated DMC command.
' BinResultLength          Length of the buffer.
' BinResultReturnedLength  Length of the translated DMC command.

Public Declare Function DMCFile_AsciiToBinary Lib "dmc16.dll" (ByVal hDmc As Long, ByVal InputFileName As String, ByVal OutputFileName As String) As Long
' Convert a file consisting of ascii commands to a file consisting of binary commands.

' hDmc              Handle to the Galil controller.
' InputFileName     File name for the input ascii file.
' OutputFileName    File name for the output binary file.

Public Declare Function DMCFile_BinaryToAscii Lib "dmc16.dll" (ByVal hDmc As Long, ByVal InputFileName As String, ByVal OutputFileName As String) As Long
' Convert a file consisting of binary commands to a file consisting of ascii commands.

' hDmc              Handle to the Galil controller.
' InputFileName     File name for the input binary file.
' OutputFileName    File name for the output ascii file.

Public Declare Function DMCReadSpecialConversionFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String) As Long
' Read into memory a special BinaryToAscii/AsciiToBinary conversion table.

' hDmc              Handle to the Galil controller.
' FileName          File name for the special conversion file.

Public Declare Function DMCAddGalilRegistry Lib "dmc16.dll" (GALILREGISTRY As GALILREGISTRY, Controller As Integer) As Long
' Add a Galil controller to the Windows registry.

' galilregistry    Pointer to a GALILREGISTRY struct.
' Controller       Galil controller number assigned by the successful completion of this function.

Public Declare Function DMCModifyGalilRegistry Lib "dmc16.dll" (ByVal Controller As Integer, GALILREGISTRY As GALILREGISTRY) As Long
' Change a Galil controller in the Windows registry.

' Controller       Galil controller number.
' galilregistry    Pointer to a GALILREGISTRY struct.

Public Declare Function DMCDeleteGalilRegistry Lib "dmc16.dll" (ByVal Controller As Integer) As Long
' Delete a Galil controller in the Windows registry.

' Controller       Galil controller number. Use -1 to delete all Galil controllers.

Public Declare Function DMCGetGalilRegistryInfo Lib "dmc16.dll" (ByVal Controller As Integer, GALILREGISTRY As GALILREGISTRY) As Long
' Get Windows registry information for a given Galil controller.

' Controller       Galil controller number.
' galilregistry    Pointer to a GALILREGISTRY struct.

Public Declare Function DMCRegisterPnpControllers Lib "dmc16.dll" (count As Integer) As Long
' Update Windows registry for all Galil Plug-and-Play (PnP) controllers. This function
' will add new controllers to the registry or update existing ones.

' Count             Pointer to the number of Galil PnP controllers registered (and/or updated).

Public Declare Function DMCSelectController Lib "dmc16.dll" (ByVal hWnd As Integer) As Integer
' Select a Galil motion controller from a list of registered controllers. Returns the
' selected controller number or -1 if no controller was selected.
' NOTE: This function invokes a dialog window.

' hwnd              The window handle of the calling application. If NULL, the
'                   window with the current input focus is used.

Public Declare Sub DMCEditRegistry Lib "dmc16.dll" (ByVal hWnd As Integer)
' Edit the Windows registry: add, change, or delete Galil motion controllers.
' NOTE: This function invokes a dialog window.

' hwnd              The window handle of the calling application. If NULL, the
'                   window with the current input focus is used.

Public Declare Function DMCWaitForMotionComplete Lib "dmc16.dll" (ByVal hDmc As Long, ByVal Axes As String, ByVal DispatchMsgs As Integer) As Long
' Wait for motion complete by creating a thread to query the controller. The function returns
' when motion is complete.

' hDmc              Handle to the Galil controller.
' Axes              Which axes to wait for: X, Y, Z, W, E, F, G, H, or S for
'                   coordinated motion. To wait for more than one axis (other than
'                   coordinated motion), simply concatenate the axis letters in the string.
' DispatchMsgs      Set to TRUE if you want to get and dispatch Windows messages
'                   while waiting for motion complete. This flag is always TRUE for Win16.

Public Declare Function DMCDownloadFirmwareFile Lib "dmc16.dll" (ByVal hDmc As Long, ByVal FileName As String, ByVal DisplayDialog As Integer) As Long
' Update the controller's firmware. This function will open a binary firmware file and refresh
' the flash EEPROM of the controller.

' hDmc              Handle to the Galil controller.
' FileName          File name to download to the Galil controller.
' DisplayDialog     Display a progress dialog to the user.
#End If
