Attribute VB_Name = "moduleGlobal"
Option Explicit

Global Const gcCaptionString = "Ingersoll Rand Test Tool Version "
Global Const gcVersionString = "1.2d"

'Enum UartProtocolBytes
'  Start
'  Length
'  MsgType
'  Payload
'  Checksum
'  EndOfTransmission
'End Enum

Global Const gcStartIndex = 0
Global Const gcLengthIndex = 1
Global Const gcMsgTypeIndex = 2
Global Const gcMinMsgLength = 5
Global Const gcNumberOfMessageTypes = 21
Global Const gcMessageOverhead = 5

Enum eMessageType
    InvalidMsg = 0
    QueryVersionMsg
    StaticTestModeMsg
    SetRadioRegisterMsg
    QueryRadioRegisterMsg
    SetMicroRegisterMsg
    QueryMicroRegisterMsg
    ResumeScan
    Reserved2Msg
    QueryStatisticsMsg
    ClearStatisticsMsg
    TransmitRfMsg
    ReceiveRfMsg ' from module only
    TransmitPERTMsg
    TransmitPERTDoneMsg ' from module only
    ReceivePERTMsg
    ReceivePERTDoneMsg ' from module only
    EDScanRequestMsg
    EDScanResultMsg ' from module only
    SaveToNvalMsg
    SetSleepModeMsg
    TOTAL_MESSAGE_TYPES
End Enum

Global Const gcMessageTypesBigString = "" & _
  "InvalidMsg," & _
  "QueryVersionMsg," & _
  "StaticTestModeMsg," & _
  "SetRadioRegisterMsg," & _
  "QueryRadioRegisterMsg," & _
  "SetMicroRegisterMsg," & _
  "QueryMicroRegisterMsg," & _
  "ResumeScanMsg," & _
  "Reserved2Msg," & _
  "QueryStatisticsMsg," & _
  "ClearStatisticsMsg," & _
  "TransmitRfMsg," & _
  "ReceiveRfMsg," & _
  "TransmitPERTMsg," & _
  "TransmitPERTDoneMsg," & _
  "ReceivePERTMsg," & _
  "ReceivePERTDoneMsg," & _
  "EDScanRequestMsg," & _
  "EDScanResultMsg," & _
  "SaveToNvalMsg," & _
  "SetSleepModeMsg"
  
Global gMsgTypeStrings() As String

Global Const gcValidHexadecimalCharacters = "0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F"
Global gValidHexChars() As String

Global Const gcMinimumTxPayloadSize = "0,0,3,2,2,2,2,0,0,0,0,6,8,8,2,9,6,1,10,0,1,0,0,0"
Global gMinTxPayloadSize() As String

Enum RxStates
  Idle
  PreprocessMessage
  ValidateMessage
  ProcessMessage
  InvalidateMessage
End Enum
  
Global geRxState As RxStates
Global geNextRxState As RxStates

' pieces of a receive the message
Global gRxBuffer As String
Global gRxMessage As String
Global gStartPosition As Integer
Global gRxEndOfFramePosition As Integer
Global gRxMessageLength As Integer
Global gRxMessageType As Integer
Global gRxMessagePayload As String
Global gRxMessageChecksum As Integer
Global gRxBufferAsBytes() As Byte
Global gRxMessageAsBytes() As Byte

' transmit message
Global Const gcTxTxtBoxes = 42
Global gTxMessage As String
Global gTxMessageType As Integer
Global gBytes() As Byte
Global gTxMessageSelected As Integer
Global gTxNumberOfBytes As Integer

Global gCommPort As Integer
Global Const gcNumberOfComPorts = 20

' keep all messages the same length
' so that things line up in the output window
Global Const gcRxInvalidMsg = "Rx Invalid"
Global Const gcRxValidMsg = "Rx Valid  "
Global Const gcTransmitMsg = "Transmit  "
Global Const gcTransmitRaw = "Tx Raw    "

Global Const gcPcbVersionString = "PCB Version: "
Global Const gcIcPartNumberString = "RF IC Part Number: "
Global Const gcVersionNumberString = "RF IC Version Number: "
Global Const gcAppFirmwareVersionString = "Firmware Version: "
Global Const gcAppFirmwareDateString = "Firmware Date: "
Global Const gcAppFirmwareDescription = "Firmware Description: "
  
Global Const gcTransmittedPacketsString = "Transmit Packet Count: "
Global Const gcReceivedPacketsString = "Receive Packet Count: "
Global Const gcTxNotAcknowledgedString = "Transmit Packets Not Acknowledged: "
Global Const gcTransmitFailuresString = "Transmit Failure Count: "
Global Const gcSpareStatistic1String = "Spare Statistic1: "
Global Const gcSpareStatistic2String = "Spare Statistic2: "

Global Const gcReceiveMsgValidityString = "Receive Message is "
Global Const gcReceivedMessageTypeString = "Received Message Type is "
Global Const gcNumberOfReceiveTxtBoxes = 42

' PERT
Global gRxMsgTypeForPert As eMessageType
Global gPertPimId As Long
Global gPertSourceAddr As Long
Global gPertDestAddr As Long
Global gPertNumberOfPackets As Long
Global gPertPacketSize As Byte
Global gPertRetries As Byte

Global Const gcPertPimId = &HFACE
Global Const gcPertSourceAddr = &H1
Global Const gcPertDestAddr = &H1
Global Const gcPertNumberOfPackets = 100
Global Const gcPertPacketSize = 5
Global Const gcMaxPacketSize = 100

Enum PertStates
  PertIdle = 0
  TxRetryMode
  WaitForRetryModeResponse
  TxAckMode
  WaitForAckModeResponse
  TxPertTxCommand
  WaitForPertTxResponse
  WaitForPertDoneResponse
  TxPertRxCommand
  WaitForPertRxResponse
  WaitForPertRxDoneResponse
End Enum
 
Global gPertState As PertStates

Global Const gcPacketsNotAckedString = "Packets Not Acked = 0x"
Global Const gcPertReceivedPacketsString = "Received Packets = 0x"
Global Const gcPertExpectedPacketsString = "Expected Packets = 0x"
Global Const gcPertMatchingPacketsString = "Matching Packets = 0x"
  
' microcontroller registers/tokens

'Global Const gcRegisterNamesBigString = "NVAL Version,Operating Mode,Device Type,Own PIM ID,IEEE Address,Own Source Address,Default Destination Address,RF Power,Channel,Channel Set,DCS Enable,Encryption Mode,Encryption Key,Reserved,Reserved,Reserved,Rx Filter Mode,RF Request Ack Mode,RF Transmit Ack Mode,Retry Attempts,Wake Up Mode"
'Global gRegisterNameStrings() As String

Global Const gcIeeeAddressSize = 8
Global Const gcEncryptionKeySize = 16
Global Const gcIvSize = 8
Global Const gc32Size = 4

Type eTokenType
  NvalVersion As Byte
  OpMode As Byte
  DeviceType As Byte
  PimPanId As Long
  IeeeAddr(0 To gcIeeeAddressSize - 1) As Byte
  SourceAddr As Long
  DestAddr As Long
  RfPower As Byte
  Channel As Byte
  ChannelSet As Byte
  Reserved0 As Byte
  EncryptionMode As Byte
  EncryptionKey(0 To gcEncryptionKeySize - 1) As Byte
  Reserved1 As Byte
  Reserved2 As Byte
  Reserved3 As Byte
  RxFilterMode As Byte
  RfReqAckMode As Byte
  RfTxAckMode As Byte
  RetryAttempts As Byte
  WakeUpMode As Byte
  WakeUpMsgSize As Byte
  EncryptionIv(0 To gcIvSize - 1) As Byte
  HeaderSize As Byte
  SequenceNumber(0 To gc32Size - 1) As Byte
  StartFrameDelimiter As Byte
  NvalWriteCount(0 To gc32Size - 1) As Byte
  Reserved4 As Byte
  TokenSum As Byte
End Type
    
Global gToken As eTokenType

Enum eRegisterNumber
  NvalVersion = 0
  OpMode
  DeviceType
  PimPanId
  IeeeAddr
  SourceAddr
  DestAddr
  RfPower
  Channel
  ChannelSet
  Reserved0
  EncryptionMode
  EncryptionKey
  Reserved1
  Reserved2
  Reserved3
  RxFilterMode
  RfReqAckMode
  RfTxAckMode
  RetryAttempts
  WakeUpMode
  WakeUpMsgSize
  EncryptionIv
  HeaderSize
  SequenceNumber
  StartFrameDelimiter
  NvalWriteCount
  Reserved4
  TokenSum
End Enum

' register read state machine
Enum eReadRegisterStates
  ReadRegisterIdle
  ReadRegisterStart
  ReadNextRegister
  ReadRegisterWaitState
End Enum
  
Global gReadRegisterState As eReadRegisterStates
Global gRxMsgTypeForReadRegisters As eMessageType

Global gTestMode As Byte
Global gTestModePower As Byte
Global gTestModeChannel As Byte
Global gModuleMode As Byte
Global gTxModPattern As Byte
   
Global Const LB_SETHORIZONTALEXTENT = &H194


Enum SyncTestStates
  SyncTestIdleState
  SyncTestTxState
  SyncTestWaitState
End Enum

Global gSyncTestEnable As Boolean
Global gRxMsgTypeForSyncTest As eMessageType
Global gSyncTestState As SyncTestStates
  
  
Global Const gcStatusCodeString = "Status Code: 0x"
Global Const gcStatusCodeDescriptionString = "Status Code Description: "
Global Const gcStatusCodeFailCountString = "Status Code != Success Count: 0x"

'
' this is the retval_t type from the Atmel code
' many are never returned to the host, but if we add some we don't want them
' to be the same as values that are already used
'
Enum eStatusCodes
  Success = &H0
  InvalidPayloadLength = &H70
  BeaconTxSSuccess = &H80
  TrxAsleep = &H81
  TrxAwake = &H82
  CrcCorrect = &H83
  CrcIncorrect = &H84
  Failure = &H85
  Busy = &H86
  TalFramePending = &H87
  AlreadyRuning = &H88
  NotRunning = &H89
  InvalidId = &H8A
  InvalidTimeout = &H8B
  InvalidParameter = &HE8
  QueueFull = &H8C
  CsmaCaInProgress = &H8D
  NoFrameTransmission = &H8E
  ChannelAccessFailure = &HE1
  NoAck = &HE9
  UnsupportedAttribute = &HF4
End Enum

Global gStatusCode As eStatusCodes
Global gStatusCodeFailCount As Long

Global Const gcNumOfStatistics = 8

Global Const gcStatisticStrings = "" & _
  "Transmit Packet Count: 0x," & _
  "Receive Packet Count: 0x," & _
  "Transmit Packets Not Acknowledged: 0x," & _
  "Transmit Failure Count: 0x," & _
  "Invalid MIC Count: 0x," & _
  "Transmit to Host Lost Byte Count: 0x," & _
  "Buffer Not Free Count: 0x," & _
  "Queue Full Count: 0x"
  
Global gStatisticStrings() As String

Global gFnumLog As Integer
  
