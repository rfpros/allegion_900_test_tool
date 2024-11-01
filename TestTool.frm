VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form TestTool 
   Caption         =   "Ingersoll Rand Test Tool"
   ClientHeight    =   9732
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   16644
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9732
   ScaleWidth      =   16644
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab2 
      Height          =   9732
      Left            =   8640
      TabIndex        =   0
      Top             =   0
      Width           =   8652
      _ExtentX        =   15261
      _ExtentY        =   17166
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "TestTool.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListBoxClear"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAppendLog"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Comm Port Selection"
      TabPicture(1)   =   "TestTool.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCommPortRate"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSelectNewCommPort"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRescanCommPorts"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdGetBaudRate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Raw Receive"
      TabPicture(2)   =   "TestTool.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame26"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Blank"
      TabPicture(3)   =   "TestTool.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Blank"
      TabPicture(4)   =   "TestTool.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "QueryTimer"
      Tab(4).Control(1)=   "tmrMsgTimeout"
      Tab(4).Control(2)=   "RxFsmTimer"
      Tab(4).Control(3)=   "Frame24"
      Tab(4).Control(4)=   "MSComm1"
      Tab(4).ControlCount=   5
      Begin VB.CheckBox chkAppendLog 
         Caption         =   "Append Log File"
         Height          =   372
         Left            =   -72480
         TabIndex        =   441
         Top             =   9000
         Width           =   2292
      End
      Begin VB.Timer QueryTimer 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   -73920
         Top             =   6840
      End
      Begin VB.Frame Frame26 
         Height          =   1212
         Left            =   -74040
         TabIndex        =   424
         Top             =   5400
         Width           =   6372
         Begin VB.Label lblReceivedMessageType 
            Caption         =   "Received Message type is"
            Height          =   252
            Left            =   240
            TabIndex        =   426
            Top             =   720
            Width           =   5892
         End
         Begin VB.Label lblReceiveMsgValidity 
            Caption         =   "Received Message is"
            Height          =   252
            Left            =   240
            TabIndex        =   425
            Top             =   360
            Width           =   5892
         End
      End
      Begin VB.Timer tmrMsgTimeout 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   -74520
         Top             =   6840
      End
      Begin VB.Timer RxFsmTimer 
         Interval        =   10
         Left            =   -74520
         Top             =   7320
      End
      Begin VB.Frame Frame24 
         Caption         =   "Energy Detect"
         Height          =   5532
         Left            =   -74520
         TabIndex        =   345
         Top             =   1080
         Visible         =   0   'False
         Width           =   7572
         Begin MSChart20Lib.MSChart EnergyDetectChart 
            Height          =   4572
            Left            =   480
            OleObjectBlob   =   "TestTool.frx":008C
            TabIndex        =   346
            Top             =   480
            Width           =   6972
         End
      End
      Begin VB.CommandButton cmdGetBaudRate 
         Caption         =   "Get Baud Rate"
         Height          =   372
         Left            =   2400
         TabIndex        =   71
         Top             =   5400
         Width           =   3852
      End
      Begin VB.Frame Frame19 
         Caption         =   "Comm Ports"
         Height          =   2892
         Left            =   1320
         TabIndex        =   50
         Top             =   840
         Width           =   6012
         Begin VB.OptionButton Option4 
            Caption         =   "COM1"
            Height          =   372
            Index           =   0
            Left            =   840
            TabIndex        =   70
            Top             =   600
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM2"
            Height          =   372
            Index           =   1
            Left            =   840
            TabIndex        =   69
            Top             =   960
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM3"
            Height          =   372
            Index           =   2
            Left            =   840
            TabIndex        =   68
            Top             =   1320
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM4"
            Height          =   372
            Index           =   3
            Left            =   840
            TabIndex        =   67
            Top             =   1680
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM5"
            Height          =   372
            Index           =   4
            Left            =   840
            TabIndex        =   66
            Top             =   2040
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM6"
            Height          =   372
            Index           =   5
            Left            =   2040
            TabIndex        =   65
            Top             =   600
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM7"
            Height          =   372
            Index           =   6
            Left            =   2040
            TabIndex        =   64
            Top             =   960
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM8"
            Height          =   372
            Index           =   7
            Left            =   2040
            TabIndex        =   63
            Top             =   1320
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM9"
            Height          =   372
            Index           =   8
            Left            =   2040
            TabIndex        =   62
            Top             =   1680
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM10"
            Height          =   372
            Index           =   9
            Left            =   2040
            TabIndex        =   61
            Top             =   2040
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM20"
            Height          =   372
            Index           =   19
            Left            =   4080
            TabIndex        =   60
            Top             =   2040
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM19"
            Height          =   372
            Index           =   18
            Left            =   4080
            TabIndex        =   59
            Top             =   1680
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM18"
            Height          =   372
            Index           =   17
            Left            =   4080
            TabIndex        =   58
            Top             =   1320
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM17"
            Height          =   372
            Index           =   16
            Left            =   4080
            TabIndex        =   57
            Top             =   960
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM16"
            Height          =   372
            Index           =   15
            Left            =   4080
            TabIndex        =   56
            Top             =   600
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM15"
            Height          =   372
            Index           =   14
            Left            =   3000
            TabIndex        =   55
            Top             =   2040
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM14"
            Height          =   372
            Index           =   13
            Left            =   3000
            TabIndex        =   54
            Top             =   1680
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM13"
            Height          =   372
            Index           =   12
            Left            =   3000
            TabIndex        =   53
            Top             =   1320
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM12"
            Height          =   372
            Index           =   11
            Left            =   3000
            TabIndex        =   52
            Top             =   960
            Width           =   972
         End
         Begin VB.OptionButton Option4 
            Caption         =   "COM11"
            Height          =   372
            Index           =   10
            Left            =   3000
            TabIndex        =   51
            Top             =   600
            Width           =   972
         End
      End
      Begin VB.CommandButton cmdRescanCommPorts 
         Caption         =   "Rescan Comm Ports"
         Height          =   372
         Left            =   2400
         TabIndex        =   49
         Top             =   4680
         Width           =   3852
      End
      Begin VB.Frame Frame8 
         Caption         =   "Serial Data"
         Height          =   8052
         Left            =   -74760
         TabIndex        =   47
         Top             =   720
         Width           =   7932
         Begin VB.ListBox ListBox 
            Height          =   7008
            ItemData        =   "TestTool.frx":1A1D
            Left            =   360
            List            =   "TestTool.frx":1A24
            TabIndex        =   48
            Top             =   600
            Width           =   7212
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Receive Data"
         Height          =   4452
         Left            =   -74040
         TabIndex        =   4
         Top             =   720
         Width           =   6372
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   1
            Left            =   1200
            TabIndex        =   45
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   2
            Left            =   2040
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   3
            Left            =   2880
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   4
            Left            =   3720
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   5
            Left            =   4560
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   6
            Left            =   5400
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   8
            Left            =   1200
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   9
            Left            =   2040
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   10
            Left            =   2880
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   11
            Left            =   3720
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   12
            Left            =   4560
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   13
            Left            =   5400
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   14
            Left            =   360
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   15
            Left            =   1200
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   16
            Left            =   2040
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   17
            Left            =   2880
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   18
            Left            =   3720
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   19
            Left            =   4560
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   20
            Left            =   5400
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   1800
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   21
            Left            =   360
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   22
            Left            =   1200
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   23
            Left            =   2040
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   24
            Left            =   2880
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   25
            Left            =   3720
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   26
            Left            =   4560
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   27
            Left            =   5400
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   28
            Left            =   360
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   29
            Left            =   1200
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   30
            Left            =   2040
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   31
            Left            =   2880
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   32
            Left            =   3720
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   33
            Left            =   4560
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   34
            Left            =   5400
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   3000
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   35
            Left            =   360
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   36
            Left            =   1200
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   37
            Left            =   2040
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   38
            Left            =   2880
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   39
            Left            =   3720
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   40
            Left            =   4560
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
         Begin VB.TextBox txtRxByte 
            Height          =   372
            Index           =   41
            Left            =   5400
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   3600
            Width           =   700
         End
      End
      Begin VB.CommandButton ListBoxClear 
         Caption         =   "Clear List Box"
         Height          =   612
         Left            =   -74760
         TabIndex        =   2
         Top             =   8880
         Width           =   2052
      End
      Begin VB.CommandButton cmdSelectNewCommPort 
         Caption         =   "Select New Comm Port"
         Height          =   372
         Left            =   2400
         TabIndex        =   1
         Top             =   3960
         Width           =   3852
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   -74520
         Top             =   7800
         _ExtentX        =   804
         _ExtentY        =   804
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Label lblCommPortRate 
         Caption         =   "Comm Port Rate: "
         Height          =   252
         Left            =   2400
         TabIndex        =   3
         ToolTipText     =   "Some ports may not support this baud rate."
         Top             =   6120
         Width           =   3852
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9732
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   8652
      _ExtentX        =   15261
      _ExtentY        =   17166
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Test Mode Controls"
      TabPicture(0)   =   "TestTool.frx":1A31
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSetModuleMode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSetChannel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSetPower"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSendTestMode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame21"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame25"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdStartStopSyncTest"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdRepeatedQuery"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame27"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Queries"
      TabPicture(1)   =   "TestTool.frx":1A4D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClearStatistics"
      Tab(1).Control(1)=   "cmdQueryStatistics"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "cmdClearVersion"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "cmdQueryVersion"
      Tab(1).Control(6)=   "Frame20"
      Tab(1).Control(7)=   "cmdClearStatusCodeFailCount"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Raw Tx"
      TabPicture(2)   =   "TestTool.frx":1A69
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame23"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "frmMessageTypes"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "PERT"
      TabPicture(3)   =   "TestTool.frx":1A85
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame18"
      Tab(3).Control(1)=   "cmdResetPert"
      Tab(3).Control(2)=   "cmdRxPert"
      Tab(3).Control(3)=   "cmdTxPert"
      Tab(3).Control(4)=   "Frame15"
      Tab(3).Control(5)=   "Frame17"
      Tab(3).Control(6)=   "Frame13"
      Tab(3).Control(7)=   "Frame12"
      Tab(3).Control(8)=   "Frame11"
      Tab(3).Control(9)=   "Frame10"
      Tab(3).Control(10)=   "Frame9"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Register Tab 1"
      TabPicture(4)   =   "TestTool.frx":1AA1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblMicroReg(16)"
      Tab(4).Control(1)=   "lblMicroReg(17)"
      Tab(4).Control(2)=   "lblMicroReg(18)"
      Tab(4).Control(3)=   "lblMicroReg(19)"
      Tab(4).Control(4)=   "lblMicroReg(20)"
      Tab(4).Control(5)=   "lblMicroReg(5)"
      Tab(4).Control(6)=   "lblMicroReg(6)"
      Tab(4).Control(7)=   "lblMicroReg(7)"
      Tab(4).Control(8)=   "lblMicroReg(8)"
      Tab(4).Control(9)=   "lblMicroReg(9)"
      Tab(4).Control(10)=   "lblMicroReg(3)"
      Tab(4).Control(11)=   "lblMicroReg(2)"
      Tab(4).Control(12)=   "lblMicroReg(1)"
      Tab(4).Control(13)=   "lblMicroReg(21)"
      Tab(4).Control(14)=   "Frame22"
      Tab(4).Control(15)=   "cmdReadRegisters"
      Tab(4).Control(16)=   "txtMicroReg(16)"
      Tab(4).Control(17)=   "txtMicroReg(17)"
      Tab(4).Control(18)=   "txtMicroReg(18)"
      Tab(4).Control(19)=   "txtMicroReg(19)"
      Tab(4).Control(20)=   "txtMicroReg(20)"
      Tab(4).Control(21)=   "cmdSetMicroReg(16)"
      Tab(4).Control(22)=   "cmdSetMicroReg(17)"
      Tab(4).Control(23)=   "cmdSetMicroReg(18)"
      Tab(4).Control(24)=   "cmdSetMicroReg(19)"
      Tab(4).Control(25)=   "cmdSetMicroReg(20)"
      Tab(4).Control(26)=   "cmdGetMicroRegister(16)"
      Tab(4).Control(27)=   "cmdGetMicroRegister(17)"
      Tab(4).Control(28)=   "cmdGetMicroRegister(18)"
      Tab(4).Control(29)=   "cmdGetMicroRegister(19)"
      Tab(4).Control(30)=   "cmdGetMicroRegister(20)"
      Tab(4).Control(31)=   "txtMicroReg(5)"
      Tab(4).Control(32)=   "txtMicroReg(6)"
      Tab(4).Control(33)=   "txtMicroReg(7)"
      Tab(4).Control(34)=   "txtMicroReg(8)"
      Tab(4).Control(35)=   "txtMicroReg(9)"
      Tab(4).Control(36)=   "cmdSetMicroReg(5)"
      Tab(4).Control(37)=   "cmdSetMicroReg(6)"
      Tab(4).Control(38)=   "cmdSetMicroReg(7)"
      Tab(4).Control(39)=   "cmdSetMicroReg(8)"
      Tab(4).Control(40)=   "cmdSetMicroReg(9)"
      Tab(4).Control(41)=   "cmdGetMicroRegister(5)"
      Tab(4).Control(42)=   "cmdGetMicroRegister(6)"
      Tab(4).Control(43)=   "cmdGetMicroRegister(7)"
      Tab(4).Control(44)=   "cmdGetMicroRegister(8)"
      Tab(4).Control(45)=   "cmdGetMicroRegister(9)"
      Tab(4).Control(46)=   "cmdGetMicroRegister(3)"
      Tab(4).Control(47)=   "cmdGetMicroRegister(2)"
      Tab(4).Control(48)=   "cmdGetMicroRegister(1)"
      Tab(4).Control(49)=   "cmdSetMicroReg(3)"
      Tab(4).Control(50)=   "cmdSetMicroReg(2)"
      Tab(4).Control(51)=   "cmdSetMicroReg(1)"
      Tab(4).Control(52)=   "txtMicroReg(3)"
      Tab(4).Control(53)=   "txtMicroReg(2)"
      Tab(4).Control(54)=   "txtMicroReg(1)"
      Tab(4).Control(55)=   "txtMicroReg(21)"
      Tab(4).Control(56)=   "cmdSetMicroReg(21)"
      Tab(4).Control(57)=   "cmdGetMicroRegister(21)"
      Tab(4).ControlCount=   58
      TabCaption(5)   =   "Register Tab 2"
      TabPicture(5)   =   "TestTool.frx":1ABD
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblMicroReg(28)"
      Tab(5).Control(1)=   "lblMicroReg(26)"
      Tab(5).Control(2)=   "lblMicroReg(25)"
      Tab(5).Control(3)=   "lblMicroReg(24)"
      Tab(5).Control(4)=   "lblMicroReg(23)"
      Tab(5).Control(5)=   "lblMicroReg(22)"
      Tab(5).Control(6)=   "lblMicroReg(4)"
      Tab(5).Control(7)=   "lblMicroReg(12)"
      Tab(5).Control(8)=   "lblMicroReg(11)"
      Tab(5).Control(9)=   "lblMicroReg(27)"
      Tab(5).Control(10)=   "lblMicroReg(10)"
      Tab(5).Control(11)=   "lblMicroReg(15)"
      Tab(5).Control(12)=   "lblMicroReg(14)"
      Tab(5).Control(13)=   "lblMicroReg(13)"
      Tab(5).Control(14)=   "lblMicroReg(0)"
      Tab(5).Control(15)=   "txtMicroReg(26)"
      Tab(5).Control(16)=   "txtMicroReg(25)"
      Tab(5).Control(17)=   "txtMicroReg(24)"
      Tab(5).Control(18)=   "txtMicroReg(23)"
      Tab(5).Control(19)=   "txtMicroReg(22)"
      Tab(5).Control(20)=   "cmdGetMicroRegister(26)"
      Tab(5).Control(21)=   "cmdGetMicroRegister(25)"
      Tab(5).Control(22)=   "cmdGetMicroRegister(24)"
      Tab(5).Control(23)=   "cmdGetMicroRegister(23)"
      Tab(5).Control(24)=   "cmdGetMicroRegister(22)"
      Tab(5).Control(25)=   "cmdSetMicroReg(26)"
      Tab(5).Control(26)=   "cmdSetMicroReg(25)"
      Tab(5).Control(27)=   "cmdSetMicroReg(24)"
      Tab(5).Control(28)=   "cmdSetMicroReg(23)"
      Tab(5).Control(29)=   "cmdSetMicroReg(22)"
      Tab(5).Control(30)=   "txtMicroReg(28)"
      Tab(5).Control(31)=   "cmdSetMicroReg(28)"
      Tab(5).Control(32)=   "cmdGetMicroRegister(28)"
      Tab(5).Control(33)=   "cmdGetMicroRegister(4)"
      Tab(5).Control(34)=   "cmdSetMicroReg(4)"
      Tab(5).Control(35)=   "txtMicroReg(4)"
      Tab(5).Control(36)=   "cmdGetMicroRegister(12)"
      Tab(5).Control(37)=   "cmdGetMicroRegister(11)"
      Tab(5).Control(38)=   "cmdSetMicroReg(12)"
      Tab(5).Control(39)=   "cmdSetMicroReg(11)"
      Tab(5).Control(40)=   "txtMicroReg(12)"
      Tab(5).Control(41)=   "txtMicroReg(11)"
      Tab(5).Control(42)=   "txtMicroReg(27)"
      Tab(5).Control(43)=   "cmdSetMicroReg(27)"
      Tab(5).Control(44)=   "cmdGetMicroRegister(27)"
      Tab(5).Control(45)=   "cmdGetMicroRegister(10)"
      Tab(5).Control(46)=   "cmdSetMicroReg(10)"
      Tab(5).Control(47)=   "txtMicroReg(10)"
      Tab(5).Control(48)=   "cmdGetMicroRegister(15)"
      Tab(5).Control(49)=   "cmdGetMicroRegister(14)"
      Tab(5).Control(50)=   "cmdGetMicroRegister(13)"
      Tab(5).Control(51)=   "cmdSetMicroReg(15)"
      Tab(5).Control(52)=   "cmdSetMicroReg(14)"
      Tab(5).Control(53)=   "cmdSetMicroReg(13)"
      Tab(5).Control(54)=   "txtMicroReg(15)"
      Tab(5).Control(55)=   "txtMicroReg(14)"
      Tab(5).Control(56)=   "txtMicroReg(13)"
      Tab(5).Control(57)=   "txtMicroReg(0)"
      Tab(5).Control(58)=   "cmdSetMicroReg(0)"
      Tab(5).Control(59)=   "cmdGetMicroRegister(0)"
      Tab(5).ControlCount=   60
      Begin VB.Frame Frame27 
         Caption         =   "Tx Mod Pattern"
         Height          =   2292
         Left            =   5280
         TabIndex        =   442
         Top             =   6840
         Width           =   3012
         Begin VB.OptionButton optTxModPattern 
            Caption         =   "0xAA"
            Height          =   372
            Index           =   4
            Left            =   120
            TabIndex        =   447
            Top             =   1800
            Width           =   2772
         End
         Begin VB.OptionButton optTxModPattern 
            Caption         =   "0xFF"
            Height          =   372
            Index           =   3
            Left            =   120
            TabIndex        =   446
            Top             =   1440
            Width           =   2772
         End
         Begin VB.OptionButton optTxModPattern 
            Caption         =   "0x00"
            Height          =   372
            Index           =   2
            Left            =   120
            TabIndex        =   445
            Top             =   1080
            Width           =   2772
         End
         Begin VB.OptionButton optTxModPattern 
            Caption         =   "Psuedo Random"
            Height          =   372
            Index           =   1
            Left            =   120
            TabIndex        =   444
            Top             =   720
            Width           =   2772
         End
         Begin VB.OptionButton optTxModPattern 
            Caption         =   "1,2,3,4,...,127 "
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   443
            Top             =   360
            Width           =   2772
         End
      End
      Begin VB.CommandButton cmdRepeatedQuery 
         Caption         =   "Query Version Test"
         Height          =   612
         Left            =   5280
         TabIndex        =   440
         ToolTipText     =   "Pressing this button will set the Channel Register in the micro and the radio"
         Top             =   6000
         Width           =   3012
      End
      Begin VB.CommandButton cmdClearStatusCodeFailCount 
         Caption         =   "Clear Status Code Fail Count"
         Height          =   372
         Left            =   -72840
         TabIndex        =   431
         ToolTipText     =   "Send Clear Statistic Message to Module"
         Top             =   9120
         Width           =   4572
      End
      Begin VB.Frame Frame20 
         Height          =   1332
         Left            =   -74640
         TabIndex        =   427
         Top             =   7680
         Width           =   6372
         Begin VB.Label lblStatusCode 
            Caption         =   "Status Code:"
            Height          =   252
            Left            =   240
            TabIndex        =   430
            Top             =   240
            Width           =   3612
         End
         Begin VB.Label lblStatusCodeDescription 
            Caption         =   "Status Code Description:"
            Height          =   252
            Left            =   240
            TabIndex        =   429
            Top             =   600
            Width           =   5532
         End
         Begin VB.Label lblStatusCodeFailCount 
            Caption         =   "Status Code != Success Count:"
            Height          =   252
            Left            =   240
            TabIndex        =   428
            Top             =   960
            Width           =   5892
         End
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   0
         Left            =   -67800
         TabIndex        =   421
         Top             =   7680
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   0
         Left            =   -68400
         TabIndex        =   420
         Top             =   7680
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   0
         Left            =   -69600
         TabIndex        =   419
         Text            =   "Text1"
         Top             =   7680
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   13
         Left            =   -69600
         TabIndex        =   414
         Text            =   "Text1"
         Top             =   1440
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   14
         Left            =   -69600
         TabIndex        =   413
         Text            =   "Text1"
         Top             =   1920
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   15
         Left            =   -69600
         TabIndex        =   412
         Text            =   "Text1"
         Top             =   2400
         Width           =   972
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   13
         Left            =   -68400
         TabIndex        =   411
         Top             =   1440
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   14
         Left            =   -68400
         TabIndex        =   410
         Top             =   1920
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   15
         Left            =   -68400
         TabIndex        =   409
         Top             =   2400
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   13
         Left            =   -67800
         TabIndex        =   408
         Top             =   1440
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   14
         Left            =   -67800
         TabIndex        =   407
         Top             =   1920
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   15
         Left            =   -67800
         TabIndex        =   406
         Top             =   2400
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   10
         Left            =   -69600
         TabIndex        =   405
         Text            =   "Text1"
         Top             =   960
         Width           =   972
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   10
         Left            =   -68400
         TabIndex        =   404
         Top             =   960
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   10
         Left            =   -67800
         TabIndex        =   403
         Top             =   960
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   21
         Left            =   -70440
         TabIndex        =   401
         Top             =   7560
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   21
         Left            =   -71040
         TabIndex        =   400
         Top             =   7560
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   21
         Left            =   -72240
         TabIndex        =   399
         Text            =   "Text1"
         Top             =   7560
         Width           =   972
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   27
         Left            =   -67800
         TabIndex        =   397
         Top             =   2880
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   27
         Left            =   -68400
         TabIndex        =   396
         Top             =   2880
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   27
         Left            =   -69600
         TabIndex        =   395
         Text            =   "Text1"
         Top             =   2880
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   11
         Left            =   -69600
         TabIndex        =   391
         Text            =   "Text1"
         Top             =   4320
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   12
         Left            =   -72600
         TabIndex        =   390
         Text            =   "0123456789ABCDEF0123456789ABCDEF"
         Top             =   4800
         Width           =   3972
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   11
         Left            =   -68400
         TabIndex        =   389
         Top             =   4320
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   12
         Left            =   -68400
         TabIndex        =   388
         Top             =   4800
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   11
         Left            =   -67800
         TabIndex        =   387
         Top             =   4320
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   12
         Left            =   -67800
         TabIndex        =   386
         Top             =   4800
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   4
         Left            =   -70680
         TabIndex        =   385
         Text            =   "????????????????"
         Top             =   3840
         Width           =   2052
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   4
         Left            =   -68400
         TabIndex        =   384
         Top             =   3840
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   4
         Left            =   -67800
         TabIndex        =   383
         Top             =   3840
         Width           =   492
      End
      Begin VB.CommandButton cmdQueryVersion 
         Caption         =   "Query Version"
         Height          =   372
         Left            =   -72960
         TabIndex        =   382
         ToolTipText     =   "Send Query Version Message to Module"
         Top             =   3360
         Width           =   2172
      End
      Begin VB.Frame Frame6 
         Caption         =   "Version Information"
         Height          =   2532
         Left            =   -74640
         TabIndex        =   375
         Top             =   720
         Width           =   6372
         Begin VB.Label lblPcbVersion 
            Caption         =   "PCB Hardware Version"
            Height          =   252
            Left            =   240
            TabIndex        =   381
            Top             =   360
            Width           =   4452
         End
         Begin VB.Label lblIcVersionNumber 
            Caption         =   "RF IC Version Number"
            Height          =   252
            Left            =   240
            TabIndex        =   380
            ToolTipText     =   "If the radio is sleeping this will be zero"
            Top             =   720
            Width           =   4452
         End
         Begin VB.Label lblIcPartNumber 
            Caption         =   "RF IC Part Number"
            Height          =   252
            Left            =   240
            TabIndex        =   379
            ToolTipText     =   "If the radio is sleeping this will be zero"
            Top             =   1080
            Width           =   4452
         End
         Begin VB.Label lblAppFirmwareDescription 
            Caption         =   "Application Version String"
            Height          =   252
            Left            =   240
            TabIndex        =   378
            Top             =   2160
            Width           =   4452
         End
         Begin VB.Label lblAppFirmwareVersion 
            Caption         =   "Application Firmware Version"
            Height          =   252
            Left            =   240
            TabIndex        =   377
            Top             =   1440
            Width           =   4452
         End
         Begin VB.Label lblAppFirmwareDate 
            Caption         =   "Application Firmware Date"
            Height          =   252
            Left            =   240
            TabIndex        =   376
            Top             =   1800
            Width           =   4452
         End
      End
      Begin VB.CommandButton cmdClearVersion 
         Caption         =   "Clear Version"
         Height          =   372
         Left            =   -70560
         TabIndex        =   374
         ToolTipText     =   "Clear Version Information on this Form"
         Top             =   3360
         Width           =   2292
      End
      Begin VB.Frame Frame1 
         Caption         =   "Statistics"
         Height          =   3252
         Left            =   -74640
         TabIndex        =   373
         Top             =   3840
         Width           =   6372
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   7
            Left            =   240
            TabIndex        =   439
            Top             =   2880
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   6
            Left            =   240
            TabIndex        =   438
            Top             =   2520
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   437
            Top             =   2160
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   436
            Top             =   1800
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   435
            Top             =   1440
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   434
            Top             =   1080
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   433
            Top             =   720
            Width           =   5892
         End
         Begin VB.Label lblStatistic 
            Caption         =   "Statistic:"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   432
            Top             =   360
            Width           =   5892
         End
      End
      Begin VB.CommandButton cmdQueryStatistics 
         Caption         =   "Query Statistics"
         Height          =   372
         Left            =   -72840
         TabIndex        =   372
         ToolTipText     =   "Send Query Statistic Message to Module"
         Top             =   7200
         Width           =   2172
      End
      Begin VB.CommandButton cmdClearStatistics 
         Caption         =   "Clear Statistics"
         Height          =   372
         Left            =   -70560
         TabIndex        =   371
         ToolTipText     =   "Send Clear Statistic Message to Module"
         Top             =   7200
         Width           =   2292
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   28
         Left            =   -67800
         TabIndex        =   370
         Top             =   8640
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   28
         Left            =   -68400
         TabIndex        =   363
         Top             =   8640
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   28
         Left            =   -69600
         TabIndex        =   362
         Text            =   "Text1"
         Top             =   8640
         Width           =   972
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   22
         Left            =   -68400
         TabIndex        =   361
         Top             =   5280
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   23
         Left            =   -68400
         TabIndex        =   360
         Top             =   5760
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   24
         Left            =   -68400
         TabIndex        =   359
         Top             =   6240
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   25
         Left            =   -68400
         TabIndex        =   358
         Top             =   6720
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   26
         Left            =   -68400
         TabIndex        =   357
         Top             =   8160
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   22
         Left            =   -67800
         TabIndex        =   356
         Top             =   5280
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   23
         Left            =   -67800
         TabIndex        =   355
         Top             =   5760
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   24
         Left            =   -67800
         TabIndex        =   354
         Top             =   6240
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   25
         Left            =   -67800
         TabIndex        =   353
         Top             =   6720
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   26
         Left            =   -67800
         TabIndex        =   352
         Top             =   8160
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   22
         Left            =   -70680
         TabIndex        =   351
         Text            =   "Text1"
         Top             =   5280
         Width           =   2052
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   23
         Left            =   -69600
         TabIndex        =   350
         Text            =   "Text1"
         Top             =   5760
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   24
         Left            =   -70680
         TabIndex        =   349
         Text            =   "Text1"
         Top             =   6240
         Width           =   2052
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   25
         Left            =   -69600
         TabIndex        =   348
         Text            =   "Text1"
         Top             =   6720
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   26
         Left            =   -70680
         TabIndex        =   347
         Text            =   "Text1"
         Top             =   8160
         Width           =   2052
      End
      Begin VB.CommandButton cmdStartStopSyncTest 
         Caption         =   "Start/Stop Sync Test"
         Height          =   612
         Left            =   5280
         TabIndex        =   344
         ToolTipText     =   "Pressing this button will set the Channel Register in the micro and the radio"
         Top             =   5280
         Width           =   3012
      End
      Begin VB.Frame Frame25 
         Caption         =   "Sync Test"
         Height          =   972
         Left            =   5280
         TabIndex        =   341
         Top             =   4200
         Width           =   3012
         Begin VB.OptionButton optSyncTestState 
            Caption         =   "Idle"
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   343
            Top             =   360
            Width           =   1092
         End
         Begin VB.OptionButton optSyncTestState 
            Caption         =   "Running"
            Height          =   372
            Index           =   1
            Left            =   1320
            TabIndex        =   342
            Top             =   360
            Width           =   1212
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Module Mode"
         Height          =   2412
         Left            =   5280
         TabIndex        =   323
         Top             =   840
         Width           =   3012
         Begin VB.OptionButton optModuleMode 
            Caption         =   "Micro On Radio Rx"
            Height          =   492
            Index           =   0
            Left            =   120
            TabIndex        =   327
            Top             =   360
            Width           =   2652
         End
         Begin VB.OptionButton optModuleMode 
            Caption         =   "Micro Off Radio Off"
            Height          =   492
            Index           =   1
            Left            =   120
            TabIndex        =   326
            Top             =   840
            Width           =   2652
         End
         Begin VB.OptionButton optModuleMode 
            Caption         =   "Micro On Radio Off"
            Height          =   492
            Index           =   2
            Left            =   120
            TabIndex        =   325
            Top             =   1320
            Width           =   2652
         End
         Begin VB.OptionButton optModuleMode 
            Caption         =   "Micro On Radio On"
            Height          =   492
            Index           =   3
            Left            =   120
            TabIndex        =   324
            Top             =   1800
            Width           =   2652
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Power Level"
         Height          =   8772
         Left            =   240
         TabIndex        =   299
         Top             =   840
         Width           =   1572
         Begin VB.OptionButton Option1 
            Caption         =   "11"
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   322
            Top             =   240
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "10"
            Height          =   372
            Index           =   1
            Left            =   240
            TabIndex        =   321
            Top             =   600
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "9"
            Height          =   372
            Index           =   2
            Left            =   240
            TabIndex        =   320
            Top             =   960
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "8"
            Height          =   372
            Index           =   3
            Left            =   240
            TabIndex        =   319
            Top             =   1320
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "7"
            Height          =   372
            Index           =   4
            Left            =   240
            TabIndex        =   318
            Top             =   1680
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "6"
            Height          =   372
            Index           =   5
            Left            =   240
            TabIndex        =   317
            Top             =   2040
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "5"
            Height          =   372
            Index           =   6
            Left            =   240
            TabIndex        =   316
            Top             =   2400
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "4"
            Height          =   372
            Index           =   7
            Left            =   240
            TabIndex        =   315
            Top             =   2760
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3"
            Height          =   372
            Index           =   8
            Left            =   240
            TabIndex        =   314
            Top             =   3120
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   372
            Index           =   9
            Left            =   240
            TabIndex        =   313
            Top             =   3480
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Height          =   372
            Index           =   10
            Left            =   240
            TabIndex        =   312
            Top             =   3840
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
            Height          =   372
            Index           =   11
            Left            =   240
            TabIndex        =   311
            Top             =   4200
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-1"
            Height          =   372
            Index           =   12
            Left            =   240
            TabIndex        =   310
            Top             =   4560
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-2"
            Height          =   372
            Index           =   13
            Left            =   240
            TabIndex        =   309
            Top             =   4920
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-3"
            Height          =   372
            Index           =   14
            Left            =   240
            TabIndex        =   308
            Top             =   5280
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-4"
            Height          =   372
            Index           =   15
            Left            =   240
            TabIndex        =   307
            Top             =   5640
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-5"
            Height          =   372
            Index           =   16
            Left            =   240
            TabIndex        =   306
            Top             =   6000
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-6"
            Height          =   372
            Index           =   17
            Left            =   240
            TabIndex        =   305
            Top             =   6360
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-7"
            Height          =   372
            Index           =   18
            Left            =   240
            TabIndex        =   304
            Top             =   6720
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-8"
            Height          =   372
            Index           =   19
            Left            =   240
            TabIndex        =   303
            Top             =   7080
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-9"
            Height          =   372
            Index           =   20
            Left            =   240
            TabIndex        =   302
            Top             =   7440
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-10"
            Height          =   372
            Index           =   21
            Left            =   240
            TabIndex        =   301
            Top             =   7800
            Width           =   972
         End
         Begin VB.OptionButton Option1 
            Caption         =   "-11"
            Height          =   372
            Index           =   22
            Left            =   240
            TabIndex        =   300
            Top             =   8160
            Width           =   972
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Test Mode"
         Height          =   3852
         Left            =   2040
         TabIndex        =   291
         Top             =   840
         Width           =   2892
         Begin VB.OptionButton Option2 
            Caption         =   "Idle"
            Height          =   372
            Index           =   0
            Left            =   120
            TabIndex        =   298
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Rx"
            Height          =   372
            Index           =   1
            Left            =   120
            TabIndex        =   297
            Top             =   840
            Width           =   972
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tx Unmod +0.25 MHz"
            Height          =   372
            Index           =   2
            Left            =   120
            TabIndex        =   296
            Top             =   1320
            Width           =   2532
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tx Unmod -0.25 MHz"
            Height          =   372
            Index           =   3
            Left            =   120
            TabIndex        =   295
            Top             =   1800
            Width           =   2532
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tx Unmod +0.10 MHz"
            Height          =   372
            Index           =   4
            Left            =   120
            TabIndex        =   294
            Top             =   2280
            Width           =   2532
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tx Unmod -0.10 MHz"
            Height          =   372
            Index           =   5
            Left            =   120
            TabIndex        =   293
            Top             =   2760
            Width           =   2652
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Tx Mod"
            Height          =   372
            Index           =   6
            Left            =   120
            TabIndex        =   292
            Top             =   3240
            Width           =   1692
         End
      End
      Begin VB.CommandButton cmdSendTestMode 
         Caption         =   "Set Test Mode"
         Height          =   612
         Left            =   2040
         TabIndex        =   290
         ToolTipText     =   $"TestTool.frx":1AD9
         Top             =   4800
         Width           =   2892
      End
      Begin VB.Frame Frame5 
         Caption         =   "Channel"
         Height          =   2292
         Left            =   2040
         TabIndex        =   279
         Top             =   5760
         Width           =   2892
         Begin VB.OptionButton Option3 
            Caption         =   "1"
            Height          =   372
            Index           =   0
            Left            =   480
            TabIndex        =   289
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton Option3 
            Caption         =   "2"
            Height          =   372
            Index           =   1
            Left            =   480
            TabIndex        =   288
            Top             =   720
            Width           =   972
         End
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Height          =   372
            Index           =   2
            Left            =   480
            TabIndex        =   287
            Top             =   1080
            Width           =   972
         End
         Begin VB.OptionButton Option3 
            Caption         =   "4"
            Height          =   372
            Index           =   3
            Left            =   480
            TabIndex        =   286
            Top             =   1440
            Width           =   972
         End
         Begin VB.OptionButton Option3 
            Caption         =   "5"
            Height          =   372
            Index           =   4
            Left            =   480
            TabIndex        =   285
            Top             =   1800
            Width           =   972
         End
         Begin VB.OptionButton Option3 
            Caption         =   "6"
            Height          =   372
            Index           =   5
            Left            =   1800
            TabIndex        =   284
            Top             =   360
            Width           =   612
         End
         Begin VB.OptionButton Option3 
            Caption         =   "7"
            Height          =   372
            Index           =   6
            Left            =   1800
            TabIndex        =   283
            Top             =   720
            Width           =   492
         End
         Begin VB.OptionButton Option3 
            Caption         =   "8"
            Height          =   372
            Index           =   7
            Left            =   1800
            TabIndex        =   282
            Top             =   1080
            Width           =   492
         End
         Begin VB.OptionButton Option3 
            Caption         =   "9"
            Height          =   372
            Index           =   8
            Left            =   1800
            TabIndex        =   281
            Top             =   1440
            Width           =   612
         End
         Begin VB.OptionButton Option3 
            Caption         =   "10"
            Height          =   372
            Index           =   9
            Left            =   1800
            TabIndex        =   280
            Top             =   1800
            Width           =   612
         End
      End
      Begin VB.Frame frmMessageTypes 
         Caption         =   "Message Types"
         Height          =   3300
         Left            =   -74760
         TabIndex        =   254
         Top             =   6240
         Width           =   8052
         Begin VB.OptionButton optMsgType 
            Caption         =   "Invalid"
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   278
            Top             =   360
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Query Version"
            Height          =   372
            Index           =   1
            Left            =   240
            TabIndex        =   277
            Top             =   720
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Static Test Mode"
            Height          =   372
            Index           =   2
            Left            =   240
            TabIndex        =   276
            Top             =   1080
            Width           =   2412
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Set Radio Register"
            Height          =   372
            Index           =   3
            Left            =   240
            TabIndex        =   275
            Top             =   1440
            Width           =   2532
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Query Radio Register"
            Height          =   372
            Index           =   4
            Left            =   240
            TabIndex        =   274
            Top             =   1800
            Width           =   2772
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Set Micro Register"
            Height          =   372
            Index           =   5
            Left            =   240
            TabIndex        =   273
            Top             =   2160
            Width           =   2652
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Query Micro Register"
            Height          =   372
            Index           =   6
            Left            =   240
            TabIndex        =   272
            Top             =   2520
            Width           =   2772
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Reserved"
            Height          =   372
            Index           =   7
            Left            =   5880
            TabIndex        =   271
            Top             =   2160
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Reserved"
            Height          =   372
            Index           =   8
            Left            =   5880
            TabIndex        =   270
            Top             =   2520
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "ED Scan Done"
            Height          =   372
            Index           =   18
            Left            =   3120
            TabIndex        =   269
            Top             =   2880
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Save to Nval"
            Height          =   372
            Index           =   19
            Left            =   5880
            TabIndex        =   268
            Top             =   360
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Set Sleep Mode"
            Height          =   372
            Index           =   20
            Left            =   5880
            TabIndex        =   267
            Top             =   720
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Reserved"
            Height          =   372
            Index           =   21
            Left            =   5880
            TabIndex        =   266
            Top             =   1080
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Reserved"
            Height          =   372
            Index           =   22
            Left            =   5880
            TabIndex        =   265
            Top             =   1440
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Reserved"
            Height          =   372
            Index           =   23
            Left            =   5880
            TabIndex        =   264
            Top             =   1800
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "ED Scan Request"
            Height          =   372
            Index           =   17
            Left            =   3120
            TabIndex        =   263
            Top             =   2520
            Width           =   2532
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "PERT Recieve Done"
            Height          =   372
            Index           =   16
            Left            =   3120
            TabIndex        =   262
            Top             =   2160
            Width           =   2412
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "PERT Receive"
            Height          =   372
            Index           =   15
            Left            =   3120
            TabIndex        =   261
            Top             =   1800
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "PERT Transmit Done"
            Height          =   372
            Index           =   14
            Left            =   3120
            TabIndex        =   260
            Top             =   1440
            Width           =   2652
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "PERT Transmit"
            Height          =   372
            Index           =   13
            Left            =   3120
            TabIndex        =   259
            Top             =   1080
            Width           =   2052
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Recieve RF Message"
            Height          =   360
            Index           =   12
            Left            =   3120
            TabIndex        =   258
            Top             =   720
            Width           =   2652
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Transmit RF Message"
            Height          =   372
            Index           =   11
            Left            =   3120
            TabIndex        =   257
            Top             =   360
            Width           =   2652
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Clear Statistics"
            Height          =   372
            Index           =   10
            Left            =   240
            TabIndex        =   256
            Top             =   2880
            Width           =   2412
         End
         Begin VB.OptionButton optMsgType 
            Caption         =   "Query Statistics"
            Height          =   372
            Index           =   9
            Left            =   5520
            TabIndex        =   255
            Top             =   2880
            Width           =   2412
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transmit Data (Hex)"
         Height          =   5532
         Left            =   -74760
         TabIndex        =   169
         Top             =   720
         Width           =   5772
         Begin VB.OptionButton optByteCount 
            Caption         =   "0"
            Height          =   372
            Index           =   0
            Left            =   360
            TabIndex        =   253
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "1"
            Height          =   372
            Index           =   1
            Left            =   1080
            TabIndex        =   252
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "2"
            Height          =   372
            Index           =   2
            Left            =   1800
            TabIndex        =   251
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "3"
            Height          =   372
            Index           =   3
            Left            =   2520
            TabIndex        =   250
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "4"
            Height          =   372
            Index           =   4
            Left            =   3240
            TabIndex        =   249
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "5"
            Height          =   372
            Index           =   5
            Left            =   3960
            TabIndex        =   248
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "6"
            Height          =   372
            Index           =   6
            Left            =   4680
            TabIndex        =   247
            Top             =   360
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "7"
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   246
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "8"
            Height          =   372
            Index           =   8
            Left            =   1080
            TabIndex        =   245
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "9"
            Height          =   372
            Index           =   9
            Left            =   1800
            TabIndex        =   244
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "10"
            Height          =   372
            Index           =   10
            Left            =   2520
            TabIndex        =   243
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "11"
            Height          =   372
            Index           =   11
            Left            =   3240
            TabIndex        =   242
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "12"
            Height          =   372
            Index           =   12
            Left            =   3960
            TabIndex        =   241
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "13"
            Height          =   372
            Index           =   13
            Left            =   4680
            TabIndex        =   240
            Top             =   1200
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "14"
            Height          =   372
            Index           =   14
            Left            =   360
            TabIndex        =   239
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "15"
            Height          =   372
            Index           =   15
            Left            =   1080
            TabIndex        =   238
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "16"
            Height          =   372
            Index           =   16
            Left            =   1800
            TabIndex        =   237
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "17"
            Height          =   372
            Index           =   17
            Left            =   2520
            TabIndex        =   236
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "18"
            Height          =   372
            Index           =   18
            Left            =   3240
            TabIndex        =   235
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "19"
            Height          =   372
            Index           =   19
            Left            =   3960
            TabIndex        =   234
            Top             =   2040
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "20"
            Height          =   372
            Index           =   20
            Left            =   4680
            TabIndex        =   233
            Top             =   2040
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   360
            Index           =   0
            Left            =   360
            TabIndex        =   232
            Text            =   "0x00"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   1
            Left            =   1080
            TabIndex        =   231
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   2
            Left            =   1800
            TabIndex        =   230
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   3
            Left            =   2520
            TabIndex        =   229
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   4
            Left            =   3240
            TabIndex        =   228
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   5
            Left            =   3960
            TabIndex        =   227
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   11
            Left            =   3240
            TabIndex        =   226
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   12
            Left            =   3960
            TabIndex        =   225
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   13
            Left            =   4680
            TabIndex        =   224
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   14
            Left            =   360
            TabIndex        =   223
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   15
            Left            =   1080
            TabIndex        =   222
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   16
            Left            =   1800
            TabIndex        =   221
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   17
            Left            =   2520
            TabIndex        =   220
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   18
            Left            =   3240
            TabIndex        =   219
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   19
            Left            =   3960
            TabIndex        =   218
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   20
            Left            =   4680
            TabIndex        =   217
            Text            =   "Text1"
            Top             =   2400
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   6
            Left            =   4680
            TabIndex        =   216
            Text            =   "Text1"
            Top             =   720
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   7
            Left            =   360
            TabIndex        =   215
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   8
            Left            =   1080
            TabIndex        =   214
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   9
            Left            =   1800
            TabIndex        =   213
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   10
            Left            =   2520
            TabIndex        =   212
            Text            =   "Text1"
            Top             =   1560
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   21
            Left            =   360
            TabIndex        =   211
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   22
            Left            =   1080
            TabIndex        =   210
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   23
            Left            =   1800
            TabIndex        =   209
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   24
            Left            =   2520
            TabIndex        =   208
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   25
            Left            =   3240
            TabIndex        =   207
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   26
            Left            =   3960
            TabIndex        =   206
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   27
            Left            =   4680
            TabIndex        =   205
            Text            =   "Text1"
            Top             =   3240
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   28
            Left            =   360
            TabIndex        =   204
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   29
            Left            =   1080
            TabIndex        =   203
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   30
            Left            =   1800
            TabIndex        =   202
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   31
            Left            =   2520
            TabIndex        =   201
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   32
            Left            =   3240
            TabIndex        =   200
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   33
            Left            =   3960
            TabIndex        =   199
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   34
            Left            =   4680
            TabIndex        =   198
            Text            =   "Text1"
            Top             =   4080
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   35
            Left            =   360
            TabIndex        =   197
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   36
            Left            =   1080
            TabIndex        =   196
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   37
            Left            =   1800
            TabIndex        =   195
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   38
            Left            =   2520
            TabIndex        =   194
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   39
            Left            =   3240
            TabIndex        =   193
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   40
            Left            =   3960
            TabIndex        =   192
            Text            =   "Text1"
            Top             =   4920
            Width           =   700
         End
         Begin VB.TextBox txtByte 
            Height          =   372
            Index           =   41
            Left            =   4680
            TabIndex        =   191
            Text            =   "0x00"
            Top             =   4920
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "20"
            Height          =   372
            Index           =   21
            Left            =   360
            TabIndex        =   190
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "19"
            Height          =   372
            Index           =   22
            Left            =   1080
            TabIndex        =   189
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "18"
            Height          =   372
            Index           =   23
            Left            =   1800
            TabIndex        =   188
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "17"
            Height          =   372
            Index           =   24
            Left            =   2520
            TabIndex        =   187
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "16"
            Height          =   372
            Index           =   25
            Left            =   3240
            TabIndex        =   186
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "15"
            Height          =   372
            Index           =   26
            Left            =   3960
            TabIndex        =   185
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "14"
            Height          =   372
            Index           =   27
            Left            =   4680
            TabIndex        =   184
            Top             =   2880
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "13"
            Height          =   372
            Index           =   28
            Left            =   360
            TabIndex        =   183
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "12"
            Height          =   372
            Index           =   29
            Left            =   1080
            TabIndex        =   182
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "11"
            Height          =   372
            Index           =   30
            Left            =   1800
            TabIndex        =   181
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "10"
            Height          =   372
            Index           =   31
            Left            =   2520
            TabIndex        =   180
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "9"
            Height          =   372
            Index           =   32
            Left            =   3240
            TabIndex        =   179
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "8"
            Height          =   372
            Index           =   33
            Left            =   3960
            TabIndex        =   178
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "7"
            Height          =   372
            Index           =   34
            Left            =   4680
            TabIndex        =   177
            Top             =   3720
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "6"
            Height          =   372
            Index           =   35
            Left            =   360
            TabIndex        =   176
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "5"
            Height          =   372
            Index           =   36
            Left            =   1080
            TabIndex        =   175
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "4"
            Height          =   372
            Index           =   37
            Left            =   1800
            TabIndex        =   174
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "3"
            Height          =   372
            Index           =   38
            Left            =   2520
            TabIndex        =   173
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "2"
            Height          =   372
            Index           =   39
            Left            =   3240
            TabIndex        =   172
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "1"
            Height          =   372
            Index           =   40
            Left            =   3960
            TabIndex        =   171
            Top             =   4560
            Width           =   700
         End
         Begin VB.OptionButton optByteCount 
            Caption         =   "0"
            Height          =   372
            Index           =   41
            Left            =   4680
            TabIndex        =   170
            Top             =   4560
            Width           =   700
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "PimID"
         Height          =   972
         Left            =   -74640
         TabIndex        =   167
         Top             =   840
         Width           =   1692
         Begin VB.TextBox txtPertPimId 
            Height          =   372
            Left            =   360
            TabIndex        =   168
            Text            =   "Text1"
            Top             =   360
            Width           =   972
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Source Address"
         Height          =   972
         Left            =   -72600
         TabIndex        =   165
         Top             =   840
         Width           =   2052
         Begin VB.TextBox txtPertSourceAddr 
            Height          =   372
            Left            =   360
            TabIndex        =   166
            Text            =   "Text1"
            Top             =   360
            Width           =   972
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Destination Address"
         Height          =   972
         Left            =   -70320
         TabIndex        =   163
         Top             =   840
         Width           =   2652
         Begin VB.TextBox txtPertDestAddr 
            Height          =   372
            Left            =   360
            TabIndex        =   164
            Text            =   "Text1"
            Top             =   360
            Width           =   972
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Number of Packets"
         Height          =   972
         Left            =   -74640
         TabIndex        =   160
         Top             =   2040
         Width           =   3012
         Begin VB.TextBox txtPertNumberOfPackets 
            Height          =   372
            Left            =   360
            TabIndex        =   161
            Text            =   "Text1"
            Top             =   360
            Width           =   972
         End
         Begin VB.Label lblPertNumberPacketsDecimal 
            Caption         =   "= 0"
            Height          =   372
            Left            =   1680
            TabIndex        =   162
            Top             =   360
            Width           =   732
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Packet Size"
         Height          =   972
         Left            =   -71520
         TabIndex        =   157
         Top             =   2040
         Width           =   2652
         Begin VB.TextBox txtPertPacketSize 
            Height          =   372
            Left            =   360
            TabIndex        =   158
            Text            =   "Text1"
            Top             =   360
            Width           =   972
         End
         Begin VB.Label lblPertPacketSizeDecimal 
            Caption         =   "= 0"
            Height          =   372
            Left            =   1680
            TabIndex        =   159
            Top             =   360
            Width           =   732
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Receive"
         Height          =   2772
         Left            =   -74640
         TabIndex        =   150
         Top             =   6600
         Width           =   3612
         Begin VB.Frame Frame16 
            Caption         =   "RF Messages to Host"
            Height          =   972
            Left            =   240
            TabIndex        =   151
            Top             =   360
            Width           =   3012
            Begin VB.OptionButton optRfToHost 
               Caption         =   "Default"
               Height          =   372
               Index           =   0
               Left            =   120
               TabIndex        =   153
               Top             =   360
               Width           =   1572
            End
            Begin VB.OptionButton optRfToHost 
               Caption         =   "No"
               Height          =   372
               Index           =   1
               Left            =   1920
               TabIndex        =   152
               Top             =   360
               Width           =   972
            End
         End
         Begin VB.Label lblPertMatchingPackets 
            Caption         =   "Label1"
            Height          =   252
            Left            =   240
            TabIndex        =   156
            Top             =   2280
            Width           =   3012
         End
         Begin VB.Label lblPertExpectedPackets 
            Caption         =   "Label1"
            Height          =   252
            Left            =   240
            TabIndex        =   155
            Top             =   1920
            Width           =   3012
         End
         Begin VB.Label lblPertReceivedPackets 
            Caption         =   "Label1"
            Height          =   252
            Left            =   240
            TabIndex        =   154
            Top             =   1560
            Width           =   3012
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Transmit"
         Height          =   3132
         Left            =   -74640
         TabIndex        =   138
         Top             =   3240
         Width           =   3612
         Begin VB.Frame ackFrame 
            Caption         =   "RF Request Ack"
            Height          =   972
            Left            =   240
            TabIndex        =   146
            Top             =   360
            Width           =   3132
            Begin VB.OptionButton optRequestAck 
               Caption         =   "Disabled"
               Height          =   372
               Index           =   0
               Left            =   120
               TabIndex        =   148
               Top             =   360
               Width           =   1452
            End
            Begin VB.OptionButton optRequestAck 
               Caption         =   "Enabled"
               Height          =   372
               Index           =   1
               Left            =   1560
               TabIndex        =   147
               Top             =   360
               Width           =   1332
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Retry Mode"
            Height          =   972
            Left            =   240
            TabIndex        =   139
            Top             =   1440
            Width           =   3132
            Begin VB.OptionButton optRetry 
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   3
               Left            =   1560
               TabIndex        =   145
               Top             =   360
               Width           =   492
            End
            Begin VB.OptionButton optRetry 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   4
               Left            =   2040
               TabIndex        =   144
               Top             =   360
               Width           =   492
            End
            Begin VB.OptionButton optRetry 
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   0
               Left            =   120
               TabIndex        =   143
               Top             =   360
               Width           =   492
            End
            Begin VB.OptionButton optRetry 
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   1
               Left            =   600
               TabIndex        =   142
               Top             =   360
               Width           =   492
            End
            Begin VB.OptionButton optRetry 
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   2
               Left            =   1080
               TabIndex        =   141
               Top             =   360
               Width           =   492
            End
            Begin VB.OptionButton optRetry 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   7.8
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   5
               Left            =   2520
               TabIndex        =   140
               Top             =   360
               Width           =   492
            End
         End
         Begin VB.Label lblPacketsNotAcked 
            Caption         =   "Label"
            Height          =   252
            Left            =   240
            TabIndex        =   149
            Top             =   2640
            Width           =   3132
         End
      End
      Begin VB.CommandButton cmdTxPert 
         Caption         =   "Start PERT Tx"
         Height          =   852
         Left            =   -70440
         TabIndex        =   137
         ToolTipText     =   "Send Start Tx PERT command to Micro"
         Top             =   3240
         Width           =   1932
      End
      Begin VB.CommandButton cmdRxPert 
         Caption         =   "Start PERT Rx"
         Height          =   852
         Left            =   -70440
         TabIndex        =   136
         ToolTipText     =   "Send Start Rx PERT command to Micro"
         Top             =   6720
         Width           =   1932
      End
      Begin VB.CommandButton cmdResetPert 
         Caption         =   "Reset Pert"
         Height          =   852
         Left            =   -70440
         TabIndex        =   135
         ToolTipText     =   "Reset the PERT state machine in this Application"
         Top             =   8520
         Width           =   1932
      End
      Begin VB.Frame Frame18 
         Caption         =   "Pert State"
         Height          =   1932
         Left            =   -70440
         TabIndex        =   131
         Top             =   4440
         Width           =   1932
         Begin VB.OptionButton optPertState 
            Caption         =   "Busy Tx"
            Height          =   372
            Index           =   1
            Left            =   240
            TabIndex        =   134
            Top             =   840
            Width           =   1452
         End
         Begin VB.OptionButton optPertState 
            Caption         =   "Idle"
            Height          =   372
            Index           =   0
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton optPertState 
            Caption         =   "Busy Rx"
            Height          =   372
            Index           =   2
            Left            =   240
            TabIndex        =   132
            Top             =   1320
            Width           =   1212
         End
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   1
         Left            =   -72240
         TabIndex        =   130
         Text            =   "Text1"
         Top             =   1320
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   2
         Left            =   -72240
         TabIndex        =   129
         Text            =   "Text1"
         Top             =   1800
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   3
         Left            =   -72240
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   2280
         Width           =   972
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   1
         Left            =   -71040
         TabIndex        =   127
         Top             =   1320
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   2
         Left            =   -71040
         TabIndex        =   126
         Top             =   1800
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   3
         Left            =   -71040
         TabIndex        =   125
         Top             =   2280
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   1
         Left            =   -70440
         TabIndex        =   124
         Top             =   1320
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   2
         Left            =   -70440
         TabIndex        =   123
         Top             =   1800
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   3
         Left            =   -70440
         TabIndex        =   122
         Top             =   2280
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   9
         Left            =   -70440
         TabIndex        =   121
         Top             =   4680
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   8
         Left            =   -70440
         TabIndex        =   120
         Top             =   4200
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   7
         Left            =   -70440
         TabIndex        =   119
         Top             =   3720
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   6
         Left            =   -70440
         TabIndex        =   118
         Top             =   3240
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   5
         Left            =   -70440
         TabIndex        =   117
         Top             =   2760
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   9
         Left            =   -71040
         TabIndex        =   116
         Top             =   4680
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   8
         Left            =   -71040
         TabIndex        =   115
         Top             =   4200
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   7
         Left            =   -71040
         TabIndex        =   114
         Top             =   3720
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   6
         Left            =   -71040
         TabIndex        =   113
         Top             =   3240
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   5
         Left            =   -71040
         TabIndex        =   112
         Top             =   2760
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   9
         Left            =   -72240
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   4680
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   8
         Left            =   -72240
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   4200
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   7
         Left            =   -72240
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   3720
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   6
         Left            =   -72240
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   3240
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   5
         Left            =   -72240
         TabIndex        =   107
         Text            =   "Text1"
         Top             =   2760
         Width           =   972
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   20
         Left            =   -70440
         TabIndex        =   106
         Top             =   7080
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   19
         Left            =   -70440
         TabIndex        =   105
         Top             =   6600
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   18
         Left            =   -70440
         TabIndex        =   104
         Top             =   6120
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   17
         Left            =   -70440
         TabIndex        =   103
         Top             =   5640
         Width           =   492
      End
      Begin VB.CommandButton cmdGetMicroRegister 
         Caption         =   "Get"
         Height          =   372
         Index           =   16
         Left            =   -70440
         TabIndex        =   102
         Top             =   5160
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   20
         Left            =   -71040
         TabIndex        =   101
         Top             =   7080
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   19
         Left            =   -71040
         TabIndex        =   100
         Top             =   6600
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   18
         Left            =   -71040
         TabIndex        =   99
         Top             =   6120
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   17
         Left            =   -71040
         TabIndex        =   98
         Top             =   5640
         Width           =   492
      End
      Begin VB.CommandButton cmdSetMicroReg 
         Caption         =   "Set"
         Height          =   372
         Index           =   16
         Left            =   -71040
         TabIndex        =   97
         Top             =   5160
         Width           =   492
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   20
         Left            =   -72240
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   7080
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   19
         Left            =   -72240
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   6600
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   18
         Left            =   -72240
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   6120
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   17
         Left            =   -72240
         TabIndex        =   93
         Text            =   "Text1"
         Top             =   5640
         Width           =   972
      End
      Begin VB.TextBox txtMicroReg 
         Height          =   372
         Index           =   16
         Left            =   -72240
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   5160
         Width           =   972
      End
      Begin VB.CommandButton cmdReadRegisters 
         Caption         =   "Read All Micro Registers"
         Height          =   1332
         Left            =   -68280
         TabIndex        =   91
         Top             =   6480
         Width           =   1452
      End
      Begin VB.CommandButton cmdSetPower 
         Caption         =   "Set Power"
         Height          =   612
         Left            =   2040
         TabIndex        =   90
         ToolTipText     =   "Pressing this button will set the Power Level Register in the micro and the radio"
         Top             =   9000
         Width           =   1572
      End
      Begin VB.CommandButton cmdSetChannel 
         Caption         =   "Set Channel"
         Height          =   612
         Left            =   2040
         TabIndex        =   89
         ToolTipText     =   "Pressing this button will set the Channel Register in the micro and the radio"
         Top             =   8160
         Width           =   2892
      End
      Begin VB.CommandButton cmdSetModuleMode 
         Caption         =   "Set Module Mode"
         Height          =   612
         Left            =   5280
         TabIndex        =   88
         Top             =   3360
         Width           =   3012
      End
      Begin VB.Frame Frame22 
         Caption         =   "Background Color Key"
         Height          =   1812
         Left            =   -69600
         TabIndex        =   83
         Top             =   1080
         Width           =   3012
         Begin VB.Label Label1 
            Caption         =   "White = Read Value"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   87
            ToolTipText     =   "Value read from micro"
            Top             =   360
            Width           =   2292
         End
         Begin VB.Label Label1 
            Caption         =   "Green = Valid Value"
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   86
            ToolTipText     =   "The value entered by the user is a valid hex value"
            Top             =   720
            Width           =   2412
         End
         Begin VB.Label Label1 
            Caption         =   "Yellow = Set Value"
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   85
            ToolTipText     =   "Set Register X message sent to micro "
            Top             =   1080
            Width           =   2292
         End
         Begin VB.Label Label1 
            Caption         =   "Red = Invalid Value"
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   84
            ToolTipText     =   "Invalid hex value entered"
            Top             =   1440
            Width           =   2292
         End
      End
      Begin VB.Frame Frame23 
         Height          =   5532
         Left            =   -68640
         TabIndex        =   73
         Top             =   720
         Width           =   1932
         Begin VB.CommandButton cmdValidate 
            Caption         =   "Validate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   423
            Top             =   2280
            Width           =   1452
         End
         Begin VB.CommandButton cmdSendMessage 
            Caption         =   "Send Message"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   82
            ToolTipText     =   "Send message in Payload boxes"
            Top             =   4200
            Width           =   1452
         End
         Begin VB.CommandButton cmdComputeChecksum 
            Caption         =   "Calc Checksum"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   81
            Top             =   2760
            Width           =   1452
         End
         Begin VB.CommandButton cmdUpdateLength 
            Caption         =   "Update Length"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   80
            Top             =   1320
            Width           =   1452
         End
         Begin VB.CommandButton cmdAddStart 
            Caption         =   "Add Start"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   79
            Top             =   840
            Width           =   1452
         End
         Begin VB.CommandButton cmdAddStop 
            Caption         =   "Add Stop"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   78
            Top             =   3240
            Width           =   1452
         End
         Begin VB.CommandButton cmdClearRawTx 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   77
            Top             =   360
            Width           =   1452
         End
         Begin VB.CommandButton cmdUpdateType 
            Caption         =   "Update Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   76
            Top             =   1800
            Width           =   1452
         End
         Begin VB.CheckBox chkCopyTx 
            Caption         =   "Copy Tx"
            Height          =   372
            Left            =   120
            TabIndex        =   75
            ToolTipText     =   "When checked the messages transmitted in other tabs will be copied to the Raw Tx tab"
            Top             =   4920
            Width           =   1572
         End
         Begin VB.CommandButton cmdUpdateAndSend 
            Caption         =   "Update And Send"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   240
            TabIndex        =   74
            ToolTipText     =   "Update Start, Length, Type, Validate Data, Checksum, and Stop bytes. Then send the message."
            Top             =   3720
            Width           =   1452
         End
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Nval Version"
         Height          =   252
         Index           =   0
         Left            =   -74640
         TabIndex        =   422
         ToolTipText     =   "0"
         Top             =   7800
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Reserved1"
         Height          =   252
         Index           =   13
         Left            =   -74640
         TabIndex        =   418
         ToolTipText     =   "13"
         Top             =   1560
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Reserved2"
         Height          =   252
         Index           =   14
         Left            =   -74640
         TabIndex        =   417
         ToolTipText     =   "14"
         Top             =   2040
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Reserved3"
         Height          =   252
         Index           =   15
         Left            =   -74640
         TabIndex        =   416
         ToolTipText     =   "15"
         Top             =   2520
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Reserved0"
         Height          =   252
         Index           =   10
         Left            =   -74640
         TabIndex        =   415
         ToolTipText     =   "10"
         Top             =   1080
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Wake Up Msg Size"
         Height          =   252
         Index           =   21
         Left            =   -74640
         TabIndex        =   402
         ToolTipText     =   "21"
         Top             =   7680
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Reserved4"
         Height          =   252
         Index           =   27
         Left            =   -74640
         TabIndex        =   398
         ToolTipText     =   "27"
         Top             =   3000
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Encryption Mode"
         Height          =   252
         Index           =   11
         Left            =   -74640
         TabIndex        =   394
         ToolTipText     =   "11"
         Top             =   4440
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Encryption Key"
         Height          =   252
         Index           =   12
         Left            =   -74640
         TabIndex        =   393
         ToolTipText     =   "12"
         Top             =   4920
         Width           =   1932
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "IEEE Address"
         Height          =   252
         Index           =   4
         Left            =   -74640
         TabIndex        =   392
         ToolTipText     =   "4"
         Top             =   3960
         Width           =   2412
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Encryption IV"
         Height          =   252
         Index           =   22
         Left            =   -74640
         TabIndex        =   369
         ToolTipText     =   "22"
         Top             =   5400
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Header Size"
         Height          =   252
         Index           =   23
         Left            =   -74640
         TabIndex        =   368
         ToolTipText     =   "23"
         Top             =   5880
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Sequence Number"
         Height          =   252
         Index           =   24
         Left            =   -74640
         TabIndex        =   367
         ToolTipText     =   "24"
         Top             =   6360
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "SFD"
         Height          =   252
         Index           =   25
         Left            =   -74640
         TabIndex        =   366
         ToolTipText     =   "Start Frame Delimiter (25)"
         Top             =   6840
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Write Count"
         Height          =   252
         Index           =   26
         Left            =   -74640
         TabIndex        =   365
         ToolTipText     =   "26"
         Top             =   8280
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Token Sum"
         Height          =   252
         Index           =   28
         Left            =   -74640
         TabIndex        =   364
         ToolTipText     =   "28"
         Top             =   8760
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Operating Mode"
         Height          =   252
         Index           =   1
         Left            =   -74640
         TabIndex        =   340
         ToolTipText     =   "1"
         Top             =   1440
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Device Type"
         Height          =   252
         Index           =   2
         Left            =   -74640
         TabIndex        =   339
         ToolTipText     =   "2"
         Top             =   1920
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Own PIM ID"
         Height          =   252
         Index           =   3
         Left            =   -74640
         TabIndex        =   338
         ToolTipText     =   "3"
         Top             =   2400
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Channel Set"
         Height          =   252
         Index           =   9
         Left            =   -74640
         TabIndex        =   337
         ToolTipText     =   "9"
         Top             =   4800
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Channel"
         Height          =   252
         Index           =   8
         Left            =   -74640
         TabIndex        =   336
         ToolTipText     =   "8"
         Top             =   4320
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "RF Power"
         Height          =   252
         Index           =   7
         Left            =   -74640
         TabIndex        =   335
         ToolTipText     =   "7"
         Top             =   3840
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Default Destination Address"
         Height          =   252
         Index           =   6
         Left            =   -74640
         TabIndex        =   334
         ToolTipText     =   "6"
         Top             =   3360
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Own Source Address"
         Height          =   252
         Index           =   5
         Left            =   -74640
         TabIndex        =   333
         ToolTipText     =   "5"
         Top             =   2880
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Wake Up Mode"
         Height          =   252
         Index           =   20
         Left            =   -74640
         TabIndex        =   332
         ToolTipText     =   "20"
         Top             =   7200
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Retry Attempts"
         Height          =   252
         Index           =   19
         Left            =   -74640
         TabIndex        =   331
         ToolTipText     =   "19"
         Top             =   6720
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "RF Tx Ack Mode"
         Height          =   252
         Index           =   18
         Left            =   -74640
         TabIndex        =   330
         ToolTipText     =   "18"
         Top             =   6240
         Width           =   2412
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "RF Request Ack Mode"
         Height          =   252
         Index           =   17
         Left            =   -74640
         TabIndex        =   329
         ToolTipText     =   "17"
         Top             =   5760
         Width           =   2292
      End
      Begin VB.Label lblMicroReg 
         Caption         =   "Rx Filter Mode"
         Height          =   252
         Index           =   16
         Left            =   -74640
         TabIndex        =   328
         ToolTipText     =   "16"
         Top             =   5280
         Width           =   2292
      End
   End
End
Attribute VB_Name = "TestTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Private Sub cmdRepeatedQuery_Click()
  If QueryTimer.Enabled = True Then
    QueryTimer.Enabled = False
  Else
    QueryTimer.Enabled = True
  End If
End Sub

Private Sub MSComm1_OnComm()

  Select Case MSComm1.CommEvent
    ' Handle each event or error by placing
    ' code below each case statement.
    
    ' This template is found in the Example
    ' section of the OnComm event Help topic
    ' in VB Help.
    
    ' Errors
    Case comEventBreak   ' A Break was received.
    Case comEventCDTO    ' CD (RLSD) Timeout.
    Case comEventCTSTO   ' CTS Timeout.
    Case comEventDSRTO   ' DSR Timeout.
    Case comEventFrame   ' Framing Error.
    Case comEventOverrun ' Data Lost.
    Case comEventRxOver  ' Receive buffer overflow.
    Case comEventRxParity   ' Parity Error.
    Case comEventTxFull  ' Transmit buffer full.
    Case comEventDCB     ' Unexpected error retrieving DCB]
  
    ' Events
    Case comEvCD   ' Change in the CD line.
    Case comEvCTS  ' Change in the CTS line.
    Case comEvDSR  ' Change in the DSR line.
    Case comEvRing ' Change in the Ring Indicator.
    Case comEvReceive ' Received RThreshold # of chars.
      gRxBuffer = gRxBuffer + MSComm1.Input
      
    Case comEvSend ' There are SThreshold number of
                  ' characters in the transmit buffer.
    Case comEvEOF  ' An EOF character was found in the
                  ' input stream.
    End Select

End Sub

Private Sub LookForEndOfFrame()
  
  Dim Length As Integer
  
  Length = Len(gRxBuffer)
  
  gRxMessageLength = 0
      
  If (Length > gcMinMsgLength - 1) Then
        
    ' locate start character
    gStartPosition = InStr(1, gRxBuffer, Chr$(1))
    
    If gStartPosition > 0 Then
      
      ' determine how long the message should be
      gRxMessageLength = Asc(Mid$(gRxBuffer, gStartPosition + 1, 1))
      
    End If
        
  End If
    
    
  ' debug code
  Dim i As Integer
  
  ReDim gRxBufferAsBytes(1 To Length)
  
  For i = 1 To Length
    gRxBufferAsBytes(i) = Asc(Mid$(gRxBuffer, i, 1))
  Next
  ' end debug code
  
  
  If gRxMessageLength > 0 Then
  
    ' have we received the EOT character where it should be ?
    gRxEndOfFramePosition = InStr(gRxMessageLength, gRxBuffer, Chr$(4))
  
    ' if we have then process the message
    If gRxEndOfFramePosition > 0 Then
  
      geNextRxState = RxStates.PreprocessMessage
    
    End If
    
  End If
    
End Sub


Private Sub RxStateMachine()

  Select Case geRxState
  
    Case RxStates.Idle
    
      If Len(gRxBuffer) > 0 Then
        LookForEndOfFrame
      End If
      
    Case RxStates.PreprocessMessage
    
      PreprocessMessage
      
    Case RxStates.ValidateMessage
    
      ValidateMessage
      
    Case RxStates.ProcessMessage
    
      ProcessRxMessage
      
    Case RxStates.InvalidateMessage
    
      InvalidateMessage
      
  End Select
    
End Sub

Private Sub PreprocessMessage()
  
  Dim Length As Integer
  
  Length = Len(gRxBuffer)
     
  ' put the message to be processed into another string
  gRxMessage = Mid$(gRxBuffer, gStartPosition, gRxEndOfFramePosition)
    
  ' keep any characters that could be part of the next message
  If (Length - gRxEndOfFramePosition > 1) Then
    gRxBuffer = Mid$(gRxBuffer, gRxEndOfFramePosition + 1, Length)
  Else
    gRxBuffer = vbNullString
  End If
  
  geNextRxState = RxStates.ValidateMessage
  
  
End Sub

Private Sub ValidateMessage()
  
  Dim ChecksumIndex As Integer
  Dim ActualLength As Integer
  
  ActualLength = Len(gRxMessage)
  
  
  'debug code
  ReDim gRxMessageAsBytes(1 To ActualLength)
  Dim i As Integer
  For i = 1 To ActualLength
    gRxMessageAsBytes(i) = Asc(Mid$(gRxMessage, i, 1))
  Next
  'end debug code
  
  
  If ActualLength <> gRxMessageLength Then
  
    geNextRxState = RxStates.InvalidateMessage
    
  Else
  
    ' subtract 0x80 because all messages from module are > 0x80
    gRxMessageType = Asc(Mid$(gRxMessage, 3, 1)) - &H80
    
    ' check the message type
    If gRxMessageType > gcNumberOfMessageTypes - 1 Or gRxMessageType < 0 Then
    
      geNextRxState = RxStates.InvalidateMessage
      
    Else
      
      ' subtract message overhead and place payload into a string
      gRxMessagePayload = Mid$(gRxMessage, 4, gRxMessageLength - gcMessageOverhead)
      
      ChecksumIndex = gRxMessageLength - 2 + 1
      gRxMessageChecksum = Asc(Mid$(gRxMessage, ChecksumIndex, 1))
    
      Dim CalculatedChecksum As Byte
      
      CalculatedChecksum = CalcMsgChecksum(Left$(gRxMessage, gRxMessageLength - 2))
      
      If CalculatedChecksum <> gRxMessageChecksum Then
        geNextRxState = RxStates.InvalidateMessage
      Else
        geNextRxState = RxStates.ProcessMessage
      End If
      
      If gRxMessageType = gTxMessageType Then
        StopMessageTimeoutTimer
      End If
          
    End If
    
  End If
  
End Sub

Private Sub ProcessRxMessage()
  
  Dim FormattedString As String
  
  FormattedString = FormatAndSendStringToListBox(gRxMessage, gcRxValidMsg, gMsgTypeStrings(gRxMessageType))
  
  UpdateReceiveTab gcRxValidMsg, FormattedString, gMsgTypeStrings(gRxMessageType)
    
  Select Case gRxMessageType
  
    Case QueryVersionMsg
      
      ProcessQueryVersionMsg
      
    Case QueryStatisticsMsg
    
      ProcessQueryStatistics
      
    Case ClearStatisticsMsg
    
      ProcessQueryStatistics
      
    Case QueryMicroRegisterMsg
      
      ProcessReadRegisters
      
    Case TransmitPERTDoneMsg
  
      ProcessPertTransmitDone
      
    Case ReceivePERTDoneMsg
    
      ProcessReceivePertDone
      
    Case SetMicroRegisterMsg, TransmitRfMsg, StaticTestModeMsg, TransmitPERTMsg
    
      ProcessStatusCode
    
  End Select
  
  gRxMsgTypeForPert = gRxMessageType
  gRxMsgTypeForReadRegisters = gRxMessageType
  gRxMsgTypeForSyncTest = gRxMessageType
  
  geNextRxState = RxStates.Idle
      
End Sub

Private Sub InvalidateMessage()

  Dim FormattedString As String
  
  If gRxMessageType < eMessageType.TOTAL_MESSAGE_TYPES And gRxMessageType >= 0 Then
  
    FormattedString = FormatAndSendStringToListBox(gRxMessage, gcRxInvalidMsg, gMsgTypeStrings(gRxMessageType))
    
    UpdateReceiveTab gcRxInvalidMsg, FormattedString, gMsgTypeStrings(gRxMessageType)
  
  Else
  
    FormattedString = FormatAndSendStringToListBox(gRxMessage, gcRxInvalidMsg, "Unknown Message Type")
    
    UpdateReceiveTab gcRxInvalidMsg, FormattedString, "Unknown Message Type"
    
  End If
  
  
  gRxMessage = vbNullString
  geNextRxState = RxStates.Idle

End Sub

Private Sub SetupComm()

  With MSComm1
    '.Handshaking = 2 - comRTS
    .RThreshold = 1
    .RTSEnable = True
    .Settings = "38400,n,8,1"
    .SThreshold = 1
    ' Leave all other settings as default values
  End With
  
End Sub

Private Sub SetCommPort()

  MSComm1.CommPort = gCommPort
  
End Sub
Private Sub OpenComm()

  If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
  End If
    
  ' set the baud rate to the non-standard value of 125000
  SetBaudRate MSComm1, 125000
  
End Sub
Private Sub CloseComm()

  If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
  End If

End Sub


Private Sub optRetry_Click(Index As Integer)
  gPertRetries = Index
End Sub




Private Sub optTxModPattern_Click(Index As Integer)
  gTxModPattern = Index
End Sub

Private Sub QueryTimer_Timer()
  
  StartNewTxMessage (QueryVersionMsg)

  TransmitMessage
  
End Sub

Private Sub RxFsmTimer_Timer()
  
  geRxState = geNextRxState
  
  RxStateMachine
  PertFsm
  ReadRegisterFsm
  SyncTestFsm
  
End Sub

Private Sub cmdSendTestMode_Click()
 
  StartNewTxMessage (StaticTestModeMsg)
  
  ' mode , channel , power , tx modulation pattern
  
  AddByteToTxMessage gTestMode
    
  AddByteToTxMessage gTestModeChannel
  
  AddByteToTxMessage gTestModePower
  
  AddByteToTxMessage gTxModPattern
  
  TransmitMessage
  
End Sub

Sub StartNewTxMessage(MsgType As Integer)
  
  ' start and length bytes will be added later
  gTxMessage = ""
  
  ' message type
  gTxMessage = gTxMessage + Chr$(MsgType)
  
  ' set global varaible
  gTxMessageType = MsgType
  
End Sub

Sub TransmitMessage()

  Dim Length As Integer
  
  ' calculate length pluse the bytes that are going to be added next
  Length = Len(gTxMessage) + 4
  
  ' add start and length
  gTxMessage = Chr$(1) + Chr$(Length) + gTxMessage
  
  ' calculate and add checksum
  gTxMessage = gTxMessage + Chr$(CalcMsgChecksum(gTxMessage))
  
  ' stop character
  gTxMessage = gTxMessage + Chr$(4)
  
  Dim FormattedString As String
  
  FormattedString = FormatAndSendStringToListBox(gTxMessage, gcTransmitMsg, gMsgTypeStrings(gTxMessageType))
  
  ' now put the message into the Raw Transmit Byte Boxes
  If chkCopyTx.Value = Checked Then
  
    ClearTxBoxes
  
    Dim i As Byte
  
    For i = 0 To Length - 1
  
      txtByte(i).Text = Trim$(Mid$(FormattedString, (i * 3) + 1, 3))
  
    Next
    
    optMsgType(gTxMessageType).Value = True
      
  End If
    
  
  ' send to port
  MSComm1.Output = gTxMessage
  
  ' start timeout timer
  StartMessageTimeoutTimer
  
End Sub

Private Sub cmdQueryVersion_Click()
  
  StartNewTxMessage (QueryVersionMsg)

  TransmitMessage
    
End Sub

Private Sub ProcessQueryVersionMsg()

  ' take byte from string
  ' convert it to a number
  ' convert it back to a string
  ' remove any leading spaces
  
  ' update the labels with the version information
  
  lblPcbVersion = gcPcbVersionString + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 1, 1))))
  
  lblIcPartNumber = gcIcPartNumberString + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 2, 1))))
  
  lblIcVersionNumber = gcVersionNumberString + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 3, 1))))
  
  lblAppFirmwareVersion = gcAppFirmwareVersionString + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 4, 1))))
  lblAppFirmwareVersion = lblAppFirmwareVersion + "." + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 5, 1))))
  
  lblAppFirmwareDate = gcAppFirmwareDateString + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 6, 1))))
  lblAppFirmwareDate = lblAppFirmwareDate + "-" + LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 7, 1))))
  
  Dim Year As String
  Year = LTrim$(Str$(Asc(Mid$(gRxMessagePayload, 8, 1))))
  
  ' handle case when year is < 10
  If Len(Year) < 2 Then
    Year = "0" + Year
  End If
  
  lblAppFirmwareDate = lblAppFirmwareDate + "-20" + Year
  
  ' discard copyright string length (index 9)
  
  lblAppFirmwareDescription = gcAppFirmwareDescription + Mid$(gRxMessagePayload, 10, gRxMessageLength - 5 - 9)
    
End Sub

Public Function CalcMsgChecksum(sMsg As String) As Byte
    
  Dim Checksum As Integer
  Dim i As Integer
  
  Checksum = 0

  For i = 1 To Len(sMsg)
    Checksum = Checksum + Asc(Mid$(sMsg, i, 1))
  Next
      
  CalcMsgChecksum = Checksum Mod 256
    
End Function



Sub Form_Load()

  Dim i As Integer
  
  TestTool.Caption = gcCaptionString + gcVersionString
  
  ReDim gBytes(txtByte.lbound To txtByte.ubound)
  
  ListBox.Clear
  
  'get message type strings
  gMsgTypeStrings() = Split(gcMessageTypesBigString, ",")
  
  ' valid hexadecimal characters
  gValidHexChars() = Split(gcValidHexadecimalCharacters, ",")
  
  ' minimum payload sizes for transmit messages
  gMinTxPayloadSize() = Split(gcMinimumTxPayloadSize, ",")
  
  'gRegisterNameStrings() = Split(gcRegisterNamesBigString, ",")

  'initialize labels
  ClearVersion
  
  InitStatistics
  
  'register page
  optByteCount(0) = True
  optMsgType(0) = True
  
  ClearReceiveTab
  
  InitPertTxTab

  InitRegisterTab
  
  InitRawTxTab
  
  InitTestModeTab
  
  InitRawReceiveTab
  
  gCommPort = GetSetting(App.Title, "Properties", "COMPort", 1)
  
  UpdateComm
  
  cmdGetBaudRate_Click
  
  chkAppendLog.Value = GetSetting(App.Title, "Properties", "AppendLog", Unchecked)
  
  OpenLog
  
End Sub
  
Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Properties", "AppendLog", chkAppendLog.Value
  CloseLog
End Sub

Private Sub InitTestModeTab()
  
  ' set a default for the options
  Option1(Option1.lbound) = True
  Option2(Option2.lbound) = True
  Option3(Option3.lbound) = True
  optTxModPattern(0) = True
  
  optModuleMode(0) = True
  
  optSyncTestState(0) = True
  
  gSyncTestEnable = False
  gSyncTestState = SyncTestIdleState
  
End Sub
  
Private Sub InitRawTxTab()

  ClearTxBoxes
  
End Sub

Private Sub ClearTxBoxes()
    
  Dim i As Integer
  
  For i = txtByte.lbound To txtByte.ubound
    txtByte(i).Text = 0
    optByteCount(i).Caption = i + 1
  Next
  
End Sub
  
Private Sub cmdRescanCommPorts_Click()
  UpdateComm
End Sub
  
Private Sub UpdateComm()
    
  CloseComm
  
  ' Enable error handler.
  On Error GoTo PortErrorHandler
  
  Dim CommCount As Integer
    
  ' Display current settings.
  For CommCount = 1 To gcNumberOfComPorts Step 1
    Option4(CommCount - 1).Enabled = True
    MSComm1.CommPort = CommCount
    MSComm1.PortOpen = True
    MSComm1.PortOpen = False
  Next
  
  ' Disable error handler.
  On Error Resume Next
      
  ' Set the option button
  If Option4(gCommPort - 1).Enabled = True Then
  
    Option4(gCommPort - 1).Value = True
  
  Else
  
      ' select the first open port
      
      For CommCount = 1 To gcNumberOfComPorts Step 1
        
        If Option4(CommCount - 1).Enabled = True Then
          gCommPort = CommCount
          Option4(gCommPort - 1).Value = True
          Exit For
        End If
          
      Next
    
  End If
    
  SetupComm
    
  SetCommPort
  
  OpenComm
    
Exit Sub

PortErrorHandler:
  Option4(CommCount - 1).Enabled = False
  Resume Next

End Sub

Private Sub cmdSelectNewCommPort_Click()
      
  Dim CommCount As Integer
  Dim OldCommPort As Integer
  
  ' disable error handler
  On Error Resume Next
  
  CloseComm
  
  OldCommPort = gCommPort
  
  ' get comm port value from form
  For CommCount = 1 To gcNumberOfComPorts
  
    If Option4(CommCount - 1).Value = True Then
      gCommPort = CommCount
    End If
  
  Next
      
  SetCommPort
  
  OpenComm
  
  If Err Then
    
    MsgBox Error$, vbOKOnly, "Invalid Comm Port Selected"
  
  Else
    
    MsgBox Error$, vbOKOnly, "Valid Comm Selected"
    
    ' Save Registry Settings.
    SaveSetting App.Title, "Properties", "COMPort", gCommPort
  
  End If
  
End Sub
   
'the list box index starts at Zero
Private Sub AddOutputToList(lbListBox As ListBox, sNewItem As String)

  Dim iListBoxMessages As Integer

  'If the list box is full, then remove one
  iListBoxMessages = lbListBox.ListCount
  
  ' limit the size of the list box to 120
  If iListBoxMessages > 120 - 1 Then
    lbListBox.RemoveItem iListBoxMessages - 1
  End If

  lbListBox.AddItem sNewItem, 0
  
  ' "focus" at bottom
  'lbListBox.TopIndex = lbListBox.ListCount - 1
  
  ' focus at top of list box
  lbListBox.TopIndex = 0
  
  ' add a horizontal scrollbar to the list box window
  Static LongestLine As Long
  Dim Length As Long
  
  Length = TextWidth(sNewItem)
  
  If Length > LongestLine Then
    
    LongestLine = Length
    
    ' if twips change to pixels
    If ScaleMode = vbTwips Then
      Length = Length / Screen.TwipsPerPixelX
    End If
    
    SendMessageByNum lbListBox.hwnd, LB_SETHORIZONTALEXTENT, Length, 0
    
  End If
    
End Sub

Private Sub ListBoxClear_Click()
  ListBox.Clear
End Sub

' This prints out the received message
' The comment string identifies things like valid/invalid message
' The Type message is used to show what kind of message was received

Private Function FormatAndSendStringToListBox(RawSerialMsg As String, Comment As String, TypeMsg As String) As String
    
    Dim i As Integer
    Dim TempString As String
    Dim FinalString As String
    
    ' build the message in hexadecimal strings of each byte
    For i = 1 To Len(RawSerialMsg)
      
      TempString = Mid$(RawSerialMsg, i, 1)
      
      TempString = Hex$(Asc(TempString))
      
      ' add leading zero if required
      ' otherwise keep first things first
      If (Len(TempString) < 2) Then
        FinalString = FinalString + "0" + TempString + " "
      Else
        FinalString = FinalString + TempString + " "
      End If
      
    Next

    Dim ListBoxString As String
    
    ListBoxString = Comment + " - " + FinalString + " ( " & TypeMsg & " )"
    
    AddOutputToList ListBox, ListBoxString
    
        
    If chkAppendLog.Value = Checked Then
      AppendLog ListBoxString
    End If
    
    FormatAndSendStringToListBox = FinalString
    
End Function

Private Sub StartMessageTimeoutTimer()
    tmrMsgTimeout.Interval = 500
    tmrMsgTimeout.Enabled = True
End Sub

Private Sub StopMessageTimeoutTimer()
    tmrMsgTimeout.Enabled = False
End Sub



Private Sub tmrMsgTimeout_Timer()

  StopMessageTimeoutTimer
  AddOutputToList ListBox, "Rx Message Timeout ( " & gMsgTypeStrings(gTxMessageType) & " )"
  lblReceiveMsgValidity = gcReceiveMsgValidityString + "Invalid (Message Timeout)"
  lblReceivedMessageType = gcReceivedMessageTypeString
  
End Sub

Private Sub cmdClearVersion_Click()
  ClearVersion
End Sub

Private Sub ClearVersion()
  
  lblPcbVersion = gcPcbVersionString
  lblIcPartNumber = gcIcPartNumberString
  lblIcVersionNumber = gcVersionNumberString
  lblAppFirmwareVersion = gcAppFirmwareVersionString
  lblAppFirmwareDate = gcAppFirmwareDateString
  lblAppFirmwareDescription = gcAppFirmwareDescription
  
End Sub

Private Sub InitStatistics()

  gStatisticStrings() = Split(gcStatisticStrings, ",")
  
  Dim i As Byte
  
  For i = 0 To gcNumOfStatistics - 1
    lblStatistic(i) = gStatisticStrings(i)
  Next
  
End Sub

Private Sub cmdClearStatistics_Click()

    StartNewTxMessage (ClearStatisticsMsg)
  
    TransmitMessage
  
End Sub

Private Sub cmdQueryStatistics_Click()

    StartNewTxMessage (QueryStatisticsMsg)
  
    TransmitMessage
  
End Sub

Private Sub ProcessQueryStatistics()

  Dim i As Byte
  Dim j As Byte
  Dim PayloadIndex
  Dim txt As String
  
  PayloadIndex = 1
  
  ' check that there are enough bytes in the message
  If Len(gRxMessagePayload) = 32 Then
     
    For i = 0 To gcNumOfStatistics - 1
      
      txt = ""
      
      For j = 0 To 3
        txt = Pad(Hex$(Asc(Mid$(gRxMessagePayload, PayloadIndex, 1))), 2) + txt
        PayloadIndex = PayloadIndex + 1
      Next
        
      lblStatistic(i) = gStatisticStrings(i) + txt
      
    Next
    
  End If
   
End Sub

Private Function Pad(BytesAsString As String, DesiredSize As Byte) As String

  Dim result As String
  Dim Length As Byte
  Dim i As Byte
  
  result = ""
  Length = Len(BytesAsString)
  
  If Length > DesiredSize Then
  
    result = Right$(BytesAsString, DesiredSize)
    
  Else
  
    For i = 1 To DesiredSize - Length
    
      result = result + "0"
      
    Next
    
    result = result + BytesAsString
    
  End If
    
  Pad = result
  
End Function


' validate that the characters entered are hexadecimal

Private Sub txtByte_Change(Index As Integer)

  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtByte(Index).Text))
  
  valid = ValidateHexString(txt, 2)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  If valid = False Then
    
    txtByte(Index).BackColor = vbRed
  
  Else
  
    txtByte(Index).BackColor = vbGreen
    
  End If
  
End Sub

  
Private Sub optMsgType_Click(Index As Integer)

  Dim i As Integer
  
  gTxNumberOfBytes = gcMessageOverhead + Val(gMinTxPayloadSize(Index))
  
  ' set the (min) number of bytes for the message
  optByteCount(gTxNumberOfBytes - 1).Value = True
  
  ' set message byte
  txtByte(2).Text = Hex$(Index)
  
  gTxMessageSelected = Index
  
End Sub


Private Sub cmdSendMessage_Click()
  
  Dim i As Integer
  Dim BytesAreValid As Boolean
   
  ' check the byte boxes
  BytesAreValid = True
  
  For i = 1 To gTxNumberOfBytes - 1
  
    If txtByte(i).BackColor = vbRed Then
      BytesAreValid = False
      Exit For
    End If
      
  Next
  
  If BytesAreValid Then
    BuildnSendTxMessage
  Else
    cmdSendMessage.BackColor = vbYellow
  End If
  

End Sub

Private Sub BuildnSendTxMessage()

  Dim i As Integer
  
  gTxMessage = ""
  
  ' build up message using the bytes in the message boxes
  For i = 0 To gTxNumberOfBytes - 1
  
    AddByteToTxMessage gBytes(i)
  
  Next
  
  TransmitRawMessage
  
End Sub

Private Sub optByteCount_Click(Index As Integer)
  gTxNumberOfBytes = Index + 1
End Sub

Private Sub ClearReceiveTab()

  Dim i As Integer
  
  For i = txtRxByte.lbound To txtRxByte.ubound
    txtRxByte(i).Text = " "
  Next
    
  lblReceiveMsgValidity = gcReceiveMsgValidityString
  lblReceivedMessageType = gcReceivedMessageTypeString

End Sub

Private Sub UpdateReceiveTab(Comment As String, Message As String, TypeMsg As String)

  ClearReceiveTab
  
  lblReceiveMsgValidity = gcReceiveMsgValidityString + Comment
  lblReceivedMessageType = gcReceivedMessageTypeString + TypeMsg

  Dim MessageInBytes() As String
  Dim MsgByte As Variant
  Dim Index As Integer
  
  MessageInBytes() = Split(Message, " ")
  Index = 0
  
  For Each MsgByte In MessageInBytes
  
    txtRxByte(Index).Text = MsgByte
    
    Index = Index + 1
    
    If Index > gcNumberOfReceiveTxtBoxes - 1 Then
      Exit For
    End If
    
  Next

End Sub

Function ValidateHexString(HexString As String, Size As Byte) As Boolean

  Dim valid As Boolean
  Dim Index As Integer
  Dim HexChar As Variant
  Dim HexCharFound() As Boolean
    
  ReDim HexCharFound(1 To Size)
  
  valid = True
  
  If Len(HexString) > Size Or Len(HexString) = 0 Then
      
    valid = False
  
  Else
  
    ' handle strings smaller than the size
    For Index = Len(HexString) To Size
      HexCharFound(Index) = True
    Next
      
      
    For Index = 1 To Len(HexString)
    
      HexCharFound(Index) = False
      
      For Each HexChar In gValidHexChars()
       
        If InStr(1, Mid$(HexString, Index, 1), HexChar) > 0 Then
          HexCharFound(Index) = True
          Exit For
        End If
      
      Next
      
    Next
         
  End If
       
  For Index = 1 To Size
  
    valid = valid And HexCharFound(Index)
    
  Next
  
  ValidateHexString = valid
  
End Function

Private Sub InitRegisterTab()

  Dim i As Byte
  
  For i = txtMicroReg.lbound To txtMicroReg.ubound
  
    txtMicroReg(i).Text = "??"
  
  Next
  
End Sub

Private Sub txtMicroReg_Change(Index As Integer)

  Dim txt As String
  Dim valid As Boolean
  Dim Size As Byte
  
  txt = UCase$(Trim$(txtMicroReg(Index)))
  
  Size = GetRegisterSizeinHexChars(Index)
  
  valid = ValidateHexString(txt, Size)
    
  '
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  If valid = False Then
    
    txtMicroReg(Index).BackColor = vbRed
  
  Else
  
    txtMicroReg(Index).BackColor = vbGreen
  
  End If
  
End Sub


Private Sub cmdSetMicroReg_Click(Index As Integer)

  Dim i As Integer
  Dim Length As Integer
  
  ClearStatusLabels
  
  If txtMicroReg(Index).BackColor <> vbRed Then
    
    txtMicroReg(Index).Text = Pad(txtMicroReg(Index).Text, GetRegisterSizeinHexChars(Index))
  
    txtMicroReg(Index).BackColor = vbYellow
    
    StartNewTxMessage (SetMicroRegisterMsg)
   
    ' register address
    AddByteToTxMessage CByte(Index)
  
    ' data (convert from ascii (hex) string to bytes)
    ' send lsbyte first
    Length = Len(txtMicroReg(Index).Text)
    
    For i = 1 To Length Step 2
    
      AddByteToTxMessage CByte("&H" & Mid$(txtMicroReg(Index).Text, Length - i, 2))
      
    Next
    
    TransmitMessage
      
  End If

End Sub

Private Sub cmdGetMicroRegister_Click(Index As Integer)

  StartNewTxMessage (QueryMicroRegisterMsg)
  
  ' register address
  AddByteToTxMessage CByte(Index)
  
  TransmitMessage

End Sub

Private Function GetRegisterSizeinHexChars(ByVal RegisterNumber As Integer) As Byte

  GetRegisterSizeinHexChars = 2 * GetRegisterSizeinBytes(RegisterNumber)

End Function

Private Function GetRegisterSizeinBytes(ByVal RegisterNumber As Integer) As Byte

  Dim RegisterSizeInBytes As Byte
  
  Select Case RegisterNumber
  
    Case PimPanId
      RegisterSizeInBytes = 2
    
    Case SourceAddr
      RegisterSizeInBytes = 2
    
    Case DestAddr
      RegisterSizeInBytes = 2
    
    Case IeeeAddr
      RegisterSizeInBytes = gcIeeeAddressSize
    
    Case EncryptionKey
      RegisterSizeInBytes = gcEncryptionKeySize
    
    Case EncryptionIv
      RegisterSizeInBytes = gcIvSize
    
    Case SequenceNumber
      RegisterSizeInBytes = gc32Size
      
    Case NvalWriteCount
      RegisterSizeInBytes = gc32Size
  
  Case Else
    
    RegisterSizeInBytes = 1
    
  End Select

  GetRegisterSizeinBytes = RegisterSizeInBytes
  
End Function

Private Sub txtPertDestAddr_Change()

  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtPertDestAddr.Text))
  
  valid = ValidateHexString(txt, 4)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  ' if it is OK then update the global variable
  '
  If valid = False Then
    
    txtPertDestAddr.BackColor = vbRed
  
  Else
  
    txtPertDestAddr.BackColor = vbGreen
  
  End If

End Sub

Private Sub txtPertPimId_Change()

  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtPertPimId.Text))
  
  valid = ValidateHexString(txt, 4)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  ' if it is OK then update the global variable
  '
  If valid = False Then
    
    txtPertPimId.BackColor = vbRed
  
  Else
    
    txtPertPimId.BackColor = vbGreen
  
  End If
  
End Sub


Private Sub txtPertNumberOfPackets_Change()
  
  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtPertNumberOfPackets.Text))
  
  valid = ValidateHexString(txt, 4)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  ' if it is OK then update the global variable
  '
  If valid = False Then
    
    txtPertNumberOfPackets.BackColor = vbRed
  
  Else
  
    txtPertNumberOfPackets.BackColor = vbGreen
        
  End If
  
End Sub


Private Sub txtPertPacketSize_Change()

  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtPertPacketSize.Text))
  
  valid = ValidateHexString(txt, 2)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  ' if it is OK then update the global variable
  '
  If valid = False Then
    
    txtPertPacketSize.BackColor = vbRed
  
  Else
  
    txtPertPacketSize.BackColor = vbGreen

  End If
     

End Sub

Private Sub txtPertSourceAddr_Change()
    
  Dim txt As String
  Dim valid As Boolean

  txt = UCase$(Trim$(txtPertSourceAddr.Text))
  
  valid = ValidateHexString(txt, 4)
    
  ' if the value is invalid then change the box color to red
  ' but don't change the value in the box
  '
  ' if it is OK then update the global variable
  '
  If valid = False Then
    
    txtPertSourceAddr.BackColor = vbRed
  
  Else

    txtPertSourceAddr.BackColor = vbGreen
  
  End If
  
End Sub


Private Sub InitPertTxTab()

  gPertPimId = gcPertPimId
  txtPertPimId.Text = Hex$(gcPertPimId) ' f's in front (negative) ??
  
  gPertSourceAddr = gcPertSourceAddr
  txtPertSourceAddr.Text = Hex$(gPertSourceAddr)
  
  gPertDestAddr = gcPertDestAddr
  txtPertDestAddr.Text = Hex$(gPertDestAddr)
  
  gPertNumberOfPackets = gcPertNumberOfPackets
  txtPertNumberOfPackets.Text = Hex$(gPertNumberOfPackets)
  
  gPertPacketSize = gcPertPacketSize
  txtPertPacketSize.Text = Hex$(gPertPacketSize)
  
  ' Transmit
  
  ' 0 = disabled, 1 = enabled
  optRequestAck(0) = True

  optRetry(0).Value = True
  
  ' Receive
  
  ' 0 = use filter reg, 1 = no
  optRfToHost(1).Value = True
  
  ' both
  ClearPertResults
  
End Sub

Private Sub ClearPertResults()

  lblPacketsNotAcked = gcPacketsNotAckedString
  lblPertReceivedPackets = gcPertReceivedPacketsString
  lblPertExpectedPackets = gcPertExpectedPacketsString
  lblPertMatchingPackets = gcPertMatchingPacketsString
  
End Sub

Private Function ValidatePertBoxes() As Boolean
   
  Dim OptionsAreValid As Boolean
    
  ' check the byte boxes
  OptionsAreValid = True
  
  If txtPertPimId.BackColor = vbRed Then
    
    OptionsAreValid = False
  
  Else
        
    gPertPimId = CLng("&H" & txtPertPimId.Text)
  
    txtPertPimId.Text = Pad(Hex$(gPertPimId), 4)

  End If
    
  If txtPertNumberOfPackets.BackColor = vbRed Then
    
    OptionsAreValid = False
  
  Else
    
    gPertNumberOfPackets = CLng("&H" & txtPertNumberOfPackets.Text)
  
    txtPertNumberOfPackets.Text = Pad(Hex$(gPertNumberOfPackets), 4)

    lblPertNumberPacketsDecimal = "= " & gPertNumberOfPackets
    
  End If
    
  If txtPertSourceAddr.BackColor = vbRed Then
    
    OptionsAreValid = False
  
  Else
                 
    gPertSourceAddr = CLng("&H" & txtPertSourceAddr.Text)
  
    txtPertSourceAddr.Text = Pad(Hex$(gPertSourceAddr), 4)
    
  End If
  
  If txtPertDestAddr.BackColor = vbRed Then
    
    OptionsAreValid = False
  
  Else
      
    gPertDestAddr = CLng("&H" & txtPertDestAddr.Text)
  
    txtPertDestAddr.Text = Pad(Hex$(gPertDestAddr), 4)
  
  End If
    
  If txtPertPacketSize.BackColor = vbRed Then
    
    OptionsAreValid = False
  
  Else
  
    gPertPacketSize = CByte("&H" & txtPertPacketSize.Text)
  
    txtPertPacketSize.Text = Pad(Hex$(gPertPacketSize), 2)
  
    lblPertPacketSizeDecimal = "= " & gPertPacketSize
    
    ' check size
    If gPertPacketSize > gcMaxPacketSize Then
      txtPertPacketSize.BackColor = vbRed
      OptionsAreValid = False
    End If

  End If
    
  ' enabled when 0 is true
  If optRequestAck(0).Value = True Then
    lblPacketsNotAcked.FontStrikethru = True
  Else
    lblPacketsNotAcked.FontStrikethru = False
  End If
      
  ValidatePertBoxes = OptionsAreValid
  
End Function

Private Sub cmdTxPert_Click()

  If gPertState = PertIdle Then
  
    If ValidatePertBoxes Then
      
      gPertState = TxRetryMode
      
      ClearPertResults
      
      ' indicate tx busy state
      optPertState(1).Value = True
    End If
    
  End If

End Sub

Private Sub cmdResetPert_Click()
    
  gPertState = Idle
  
  ClearPertResults
  
End Sub

Private Sub cmdRxPert_Click()

  If gPertState = PertIdle Then
    
    If ValidatePertBoxes Then
      
      gPertState = TxPertRxCommand
      
      ClearPertResults
      
      ' indicate rx pert busy
      optPertState(2).Value = True
    
    End If
    
  End If

End Sub

Private Sub PertFsm()

  Dim i As Byte
  
  Select Case gPertState
  
    Case Idle
      
      gPertState = Idle
      
      gRxMsgTypeForPert = InvalidMsg
      
      optPertState(0).Value = True
      
    Case TxRetryMode
      
      StartNewTxMessage (SetMicroRegisterMsg)
   
      ' register address
      AddByteToTxMessage eRegisterNumber.RetryAttempts
  
      ' data
      AddByteToTxMessage gPertRetries
      
      TransmitMessage
      
      gPertState = WaitForRetryModeResponse
      
    Case WaitForRetryModeResponse
    
      If gRxMsgTypeForPert = SetMicroRegisterMsg Then
        gPertState = TxAckMode
        gRxMsgTypeForPert = InvalidMsg
      End If
      
    Case TxAckMode
    
      StartNewTxMessage (SetMicroRegisterMsg)
   
      ' register address
      AddByteToTxMessage eRegisterNumber.RfReqAckMode
  
      ' data
      If optRequestAck(0).Value = True Then
        AddByteToTxMessage 0
      Else
        AddByteToTxMessage 1
      End If
      
      TransmitMessage
      
      gPertState = WaitForAckModeResponse
      
    Case WaitForAckModeResponse
    
      If gRxMsgTypeForPert = SetMicroRegisterMsg Then
        gPertState = TxPertTxCommand
        gRxMsgTypeForPert = InvalidMsg
      End If
      
    Case TxPertTxCommand
    
      BuildnSendTxPert
      
      gPertState = WaitForPertTxResponse
      
    Case WaitForPertTxResponse
      
      If gRxMsgTypeForPert = TransmitPERTMsg Then
        gPertState = WaitForPertDoneResponse
        gRxMsgTypeForPert = InvalidMsg
      End If
          
    Case WaitForPertDoneResponse
    
      If gRxMsgTypeForPert = TransmitPERTDoneMsg Then
        gPertState = PertIdle
        gRxMsgTypeForPert = InvalidMsg
      End If
      
    Case TxPertRxCommand
    
      BuildnSendRxPert
      
      gPertState = WaitForPertRxResponse
      gRxMsgTypeForPert = InvalidMsg
      
    Case WaitForPertRxResponse
    
      If gRxMsgTypeForPert = ReceivePERTMsg Then
        gPertState = WaitForPertRxDoneResponse
        gRxMsgTypeForPert = InvalidMsg
      End If
      
    Case WaitForPertRxDoneResponse
    
      If gRxMsgTypeForPert = ReceivePERTDoneMsg Then
        gPertState = PertIdle
        gRxMsgTypeForPert = InvalidMsg
      End If
      
  End Select
  
End Sub


Private Sub BuildnSendTxPert()

  StartNewTxMessage (TransmitPERTMsg)
   
  AddLsbMsbToTxMessage (gPertPimId)
  AddLsbMsbToTxMessage (gPertSourceAddr)
  AddLsbMsbToTxMessage (gPertDestAddr)
  AddLsbMsbToTxMessage (gPertNumberOfPackets)
  
  Dim i As Byte
  
  For i = 1 To gPertPacketSize
  
    AddByteToTxMessage (i)
  
  Next
  
  TransmitMessage

End Sub

Private Sub BuildnSendRxPert()

  StartNewTxMessage (ReceivePERTMsg)
   
  AddLsbMsbToTxMessage gPertPimId
  AddLsbMsbToTxMessage gPertSourceAddr
  AddLsbMsbToTxMessage gPertDestAddr
  AddLsbMsbToTxMessage gPertNumberOfPackets
  
  If optRfToHost(0).Value = True Then
    
    ' module will use value in filter register
    AddByteToTxMessage 0
    
  Else
  
    ' don't pass messages to host
    AddByteToTxMessage 1
    
  End If
    
  
  Dim i As Byte
  
  For i = 1 To gPertPacketSize
  
    AddByteToTxMessage (i)
  
  Next
  
  TransmitMessage

End Sub


Private Sub AddByteToTxMessage(Msg As Byte)
  gTxMessage = gTxMessage + Chr$(Msg)
End Sub

Private Sub AddLsbMsbToTxMessage(Msg As Long)

  gTxMessage = gTxMessage + Chr$(Msg Mod 256) + Chr$(Msg \ 256)
  
End Sub

Private Sub cmdReadRegisters_Click()
    
  InitRegisterTab
  
  gReadRegisterState = ReadRegisterStart

End Sub

Private Sub ReadRegisterFsm()

  Static RegisterNumber As Byte
  
  Select Case gReadRegisterState
  
    Case ReadRegisterIdle
    
      'gReadRegisterState = ReadRegisterIdle
      
    Case ReadRegisterStart
    
      RegisterNumber = 0
      
      gReadRegisterState = ReadNextRegister
        
    Case ReadNextRegister
    
      If RegisterNumber > eRegisterNumber.TokenSum Then
      
        gReadRegisterState = ReadRegisterIdle
        
      Else
      
        gRxMsgTypeForReadRegisters = InvalidMsg
        
        StartNewTxMessage QueryMicroRegisterMsg
    
        AddByteToTxMessage RegisterNumber
        
        TransmitMessage
          
        gReadRegisterState = ReadRegisterWaitState
        
        RegisterNumber = RegisterNumber + 1
        
      End If
          
      
    Case ReadRegisterWaitState
  
      If gRxMsgTypeForReadRegisters = QueryMicroRegisterMsg Then
      
        gReadRegisterState = ReadNextRegister
        
        gRxMsgTypeForReadRegisters = InvalidMsg
      
      End If
  
  End Select
  
End Sub

Private Sub ProcessReadRegisters()

  Dim RegisterNumber As Byte
  Dim ReportedCount As Byte
  Dim ActualCount As Byte
  
  RegisterNumber = Asc(Mid$(gRxMessagePayload, 1, 1))
  ReportedCount = Asc(Mid$(gRxMessagePayload, 2, 1))

  ActualCount = Len(gRxMessagePayload) - 2
  
  ' validate the length
  
  Dim RegisterSizeInBytes As Byte
  Dim ValidNumberOfBytes As Boolean
  
  ' determine how many bytes the register should have
  RegisterSizeInBytes = GetRegisterSizeinBytes(RegisterNumber)
  
  ' check the number of bytes
  If ReportedCount <> RegisterSizeInBytes Or ActualCount <> RegisterSizeInBytes Then
    ValidNumberOfBytes = False
  Else
    ValidNumberOfBytes = True
  End If
      
  Dim i As Integer
  Dim temp As Byte
  
  temp = Asc(Mid$(gRxMessagePayload, 3, 1))
   
  If ValidNumberOfBytes Then
  
    Select Case RegisterNumber
    
      Case NvalVersion
        gToken.NvalVersion = temp
        
      Case OpMode
        gToken.OpMode = temp
        
      Case DeviceType
        gToken.DeviceType = temp
        
      Case PimPanId
        gToken.PimPanId = CLng(temp) + CLng(256) * CLng(Asc(Mid$(gRxMessagePayload, 4, 1)))
        
      Case IeeeAddr
      
        For i = 0 To gcIeeeAddressSize - 1
          gToken.IeeeAddr(i) = Asc(Mid$(gRxMessagePayload, i + 3, 1))
        Next
          
      Case SourceAddr
        gToken.SourceAddr = CLng(temp) + CLng(256) * CLng(Asc(Mid$(gRxMessagePayload, 4, 1)))
        
      Case DestAddr
     
        gToken.DestAddr = CLng(temp) + CLng(256) * CLng(Asc(Mid$(gRxMessagePayload, 4, 1)))
        
      Case RfPower
        gToken.RfPower = temp
        
      Case Channel
        gToken.Channel = temp
          
      Case ChannelSet
        gToken.ChannelSet = temp
        
      Case Reserved0
        gToken.Reserved0 = temp
        
      Case EncryptionMode
        gToken.EncryptionMode = temp
        
      Case EncryptionKey
      
        For i = 0 To gcEncryptionKeySize - 1
          gToken.EncryptionKey(i) = Asc(Mid$(gRxMessagePayload, i + 3, 1))
        Next

      Case Reserved1
        gToken.Reserved1 = temp
        
      Case Reserved2
        gToken.Reserved2 = temp
        
      Case Reserved3
        gToken.Reserved3 = temp
        
      Case RxFilterMode
        gToken.RxFilterMode = temp
        
      Case RfReqAckMode
        gToken.RfReqAckMode = temp
        
      Case RfTxAckMode
        gToken.RfTxAckMode = temp
        
      Case RetryAttempts
        gToken.RetryAttempts = temp
        
      Case WakeUpMode
        gToken.WakeUpMode = temp
        
      Case WakeUpMsgSize
        gToken.WakeUpMsgSize = temp
        
      Case EncryptionIv
            
        For i = 0 To gcIvSize - 1
          gToken.EncryptionIv(i) = Asc(Mid$(gRxMessagePayload, i + 3, 1))
        Next
        
      Case HeaderSize
      
        gToken.HeaderSize = temp
      
      Case SequenceNumber
        
        For i = 0 To gc32Size - 1
          gToken.SequenceNumber(i) = Asc(Mid$(gRxMessagePayload, i + 3, 1))
        Next
          
      Case StartFrameDelimiter
        gToken.StartFrameDelimiter = temp
      
      Case NvalWriteCount
      
        For i = 0 To gc32Size - 1
          gToken.NvalWriteCount(i) = Asc(Mid$(gRxMessagePayload, i + 3, 1))
        Next
        
      Case Reserved4
        gToken.Reserved4 = temp
      
      Case TokenSum
        gToken.TokenSum = temp
        
    End Select
      
    UpdateMicroRegTextBoxes CLng(RegisterNumber)

  Else
  
    AddOutputToList ListBox, "Invalid Read Count"
   
  End If
  
  
End Sub

'Private Sub InitRegisterLabels()
'
'  Dim i As Byte
'
'  For i = 0 To eRegisterNumber.WakeUpMode
'
'    lblReg(i) = gRegisterNameStrings(i)
'
'  Next
'
'End Sub

Private Sub UpdateMicroRegTextBoxes(RegisterNumber As Integer)

  Dim i As Byte
  Dim Size As Byte
  Dim txt As String
  
  Size = GetRegisterSizeinHexChars(RegisterNumber)
      
  Select Case RegisterNumber
  
    Case NvalVersion
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.NvalVersion), Size)
      
    Case OpMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.OpMode), Size)
      
    Case DeviceType
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.DeviceType), Size)
      
    Case PimPanId
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.PimPanId), Size)
      
    Case IeeeAddr
    
      txt = ""
      
      For i = 0 To gcIeeeAddressSize - 1
        txt = Pad(Hex$(gToken.IeeeAddr(i)), 2) + txt
      Next
        
      txtMicroReg(RegisterNumber).Text = txt
       
    Case SourceAddr
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.SourceAddr), Size)
      
    Case DestAddr
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.DestAddr), Size)
      
    Case RfPower
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.RfPower), Size)
      'lblPower = gToken.RfPower - 11
      
    Case Channel
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Channel), Size)
        
    Case ChannelSet
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.ChannelSet), Size)
      
    Case Reserved0
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Reserved0), Size)
      
    Case EncryptionMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.EncryptionMode), Size)
      
    Case EncryptionKey
    
      txt = ""
      
      For i = 0 To gcEncryptionKeySize - 1
        txt = Pad(Hex$(gToken.EncryptionKey(i)), 2) + txt
      Next
        
      txtMicroReg(RegisterNumber).Text = txt

    Case Reserved1
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Reserved1), Size)
      
    Case Reserved2
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Reserved2), Size)
      
    Case Reserved3
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Reserved3), Size)
      
    Case RxFilterMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.RxFilterMode), Size)
      
    Case RfReqAckMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.RfReqAckMode), Size)
      
    Case RfTxAckMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.RfTxAckMode), Size)
      
    Case RetryAttempts
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.RetryAttempts), Size)
      
    Case WakeUpMode
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.WakeUpMode), Size)
      
    Case WakeUpMsgSize
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.WakeUpMsgSize), Size)
        
    Case EncryptionIv
          
      txt = ""
      
      For i = 0 To gcIvSize - 1
        txt = Pad(Hex$(gToken.EncryptionIv(i)), 2) + txt
      Next
        
      txtMicroReg(RegisterNumber).Text = txt
      
    Case HeaderSize
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.HeaderSize), Size)
    
    Case SequenceNumber
      
      txt = ""
      
      For i = 0 To gc32Size - 1
        txt = Pad(Hex$(gToken.SequenceNumber(i)), 2) + txt
      Next
        
      txtMicroReg(RegisterNumber).Text = txt

    Case StartFrameDelimiter
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.StartFrameDelimiter), Size)
        
    Case NvalWriteCount
              
      txt = ""
      
      For i = 0 To gc32Size - 1
        txt = Pad(Hex$(gToken.NvalWriteCount(i)), 2) + txt
      Next
        
      txtMicroReg(RegisterNumber).Text = txt

      
    Case Reserved4
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.Reserved4), Size)
      
    
    Case TokenSum
      txtMicroReg(RegisterNumber).Text = Pad(Hex$(gToken.TokenSum), Size)
      
      
      
      
      
  End Select
    
    
  ' black means the value has been read successfully
  ' this must be set after the text box is changed
  txtMicroReg(RegisterNumber).BackColor = vbWhite
  
End Sub

Private Sub ProcessPertTransmitDone()

  Dim PacketsNotAcked As Integer
  
  PacketsNotAcked = Asc(Mid$(gRxMessagePayload, 1, 1)) + 256 * Asc(Mid$(gRxMessagePayload, 2, 1))
  
  lblPacketsNotAcked = gcPacketsNotAckedString + Pad(Hex$(PacketsNotAcked), 4)
  
End Sub

Private Sub ProcessReceivePertDone()

  Dim PertReceivedPackets As Long
  Dim PertExpectedPackets As Long
  Dim PertMatchingPackets As Long

  PertReceivedPackets = Asc(Mid$(gRxMessagePayload, 1, 1)) + 256 * Asc(Mid$(gRxMessagePayload, 2, 1))
  PertExpectedPackets = Asc(Mid$(gRxMessagePayload, 3, 1)) + 256 * Asc(Mid$(gRxMessagePayload, 4, 1))
  PertMatchingPackets = Asc(Mid$(gRxMessagePayload, 5, 1)) + 256 * Asc(Mid$(gRxMessagePayload, 6, 1))
  
  lblPertReceivedPackets = gcPertReceivedPacketsString + Hex$(PertReceivedPackets)
  lblPertExpectedPackets = gcPertExpectedPacketsString + Hex$(PertExpectedPackets)
  lblPertMatchingPackets = gcPertMatchingPacketsString + Hex$(PertMatchingPackets)
    
End Sub

' 0 = +11 dBm , 1 = +10 dBm, ... , 22 = -11 dBm
  
Private Sub Option1_Click(Index As Integer)

  gTestModePower = Index
  
End Sub

Private Sub Option2_Click(Index As Integer)
  
  gTestMode = Index
  
End Sub

Private Sub Option3_Click(Index As Integer)
  
  gTestModeChannel = Index + 1
  
End Sub


' don't update the global variables

Private Sub cmdSetPower_Click()

  StartNewTxMessage (SetMicroRegisterMsg)
  
  ' register address
  AddByteToTxMessage CByte(eRegisterNumber.RfPower)
  
  AddByteToTxMessage gTestModePower
  
  TransmitMessage

End Sub

Private Sub cmdSetChannel_Click()
  
  StartNewTxMessage (SetMicroRegisterMsg)
  
  ' register address
  AddByteToTxMessage CByte(eRegisterNumber.Channel)
  
  AddByteToTxMessage gTestModeChannel
  
  TransmitMessage
  
End Sub

Private Sub optModuleMode_Click(Index As Integer)
  gModuleMode = Index
End Sub
 
Private Sub cmdSetModuleMode_Click()
  
  StartNewTxMessage (eMessageType.SetSleepModeMsg)
  
  AddByteToTxMessage gModuleMode
  
  TransmitMessage
  
End Sub

Sub TransmitRawMessage()

  Dim JunkString As String
  
  ' set global variable
  
  If gBytes(2) >= eMessageType.InvalidMsg And gBytes(2) < eMessageType.TOTAL_MESSAGE_TYPES Then
    gTxMessageType = gBytes(2)
    JunkString = FormatAndSendStringToListBox(gTxMessage, gcTransmitRaw, gMsgTypeStrings(gTxMessageType))
  Else
    gTxMessageType = eMessageType.InvalidMsg
    JunkString = FormatAndSendStringToListBox(gTxMessage, gcTransmitRaw, "UnknownMsgType")
  End If
       
  ' send to port
  MSComm1.Output = gTxMessage
  
  ' start timeout timer
  StartMessageTimeoutTimer
  
End Sub

Private Sub cmdValidate_Click()

  Dim i As Byte
  
  For i = 0 To gTxNumberOfBytes - 1
  
    '
    ' if it is OK then update the global variable
    '
    
    If txtByte(i).BackColor = vbGreen Then
    
      txtByte(i).Text = Pad(txtByte(i).Text, 2)
      
      gBytes(i) = CByte("&H" & (txtByte(i).Text))
      
    End If
    
  Next
   
End Sub

Private Sub cmdComputeChecksum_Click()

  Dim Checksum As Integer
  Dim i As Byte
  
  Checksum = 0
  
  ' to number of bytes -2 -1 (zero indexed)
  For i = 0 To gTxNumberOfBytes - 2 - 1
  
    Checksum = Checksum + gBytes(i)
  
  Next
  
  txtByte(gTxNumberOfBytes - 2).Text = Hex$(Checksum Mod 256)
   
  ' update the byte value (because validate bytes may have already been run)
  gBytes(gTxNumberOfBytes - 2) = Checksum Mod 256
       
End Sub

Private Sub cmdUpdateLength_Click()

  txtByte(gcLengthIndex).Text = Hex$(gTxNumberOfBytes)
  
End Sub

Private Sub cmdAddStart_Click()

  txtByte(gcStartIndex).Text = Hex$(1)
  
End Sub

Private Sub cmdAddStop_Click()

  txtByte(gTxNumberOfBytes - 1).Text = Hex$(4)
  
End Sub

Private Sub cmdUpdateType_Click()
  
  txtByte(gcMsgTypeIndex).Text = Hex$(gTxMessageSelected)
  
End Sub

Private Sub cmdClearRawTx_Click()
  ClearTxBoxes
End Sub


Private Sub cmdUpdateAndSend_Click()

  cmdAddStart_Click
  cmdUpdateLength_Click
  cmdUpdateType_Click
  cmdAddStop_Click
  cmdValidate_Click
  cmdComputeChecksum_Click
  cmdSendMessage_Click
  
End Sub


Private Sub cmdGetBaudRate_Click()
  
  Dim BaudRate As Long
  
  BaudRate = GetBaudRate(MSComm1)

  lblCommPortRate = "Comm Port Rate: " + Str$(BaudRate)
  
End Sub

' ****************************************
' The PortOpen property should be set to True before calling.
' May raise the following errors:
' comPortNotOpen the PortOpen property has not been set to True
' comDCBError failed to read current state of the port
' comSetCommStateFailed failed to set new baud rate
Sub SetBaudRate(Com As MSComm, baud As Long)
  
  Dim ComDcb As dcb
  Dim ret As Long
  ' Check port is open
  If Not Com.PortOpen Then
    Err.Raise comPortNotOpen, Com.Name, _
    "Operation valid only when the port is open"
    Exit Sub
  End If
  
  ' Get existing Comm state
  ret = GetCommState(Com.CommID, ComDcb)
  
  If ret = 0 Then
    Err.Raise comDCBError, Com.Name, _
    "Could not read current state of the port"
    Exit Sub
  End If
  
  ' Modify state with new baud rate
  ComDcb.BaudRate = baud
  
  ' Set the new Comm state
  ret = SetCommState(Com.CommID, ComDcb)
  
  If ret = 0 Then
    Err.Raise comSetCommStateFailed, Com.Name, _
    "Could not set port to specified baud rate"
    Exit Sub
  End If
  

End Sub


' Get baud rate using Win32 API
' The PortOpen property should be set to True before calling.
' May raise the following errors:
' comPortNotOpen the PortOpen property has not been set to True
' comDCBError failed to read current state of the port
Function GetBaudRate(Com As MSComm) As Long
  
  Dim ComDcb As dcb
  Dim ret As Long

  GetBaudRate = 0

  ' Check port is open
  If Not Com.PortOpen Then
    Err.Raise comPortNotOpen, Com.Name, _
    "Operation valid only when the port is open"
  Exit Function
  End If

  ' Get Comm state
  ret = GetCommState(Com.CommID, ComDcb)
  
  If ret = 0 Then
    Err.Raise comDCBError, Com.Name, _
    "Could not read current state of the port"
    Exit Function
  End If
  
  ' Extract baud rate
  GetBaudRate = ComDcb.BaudRate

End Function

Private Sub cmdStartStopSyncTest_Click()
  
  If gSyncTestEnable Then
    gSyncTestEnable = False
  Else
    gSyncTestEnable = True
  End If
  
End Sub


Private Sub SyncTestFsm()

  ' reset condition
  If gSyncTestEnable = False Then
    gSyncTestState = SyncTestIdleState
    optSyncTestState(0) = True
  End If
    
  Select Case gSyncTestState
    
      Case SyncTestIdleState
  
        If gSyncTestEnable Then
          gSyncTestState = SyncTestTxState
          optSyncTestState(1) = True
        End If
        
        gRxMsgTypeForSyncTest = InvalidMsg
        
      Case SyncTestTxState
      
        SyncTestTx
        
        gSyncTestState = SyncTestWaitState
        
      Case SyncTestWaitState
      
        If gRxMsgTypeForSyncTest = TransmitRfMsg Then
          gSyncTestState = SyncTestIdleState
        End If
    
  End Select

End Sub

Private Sub SyncTestTx()

  StartNewTxMessage (TransmitRfMsg)
   
  ' use default panid, source, and destination
  AddByteToTxMessage (0)
  AddByteToTxMessage (0)
  
  AddByteToTxMessage (0)
  AddByteToTxMessage (0)

  AddByteToTxMessage (0)
  AddByteToTxMessage (0)

  ' three payload bytes
  AddByteToTxMessage (&H11)
  AddByteToTxMessage (&H22)
  AddByteToTxMessage (&H33)

  TransmitMessage

End Sub

Private Sub InitRawReceiveTab()

  ClearStatusLabels
  lblStatusCodeFailCount = gcStatusCodeFailCountString
  
  gStatusCode = eStatusCodes.Success
  gStatusCodeFailCount = 0
  
End Sub
  
' get the status code from the message and display it
' may want to put code in log
Private Sub ProcessStatusCode()
  
  ' make partially compatible with old code
  If Len(gRxMessagePayload) > 1 Then
    
    gStatusCode = Asc(Mid$(gRxMessagePayload, 1, 1))
  
    UpdateStatusCodeLabels
  
  End If
    
End Sub
    
' only clear the ones that are cleared each time the set button is pressed
Private Sub ClearStatusLabels()
  
  lblStatusCode = gcStatusCodeString
  lblStatusCodeDescription = gcStatusCodeDescriptionString
  
End Sub
    
Private Sub UpdateStatusCodeLabels()

  Dim StatusMsg As String
  
  lblStatusCode = gcStatusCodeString + Pad(Hex$(gStatusCode), 2)
  
  StatusMsg = GetStatusCodeDescriptionString
  lblStatusCodeDescription = gcStatusCodeDescriptionString + StatusMsg
  
  If gStatusCode <> eStatusCodes.Success Then
    gStatusCodeFailCount = gStatusCodeFailCount + 1
    AddOutputToList ListBox, "StatusCode - 0x" + Pad(Hex$(gStatusCode), 2) + " ( " + StatusMsg + " )"
  End If
  
  lblStatusCodeFailCount = gcStatusCodeFailCountString + Pad(Hex$(gStatusCodeFailCount), 4) + _
                            " (" + Str$(gStatusCodeFailCount) + " )"
    
End Sub

Private Sub cmdClearStatusCodeFailCount_Click()
  
  gStatusCodeFailCount = 0
  
  lblStatusCodeFailCount = gcStatusCodeFailCountString + Pad(Hex$(gStatusCodeFailCount), 4) + _
                            " (" + Str$(gStatusCodeFailCount) + " )"

End Sub

Function GetStatusCodeDescriptionString() As String

  Dim result As String
  
  Select Case gStatusCode
  
    Case Success
      result = "Success"
      
    'Case BeaconTxSSuccess
    'Case TrxAsleep
    'Case TrxAwake
    'Case CrcCorrect
    'Case CrcIncorrect
    
    Case Failure
      result = "Failure"
      
    Case Busy
      result = "Busy"
      
    'Case TalFramePending
    'Case AlreadyRuning
    'Case NotRunning
    'Case InvalidId
    'Case InvalidTimeout
    
    Case InvalidParameter
      result = "Invalid Parameter"
      
    Case QueueFull
      result = "Queue Full"
      
    'Case CsmaCaInProgress
    'Case NoFrameTransmission
    
    Case ChannelAccessFailure
      result = "Channel Access Failure"
      
    Case NoAck
      result = "No Ack"
      
    'Case UnsupportedAttribute
     
  Case Else
  
    result = "Unknown Status Code"
  
  End Select

  GetStatusCodeDescriptionString = result
  
End Function
  
''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Log File
''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub OpenLog()

  Dim FileName As String
  FileName = "c:\TestToolLog.txt"
  
  ' get next free file number
  gFnumLog = FreeFile()
  
  Open FileName For Append As #gFnumLog
      
End Sub

Private Sub AppendLog(ByVal Line As String)

  Print #gFnumLog, Line
  
End Sub

Private Sub CloseLog()
  
  Close #gFnumLog
  
End Sub

