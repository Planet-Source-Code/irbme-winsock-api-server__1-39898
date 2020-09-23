VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarFriendlyName    As String
Private mvarUserName        As String
Private mvarPingTime        As Long
Private mvarUpTimeSeconds   As Integer
Private mvarUpTimeMinutes   As Integer
Private mvarUpTimeHours     As Long
Private mvarRemotePort      As Integer
Private mvarRemoteHost      As String
Private mvarRemoteIP        As String
Private mVarSocketHandle    As Integer
Private mVarLogOnTime       As String


Public Property Get UpTimeHours() As Long
    UpTimeHours = mvarUpTimeHours
End Property



Private Property Let UpTimeMinutes(ByVal vData As Integer)
    mvarUpTimeMinutes = vData
    mvarUpTimeHours = mvarUpTimeHours + (UpTimeMinutes \ 60)
    mvarUpTimeMinutes = UpTimeMinutes - (UpTimeMinutes \ 60) * 60
End Property

Public Property Get UpTimeMinutes() As Integer
    UpTimeMinutes = mvarUpTimeMinutes
End Property



Public Property Let UpTimeSeconds(ByVal vData As Integer)
    mvarUpTimeSeconds = vData
    UpTimeMinutes = mvarUpTimeMinutes + (mvarUpTimeSeconds \ 60)
    mvarUpTimeSeconds = mvarUpTimeSeconds - (mvarUpTimeSeconds \ 60) * 60
End Property

Public Property Get UpTimeSeconds() As Integer
    UpTimeSeconds = mvarUpTimeSeconds
End Property



Public Property Let PingTime(ByVal vData As Long)
    mvarPingTime = vData
End Property

Public Property Get PingTime() As Long
    PingTime = mvarPingTime
End Property


Public Property Let RemoteHost(ByVal vData As String)
    mvarRemoteHost = vData
End Property

Public Property Get RemoteHost() As String
    RemoteHost = mvarRemoteHost
End Property



Public Property Let RemoteIP(ByVal vData As String)
    mvarRemoteIP = vData
End Property

Public Property Get RemoteIP() As String
    RemoteIP = mvarRemoteIP
End Property



Public Property Let RemotePort(ByVal vData As String)
    mvarRemotePort = vData
End Property

Public Property Get RemotePort() As String
    RemotePort = mvarRemotePort
End Property



Public Property Get SocketHandle() As Integer
    SocketHandle = mVarSocketHandle
End Property

Public Property Let SocketHandle(ByVal vData As Integer)
    mVarSocketHandle = vData
End Property



Public Property Get LogOnTime() As String
    LogOnTime = mVarLogOnTime
End Property

Public Property Let LogOnTime(ByVal vData As String)
    mVarLogOnTime = vData
End Property

