VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fServer 
   Caption         =   "SocketWise Server"
   ClientHeight    =   6480
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11775
   Icon            =   "fServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "All Commands"
      Height          =   2115
      Left            =   10080
      TabIndex        =   7
      Top             =   3255
      Width           =   1590
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   105
         ScaleHeight     =   1695
         ScaleWidth      =   1275
         TabIndex        =   8
         Top             =   210
         Width           =   1275
         Begin VB.CommandButton cmdPingAll 
            Caption         =   "Ping"
            Height          =   435
            Left            =   105
            TabIndex        =   10
            Top             =   735
            Width           =   1065
         End
         Begin VB.CommandButton cmdDisconnectAll 
            Caption         =   "Disconnect"
            Height          =   435
            Left            =   105
            TabIndex        =   9
            Top             =   210
            Width           =   1065
         End
      End
   End
   Begin VB.Frame frmServer 
      Caption         =   "Server Commands"
      Height          =   2115
      Left            =   8400
      TabIndex        =   2
      Top             =   3360
      Width           =   1590
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   105
         ScaleHeight     =   1695
         ScaleWidth      =   1275
         TabIndex        =   3
         Top             =   210
         Width           =   1275
         Begin VB.CommandButton cmdSvrConfigure 
            Caption         =   "Configure"
            Height          =   435
            Left            =   105
            TabIndex        =   6
            Top             =   1260
            Width           =   1065
         End
         Begin VB.CommandButton cmdSvrConnect 
            Caption         =   "Connect"
            Height          =   435
            Left            =   105
            TabIndex        =   5
            Top             =   735
            Width           =   1065
         End
         Begin VB.CommandButton cmdSvrDisconnect 
            Caption         =   "Disconnect"
            Height          =   435
            Left            =   105
            TabIndex        =   4
            Top             =   210
            Width           =   1065
         End
      End
   End
   Begin VB.Timer tmrPing 
      Interval        =   60000
      Left            =   840
      Top             =   3360
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   315
      Top             =   3360
   End
   Begin VB.TextBox txtLog 
      ForeColor       =   &H80000007&
      Height          =   3165
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3150
      Width           =   8115
   End
   Begin MSComctlLib.ListView lstClients 
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblUpTime 
      Caption         =   "Up Time:"
      Height          =   330
      Left            =   8505
      TabIndex        =   12
      Top             =   5985
      Width           =   2955
   End
   Begin VB.Label lblLocalHost 
      Caption         =   "Local Host:"
      Height          =   330
      Left            =   8505
      TabIndex        =   11
      Top             =   5565
      Width           =   2955
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPing 
         Caption         =   "Ping"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "fServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub InitCommonControls Lib "comctl32" ()


Private WithEvents Server   As clsServer        'Server class
Attribute Server.VB_VarHelpID = -1
Private WithEvents Ping     As clsPing          'Ping class
Attribute Ping.VB_VarHelpID = -1
Private EndPoint            As clsEndPoint      'End point class
Private Clients             As Collection       'Collection of clients
Private ListeningSocket     As Long             'Listening socket
Private FileNum             As Integer          'File number for file access
Private ServerTime          As clsUser          'Used for the server uptime
Public ServerPort           As Integer          'The port to listen on



Private Sub cmdDisconnectAll_Click()
'Disconnect all clients from the server.
  
  Dim Dummy As clsUser
    
    'Loop through each client
    For Each Dummy In Clients
        'Lets just pretend they disconnected from us, it's the same code anyhow
        Call Server_OnClose(Dummy.SocketHandle)
    Next
    
    'Reset the collection
    Set Clients = New Collection
    
End Sub


Private Sub cmdPingAll_Click()
'Ping all the clients in the list

  Dim Dummy As clsUser
    
    'Loop through each client
    For Each Dummy In Clients
        'Call the ping command
        Ping.ICMPPing Dummy.RemoteIP
    Next
    
End Sub


Private Sub cmdSvrConfigure_Click()
'Show the configuration form
    fConfig.Show vbApplicationModal
End Sub


Private Sub cmdSvrConnect_Click()
'Connect the server
    ConnectServer
End Sub


Private Sub cmdSvrDisconnect_Click()
'Disconnect the server
    DisconnectServer
End Sub


Private Sub Form_Load()
'For Windows XP users
    Call InitCommonControls
    
    'Create new instances of all classes
    Set Server = New clsServer
    Set ServerTime = New clsUser
    Set Ping = New clsPing
    Set EndPoint = New clsEndPoint
    Set Clients = New Collection
    
    'Set up the column headers for the list view
    lstClients.ColumnHeaders.Add 1, , "Socket Handle"
    lstClients.ColumnHeaders.Add 2, , "Remote Host"
    lstClients.ColumnHeaders.Add 3, , "Remote IP"
    lstClients.ColumnHeaders.Add 4, , "Remote Port"
    lstClients.ColumnHeaders.Add 5, , "Local Port"
    lstClients.ColumnHeaders.Add 6, , "Ping"
    lstClients.ColumnHeaders.Add 7, , "Log on time"
    lstClients.ColumnHeaders.Add 8, , "Up Time"
    
    Me.Show: Me.Refresh
         
    ServerPort = 4000
    
    'Start up the server
    ListeningSocket = Server.CreateSocket()
    Server.Listen ListeningSocket, CLng(ServerPort)
    
    FileNum = FreeFile()
    
    'Open a log file and start logging
    Open App.Path & "\Log - " & Replace(Date, "/", " ") & ".txt" For Output As #FileNum
    Print #1, vbCrLf & "NEW SESSION STARTED AT " & Time & " ON " & Date & vbCrLf
    
    'Display some information
    AddLog "SocketWise Server is starting" & vbCrLf
    AddLog "* Listening Socket Handle: " & ListeningSocket
    AddLog "* Server listening for connections on port " & ServerPort
    
    'Local host name
    lblLocalHost.Caption = "Local Host Name: " & EndPoint.GetLocalHost(ListeningSocket)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

  Dim Dummy         As clsUser
    
    'Unload all clients
    For Each Dummy In Clients
        Server.CloseSocket Dummy.SocketHandle
        Set Dummy = Nothing
    Next
    
    'Clean up the collection
    Set Clients = Nothing
    
    'Close the listening socket
    Server.CloseSocket ListeningSocket
    
    'Free all classes
    Set Server = Nothing
    Set ServerTime = Nothing
    Set Ping = Nothing
    Set EndPoint = Nothing
    
    'Close the log file
    Close #FileNum
    
    Unload fConfig
    
End Sub


Private Sub lstClients_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Show the pop up menu
    If Button = 2 Then PopupMenu mnuPopup, , x, y
End Sub


Private Sub mnuDisconnect_Click()
'Disconnect the selected client

    'Make sure one is selected first
    If ObjPtr(lstClients.SelectedItem) Then
        
        'Pretend they closed it
        Server_OnClose CLng(lstClients.SelectedItem)
    End If
    
End Sub


Private Sub mnuPing_Click()
'Ping the selected client

  Dim Header As ListItem
    
    'Make sure one is selected first
    If ObjPtr(lstClients.SelectedItem) Then
        Set Header = lstClients.SelectedItem
        
        'Ping their IP
        Ping.ICMPPing Header.SubItems(2)
    End If
    
End Sub


Private Sub Ping_OnError(lngRetCode As Long, strDescription As String)
'Ping error - Probably an unresolvable host - Just add it to the log
    AddLog "* PING ERROR " & lngRetCode & " - " & strDescription
End Sub


Private Sub Ping_OnPingResponce(lngRoundTripTime As Long, strIPAddress As String)
'Ping response

  Dim Dummy As clsUser
    
    'Loop through each client
    For Each Dummy In Clients
        'Find the one that matches the returned IP
        If Dummy.RemoteIP = strIPAddress Then
            'Update the ping time
            Dummy.PingTime = lngRoundTripTime
            Exit For
        End If
    Next
    
    'Update the list
    Call RefreshList
    
End Sub


Private Sub Server_OnClose(lngSocket As Long)
'A client disconnected (or we call this once or twice to disconnect them ;D)

    'Log this event
    AddLog "* " & Clients(Str(lngSocket)).RemoteIP & " Disconnected"
    AddLog "* Closing Socket Handle " & Clients(Str(lngSocket)).SocketHandle
            
    'Close the socket and remove the client
    Clients.Remove Str(lngSocket)
    Server.CloseSocket lngSocket
    
    'Refresh the list view control
    Call RefreshList

End Sub


Private Sub Server_OnConnectRequest(lngSocket As Long)
'A client is trying to connect to the server

  Dim NewSocket As Long
  Dim NewUser   As clsUser
  
    'Accept the connection
    NewSocket = Server.Accept(ListeningSocket)
    
    Set NewUser = New clsUser
    
    'Fill in the new data - See how much we can get once they are connected.
    NewUser.RemoteHost = EndPoint.GetRemoteHost(NewSocket)
    NewUser.RemoteIP = EndPoint.GetRemoteIP(NewSocket)
    NewUser.RemotePort = EndPoint.GetRemotePort(NewSocket)
    
    NewUser.LogOnTime = Str(Time)
    NewUser.UpTimeSeconds = 0
    
    Ping.ICMPPing NewUser.RemoteIP
    
    NewUser.SocketHandle = NewSocket
    
    'Add this event to the log
    AddLog "* Connection request on port " & ServerPort & " from " & NewUser.RemoteIP
    AddLog "* New socket created. Handle: " & NewUser.SocketHandle
    AddLog "* Socket connected"

    'Add the user to the collection and refresh the listview
    Clients.Add NewUser, Str(NewUser.SocketHandle)
    Call RefreshList
    
End Sub


Private Sub RefreshList()
'Refresh the listview control

  Dim Dummy         As clsUser
  Dim ListHeader    As ListItem
      
    lstClients.ListItems.Clear
      
    'Loop through each client and add all the details to the contorl
    For Each Dummy In Clients
        Set ListHeader = lstClients.ListItems.Add(, , Dummy.SocketHandle)
        ListHeader.SubItems(1) = Dummy.RemoteHost
        ListHeader.SubItems(2) = Dummy.RemoteIP
        ListHeader.SubItems(3) = Dummy.RemotePort
        ListHeader.SubItems(4) = EndPoint.GetLocalPort(CLng(Dummy.SocketHandle))
        ListHeader.SubItems(5) = Dummy.PingTime
        ListHeader.SubItems(6) = Dummy.LogOnTime
        ListHeader.SubItems(7) = Dummy.UpTimeHours & ":" & Dummy.UpTimeMinutes & ":" & Dummy.UpTimeSeconds
    Next

End Sub


Private Sub AddLog(strText As String)
'Add a log
    
    txtLog.SelStart = Len(txtLog.Text)              'Jump to the end
    txtLog.Text = txtLog.Text & strText & vbCrLf    'Add the text
    txtLog.SelStart = Len(txtLog.Text)              'Jump to the end
    
    If Len(txtLog.Text) > 5000 Then                 'If we have over 5000 chars
        txtLog.Text = Right(txtLog.Text, 1000)      'Trim off old logs
    End If
    
    'If the log check box is checked, log to the file
    If fConfig.chkLog.Value = vbChecked Then
        Print #FileNum, strText
    End If
    
End Sub


Private Sub Server_OnDataArrive(lngSocket As Long)
'Data has arrived

  Dim strData   As String
  Dim BytesRead As Long
  Dim Dummy     As clsUser

      
    BytesRead = Server.Recv(lngSocket, strData)
    AddLog BytesRead & " bytes read from " & Clients(Str(lngSocket)).RemoteIP & " - " & strData
    
    'Loop through each client and relay on the data
    For Each Dummy In Clients
        Server.Send CLng(Dummy.SocketHandle), strData
    Next
    
End Sub


Private Sub Server_OnError(lngRetCode As Long, strDescription As String)
'Server error - Add it to the log for now
    AddLog "* SERVER ERROR " & lngRetCode & " - " & strDescription
End Sub


Private Sub tmrPing_Timer()

  Dim Dummy As clsUser
    
    'Ping each client
    For Each Dummy In Clients
        Ping.ICMPPing Dummy.RemoteIP
    Next

    AddLog "* PING"

End Sub


Private Sub tmrTime_Timer()
  
  Dim Dummy         As clsUser
    
    'Update the uptime of each client
    For Each Dummy In Clients
        Dummy.UpTimeSeconds = Dummy.UpTimeSeconds + 1
    Next
    
    'Refresh the list
    Call RefreshList
    
    'Update the uptime of the server.
    ServerTime.UpTimeSeconds = ServerTime.UpTimeSeconds + 1
    lblUpTime.Caption = "Server Up Time: " & ServerTime.UpTimeHours & ":" & ServerTime.UpTimeMinutes & ":" & ServerTime.UpTimeSeconds
    
End Sub



Public Sub DisconnectServer()
'Disconnect the server

  If ListeningSocket <> 0 Then
        'Close the listening socket
        Server.CloseSocket ListeningSocket
        
        'Disconnect all clients
        cmdDisconnectAll_Click
        
        'Add to the log and disable uptime timer
        AddLog vbCrLf & "SocketWise Server is Disconnected" & vbCrLf
        tmrTime.Enabled = False
    End If
    
End Sub


Public Sub ConnectServer()
'Connect the server
    
    'If not connected,
    If ListeningSocket = 0 Then
        
        'Start the server
        ListeningSocket = Server.CreateSocket()
        Server.Listen ListeningSocket, CLng(ServerPort)
        
        'Add the log
        AddLog vbCrLf & "SocketWise Server is starting" & vbCrLf
        AddLog "* Listening Socket Handle: " & ListeningSocket
        AddLog "* Server listening for connections on port " & ServerPort
        
        'Enable the up time timer
        tmrTime.Enabled = True
    End If
    
End Sub
