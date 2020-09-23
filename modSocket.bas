Attribute VB_Name = "modSocket"
' Catalyst SocketWrench 3.5
' Copyright 1995-2000, Catalyst Development Corporation
' All rights reserved
'
' This product is licensed to you pursuant to the terms of the
' Catalyst license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'
Option Explicit

'
' General constants used with most of the controls
'
Global Const INVALID_HANDLE = -1
Global Const CONTROL_ERRIGNORE = 0
Global Const CONTROL_ERRDISPLAY = 1

'
' SocketWrench Control Actions
'
Global Const SOCKET_OPEN = 1
Global Const SOCKET_CONNECT = 2
Global Const SOCKET_LISTEN = 3
Global Const SOCKET_ACCEPT = 4
Global Const SOCKET_CANCEL = 5
Global Const SOCKET_FLUSH = 6
Global Const SOCKET_CLOSE = 7
Global Const SOCKET_DISCONNECT = 7
Global Const SOCKET_ABORT = 8
Global Const SOCKET_STARTUP = 9
Global Const SOCKET_CLEANUP = 10

'
' SocketWrench Control States
'
Global Const SOCKET_UNUSED = 0
Global Const SOCKET_IDLE = 1
Global Const SOCKET_LISTENING = 2
Global Const SOCKET_CONNECTING = 3
Global Const SOCKET_ACCEPTING = 4
Global Const SOCKET_RECEIVING = 5
Global Const SOCKET_SENDING = 6
Global Const SOCKET_CLOSING = 7

'
' Socket Address Families
'
Global Const AF_UNSPEC = 0
Global Const AF_UNIX = 1
Global Const AF_INET = 2

'
' Socket Types
'
Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2
Global Const SOCK_RAW = 3
Global Const SOCK_RDM = 4
Global Const SOCK_SEQPACKET = 5

'
' Protocol Types
'
Global Const IPPROTO_IP = 0
Global Const IPPROTO_ICMP = 1
Global Const IPPROTO_GGP = 2
Global Const IPPROTO_TCP = 6
Global Const IPPROTO_PUP = 12
Global Const IPPROTO_UDP = 17
Global Const IPPROTO_IDP = 22
Global Const IPPROTO_ND = 77
Global Const IPPROTO_RAW = 255
Global Const IPPROTO_MAX = 256

'
' Well-Known Port Numbers
'
Global Const IPPORT_ANY = 0
Global Const IPPORT_ECHO = 7
Global Const IPPORT_DISCARD = 9
Global Const IPPORT_SYSTAT = 11
Global Const IPPORT_DAYTIME = 13
Global Const IPPORT_NETSTAT = 15
Global Const IPPORT_CHARGEN = 19
Global Const IPPORT_FTP = 21
Global Const IPPORT_TELNET = 23
Global Const IPPORT_SMTP = 25
Global Const IPPORT_TIMESERVER = 37
Global Const IPPORT_NAMESERVER = 42
Global Const IPPORT_WHOIS = 43
Global Const IPPORT_MTP = 57
Global Const IPPORT_TFTP = 69
Global Const IPPORT_FINGER = 79
Global Const IPPORT_HTTP = 80
Global Const IPPORT_POP3 = 110
Global Const IPPORT_NNTP = 119
Global Const IPPORT_SNMP = 161
Global Const IPPORT_EXEC = 512
Global Const IPPORT_LOGIN = 513
Global Const IPPORT_SHELL = 514
Global Const IPPORT_RESERVED = 1024
Global Const IPPORT_USERRESERVED = 5000

'
' Network Addresses
'
Global Const INADDR_ANY = "0.0.0.0"
Global Const INADDR_LOOPBACK = "127.0.0.1"
Global Const INADDR_NONE = "255.255.255.255"

'
' Shutdown Values
'
Global Const SOCKET_READ = 0
Global Const SOCKET_WRITE = 1
Global Const SOCKET_READWRITE = 2

'
' Byte Order
'
Global Const LOCAL_BYTE_ORDER = 0
Global Const NETWORK_BYTE_ORDER = 1

'
' SocketWrench Error Response
'
Global Const SOCKET_ERRIGNORE = 0
Global Const SOCKET_ERRDISPLAY = 1

'
' SocketWrench Error Codes
'
Global Const WSABASEERR = 24000
Global Const WSAEINTR = 24004
Global Const WSAEBADF = 24009
Global Const WSAEACCES = 24013
Global Const WSAEFAULT = 24014
Global Const WSAEINVAL = 24022
Global Const WSAEMFILE = 24024
Global Const WSAEWOULDBLOCK = 24035
Global Const WSAEINPROGRESS = 24036
Global Const WSAEALREADY = 24037
Global Const WSAENOTSOCK = 24038
Global Const WSAEDESTADDRREQ = 24039
Global Const WSAEMSGSIZE = 24040
Global Const WSAEPROTOTYPE = 24041
Global Const WSAENOPROTOOPT = 24042
Global Const WSAEPROTONOSUPPORT = 24043
Global Const WSAESOCKTNOSUPPORT = 24044
Global Const WSAEOPNOTSUPP = 24045
Global Const WSAEPFNOSUPPORT = 24046
Global Const WSAEAFNOSUPPORT = 24047
Global Const WSAEADDRINUSE = 24048
Global Const WSAEADDRNOTAVAIL = 24049
Global Const WSAENETDOWN = 24050
Global Const WSAENETUNREACH = 24051
Global Const WSAENETRESET = 24052
Global Const WSAECONNABORTED = 24053
Global Const WSAECONNRESET = 24054
Global Const WSAENOBUFS = 24055
Global Const WSAEISCONN = 24056
Global Const WSAENOTCONN = 24057
Global Const WSAESHUTDOWN = 24058
Global Const WSAETOOMANYREFS = 24059
Global Const WSAETIMEDOUT = 24060
Global Const WSAECONNREFUSED = 24061
Global Const WSAELOOP = 24062
Global Const WSAENAMETOOLONG = 24063
Global Const WSAEHOSTDOWN = 24064
Global Const WSAEHOSTUNREACH = 24065
Global Const WSAENOTEMPTY = 24066
Global Const WSAEPROCLIM = 24067
Global Const WSAEUSERS = 24068
Global Const WSAEDQUOT = 24069
Global Const WSAESTALE = 24070
Global Const WSAEREMOTE = 24071
Global Const WSASYSNOTREADY = 24091
Global Const WSAVERNOTSUPPORTED = 24092
Global Const WSANOTINITIALISED = 24093
Global Const WSAHOST_NOT_FOUND = 25001
Global Const WSATRY_AGAIN = 25002
Global Const WSANO_RECOVERY = 25003
Global Const WSANO_DATA = 25004
Global Const WSANO_ADDRESS = 25004

Public Sub RecvFile()
    Dim strResult As String
    Dim byteBuffer() As Byte
    Dim cbBuffer As Integer
    Dim Buffer As Long
    Dim nResult As Integer
    Dim hFile As Integer
    Dim strLocalFile As String
    
    On Error Resume Next
    'If file exists, delete it...
    strLocalFile = frmMain.txtLocalFileName
    Kill strLocalFile
    Err.Clear
    
    'Accept automatically
    frmMain.swFileTransfer.Accept = frmMain.swFileTransfer.Handle
    
    'Create the file
    hFile = FreeFile()
    Open strLocalFile For Binary Access Write As #hFile
    
    On Error GoTo 0
    
    'Determine buffer
    Buffer = DetermineBufferSize()
    
    '
    ' ---- NOTE ----
    ' The following code writes the received data into the file.
    ' It was taken from the FTP.VB project and is almost exactly
    ' the same. I have only modified it to work with our socket,
    ' to display a true buffer (see variable Buffer), and
    ' to display the progress. The rest is kept the same.
    '
    Do
        'Make sure we're still connected (in case of cancelation by peer)
        If frmMain.swFileTransfer.Connected = False Then
            'If not, cancel the procedure
            MsgBox "File transfer has been canceled!", vbExclamation, " Error"
            ResetTransferDisplay
            Exit Do
            Exit Sub
        End If
        '
        ' Read data from the server in blocks determined by buffer size
        ' selected on main form  (Initial Settings) and write them
        ' to the file. The ReadBytes() method reads the data into
        ' a byte array, which must be adjusted to the actual
        ' number of bytes read (the Put statement determines how
        ' many bytes to write based on the size of the array).
        '
        ' Note that we redimension byteBuffer to cbBuffer - 1
        ' because the base of the array is 0, not 1 (otherwise,
        ' an extra null byte would be written after each block)
        '
        ReDim byteBuffer(Buffer)

        cbBuffer = frmMain.swFileTransfer.ReadBytes(byteBuffer, Buffer)
        
        'Check if we're done
        If cbBuffer < 1 Then Exit Do
        
        'If not, ReDim and move on
        ReDim Preserve byteBuffer(cbBuffer - 1)
        Put #hFile, , byteBuffer
        
        'Set the progress
        SetProgress FileLen(frmMain.txtLocalFileName)
        
        DoEvents
    Loop
    
    ' Close the file we're working with
    Close #hFile
End Sub

Public Sub SendData(ByVal Socket As Socket, ByVal Data As String)
    '
    ' This procedure sends data as a string to the socket.
    ' When transfering files we use the built in .Write event
    ' instead of calling this function, although it typically
    ' does the exact same thing.
    '
    ' If you insist, you can use the WriteBytes() method to send
    ' the file data.
    '
    Socket.Write Data, CInt(Len(Data))
End Sub

Public Sub SendFile(ByVal FileName As String)
    Dim iFileLen As Long
    Dim sFileName As String, sFilePath As String
    
    '
    ' This is our openning procedure to send the file.
    ' In here, we determine the filename, file size, and file path.
    '
    
    ' Use Len() here rather than FileLen() to make sure we don't get a -1
    ' or error.
    iFileLen = Len(FileName)
    If iFileLen = 0 Then
        MsgBox "Error: The file specified has been deleted or does not exist.", vbExclamation, "Socket Error"
        Exit Sub
    End If
    
    'Display status
    csClient.Status = 3
    SetClientStatus "Sending file..."
    
    'Now, let's get a REAL file size (in bytes).
    iFileLen = FileLen(FileName)
    
    'Extract filename and filepath
    sFileName = ParseFileName(FileName)
    sFilePath = ParseFilePath(FileName)
    If Right(sFilePath, 1) <> "\" Then sFilePath = sFilePath + "\"
    
    'Send the information to server
    SendData frmMain.swClient, "FILEINFO:" & iFileLen & "^" & sFileName & "^" & Replace(sFilePath, ":", "$")
    Pause 1
    
    'Prepare for our send!
    PrepareSend CStr(iFileLen), sFileName, sFilePath
End Sub
Public Sub ServerListen()
    'Set variables
    csServer.Status = 0

    'Initialize the command socket (swServer)
    With frmMain.swServer
        'Set properties
        .Binary = False
        .Blocking = False
        .BufferSize = 4096
        .LocalPort = frmMain.txtListenPort
        
        'Listen!
        .Action = SOCKET_LISTEN
    End With
    
    'Display new status
    SetServerStatus "Listening..."
    frmMain.cmdListen.Caption = "Stop listening..."
    frmMain.txtListenPort.Enabled = False
    
    'Set variables
    csServer.Status = 1
End Sub

Public Sub TransferFile()
    Dim hFile As Integer, SendCount As Long
    Dim strLocalFile$, DataChunk$
       
    
    'Tell the server to prepare the file and variables...
    SendData frmMain.swClient, "PREPAREFILE"
    Pause 1 'Wait 1 second
    strLocalFile = tFileToSend.FilePath & tFileToSend.FileName
    
    'Create the actual file
    On Error Resume Next: Err.Clear
    hFile = FreeFile()
    Open strLocalFile For Binary As #hFile
    
    On Error GoTo 0
             
        ' We use a SendCount to determine the ammount of
        ' data we've sent. This variable is then passed
        ' on to determine the progress. If the file size
        ' is 500, we add the length of DataChunk$, which
        ' is the current send (let's say it's 100 each time),
        ' to SendCount.
        '
        ' In the SetProgress() sub we display the progress by
        ' dividing SendCount by the file size, then multiplying
        ' by a hundred. The returning integer tells us the progress
        ' in a 100 percent field.
        SendCount = 0
      
        ' Loop and process the data until end of file is reached
        Do While Not EOF(hFile)
            
            'Some error checking to see if we're still connected.
            If frmMain.swFileTransfer.Connected = False Then
                MsgBox "File transfer has been canceled!", vbExclamation, " Error"
                ResetTransferDisplay
                Exit Do
                Exit Sub
            End If
            
            'Get the current chunk of data from the file & send it over
            DataChunk$ = Input(DetermineBufferSize, hFile)
            frmMain.swFileTransfer.Write DataChunk$, Len(DataChunk$)
            
            'Increase our SendCount variable by the length of current
            'chunk we sent over.
            SendCount = SendCount + Len(DataChunk$)
            
            'Set the progress meter/bar
            SetProgress SendCount
            
            DoEvents
        Loop
        
        'When done, reset the SendCount for safety purposes.
        SendCount = 0
    Close #hFile
    
    frmMain.swFileTransfer.Disconnect
    frmMain.swFileTransfer.Cleanup
    ResetTransferDisplay
    MsgBox "File transfer succesfully completed!", vbInformation, "Socket"
End Sub

