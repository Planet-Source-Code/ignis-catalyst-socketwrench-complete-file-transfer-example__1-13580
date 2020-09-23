VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SocketWrench File Transfer Example (VB6)"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdBrowse 
      Left            =   60
      Top             =   3525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Caption         =   "Initial options"
      Height          =   1350
      Left            =   435
      TabIndex        =   35
      Top             =   6045
      Width           =   5355
      Begin VB.TextBox txtCustomBuffer 
         Height          =   315
         Left            =   3720
         TabIndex        =   41
         Top             =   885
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbPresetBuffers 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   1800
         List            =   "frmMain.frx":045E
         TabIndex        =   39
         Text            =   "56 K"
         Top             =   885
         Width           =   1860
      End
      Begin MSComctlLib.Slider slidBuffer 
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   555
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   1
         Max             =   6
         TextPosition    =   1
      End
      Begin VB.Label Label12 
         Caption         =   "Preset buffer speeds:"
         Height          =   285
         Left            =   180
         TabIndex        =   38
         Top             =   915
         Width           =   1605
      End
      Begin VB.Label Label11 
         Caption         =   "Buffer speed:"
         Height          =   285
         Left            =   150
         TabIndex        =   36
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "File Transfer Progress"
      Height          =   2565
      Left            =   435
      TabIndex        =   13
      Top             =   3390
      Width           =   5355
      Begin SocketWrenchCtrl.Socket swFileTransfer 
         Left            =   120
         Top             =   225
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         AutoResolve     =   -1  'True
         Backlog         =   5
         Binary          =   -1  'True
         Blocking        =   -1  'True
         Broadcast       =   0   'False
         BufferSize      =   0
         HostAddress     =   ""
         HostFile        =   ""
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   0
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   0
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   0
         Type            =   1
         Urgent          =   0   'False
      End
      Begin VB.CommandButton cmdCancelFileTransfer 
         Caption         =   "Cancel Transfer"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3750
         TabIndex        =   40
         Top             =   2085
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar pbProgress 
         Height          =   270
         Left            =   255
         TabIndex        =   33
         Top             =   1725
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtRemoteFileName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   975
         Width           =   3780
      End
      Begin VB.TextBox txtLocalFileName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   615
         Width           =   3945
      End
      Begin VB.TextBox txtFileSize 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtFileTransferFile 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   510
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   255
         Width           =   2295
      End
      Begin VB.Label lblPercentComplete 
         Caption         =   "0%"
         Height          =   315
         Left            =   4860
         TabIndex        =   34
         Top             =   1755
         Width           =   390
      End
      Begin VB.Label lblRecvStatus 
         Caption         =   "[ No file transfer in progress. ]"
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   1380
         Width           =   4935
      End
      Begin VB.Label Label9 
         Caption         =   "Remote filename:"
         Height          =   225
         Left            =   150
         TabIndex        =   30
         Top             =   1005
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Local filename:"
         Height          =   225
         Left            =   150
         TabIndex        =   28
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "bytes"
         Height          =   315
         Left            =   4860
         TabIndex        =   27
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "File size:"
         Height          =   255
         Left            =   2895
         TabIndex        =   25
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "File:"
         Height          =   225
         Left            =   135
         TabIndex        =   23
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server"
      Height          =   3255
      Left            =   2985
      TabIndex        =   1
      Top             =   75
      Width           =   3165
      Begin SocketWrenchCtrl.Socket swServer 
         Left            =   180
         Top             =   240
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         AutoResolve     =   -1  'True
         Backlog         =   5
         Binary          =   -1  'True
         Blocking        =   -1  'True
         Broadcast       =   0   'False
         BufferSize      =   0
         HostAddress     =   ""
         HostFile        =   ""
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   0
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   0
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   0
         Type            =   1
         Urgent          =   0   'False
      End
      Begin VB.Frame Frame7 
         Caption         =   "File Receival"
         Height          =   1485
         Left            =   135
         TabIndex        =   19
         Top             =   1635
         Width           =   2895
         Begin VB.CommandButton cmdBrowseServerDir 
            Caption         =   "Browse"
            Height          =   255
            Left            =   1770
            TabIndex        =   22
            Top             =   1125
            Width           =   1005
         End
         Begin VB.ComboBox cmbRecvDirectory 
            Height          =   315
            ItemData        =   "frmMain.frx":04AB
            Left            =   120
            List            =   "frmMain.frx":04B5
            TabIndex        =   21
            Text            =   "(App Path)"
            Top             =   750
            Width           =   2670
         End
         Begin VB.Label Label3 
            Caption         =   "Place received files in the following directory:"
            Height          =   420
            Left            =   120
            TabIndex        =   20
            Top             =   255
            Width           =   2715
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Status"
         Height          =   645
         Left            =   135
         TabIndex        =   17
         Top             =   930
         Width           =   2910
         Begin VB.Label lblServerStatus 
            AutoSize        =   -1  'True
            Caption         =   "Ready."
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Begin listening..."
         Height          =   270
         Left            =   270
         TabIndex        =   16
         Top             =   585
         Width           =   2655
      End
      Begin VB.TextBox txtListenPort 
         Height          =   285
         Left            =   510
         TabIndex        =   15
         Text            =   "15330"
         Top             =   210
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Port:"
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   255
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client"
      Height          =   3255
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2835
      Begin SocketWrenchCtrl.Socket swClient 
         Left            =   120
         Top             =   270
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         AutoResolve     =   -1  'True
         Backlog         =   5
         Binary          =   -1  'True
         Blocking        =   -1  'True
         Broadcast       =   0   'False
         BufferSize      =   0
         HostAddress     =   ""
         HostFile        =   ""
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   0
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   0
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   0
         Type            =   1
         Urgent          =   0   'False
      End
      Begin VB.Frame Frame4 
         Caption         =   "File Transfer"
         Height          =   1005
         Left            =   150
         TabIndex        =   9
         Top             =   2100
         Width           =   2565
         Begin VB.CommandButton cmdSendClientFile 
            Caption         =   "Send File!"
            Height          =   255
            Left            =   105
            TabIndex        =   12
            Top             =   570
            Width           =   1290
         End
         Begin VB.CommandButton cmdBrowseClientFile 
            Caption         =   "Browse"
            Height          =   255
            Left            =   1545
            TabIndex        =   11
            Top             =   570
            Width           =   885
         End
         Begin VB.TextBox txtFilename 
            Height          =   285
            Left            =   105
            TabIndex        =   10
            Top             =   240
            Width           =   2355
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Status"
         Height          =   645
         Left            =   150
         TabIndex        =   7
         Top             =   1425
         Width           =   2565
         Begin VB.Label lblClientStatus 
            AutoSize        =   -1  'True
            Caption         =   "Ready."
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.TextBox txtConnectPort 
         Height          =   285
         Left            =   615
         TabIndex        =   6
         Text            =   "15330"
         Top             =   615
         Width           =   2130
      End
      Begin VB.TextBox txtConnectIP 
         Height          =   285
         Left            =   450
         TabIndex        =   3
         Text            =   "127.0.01"
         Top             =   270
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect..."
         Height          =   285
         Left            =   330
         TabIndex        =   2
         Top             =   990
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "IP:"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   315
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbPresetBuffers_Click()
    'Check if set on "Custom"
    If cmbPresetBuffers.ListIndex = 7 Then
        'Hide slidBuffer
        slidBuffer.Visible = False
        
        'Show custom textbox; Set focus on it
        txtCustomBuffer.Visible = True
        txtCustomBuffer.SetFocus
    Else
        'Show slidBuffer
        slidBuffer.Visible = True
        
        'Synchronize slidBuffer and cmbPresetBuffers
        slidBuffer.Value = cmbPresetBuffers.ListIndex
        slidBuffer.ToolTipText = cmbPresetBuffers.Text
        
        'Hide custom textbox
        txtCustomBuffer.Visible = False
    End If
End Sub


Private Sub cmdBrowseClientFile_Click()
    On Error GoTo ForgetIt
    'Browse for file
    With cdBrowse
        .DialogTitle = "Browse For File"
        .Filter = "All Files (*.*)|*.*"
        .Flags = &H1000 Or &H4 Or &H800
        .ShowOpen
    End With
    
    'Open file into textbox
    txtFilename.Text = cdBrowse.FileName
ForgetIt:
End Sub

Private Sub cmdBrowseServerDir_Click()
    'Show Browse dialog
    frmBrowseDir.Show vbModal, Me
End Sub


Private Sub cmdCancelFileTransfer_Click()
    'Cancel transfer
    swFileTransfer.Action = SOCKET_ABORT
    swFileTransfer.Disconnect
    swFileTransfer.Cleanup
    
    'Display status
    ResetTransferDisplay
    MsgBox "File transfer canceled!", vbInformation, "Socket"
    
    'Set variables
    csServer.Status = 2
    csServer.Status = 2
End Sub

Private Sub cmdConnect_Click()
    Select Case csClient.Status
        Case 0
            'Set variables
            csClient.Status = 0
            
            'Error checking
            If txtConnectIP.Text = "" Then
                MsgBox "Error: The IP field must be either a IP or DNS address of the server.", vbExclamation, "Socket Error"
                Exit Sub
            ElseIf (txtConnectPort.Text = "") Or (Not VBA.IsNumeric(txtConnectPort)) Or (txtConnectPort <= 1000) Then
                MsgBox "Error: The port must be a numeric integer below 1000.", vbExclamation, "Socket Error"
                Exit Sub
            End If
            
            'Initialize command socket
            With frmMain.swClient
                'Set properties
                .AddressFamily = AF_INET
                .Protocol = IPPROTO_IP
                .SocketType = SOCK_STREAM
                .Binary = False
                .Blocking = False
                .BufferSize = 4096
                .RemotePort = txtConnectPort
                .HostName = Trim(txtConnectIP)
                
                'Connect!
                .Action = SOCKET_CONNECT
            End With
            
            'Display status
            SetClientStatus "Connecting..."
            txtConnectIP.Enabled = False
            txtConnectPort.Enabled = False
            cmdConnect.Caption = "Disconnect..."
            
            'Set variables
            csClient.Status = 1 'Connecting
        Case 1
            'Cancel connect event
            If MsgBox("Are you sure you want to cancel connection event?", vbQuestion + vbYesNo, "Socket") = vbYes Then
                swClient.Action = SOCKET_ABORT
                swClient.Disconnect
                swClient.Cleanup
                
                'Display status
                SetClientStatus "Ready."
                txtConnectIP.Enabled = True
                txtConnectPort.Enabled = True
                cmdConnect.Caption = "Connect..."
                
                'Set variables
                csClient.Status = 0 'Offline
                Exit Sub
            Else
                Exit Sub
            End If
        Case 2
            'Disconnect
            If MsgBox("Are you sure you want to disconnect?", vbQuestion + vbYesNo, "Socket") = vbYes Then
                swClient.Disconnect
                swClient.Cleanup
                
                'Display status
                SetClientStatus "Ready."
                txtConnectIP.Enabled = True
                txtConnectPort.Enabled = True
                cmdConnect.Caption = "Connect..."
                
                'Set variables
                csClient.Status = 0 'Offline
            End If
        Case 3
            'Unable to abort!
            MsgBox "Error: Unable to disconnect during file transfer. Cancel the transfer first, then disconnect.", vbExclamation, "Socket Error"
            Exit Sub
    End Select
End Sub
Private Sub cmdListen_Click()
    Select Case csServer.Status
        Case 0
            'Error check port number (numeric integer above 1000)
            If (txtListenPort.Text = "") Or (Not VBA.IsNumeric(txtListenPort)) Or (txtListenPort.Text <= 1000) Then
                MsgBox "Error: Server port must be a numeric integer above 1000.", vbExclamation, "Socket Error"
                Exit Sub
            End If
            
            'Begin listening
            modSocket.ServerListen
        Case 1 To 2
            If MsgBox("Are you sure you want to shut the server down?", vbQuestion + vbYesNo, "Socket") = vbYes Then
                'Disconnect server
                swServer.Disconnect
                swServer.Cleanup
                
                'Display status
                frmMain.txtListenPort.Enabled = True
                csServer.Status = 0
                SetServerStatus "Ready."
                frmMain.cmdListen.Caption = "Begin listening..."
            Else
                Exit Sub
            End If
        Case 3 'Generate error
            MsgBox "Error: The server is in the process of file transfer, cannot close the server. Cancel the file transfer first, then close the server.", vbExclamation, "Socket Error"
            Exit Sub
    End Select
End Sub

Private Sub cmdSendClientFile_Click()
    'Error checking
    If txtFilename.Text = "" Then
        MsgBox "Error: Type in a filename to send, or click Browse.", vbExclamation, "Socket Error"
        Exit Sub
    End If
    
    'Send file
    SendFile txtFilename
End Sub

Private Sub Form_Load()
    'Set initial buffer speed
    cmbPresetBuffers.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    ' The End() method should automatically close any active sockets,
    ' however, for security, you can place default Disconnect and Cleanup()
    ' events for the sockets in place of End().
    '
    End
End Sub


Private Sub slidBuffer_Change()
    'Synchronize slidBuffer and cmbPresetBuffers
    cmbPresetBuffers.ListIndex = slidBuffer.Value
End Sub

Private Sub swClient_Connect()
    'Display status
    SetClientStatus "Connected. Ready."
    csClient.Status = 2
End Sub


Private Sub swClient_Disconnect()
    'Display status
    csClient.Status = 0
    txtConnectIP.Enabled = True
    txtConnectPort.Enabled = True
    cmdConnect.Caption = "Connect..."
    SetClientStatus "Ready."
End Sub


Private Sub swClient_Read(DataLength As Integer, IsUrgent As Integer)
    Dim upMessage As String, sData As String, sParam As String
    
    'Download the data
    swClient.RecvLen = DataLength
    upMessage = swClient.RecvData
    
    'Parse through the data
    sData = ParseData(upMessage, 1)
    sParam = ParseData(upMessage, 2)

    Debug.Print "RECV= " & upMessage
    
    'Execute commands in the data
    Select Case sData
        Case "PORT"
            'Prepare transfer socket
            With swFileTransfer
                'Set properties
                .AddressFamily = AF_INET
                .Binary = True
                .Blocking = True
                .BufferSize = DetermineBufferSize
                .HostName = Trim(txtConnectIP)
                .RemotePort = 15301
                .Protocol = IPPROTO_TCP
                .SocketType = SOCK_STREAM
                
                'Connect
                .Action = SOCKET_CONNECT
            End With
            
            'Display status
            SetTransferStatus "Sending file..."
            cmdCancelFileTransfer.Enabled = True
            
            'Execute the TransferFile() procedure, which
            'sends the actual file data
            TransferFile
            
        Case "LOCALFILE"
            'The server has sent us their local file name,
            'Display Status.
            txtRemoteFileName = Replace(sParam, "$", ":")
    End Select
End Sub

Private Sub swFileTransfer_Disconnect()
    ResetTransferDisplay
End Sub


Private Sub swServer_Accept(SocketId As Integer)
    'Accept connection
    swServer.Accept = SocketId
    
    'Display status
    SetServerStatus "Connected. Ready."
    csServer.Status = 2
End Sub


Private Sub swServer_Disconnect()
    'Display status
    csServer.Status = 1
    SetServerStatus "Listening..."
End Sub


Private Sub swServer_Read(DataLength As Integer, IsUrgent As Integer)
    Dim upMessage As String, sData As String, sParam As String
    Dim FileSize, FileName, FilePath, LocalPath
    
    'Download the data
    swServer.RecvLen = DataLength
    upMessage = swServer.RecvData
    
    'Parse through the data
    sData = ParseData(upMessage, 1)
    sParam = ParseData(upMessage, 2)
    
    Debug.Print "RECV= " & upMessage
    
    'Execute commands in the data
    Select Case sData
        Case "FILEINFO"
            'Extracts the file info from the sParam data that
            'has been sent from the client.
            FileSize = ParseData(sParam, 1, "^")
            FileName = ParseData(sParam, 2, "^")
            FilePath = Replace(ParseData(sParam, 3, "^"), "$", ":")
            
            'Creates a LocalPath variable
            If cmbRecvDirectory.Text = "(App Path)" Then
                LocalPath = App.Path
            Else
                LocalPath = cmbRecvDirectory.Text
            End If
            If Right(LocalPath, 1) <> "\" Then LocalPath = LocalPath & "\"
            
            'Display status
            txtFileTransferFile = FileName
            txtRemoteFileName = FilePath & FileName
            txtLocalFileName = LocalPath & FileName
            txtFileSize = FileSize
            
            'Returns the LocalPath to the client so it can display it
            SendData swServer, "LOCALFILE:" & Replace(txtLocalFileName, ":", "$")
            Pause 1 'Wait 1 second
            
            'Display status & set variables
            csServer.Status = 3
            SetServerStatus "Reciving file..."
            SetTransferStatus "Waiting for file to be sent..."
            
            'Prepare transfer socket
            With swFileTransfer
                'Set properties
                .BindAddress = swServer.LocalAddress
                .AddressFamily = AF_INET
                .Binary = True
                .Blocking = True
                .BufferSize = DetermineBufferSize
                .HostAddress = INADDR_ANY
                .LocalPort = 15301
                .Protocol = IPPROTO_TCP
                .Timeout = 0
                .SocketType = SOCK_STREAM
                
                'Listen
                .Action = SOCKET_LISTEN
            End With
            
            'Send back portstring
            Pause 1
            '
            ' NOTE: The PortString is not used by the client
            ' in any way right now other than determing when
            ' the server is ready to accept the file
            '
            SendData swServer, "PORT:" & swFileTransfer.PortString
            
        Case "PREPAREFILE"
            ' The client has received our PortString and is
            ' telling us to prepare the file
            SetTransferStatus "Receiving file..."
            cmdCancelFileTransfer.Enabled = True
            'Execute the RecvFile() function, which creates the
            'file and downloads all of the data.
            RecvFile
    End Select
End Sub

