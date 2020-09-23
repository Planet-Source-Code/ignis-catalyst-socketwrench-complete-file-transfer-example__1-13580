Attribute VB_Name = "modUser"
Option Explicit

'Helps us keep track of the information
Global csServer As SERVER_PROPERTIES
Global csClient As CLIENT_PROPERTIES
Global tFileToSend As FILE_PROPERTIES

'Public types used by csServer, csClient, and tFileToSend
Public Type CLIENT_PROPERTIES
    Status As Integer
    BufferSize As Long
End Type
Public Type FILE_PROPERTIES
    FileLen As Long
    FileName As String
    FilePath As String
End Type
Public Type SERVER_PROPERTIES
    Status As Integer
    BufferSize As Long
End Type
Public Function ParseData(SearchString As String, ArgNum As Integer, Optional Delim As String = ":") As String
'
' Originally created by Mike Carper (named ExtractArgument)
' Modified by Arthur Nisnevich for use with this application
'
    On Error GoTo HandleError
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    
    LastPos = 1
    If ArgNum = 1 Then Arg = SearchString

    Do While InStr(SearchString, Delim) > 0
        Pos = InStr(LastPos, SearchString, Delim)

        If Pos = 0 Then
            If ArgCount = ArgNum - 1 Then Arg = Mid(SearchString, LastPos): Exit Do
        Else
            ArgCount = ArgCount + 1

            If ArgCount = ArgNum Then
                Arg = Mid(SearchString, LastPos, Pos - LastPos): Exit Do
            End If
        End If
        
        LastPos = Pos + 1
    Loop
    
    ParseData = Arg
    Exit Function
    
HandleError:
    MsgBox "Unexpected error occured in function ParseData(). Resuming normal..." & vbNewLine & vbNewLine & _
    "Error " & Err.Number & " : " & Err.Description, vbExclamation, "Unexpected Error"
    Resume Next
End Function

Public Function ParseFileName(ByVal strPath As String) As String
    '
    ' NOTE: StrReverse() only works in VB6! If you are using VB5 or below
    ' replace all occurences of this function with the function:
    ' ExtractFileName(). The same principles and arguments apply, however
    ' this method is faster.
    '
    strPath = StrReverse(strPath)
    strPath = Left(strPath, InStr(strPath, "\") - 1)
    ParseFileName = StrReverse(strPath)
End Function
Public Function ParseFilePath(ByVal File$) As String
    '
    ' Written by "anti"
    '
    Dim a, B, C
    
    For a = 1 To Len(File$)
        B = Right(File$, a)
        If Left(B, 1) = "\" Then Exit For
    Next a
    C = Left(File$, Len(File$) - Len(B) + 1)
    ParseFilePath = C
End Function
Public Function ExtractFileName(FilePath) As String
    '
    'ExtractFileName() written by Stewart MacFarlane
    '
    Dim X
    
    For X = Len(FilePath) To 1 Step -1
        If Mid(FilePath, X, 1) = "\" Then
            ExtractFileName = Right(FilePath, Len(FilePath) - X)
            Exit Function
        End If
    Next X
End Function
Public Function DetermineBufferSize() As Long
    Dim rBufferSize As String
    '
    ' NOTE: Even though it should send files superfast to your
    ' own computer, we are using buffering techniques, which
    ' are utilized for TCP, socket streams, and over the Internet
    ' (or a cross-network). Feel free to change these values
    ' if you are planning on using file transfer through a LAN.
    '
    
    'Determines the buffer size by checking the Preset Buffers
    'listbox (and slider).
    Select Case frmMain.cmbPresetBuffers.ListIndex
        Case 0
            DetermineBufferSize = 3072      '< 33.6
        Case 1
            DetermineBufferSize = 5120      '56k
        Case 2 To 6
            '
            ' I am having trouble figuring out why buffering above this
            ' ammount does not work. Feel free to lend a hand :)
            ' For now, I have made it simple and set all buffers above 56k
            ' and below Custom to 10240, which is ISDN speed.
            '
            DetermineBufferSize = 10240     'ISDN (dual/single channel)
        Case 7
            'Get custom buffer size
            rBufferSize = frmMain.txtCustomBuffer
            
            'Do error checking on custom buffer size
            If (rBufferSize = "0") Or (rBufferSize = "") Then
                'If buffer size is not valid, set default buffer
                DetermineBufferSize = 5120  '56k
            Else
                DetermineBufferSize = rBufferSize
            End If
    End Select
End Function
Public Sub PrepareSend(FileSize As String, FileName, FilePath)
    'Prepares the Transfer Progress frame
    With frmMain
        'Sets text values
        .txtFileTransferFile = FileName
        .txtFileSize = FileSize
        .txtLocalFileName = FilePath & FileName
        
        'Sets progress bar maximum to 100 (in case altered manually)
        .pbProgress.Max = 100
    End With
    
    'Sends the tFileToSend type to match the variables passed on
    'from the SendFile() function.
    With tFileToSend
        .FileLen = CLng(FileSize)
        .FileName = CStr(FileName)
        .FilePath = CStr(FilePath)
    End With
    
    'Display status
    SetTransferStatus "Waiting for reply from server..."
End Sub
Sub Pause(Duration)
    Dim StartTime, X
    '
    ' Simple timer which lets us pause our program. The variable Duration
    ' is in seconds NOT MILLISECONDS! This is used when we want to wait for
    ' the peer (either server or client) to finish processing whatever
    ' command we sent over.
    '
    StartTime = Timer
    
    Do While Timer - StartTime < Duration
        'Just a DoEvents() so we don't Overflow or lock up.
        X = DoEvents()
    Loop
End Sub
Public Sub ResetTransferDisplay()
    'Set display of Transfer Progress frame after a file
    'transfer is complete or has been canceled.
    With frmMain
        'Reset text fields
        .txtFileTransferFile = ""
        .txtFileSize = ""
        .txtLocalFileName = ""
        .txtRemoteFileName = ""
        
        'Reset progress bar/meter
        .pbProgress.Value = 0
        .lblPercentComplete = "0%"
        
        'Set display status
        SetTransferStatus "[ No file transfer in progress. ]"
        
        'Disable "Cancel" command button
        .cmdCancelFileTransfer.Enabled = False
    End With
End Sub

Public Sub SetProgress(Progress)
    '
    ' This calculates the current progress by checking
    ' the Progress sent (which should be either the current
    ' size of file, if on server, or the current total of chunks
    ' sent, if on client) and dividing it against the total file
    ' size. The percent is also displayed in a label.
    '
    With frmMain
        .pbProgress.Value = CInt((Progress / .txtFileSize) * 100)
        .lblPercentComplete = .pbProgress.Value & "%"
    End With
End Sub
Public Sub SetServerStatus(ByVal StatusText As String)
    'Sets the server status label
    frmMain.lblServerStatus = StatusText
End Sub

Public Sub SetClientStatus(ByVal StatusText As String)
    'Sets the client status label
    frmMain.lblClientStatus = StatusText
End Sub
Public Sub SetTransferStatus(ByVal StatusText As String)
    'Sets the transfer status label
    frmMain.lblRecvStatus = StatusText
End Sub
