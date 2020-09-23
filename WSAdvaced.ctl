VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl WSAdvaced 
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   Picture         =   "WSAdvaced.ctx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   3375
   ToolboxBitmap   =   "WSAdvaced.ctx":26332
   Begin MSWinsockLib.Winsock WS 
      Index           =   0
      Left            =   2520
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "WSAdvaced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'*                                                                 *
'*   This code was written by Nok1 and is rightful property        *
'*   of Nok1.  Code Copyright 2002 Nok1 Inc.  You may use this     *
'*   code as long the name of the original author is included.     *
'*   And, oh ya, if you put this up on a website or something,     *
'*   please tell me - its not like i can stop you from using it.   *
'*   I usually dont comment my code, but if I will be glad to      *
'*   answer any questions that you have if you email them to       *
'*   N2k8000@hotmail.com and have "Program Question" as the header.*
'*   then in the email tell me the program that you are having     *
'*   problems with and I'll get back to you ASAP.                  *
'*                                                                 *
'*******************************************************************

Public Event wsSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Public Event wsSendComplete()
Public Event wsDataArrival(ByVal bytesTotal As Long)
Public Event wsConnectionRequest(ByVal requestID As Long)
Public Event wsConnect()
Public Event wsClose()
Public Event wsError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)


Public Property Get sRemotePort() As String
RemotePort = WS(0).RemotePort
End Property

Public Property Get mRemotePorts() As String
For i = WS.LBound To WS.UBound
    tmp = WS(i).RemotePort & vbCrLf
Next i
Connections = WS.UBound - WS.LBound + 1
RemotePorts = tmp
End Property

Public Property Get sRemoteHost() As String
    RemoteHost = WS(0).RemoteHost
End Property

Public Property Get mRemoteHosts() As String
For i = WS.LBound To WS.UBound
    tmp = WS(i).RemoteHost & vbCrLf
Next i
Connections = WS.UBound - WS.LBound + 1
RemoteHosts = tmp
End Property

Public Property Get sRemoteHostIP() As String
RemoteHostIP = WS(0).RemoteHostIP
End Property

Public Property Get mRemoteHostsIP() As String
For i = WS.LBound To WS.UBound
    tmp = WS(i).RemoteHostIP & vbCrLf
Next i
Connections = WS.UBound - WS.LBound + 1
RemoteHostsIP = tmp
End Property

Public Property Get sLocalPort() As Long
LocalPort = WS(0).LocalPort
End Property

Public Property Get sLocalHostName() As String
LocalHostName = WS(0).LocalHostName
End Property

Public Property Get sLocalIP() As String
LocalIP = WS(0).LocalIP
End Property

Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_Resize()
UserControl.Height = 3480
UserControl.Width = 3375
End Sub

Private Sub WS_Close(Index As Integer)
RaiseEvent wsClose
End Sub

Private Sub WS_Connect(Index As Integer)
RaiseEvent wsConnect
End Sub

Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)

If FilterIP = True Then
    For i = LBound(IPList) To UBound(IPList)
        If IPList(i) = WS(Index).RemoteHostIP Then
            WS(Index).Accept
        End If
    Next i
    Exit Sub
End If

If AcceptOne = True And WS.UBound = 1 Then
    WS(0).Accept
End If

If AcceptAll = True Then
Dim a As Integer, Connections As Integer
tryagain:
    For a = 1 To WS.UBound
        If WS(a).State = 0 Or WS(a).State = 8 Then
            WS(a).Close
            WS(a).Tag = ""
            WS(a).Accept requestID
            Log vbCrLf & "Accepted Request ID " & requestID & " by closing WS(" & a & ") IP: " & WS(a).RemoteHostIP
        End If
        DoEvents
    Next a
DoEvents
Dim num As Integer
    num = WS.UBound + 1
    On Error Resume Next
    Load WS(a)
    WS(a).Accept requestID
    Log vbCrLf & "Accepted " & WS(num).RemoteHostIP
End If
Connections = WS.UBound - WS.LBound + 1
RaiseEvent wsConnectionRequest(requestID)
End Property

Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
RaiseEvent wsDataArrival(bytesTotal)
End Sub

Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent wsError(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub WS_SendComplete(Index As Integer)
RaiseEvent wsSendComplete
End Sub

Private Sub WS_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
RaiseEvent wsSendProgress(bytesent, bytesRemaining)
End Sub

Public Sub mSendAll(ByVal data As String)
For i = LBound(WS) To UBound(WS)
    WS(i).SendData data
Next i
End Sub

Public Sub sSendData(ByVal data As String)
WS(0).SendData data
End Sub

Public Sub mCloseAll()
For i = LBound(WS) To UBound(WS)
    WS(i).Close
Next i
End Sub

Public Sub sClose()
WS(0).Close
End Sub

Public Sub sConnect(ByVal IP As Long, Optional ByVal Port As Long)
If Port <> Null Or "" Then
    WS(0).Connect IP, Port
Else
    WS(0).Connect IP
End If
End Sub

Public Sub sHost(ByVal Port As Long)
WS(0).LocalPort = Port
WS(0).Listen
End Sub
