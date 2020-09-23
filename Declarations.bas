Attribute VB_Name = "Declarations"
Public Connections As Long
Public AcceptAll2 As Boolean
Public WriteLog As Boolean
Public IPList2() As String
Public AcceptOne2 As Boolean
Public FilterIP As Boolean



Public Property Get AcceptOne() As Boolean
    AcceptOne = AcceptOne2
End Property

Public Property Let AcceptOne(ByVal b As Boolean)
    If b = True Then AcceptAll = False
    AcceptOne2 = b
End Property


Public Property Get AcceptAll() As Boolean
    AcceptAll = AcceptAll2
End Property

Public Property Let AcceptAll(ByVal b As Boolean)
    If b = True Then AcceptOne = False
    AcceptAll2 = b
End Property





Public Sub Log(ByVal data As String)
Dim ff As Integer
If WriteLog = True Then
ff = FreeFile
Open App.Path & "/Log.txt" For Append As ff
Print ff, data
Close ff
End If
End Sub
