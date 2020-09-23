VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cache As Boolean
Dim StoredCache As New Collection

Property Let DNScache(value As Boolean)
    Cache = value
End Property

Property Get DNScache() As Boolean
    DNScache = Cache
End Property

'Searches the Current Cache Database to see if the Host has already been resolved
Private Function SearchForIP(Host As String) As String

    For x = 1 To StoredCache.Count

        If Split(StoredCache(x), ":")(1) = Host Then
            SearchForIP = Split(StoredCache(x), ":")(0)
            Exit Function
        End If
    Next
    
SearchForIP = ""
    
End Function

'Searches the Current Cache Database to see if the IP has already been looked up
Private Function SearchForHost(IP As String) As String

    For x = 1 To StoredCache.Count
        
        If Split(StoredCache(x), ":")(0) = IP Then
            SearchForHost = Split(StoredCache(x), ":")(1)
            Exit Function
        End If
    Next
    
SearchForHost = ""
    
End Function

'Reads the current Cache database into a string then saves it to a file
Sub ExportCache(sFile As String)
Dim F As Long
Dim sTmp As String

For x = 1 To StoredCache.Count
    DoEvents
    sTmp = sTmp & StoredCache(x) & vbCrLf
Next

F = FreeFile                'Get Next avaliable filenumber

Open sFile For Output As #F 'Open file to write it it
    Print #F, sTmp          'Paste all the cache into the file
Close #F                    'Close the file

End Sub

Sub ImportCache(sFile As String)
Dim F As Long
Dim sTmp As String

'Removes all current cache, can be removed to add cache to the current database
'<\----Optional----/>
For x = 1 To StoredCache.Count
    DoEvents
    StoredCache.Remove (x)
Next
'</----Optional----\>

F = FreeFile                 'Get next avaliable file number

Open sFile For Input As #F   'Open file to read from it
    Do Until EOF(F)          'Loop until End Of File
        Line Input #F, sTmp  'Read line into string
        StoredCache.Add sTmp 'Add string to cache database
    Loop
Close #F                     'Close the file
End Sub


Private Sub Add2Cache(IP As String, Host As String)
    
    If SearchForHost(IP) = "" Then      'If not found in the database then
        StoredCache.Add IP & ":" & Host 'Add it
    End If
    
End Sub

Function DNSlookup(IP As String) As String

Dim exeAPI As Long
Dim HostInfo As HOSTENT
Dim TmpNm As String * 255
Dim SockStart As Integer
Dim SockInf As tagWSAData
Dim IPAdr As Long
Dim sHost As String

If DNScache = True Then     'If DNS caching enabled then

sHost = SearchForHost(IP)   'Check if current IP is in the database
    
    If sHost <> "" Then     'If found then
        DNSlookup = sHost   'Return the Hostname
        Exit Function
    End If
End If

IPAdr = AddressStringToLong(IP)              'Convert IP to Long
SockState = WSAStartup(&H101, SockInf)       'Start up new Socket
exeAPI = gethostbyaddr(IPAdr, Len(IPAdr), 2) 'get the Host by the converted IP
SockState = WSACleanup()                     'Close socket

If exeAPI <> 0 Then                          'If not invalid Host
    
    'Extract the Hostname
    RtlMoveMemory HostInfo, exeAPI, ByVal (LenB(HostInfo))
    RtlMoveMemory ByVal TmpNm, HostInfo.hName, 255
    
    If InStr(1, TmpNm, Chr$(0), vbTextCompare) > 0 Then TmpNm = Left(TmpNm, InStr(1, TmpNm, Chr$(0), vbTextCompare))
    sHost = Trim(TmpNm) 'Trim Trailing spaces
    DNSlookup = sHost   'return hostname
    If DNScache = True Then Add2Cache IP, sHost 'If caching enabled then add host
Else
    DNSlookup = ""
End If

End Function

Function ResolveHost(Host As String) As String
'Similar to above, refer to it for comments
Dim exeAPI As Long
Dim HostInfo As HOSTENT
Dim IPv4Adr As Long
Dim SockStart As Integer
Dim SockInf As tagWSAData
Dim sIP As String

If DNScache = True Then

sIP = SearchForIP(Host)
    
    If sIP <> "" Then
        ResolveHost = sIP
        Exit Function
    End If
End If


SockState = WSAStartup(&H101, SockInf)
exeAPI = gethostbyname(Host)
SockState = WSACleanup()

If exeAPI > 0 Then
    
    RtlMoveMemory HostInfo, exeAPI, ByVal (LenB(HostInfo))
    RtlMoveMemory exeAPI, ByVal HostInfo.hAddrList, 4
    RtlMoveMemory IPv4Adr, ByVal exeAPI, HostInfo.hLen
    sIP = AddressLongToString(IPv4Adr)
    ResolveHost = sIP
    
    If DNScache = True Then Add2Cache sIP, Host
Else
    
    ResolveHost = 0
End If
End Function


Private Function AddressLongToString(tmp) As String
Dim parts(0 To 4) As String

parts(0) = Right("00000000" & Hex(tmp), 8)
parts(1) = Right(parts(0), 2)
parts(2) = Mid(parts(0), Len(parts(0)) - 3, 2)
parts(3) = Mid(parts(0), Len(parts(0)) - 5, 2)
parts(4) = Left(parts(0), 2)

AddressLongToString = Val("&h" & parts(1) & "&") & "." & Val("&h" & parts(2) & "&") & "." & Val("&h" & parts(3) & "&") & "." & Val("&h" & parts(4) & "&")
End Function

Private Function AddressStringToLong(ByVal tmp As String) As Long
Dim parts(1 To 4) As String
    
    If InStr(1, tmp, ".", vbTextCompare) = 0 Then
        AddressStringToLong = gethostbyname(tmp)
    Else
    
        For x = 1 To 4
            parts(x) = Split(tmp, ".")(x - 1)
        Next
        
        AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
   End If
End Function

