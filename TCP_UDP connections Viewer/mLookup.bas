Attribute VB_Name = "mLookup"
 Const WSADESCRIPTION_LEN = 256
 Const WSASYSSTATUS_LEN = 256
 Const WSADESCRIPTION_LEN_1 = WSADESCRIPTION_LEN + 1
 Const WSASYSSTATUS_LEN_1 = WSASYSSTATUS_LEN + 1
 Const SOCKET_ERROR = -1
 
 Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
 End Type

 Type tagWSAData
        wVersion            As Integer
        wHighVersion        As Integer
        szDescription       As String * WSADESCRIPTION_LEN_1
        szSystemStatus      As String * WSASYSSTATUS_LEN_1
        iMaxSockets         As Integer
        iMaxUdpDg           As Integer
        lpVendorInfo        As String * 200
 End Type

Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequested As Integer, lpWSADATA As tagWSAData) As Integer
Declare Function WSACleanup Lib "WSOCK32" () As Integer
Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Declare Function gethostbyaddr Lib "WSOCK32" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
