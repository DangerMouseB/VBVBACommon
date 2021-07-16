Attribute VB_Name = "mWinAPI_Winsock2"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit
Option Private Module


' WSA - Windows Sockets API
' ByVal name As String => char * name


'*************************************************************************************************************************************************************************************************************************************************
' UNCLASSIFIED
'*************************************************************************************************************************************************************************************************************************************************

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Public Const MSG_PEEK As Long = &H2

Public Const ERROR_SUCCESS As Long = &H0
Public Const INADDR_NONE As Long = &HFFFF

Public Const MAXGETHOSTSTRUCT = 1024

' from example on http://binaryworld.net/Main/CodeDetail.aspx?CodeId=3591
Public Const PF_INET = 2

Public Const SOL_SOCKET As Long = &HFFFF&

Public Const SOMAXCONN_NT_WORKSTATION As Long = 5
Public Const SOMAXCONN_NT_SERVER As Long = 200
Public Const SOMAXCONN = &H7FFFFFFF             ' maximum length for a connection request queue created by the socket call listen()


'*************************************************************************************************************************************************************************************************************************************************

' General
Public Const INVALID_SOCKET As Long = -1

Public Type SOCKET
    SOCKET As Long
End Type

Public Const SOCKET_ERROR As Long = -1

Public Type ZERO_OR_SOCKET_ERROR
    ZOSE As Long
End Type

Public Type BYTES_OR_SOCKET_ERROR
    BOSE As Long
End Type

Public Type P_hostent
    phostent As Long
End Type

Public Type P_ANSI_STR
    pANSIStr As Long
End Type

Public Type IP32Bit
    IP32Bit As Long
End Type

Public Type WSA_ERROR_CODE
    WSAEC As Long
End Type

' Winsock version
Public Const SOCKET_VERSION_11 As Long = &H101
Public Const SOCKET_VERSION_22 As Long = &H202


' Address Family
Public Const AF_UNSPEC As Long = 0
Public Const AF_INET As Long = 2
Public Const AF_IPX As Long = 6
Public Const AF_APPLETALK As Long = 16
Public Const AF_NETBIOS As Long = 17
Public Const AF_INET6 As Long = 23
Public Const AF_IRDA As Long = 26
Public Const AF_BTH As Long = 32

' Socket Type
Public Const SOCK_STREAM As Long = 1
Public Const SOCK_DGRAM As Long = 2
Public Const SOCK_RAW As Long = 3
Public Const SOCK_RDM As Long = 4
Public Const SOCK_SEQPACKET As Long = 5

' Protocol
Public Const BTHPROTO_RFCOMM As Long = 3
Public Const IPPROTO_TCP As Long = 6
Public Const IPPROTO_UDP As Long = 17
Public Const IPPROTO_RM As Long = 113

' Shutdown Flags
Public Const SD_RECEIVE As Long = &H0
Public Const SD_SEND As Long = &H1
Public Const SD_BOTH As Long = &H2

' Translated Handle
Public Const TH_NETDEV As Long = &H1
Public Const TH_TAPI As Long = &H2

' Windows Sockets Error Codes
Public Const WSA_INVALID_HANDLE As Long = 6
Public Const WSA_NOT_ENOUGH_MEMORY As Long = 8
Public Const WSA_INVALID_PARAMETER As Long = 87
Public Const WSA_OPERATION_ABORTED As Long = 995
Public Const WSA_IO_INCOMPLETE As Long = 996
Public Const WSA_IO_PENDING As Long = 997

' Windows Sockets definitions of regular Microsoft C error constants
Public Const WSAEINTR As Long = 10004
Public Const WSAEBADF As Long = 10009
Public Const WSAEACCES As Long = 10013
Public Const WSAEFAULT As Long = 10014
Public Const WSAEINVAL As Long = 10022
Public Const WSAEMFILE As Long = 10024

' Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK As Long = 10035
Public Const WSAEINPROGRESS As Long = 10036
Public Const WSAEALREADY As Long = 10037
Public Const WSAENOTSOCK As Long = 10038
Public Const WSAEDESTADDRREQ As Long = 10039
Public Const WSAEMSGSIZE As Long = 10040
Public Const WSAEPROTOTYPE As Long = 10041
Public Const WSAENOPROTOOPT As Long = 10042
Public Const WSAEPROTONOSUPPORT As Long = 10043
Public Const WSAESOCKTNOSUPPORT As Long = 10044
Public Const WSAEOPNOTSUPP As Long = 10045
Public Const WSAEPFNOSUPPORT As Long = 10046
Public Const WSAEAFNOSUPPORT As Long = 10047
Public Const WSAEADDRINUSE As Long = 10048
Public Const WSAEADDRNOTAVAIL As Long = 10049
Public Const WSAENETDOWN As Long = 10050
Public Const WSAENETUNREACH As Long = 10051
Public Const WSAENETRESET As Long = 10052
Public Const WSAECONNABORTED As Long = 10053
Public Const WSAECONNRESET As Long = 10054
Public Const WSAENOBUFS As Long = 10055
Public Const WSAEISCONN As Long = 10056
Public Const WSAENOTCONN As Long = 10057
Public Const WSAESHUTDOWN As Long = 10058
Public Const WSAETOOMANYREFS As Long = 10059
Public Const WSAETIMEDOUT As Long = 10060
Public Const WSAECONNREFUSED As Long = 10061
Public Const WSAELOOP As Long = 10062
Public Const WSAENAMETOOLONG As Long = 10063
Public Const WSAEHOSTDOWN As Long = 10064
Public Const WSAEHOSTUNREACH As Long = 10065
Public Const WSAENOTEMPTY As Long = 10066
Public Const WSAEPROCLIM As Long = 10067
Public Const WSAEUSERS As Long = 10068
Public Const WSAEDQUOT As Long = 10069
Public Const WSAESTALE As Long = 10070
Public Const WSAEREMOTE As Long = 10071

' Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY As Long = 10091
Public Const WSAVERNOTSUPPORTED As Long = 10092
Public Const WSANOTINITIALISED As Long = 10093
Public Const WSAEDISCON As Long = 10101
Public Const WSAENOMORE As Long = 10102
Public Const WSAECANCELLED As Long = 10103
Public Const WSAEINVALIDPROCTABLE As Long = 10104
Public Const WSAEINVALIDPROVIDER As Long = 10105
Public Const WSAEPROVIDERFAILEDINIT As Long = 10106
Public Const WSASYSCALLFAILURE As Long = 10107
Public Const WSASERVICE_NOT_FOUND As Long = 10108
Public Const WSATYPE_NOT_FOUND As Long = 10109
Public Const WSA_E_NO_MORE As Long = 10110
Public Const WSA_E_CANCELLED As Long = 10111
Public Const WSAEREFUSED As Long = 10112

Public Const WSAHOST_NOT_FOUND As Long = 11001
Public Const WSATRY_AGAIN As Long = 11002
Public Const WSANO_RECOVERY As Long = 11003
Public Const WSANO_DATA As Long = 11004

Public Const WSA_QOS_RECEIVERS As Long = 11005
Public Const WSA_QOS_SENDERS As Long = 11006
Public Const WSA_QOS_NO_SENDERS As Long = 11007
Public Const WSA_QOS_NO_RECEIVERS As Long = 11008
Public Const WSA_QOS_REQUEST_CONFIRMED As Long = 11009
Public Const WSA_QOS_ADMISSION_FAILURE As Long = 11010
Public Const WSA_QOS_POLICY_FAILURE As Long = 11011
Public Const WSA_QOS_BAD_STYLE As Long = 11012
Public Const WSA_QOS_BAD_OBJECT As Long = 11013
Public Const WSA_QOS_TRAFFIC_CTRL_ERROR As Long = 11014
Public Const WSA_QOS_GENERIC_ERROR As Long = 11015
Public Const WSA_QOS_ESERVICETYPE As Long = 11016
Public Const WSA_QOS_EFLOWSPEC As Long = 11017
Public Const WSA_QOS_EPROVSPECBUF As Long = 11018
Public Const WSA_QOS_EFILTERSTYLE As Long = 11019
Public Const WSA_QOS_EFILTERTYPE As Long = 11020
Public Const WSA_QOS_EFILTERCOUNT As Long = 11021
Public Const WSA_QOS_EOBJLENGTH As Long = 11022
Public Const WSA_QOS_EFLOWCOUNT As Long = 11023
Public Const WSA_QOS_EUNKOWNPSOBJ As Long = 11024
Public Const WSA_QOS_EPOLICYOBJ As Long = 11025
Public Const WSA_QOS_EFLOWDESC As Long = 11026
Public Const WSA_QOS_EPSFLOWSPEC As Long = 11027
Public Const WSA_QOS_EPSFILTERSPEC As Long = 11028
Public Const WSA_QOS_ESDMODEOBJ As Long = 11029
Public Const WSA_QOS_ESHAPERATEOBJ As Long = 11030
Public Const WSA_QOS_RESERVED_PETYPE As Long = 11031

' NetworkEvents
Public Const FD_READ As Long = &H1&                      'FD_READ Wants to receive notification of readiness for reading.
Public Const FD_WRITE As Long = &H2&                  'FD_WRITE Wants to receive notification of readiness for writing.
Public Const FD_OOB As Long = &H4                        'FD_OOB Wants to receive notification of the arrival of OOB data.
Public Const FD_ACCEPT As Long = &H8                 'FD_ACCEPT Wants to receive notification of incoming connections.
Public Const FD_CONNECT As Long = &H10&          'FD_CONNECT Wants to receive notification of completed connection or multipoint join operation.
Public Const FD_CLOSE As Long = &H20&               'FD_CLOSE Wants to receive notification of socket closure.
Public Const FD_QOS As Long = &H40                    'FD_QOS Wants to receive notification of socket Quality of Service (QOS) changes.
Public Const FD_GROUP_QOS As Long = &H80        'FD_GROUP_QOS Reserved.
Public Const FD_ROUTING_INTERFACE_CHANGE As Long = &H100    'FD_ROUTING_INTERFACE_CHANGE Wants to receive notification of routing interface changes for the specified destination(s).
Public Const FD_ADDRESS_LIST_CHANGE As Long = &H200             'FD_ADDRESS_LIST_CHANGE Wants to receive notification of local address list changes for the socket's protocol family.
Public Const FD_MAX_EVENTS As Long = &H400                               'FD_MAX_EVENTS
Public Const FD_ALL_EVENTS As Long = &H800                                'FD_ALL_EVENTS

Public Const FD_SETSIZE As Long = 64

' Socket Options
Public Const SO_SNDBUF As Long = &H1001&
Public Const SO_RCVBUF As Long = &H1002&
Public Const SO_MAX_MSG_SIZE As Long = &H2003
Public Const SO_BROADCAST As Long = &H20
Public Const SO_REUSEADDR = &H4
Public Const SO_LINGER = &H80&

Public Enum IOConstants
    IOCPARM_MASK = &H7F   ' /* parameters must be < 128 bytes */
    IOC_VOID = &H20000000 ' /* no parameters */
    IOC_OUT = &H40000000   ' /* copy out parameters */
    IOC_IN = &H80000000      ' /* copy in parameters */
                                           ' /* 0x20000000 distinguishes new & old ioctl 's */
    IOC_INOUT = IOC_IN Or IOC_OUT
    IOC_UNIX = &H0
    IOC_WS2 = &H8000000
    IOC_PROTOCOL = &H10000000
    IOC_VENDOR = &H18000000
End Enum

Public Enum IOFlags
    FIONREAD = &H4004667F
    FIONBIO = &H8004667E
    FIOASYNC = &H8004667D
    
    SIOCATMARK = &H40047307
    
    SIO_ASSOCIATE_HANDLE = IOC_IN Or IOC_WS2 Or 1 ' _WSAIOW(IOC_WS2,1)
    SIO_ENABLE_CIRCULAR_QUEUEING = IOC_VOID Or IOC_WS2 Or 2 ' _WSAIO(IOC_WS2,2)
    SIO_FIND_ROUTE = IOC_OUT Or IOC_WS2 Or 3 ' _WSAIOR(IOC_WS2,3)
    SIO_FLUSH = IOC_VOID Or IOC_WS2 Or 4 ' _WSAIO(IOC_WS2,4)
    SIO_GET_BROADCAST_ADDRESS = IOC_OUT Or IOC_WS2 Or 5 ' _WSAIOR(IOC_WS2,5)
    SIO_GET_EXTENSION_FUNCTION_POINTER = IOC_INOUT Or IOC_WS2 Or 6 ' _WSAIORW(IOC_WS2,6)
    SIO_GET_QOS = IOC_INOUT Or IOC_WS2 Or 7 ' _WSAIORW(IOC_WS2,7)
    SIO_GET_GROUP_QOS = IOC_INOUT Or IOC_WS2 Or 8 ' _WSAIORW(IOC_WS2,8)
    SIO_MULTIPOINT_LOOPBACK = IOC_IN Or IOC_WS2 Or 9 ' _WSAIOW(IOC_WS2,9)
    SIO_MULTICAST_SCOPE = IOC_IN Or IOC_WS2 Or 10 ' _WSAIOW(IOC_WS2,10)
    SIO_SET_QOS = IOC_IN Or IOC_WS2 Or 11 ' _WSAIOW(IOC_WS2,11)
    SIO_SET_GROUP_QOS = IOC_IN Or IOC_WS2 Or 12 ' _WSAIOW(IOC_WS2,12)
    SIO_TRANSLATE_HANDLE = IOC_INOUT Or IOC_WS2 Or 13 ' _WSAIORW(IOC_WS2,13)
    
    SIO_ROUTING_INTERFACE_QUERY = IOC_INOUT Or IOC_WS2 Or 20 ' _WSAIORW(IOC_WS2,20)
    SIO_ROUTING_INTERFACE_CHANGE = IOC_OUT Or IOC_WS2 Or 22 ' _WSAIOR(IOC_WS2,22)
    
    SIO_ADDRESS_LIST_QUERY = IOC_VOID Or IOC_WS2 Or 23 ' _WSAIO(IOC_WS2,23)
    SIO_ADDRESS_LIST_CHANGE = IOC_VOID Or IOC_WS2 Or 23 ' _WSAIO(IOC_WS2,23)
    
    SIO_QUERY_TARGET_PNP_HANDLE = IOC_OUT Or IOC_WS2 Or 24 ' _WSAIOR(IOC_W32,24)
    
    SIO_RCVALL = IOC_VENDOR Or IOC_IN Or 1 ' _WSAIOW(IOC_VENDOR,1)
    SIO_RCVALL_MCAST = IOC_VENDOR Or IOC_IN Or 2 ' _WSAIOW(IOC_VENDOR,2)
    SIO_RCVALL_IGMPMCAST = IOC_VENDOR Or IOC_IN Or 3 ' _WSAIOW(IOC_VENDOR,3)
    SIO_KEEPALIVE_VALS = IOC_VENDOR Or IOC_IN Or 4 ' _WSAIOW(IOC_VENDOR,4)
    
    SIO_GET_INTERFACE_LIST = &H4004747F ' _WS_IOR ('t', 127, u_long)
End Enum

Public Type in_addr 'Requires Windows Sockets 2.0
    s_b(0 To 3) As Byte 'struct {u_char s_b1,s_b2,s_b3,s_b4;} // Address of the host formatted as four u_chars
    s_w(0 To 1) As Integer 'struct {u_short s_w1,s_w2;} // Address of the host formatted as two u_shorts
    S_addr As Long 'u_long // Address of the host formatted as a u_long
End Type ' NOTE: Whenver a function calls for a "IN_ADDR" structure, use LONG instead and pass the results of inet_addr("xxx.xxx.xxx.xxx")

Public Const sockaddr_len As Long = 16

Public Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(1 To 8) As Byte
'    sin_zero As String * 8
End Type

Public Type ip_mreq
    imr_multiaddr As in_addr
    imr_interface As in_addr
End Type

' Argument structure for SIO_KEEPALIVE_VALS
Public Type tcp_keepalive
    onoff As Long
    keepalivetime As Long
    keepaliveinterval As Long
End Type

Public Enum InterfaceStatus
    IFF_UP
    IFF_BROADCAST
    IFF_LOOPBACK
    IFF_POINTTOPOINT
    IFF_MULTICAST
End Enum

Public Type INTERFACE_INFO
    iiFlags As Long ' Type and status of the interface
    iiAddress As sockaddr ' Interface address
    iiBroadcastAddress As sockaddr ' Broadcast address
    iiNetmask As sockaddr ' Network mask
End Type

Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Const hostent_len As Long = 16

Public Type hostent
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

'/* Flag bit definitions for dwProviderFlags */
Public Enum PFL
    PFL_MULTIPLE_PROTO_ENTRIES = &H1
    PFL_RECOMMENDED_PROTO_ENTRY = &H2
    PFL_HIDDEN = &H4
    PFL_MATCHES_PROTOCOL_ZERO = &H8
End Enum

'/* Flag bit definitions for dwServiceFlags1 */
Public Enum XP1
    XP1_CONNECTIONLESS = &H1
    XP1_GUARANTEED_DELIVERY = &H2
    XP1_GUARANTEED_ORDER = &H4
    XP1_MESSAGE_ORIENTED = &H8
    XP1_PSEUDO_STREAM = &H10
    XP1_GRACEFUL_CLOSE = &H20
    XP1_EXPEDITED_DATA = &H40
    XP1_CONNECT_DATA = &H80
    XP1_DISCONNECT_DATA = &H100
    XP1_SUPPORT_BROADCAST = &H200
    XP1_SUPPORT_MULTIPOINT = &H400
    XP1_MULTIPOINT_CONTROL_PLANE = &H800
    XP1_MULTIPOINT_DATA_PLANE = &H1000
    XP1_QOS_SUPPORTED = &H2000
    XP1_INTERRUPT = &H4000
    XP1_UNI_SEND = &H8000
    XP1_UNI_RECV = &H10000
    XP1_IFS_HANDLES = &H20000
    XP1_PARTIAL_MESSAGE = &H40000
End Enum

Public Const MAX_PROTOCOL_CHAIN = 6

Public Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(MAX_PROTOCOL_CHAIN) As Long
End Type

Public Const WSAPROTOCOL_LEN = 256

Public Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * WSAPROTOCOL_LEN
End Type

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Public Type SOCKET_ADDRESS_LIST
    iAddressCount As Integer ' iAddressCount - number of address structures in the list;
    Address(1) As sockaddr ' Address - array of protocol family specific address structures.
End Type

Declare Function apiAccept Lib "ws2_32.dll" Alias "accept" (ByVal S As Long, ByRef name As sockaddr, Optional ByRef addrlen As Long) As SOCKET
' Permits an incoming connection attempt on a socket.

' AcceptEx
' Accepts a new connection, returns the local and remote address, and receives the first block of data sent by the client application.

Declare Function apiBind Lib "ws2_32.dll" Alias "bind" (ByVal S As Long, ByRef name As sockaddr, ByVal addrlen As Long) As ZERO_OR_SOCKET_ERROR
'Associates a local address with a socket.

Declare Function apiCloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal S As Long) As ZERO_OR_SOCKET_ERROR
' Closes an existing socket.

Declare Function apiConnect Lib "ws2_32.dll" Alias "connect" (ByVal S As Long, ByRef name As sockaddr, ByVal namelen As Long) As ZERO_OR_SOCKET_ERROR
' Establishes a connection to a specified socket.

'ConnectEx
' Establishes a connection to a specified socket, and optionally sends data once the connection is established. Only supported on connection-oriented sockets.

'DisconnectEx
' Closes a connection on a socket, and allows the socket handle to be reused.

'EnumProtocols
' Retrieves information about a specified set of network protocols that are active on a local host.

'freeaddrinfo
' Frees address information that the getaddrinfo function dynamically allocates in addrinfo structures.

'FreeAddrInfoEx
' Frees address information that the GetAddrInfoEx function dynamically allocates in addrinfoex structures.

'FreeAddrInfoW
' Frees address information that the GetAddrInfoW function dynamically allocates in addrinfoW structures.

'gai_strerror
' Assists in printing error messages based on the EAI_* errors returned by the getaddrinfo function.

'GetAcceptExSockaddrs
' Parses the data obtained from a call to the AcceptEx function.

'GetAddressByName
' Queries a namespace, or a set of default namespaces, to retrieve network address information for a specified network service. This process is known as service name resolution. A network service can also use the function to obtain local address information that it can use with the bind function.

'getaddrinfo
' Provides protocol-independent translation from an ANSI host name to an address.

'GetAddrInfoEx
' Provides protocol-independent name resolution with additional parameters to qualify which name space providers should handle the request.

'GetAddrInfoW
' Provides protocol-independent translation from a Unicode host name to an address.

Declare Function apiGetHostByAddr Lib "ws2_32.dll" Alias "gethostbyaddr" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
' Retrieves the host information corresponding to a network address.

Declare Function apiGetHostByNameDEPRECATED Lib "ws2_32.dll" Alias "gethostbyname" (ByVal name As String) As P_hostent
' Retrieves host information corresponding to a host name from a host database. Deprecated: use getaddrinfo instead.

Declare Function apiGetHostName Lib "ws2_32.dll" Alias "gethostname" (ByVal name As String, ByVal namelen As Long) As Long
' Retrieves the standard host name for the local computer.

'GetNameByType
' Retrieves the name of a network service for the specified service type.

'getnameinfo
' Provides name resolution from an IPv4 or IPv6 address to an ANSI host name and from a port number to the ANSI service name.

'GetNameInfoW
' Provides name resolution from an IPv4 or IPv6 address to a Unicode host name and from a port number to the Unicode service name.

Declare Function apiGetPeerName Lib "ws2_32.dll" Alias "getpeername" (ByVal S As Long, ByRef name As sockaddr, namelen As Long) As ZERO_OR_SOCKET_ERROR
' Retrieves the address of the peer to which a socket is connected.

'getprotobyname
' Retrieves the protocol information corresponding to a protocol name.

'getprotobynumber
' Retrieves protocol information corresponding to a protocol number.

'getservbyname
' Retrieves service information corresponding to a service name and protocol.

'getservbyport
'Retrieves service information corresponding to a port and protocol.

'GetService
' Retrieves information about a network service in the context of a set of default namespaces or a specified namespace.

Declare Function apiGetSockName Lib "ws2_32.dll" Alias "getsockname" (ByVal S As Long, ByRef name As sockaddr, namelen As Long) As ZERO_OR_SOCKET_ERROR
' Retrieves the local name for a socket.

Declare Function apiGetSockOpt Lib "ws2_32.dll" Alias "getsockopt" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByRef optlen As Long) As Long
' Retrieves a socket option.

'GetTypeByName
' Retrieves a service type GUID for a network service specified by name.

'Declare Function apiHToNL Lib "ws2_32.dll" Alias "htonl" (ByVal hostlong As Long) As Long
' Converts a u_long from host to TCP/IP network byte order (which is big-endian).

Declare Function apiHToNS Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
' Converts a u_short from host to TCP/IP network byte order (which is big-endian).

Declare Function apiInet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As IP32Bit
' Converts a string containing an (Ipv4) Internet Protocol dotted address into a proper address for the in_addr structure.

Declare Function apiInet_NToA Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal in_addr As Long) As P_ANSI_STR
' Converts an (IPv4) Internet network address into a string in Internet standard dotted format.

'InetNtop
' converts an IPv4 or IPv6 Internet network address into a string in Internet standard format. The ANSI version of this function is inet_ntop.

'InetPton
' Converts an IPv4 or IPv6 Internet network address in its standard text presentation form into its numeric binary form. The ANSI version of this function is inet_pton.

Declare Function apiIoCtlSocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal S As Long, ByVal cmd As Long, ByRef argp As Long) As ZERO_OR_SOCKET_ERROR
' Controls the I/O mode of a socket.

Declare Function apiListen Lib "ws2_32.dll" Alias "listen" (ByVal S As Long, ByVal backlog As Long) As ZERO_OR_SOCKET_ERROR
' Places a socket a state where it is listening for an incoming connection.

'ntohl
' Converts a u_long from TCP/IP network order to host byte order (which is little-endian on Intel processors).

Declare Function apiNToHS Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
' Converts a u_short from TCP/IP network byte order to host byte order (which is little-endian on Intel processors).

Declare Function apiRecv Lib "ws2_32.dll" Alias "recv" (ByVal S As Long, ByRef buf As Any, ByVal len_ As Long, ByVal flags As Long) As BYTES_OR_SOCKET_ERROR
' Receives data from a connected or bound socket.

Declare Function apiRecvFrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal S As Long, ByRef buf As Any, ByVal len_ As Long, ByVal flags As Long, ByRef from As sockaddr, ByRef fromlen As Long) As BYTES_OR_SOCKET_ERROR
' Receives a datagram and stores the source address.

'Declare Function apiSelect Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef TimeOut As Long) As Long
' Determines the status of one or more sockets, waiting if necessary, to perform synchronous I/O.

Declare Function apiSend Lib "ws2_32.dll" Alias "send" (ByVal S As Long, ByRef buf As Any, ByVal buflen As Long, ByVal fags As Long) As BYTES_OR_SOCKET_ERROR
' Sends data on a connected socket.

Declare Function apiSendTo Lib "ws2_32.dll" Alias "sendto" (ByVal S As Long, ByRef buf As Any, ByVal len_ As Long, ByVal flags As Long, ByRef to_ As sockaddr, ByVal tolen As Long) As BYTES_OR_SOCKET_ERROR
' Sends data to a specific destination.

'SetAddrInfoEx
' Registers a host and service name along with associated addresses with a specific namespace provider.

'SetService
' Registers or removes from the registry a network service within one or more namespaces. Can also add or remove a network service type within one or more namespaces.

Declare Function apiSetSockOpt Lib "ws2_32.dll" Alias "setsockopt" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As ZERO_OR_SOCKET_ERROR
' Sets a socket option.

Declare Function apiShutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As ZERO_OR_SOCKET_ERROR
' Disables sends or receives on a socket.

Declare Function apiSocket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As SOCKET
' Creates a socket that is bound to a specific service provider.

'TransmitFile
' Transmits file data over a connected socket handle.

'TransmitPackets
' Transmits in-memory data or file data over a connected socket.

Declare Function apiWSAAccept Lib "ws2_32.dll" Alias "WSAAccept" (ByVal S As Long, ByRef addr As sockaddr, ByVal addrlen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As SOCKET
' Conditionally accepts a connection based on the return value of a condition function, provides quality of service flow specifications, and allows the transfer of connection data.

'Declare Function apiWSAAddressToString Lib "ws2_32.dll" Alias "WSAAddressToString" (lpsaAddress As sockaddr, dwAddressLength As Long, lpProtocolInfo As Long, lpszAddressString As String, lpdwAddressStringLength As Long) As Long
' Converts all components of a sockaddr structure into a human-readable string representation of the address.

'WSAAsyncGetHostByAddr
' Asynchronously retrieves host information that corresponds to an address.

Declare Function apiWSAAsyncGetHostByName Lib "ws2_32.dll" Alias "WSAAsyncGetHostByName" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal name As String, ByRef buf As Any, ByVal buflen As Long) As HANDLE_OR_ZERO
' Asynchronously retrieves host information that corresponds to a host name.

'WSAAsyncGetProtoByName
' Asynchronously retrieves protocol information that corresponds to a protocol name.

'WSAAsyncGetProtoByNumber
' Asynchronously retrieves protocol information that corresponds to a protocol number.

'WSAAsyncGetServByName
' Asynchronously retrieves service information that corresponds to a service name and port.

'WSAAsyncGetServByPort
' Asynchronously retrieves service information that corresponds to a port and protocol.

Declare Function apiWSAAsyncSelect Lib "ws2_32.dll" Alias "WSAAsyncSelect" (ByVal S As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As ZERO_OR_SOCKET_ERROR
' Requests Windows message-based notification of network events for a socket.

Declare Function apiWSACancelAsyncRequest Lib "ws2_32.dll" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As ZERO_OR_SOCKET_ERROR
' Cancels an incomplete asynchronous operation.

Declare Function apiWSACleanup Lib "ws2_32.dll" Alias "WSACleanup" () As ZERO_OR_SOCKET_ERROR
' Terminates use of the Ws2_32.DLL.

'WSACloseEvent
' Closes an open event object handle.

'Declare Function apiWSAConnect Lib "ws2_32.dll" Alias "WSAConnect" (s As Long, Name As sockaddr, namelen As Long, lpCallerData As Long, lpCalleeData As Long, lpSQOS, lpGQOS) As Long
' Establishes a connection to another socket application, exchanges connect data, and specifies needed quality of service based on the specified FLOWSPEC structure.

'WSAConnectByList
' Establishes a connection to one out of a collection of possible endpoints represented by a set of destination addresses (host names and ports).

'WSAConnectByName
' Establishes a connection to another socket application on a specified host and port

'WSACreateEvent
' Creates a new event object.

'WSADeleteSocketPeerTargetName
' Removes the association between a peer target name and an IP address for a socket.

'WSADuplicateSocket
' Returns a structure that can be used to create a new socket descriptor for a shared socket.

'WSAEnumNameSpaceProviders
' Retrieves information about available namespaces.

'WSAEnumNameSpaceProvidersEx
' Retrieves information about available namespaces.

'WSAEnumNetworkEvents
' Discovers occurrences of network events for the indicated socket, clear internal network event records, and reset event objects (optional).

'WSAEnumProtocols
' Retrieves information about available transport protocols.

'Declare Function apiWSAEventSelect Lib "ws2_32.dll" Alias "WSAEventSelect" (ByVal s&, ByVal hEventObject&, ByVal lNetworkEvents&) As Long
' Specifies an event object to be associated with the specified set of FD_XXX network events.

'__WSAFDIsSet
' Specifies whether a socket is included in a set of socket descriptors.

Declare Function apiWSAGetLastError Lib "ws2_32.dll" Alias "WSAGetLastError" () As WSA_ERROR_CODE
' Returns the error status for the last operation that failed.

'WSAGetOverlappedResult
' Retrieves the results of an overlapped operation on the specified socket.

'WSAGetQOSByName
' Initializes a QOS structure based on a named template, or it supplies a buffer to retrieve an enumeration of the available template names.

'WSAGetServiceClassInfo
' Retrieves the class information (schema) pertaining to a specified service class from a specified namespace provider.

'WSAGetServiceClassNameByClassId
' Retrieves the name of the service associated with the specified type.

'WSAHtonl
' Converts a u_long from host byte order to network byte order.

'Declare Function apiWSAHtons Lib "ws2_32.dll" Alias "WSAHtons" (ByVal s As Long, ByVal hostshort As Long, ByVal lpnetshort As Long) As Integer
' Converts a u_short from host byte order to network byte order.

'WSAImpersonateSocketPeer
' Used to impersonate the security principal corresponding to a socket peer in order to perform application-level authorization.

'WSAInstallServiceClass
' Registers a service class schema within a namespace.

'Declare Function apiWSAIoctl Lib "ws2_32.dll" Alias "WSAIoctl" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
' Controls the mode of a socket.

'WSAJoinLeaf
' Joins a leaf node into a multipoint session, exchanges connect data, and specifies needed quality of service based on the specified structures.

'WSALookupServiceBegin
' Initiates a client query that is constrained by the information contained within a WSAQUERYSET structure.

'WSALookupServiceEnd
' Frees the handle used by previous calls to WSALookupServiceBegin and WSALookupServiceNext.

'WSALookupServiceNext
' Retrieve the requested service information.

'WSANSPIoctl
' Developers to make I/O control calls to a registered namespace.

'Declare Function apiWSANtohl Lib "ws2_32.dll" Alias "WSANtohl" (ByVal s As Long, ByVal netlong As Long, ByVal lphostlong As Long) As Long
' Converts a u_long from network byte order to host byte order.

'WSANtohs
' Converts a u_short from network byte order to host byte order.

'WSAPoll
' Determines status of one or more sockets.

'WSAProviderConfigChange
' Notifies the application when the provider configuration is changed.

'WSAQuerySocketSecurity
' Queries information about the security applied to a connection on a socket.

'WSARecv
' Receives data from a connected socket.

'Declare Function apiWSARecvDisconnect Lib "ws2_32.dll" Alias "WSARecvDisconnect" (ByVal s As Long, ByVal lpOutboundDisconnectData As Long) As Long
' Terminates reception on a socket, and retrieves the disconnect data if the socket is connection oriented.

'Declare Function apiWSARecvEx Lib "ws2_32.dll" Alias "WSARecvEx" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
' Receives data from a connected socket.

'WSARecvFrom
' Receives a datagram and stores the source address.

'WSARecvMsg
' Receives data and optional control information from connected and unconnected sockets.

'WSARemoveServiceClass
' Permanently removes the service class schema from the registry.

'WSAResetEvent
' Resets the state of the specified event object to nonsignaled.

'WSARevertImpersonation
' Terminates the impersonation of a socket peer.

'Declare Function apiWSASend Lib "ws2_32.dll" Alias "WSASend" (ByVal s As Long, lpBuffers As Long, dwBufferCount As Long, lpNumberOfBytesSent As Long, dwFlags As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
' Sends data on a connected socket.

'Declare Function apiWSASendDisconnect Lib "ws2_32.dll" Alias "WSASendDisconnect" (ByVal s As Long, ByVal lpOutboundDisconnectData As Long) As Long
' Initiates termination of the connection for the socket and sends disconnect data.

'WSASendMsg
' Sends data and optional control information from connected and unconnected sockets.

'WSASendTo
' Sends data to a specific destination, using overlapped I/O where applicable.

'WSASetEvent
' Sets the state of the specified event object to signaled.

'WSASetLastError
' Sets the error code.

'WSASetService
' Registers or removes from the registry a service instance within one or more namespaces.

'WSASetSocketPeerTargetName
' Used to specify the peer target name (SPN) that corresponds to a peer IP address. This target name is meant to be specified by client applications to securely identify the peer that should be authenticated.

'WSASetSocketSecurity
' Enables and applies security for a socket.

'WSASocket
' Creates a socket that is bound to a specific transport-service provider.

Declare Function apiWSAStartup Lib "ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Long, ByRef lpWSADATA As WSAData) As WSA_ERROR_CODE
' Initiates use of WS2_32.DLL by a process.

'WSAStringToAddress
' Converts a numeric string to a sockaddr structure.

'WSAWaitForMultipleEvents
' Returns either when one or all of the specified event objects are in the signaled state, or when the time-out interval expires.


' Old winsock 1.1 api
'Declare Function apiCloseSocket Lib "wsock32.dll" Alias "closesocket" (ByVal s As Long) As Long
'Declare Function apiSetSockOpt Lib "wsock32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
'Declare Function apiGetSockOpt Lib "wsock32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
'Declare Function apiWSAGetLastError Lib "wsock32.dll" Alias "WSAGetLastError" () As Long
'Declare Function apiWSACleanup Lib "wsock32.dll" Alias "WSACleanup" () As Long
'Declare Function apiWSAAsyncSelect Lib "wsock32.dll" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
'Declare Function apiWSAStartup Lib "wsock32.dll" Alias "WSAStartup" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
'Declare Function apiRecv Lib "wsock32.dll" Alias "recv" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
'Declare Function apiHtons Lib "wsock32.dll" Alias "htons" (ByVal hostshort As Long) As Integer
'Declare Function apiNtohs Lib "wsock32.dll" Alias "ntohs" (ByVal netshort As Long) As Integer
'Declare Function apiSocket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'Declare Function apiConnect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
'Declare Function apiInet_addr Lib "wsock32.dll" Alias "inet_addr" (ByVal cp As String) As Long
'Declare Function apiGetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
'Declare Function apiInet_ntoa Lib "wsock32.dll" Alias "inet_ntoa" (ByVal inn As Long) As Long
'Declare Function apiSend Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Declare Function apiWSAIsBlocking_1_1 Lib "wsock32.dll" Alias "WSAIsBlocking" () As Long
Declare Function apiWSACancelBlockingCall_1_1 Lib "wsock32.dll" Alias "WSACancelBlockingCall" () As Long
'Declare Function apiWSAAsyncSelect Lib "wsock32.dll" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long


Function WSAErrorDescription(errorCode As Long) As String
    Select Case errorCode
        Case WSA_INVALID_HANDLE: WSAErrorDescription = "Specified event object handle is invalid."
        Case WSA_NOT_ENOUGH_MEMORY: WSAErrorDescription = "Insufficient memory available."
        Case WSA_INVALID_PARAMETER: WSAErrorDescription = "One or more parameters are invalid."
        Case WSA_OPERATION_ABORTED: WSAErrorDescription = "Overlapped operation aborted."
        Case WSA_IO_INCOMPLETE: WSAErrorDescription = "Overlapped I/O event object not in signaled state."
        Case WSA_IO_PENDING: WSAErrorDescription = "Overlapped operations will complete later."
        Case WSAEINTR: WSAErrorDescription = "Interrupted function call."
        Case WSAEBADF: WSAErrorDescription = "File handle is not valid."
        Case WSAEACCES: WSAErrorDescription = "Permission denied."
        Case WSAEFAULT: WSAErrorDescription = "Bad address."
        Case WSAEINVAL: WSAErrorDescription = "Invalid argument."
        Case WSAEMFILE: WSAErrorDescription = "Too many open files."
        Case WSAEWOULDBLOCK: WSAErrorDescription = "Resource temporarily unavailable."
        Case WSAEINPROGRESS: WSAErrorDescription = "Operation now in progress."
        Case WSAEALREADY: WSAErrorDescription = "Operation already in progress."
        Case WSAENOTSOCK: WSAErrorDescription = "Socket operation on nonsocket."
        Case WSAEDESTADDRREQ: WSAErrorDescription = "Destination address required."
        Case WSAEMSGSIZE: WSAErrorDescription = "Message too long."
        Case WSAEPROTOTYPE: WSAErrorDescription = "Protocol wrong type for socket."
        Case WSAENOPROTOOPT: WSAErrorDescription = "Bad protocol option."
        Case WSAEPROTONOSUPPORT: WSAErrorDescription = "Protocol not supported."
        Case WSAESOCKTNOSUPPORT: WSAErrorDescription = "Socket type not supported."
        Case WSAEOPNOTSUPP: WSAErrorDescription = "Operation not supported."
        Case WSAEPFNOSUPPORT: WSAErrorDescription = "Protocol family not supported."
        Case WSAEAFNOSUPPORT: WSAErrorDescription = "Address family not supported by protocol family."
        Case WSAEADDRINUSE: WSAErrorDescription = "Address already in use."
        Case WSAEADDRNOTAVAIL: WSAErrorDescription = "Cannot assign requested address."
        Case WSAENETDOWN: WSAErrorDescription = "Network is down."
        Case WSAENETUNREACH: WSAErrorDescription = "Network is unreachable."
        Case WSAENETRESET: WSAErrorDescription = "Network dropped connection on reset."
        Case WSAECONNABORTED: WSAErrorDescription = "Software caused connection abort."
        Case WSAECONNRESET: WSAErrorDescription = "Connection reset by peer."
        Case WSAENOBUFS: WSAErrorDescription = "No buffer space available."
        Case WSAEISCONN: WSAErrorDescription = "Socket is already connected."
        Case WSAENOTCONN: WSAErrorDescription = "Socket is not connected."
        Case WSAESHUTDOWN: WSAErrorDescription = "Cannot send after socket shutdown."
        Case WSAETOOMANYREFS: WSAErrorDescription = "Too many references."
        Case WSAETIMEDOUT: WSAErrorDescription = "Connection timed out."
        Case WSAECONNREFUSED: WSAErrorDescription = "Connection refused."
        Case WSAELOOP: WSAErrorDescription = "Cannot translate name."
        Case WSAENAMETOOLONG: WSAErrorDescription = "Name too long."
        Case WSAEHOSTDOWN: WSAErrorDescription = "Host is down."
        Case WSAEHOSTUNREACH: WSAErrorDescription = "No route to host."
        Case WSAENOTEMPTY: WSAErrorDescription = "Directory not empty."
        Case WSAEPROCLIM: WSAErrorDescription = "Too many processes."
        Case WSAEUSERS: WSAErrorDescription = "User quota exceeded."
        Case WSAEDQUOT: WSAErrorDescription = "Disk quota exceeded."
        Case WSAESTALE: WSAErrorDescription = "Stale file handle reference."
        Case WSAEREMOTE: WSAErrorDescription = "Item is remote."
        Case WSASYSNOTREADY: WSAErrorDescription = "Network subsystem is unavailable."
        Case WSAVERNOTSUPPORTED: WSAErrorDescription = "Winsock.dll version out of range."
        Case WSANOTINITIALISED: WSAErrorDescription = "Successful WSAStartup not yet performed."
        Case WSAEDISCON: WSAErrorDescription = "Graceful shutdown in progress."
        Case WSAENOMORE: WSAErrorDescription = "No more results."
        Case WSAECANCELLED: WSAErrorDescription = "Call has been canceled."
        Case WSAEINVALIDPROCTABLE: WSAErrorDescription = "Procedure call table is invalid."
        Case WSAEINVALIDPROVIDER: WSAErrorDescription = "Service provider is invalid."
        Case WSAEPROVIDERFAILEDINIT: WSAErrorDescription = "Service provider failed to initialize."
        Case WSASYSCALLFAILURE: WSAErrorDescription = "System call failure."
        Case WSASERVICE_NOT_FOUND: WSAErrorDescription = "Service not found."
        Case WSATYPE_NOT_FOUND: WSAErrorDescription = "Class type not found."
        Case WSA_E_NO_MORE: WSAErrorDescription = "No more results."
        Case WSA_E_CANCELLED: WSAErrorDescription = "Call was canceled."
        Case WSAEREFUSED: WSAErrorDescription = "Database query was refused."
        Case WSAHOST_NOT_FOUND: WSAErrorDescription = "Host not found."
        Case WSATRY_AGAIN: WSAErrorDescription = "Nonauthoritative host not found."
        Case WSANO_RECOVERY: WSAErrorDescription = "This is a nonrecoverable error."
        Case WSANO_DATA: WSAErrorDescription = "Valid name, no data record of requested type."
        Case WSA_QOS_RECEIVERS: WSAErrorDescription = "QOS receivers."
        Case WSA_QOS_SENDERS: WSAErrorDescription = "QOS senders."
        Case WSA_QOS_NO_SENDERS: WSAErrorDescription = "No QOS senders."
        Case WSA_QOS_NO_RECEIVERS: WSAErrorDescription = "QOS no receivers."
        Case WSA_QOS_REQUEST_CONFIRMED: WSAErrorDescription = "QOS request confirmed."
        Case WSA_QOS_ADMISSION_FAILURE: WSAErrorDescription = "QOS admission error."
        Case WSA_QOS_POLICY_FAILURE: WSAErrorDescription = "QOS policy failure."
        Case WSA_QOS_BAD_STYLE: WSAErrorDescription = "QOS bad style."
        Case WSA_QOS_BAD_OBJECT: WSAErrorDescription = "QOS bad object."
        Case WSA_QOS_TRAFFIC_CTRL_ERROR: WSAErrorDescription = "QOS traffic control error."
        Case WSA_QOS_GENERIC_ERROR: WSAErrorDescription = "QOS generic error."
        Case WSA_QOS_ESERVICETYPE: WSAErrorDescription = "QOS service type error."
        Case WSA_QOS_EFLOWSPEC: WSAErrorDescription = "QOS flowspec error."
        Case WSA_QOS_EPROVSPECBUF: WSAErrorDescription = "Invalid QOS provider buffer."
        Case WSA_QOS_EFILTERSTYLE: WSAErrorDescription = "Invalid QOS filter style."
        Case WSA_QOS_EFILTERTYPE: WSAErrorDescription = "Invalid QOS filter type."
        Case WSA_QOS_EFILTERCOUNT: WSAErrorDescription = "Incorrect QOS filter count."
        Case WSA_QOS_EOBJLENGTH: WSAErrorDescription = "Invalid QOS object length."
        Case WSA_QOS_EFLOWCOUNT: WSAErrorDescription = "Incorrect QOS flow count."
        Case WSA_QOS_EUNKOWNPSOBJ: WSAErrorDescription = "Unrecognized QOS object."
        Case WSA_QOS_EPOLICYOBJ: WSAErrorDescription = "Invalid QOS policy object."
        Case WSA_QOS_EFLOWDESC: WSAErrorDescription = "Invalid QOS flow descriptor."
        Case WSA_QOS_EPSFLOWSPEC: WSAErrorDescription = "Invalid QOS provider-specific flowspec."
        Case WSA_QOS_EPSFILTERSPEC: WSAErrorDescription = "Invalid QOS provider-specific filterspec."
        Case WSA_QOS_ESDMODEOBJ: WSAErrorDescription = "Invalid QOS shape discard mode object."
        Case WSA_QOS_ESHAPERATEOBJ: WSAErrorDescription = "Invalid QOS shaping rate object."
        Case WSA_QOS_RESERVED_PETYPE: WSAErrorDescription = "Reserved policy QOS element type."
        Case Else: WSAErrorDescription = "Unknown error - " & errorCode
    End Select
End Function

Function WSAErrorDescriptionEx(errorCode As Long) As String
    Select Case errorCode
        Case WSA_INVALID_HANDLE: WSAErrorDescriptionEx = "An application attempts to use an event object, but the specified handle is not valid. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSA_NOT_ENOUGH_MEMORY: WSAErrorDescriptionEx = "An application used a Windows Sockets function that directly maps to a Windows function. The Windows function is indicating a lack of required memory resources. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSA_INVALID_PARAMETER: WSAErrorDescriptionEx = "An application used a Windows Sockets function which directly maps to a Windows function. The Windows function is indicating a problem with one or more parameters. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSA_OPERATION_ABORTED: WSAErrorDescriptionEx = "An overlapped operation was canceled due to the closure of the socket, or the execution of the SIO_FLUSH command in WSAIoctl. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSA_IO_INCOMPLETE: WSAErrorDescriptionEx = "The application has tried to determine the status of an overlapped operation which is not yet completed. Applications that use WSAGetOverlappedResult (with the fWait flag set to FALSE) in a polling mode to determine when an overlapped operation has completed, get this error code until the operation is complete. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSA_IO_PENDING: WSAErrorDescriptionEx = "The application has initiated an overlapped operation that cannot be completed immediately. A completion indication will be given later when the operation has been completed. Note that this error is returned by the operating system, so the error number may change in future releases of Windows."
        Case WSAEINTR: WSAErrorDescriptionEx = "A blocking operation was interrupted by a call to WSACancelBlockingCall."
        Case WSAEBADF: WSAErrorDescriptionEx = "The file handle supplied is not valid."
        Case WSAEACCES: WSAErrorDescriptionEx = "An attempt was made to access a socket in a way forbidden by its access permissions. An example is using a broadcast address for sendto without broadcast permission being set using setsockopt(SO_BROADCAST)." & vbCrLf & vbCrLf & _
            "Another possible reason for the WSAEACCES error is that when the bind function is called (on Windows�NT�4.0 with SP4 and later), another application, service, or kernel mode driver is bound to the same address with exclusive access. Such exclusive access is a new feature of Windows�NT�4.0 with SP4 and later, and is implemented by using the SO_EXCLUSIVEADDRUSE option."
        Case WSAEFAULT: WSAErrorDescriptionEx = "The system detected an invalid pointer address in attempting to use a pointer argument of a call. This error occurs if an application passes an invalid pointer value, or if the length of the buffer is too small. For instance, if the length of an argument, which is a sockaddr structure, is smaller than the sizeof(sockaddr)."
        Case WSAEINVAL: WSAErrorDescriptionEx = "Some invalid argument was supplied (for example, specifying an invalid level to the setsockopt function). In some instances, it also refers to the current state of the socket�for instance, calling accept on a socket that is not listening."
        Case WSAEMFILE: WSAErrorDescriptionEx = "Too many open sockets. Each implementation may have a maximum number of socket handles available, either globally, per process, or per thread."
        Case WSAEWOULDBLOCK: WSAErrorDescriptionEx = "This error is returned from operations on nonblocking sockets that cannot be completed immediately, for example recv when no data is queued to be read from the socket. It is a nonfatal error, and the operation should be retried later. It is normal for WSAEWOULDBLOCK to be reported as the result from calling connect on a nonblocking SOCK_STREAM socket, since some time must elapse for the connection to be established."
        Case WSAEINPROGRESS: WSAErrorDescriptionEx = "A blocking operation is currently executing. Windows Sockets only allows a single blocking operation�per- task or thread�to be outstanding, and if any other function call is made (whether or not it references that or any other socket) the function fails with the WSAEINPROGRESS error."
        Case WSAEALREADY: WSAErrorDescriptionEx = "An operation was attempted on a nonblocking socket with an operation already in progress�that is, calling connect a second time on a nonblocking socket that is already connecting, or canceling an asynchronous request (WSAAsyncGetXbyY) that has already been canceled or completed."
        Case WSAENOTSOCK: WSAErrorDescriptionEx = "An operation was attempted on something that is not a socket. Either the socket handle parameter did not reference a valid socket, or for select, a member of an fd_set was not valid."
        Case WSAEDESTADDRREQ: WSAErrorDescriptionEx = "A required address was omitted from an operation on a socket. For example, this error is returned if sendto is called with the remote address of ADDR_ANY."
        Case WSAEMSGSIZE: WSAErrorDescriptionEx = "A message sent on a datagram socket was larger than the internal message buffer or some other network limit, or the buffer used to receive a datagram was smaller than the datagram itself."
        Case WSAEPROTOTYPE: WSAErrorDescriptionEx = "A protocol was specified in the socket function call that does not support the semantics of the socket type requested. For example, the ARPA Internet UDP protocol cannot be specified with a socket type of SOCK_STREAM."
        Case WSAENOPROTOOPT: WSAErrorDescriptionEx = "An unknown, invalid or unsupported option or level was specified in a getsockopt or setsockopt call."
        Case WSAEPROTONOSUPPORT: WSAErrorDescriptionEx = "The requested protocol has not been configured into the system, or no implementation for it exists. For example, a socket call requests a SOCK_DGRAM socket, but specifies a stream protocol."
        Case WSAESOCKTNOSUPPORT: WSAErrorDescriptionEx = "The support for the specified socket type does not exist in this address family. For example, the optional type SOCK_RAW might be selected in a socket call, and the implementation does not support SOCK_RAW sockets at all."
        Case WSAEOPNOTSUPP: WSAErrorDescriptionEx = "The attempted operation is not supported for the type of object referenced. Usually this occurs when a socket descriptor to a socket that cannot support this operation is trying to accept a connection on a datagram socket."
        Case WSAEPFNOSUPPORT: WSAErrorDescriptionEx = "The protocol family has not been configured into the system or no implementation for it exists. This message has a slightly different meaning from WSAEAFNOSUPPORT. However, it is interchangeable in most cases, and all Windows Sockets functions that return one of these messages also specify WSAEAFNOSUPPORT."
        Case WSAEAFNOSUPPORT: WSAErrorDescriptionEx = "An address incompatible with the requested protocol was used. All sockets are created with an associated address family (that is, AF_INET for Internet Protocols) and a generic protocol type (that is, SOCK_STREAM). This error is returned if an incorrect protocol is explicitly requested in the socket call, or if an address of the wrong family is used for a socket, for example, in sendto."
        Case WSAEADDRINUSE: WSAErrorDescriptionEx = "Typically, only one usage of each socket address (protocol/IP address/port) is permitted. This error occurs if an application attempts to bind a socket to an IP address/port that has already been used for an existing socket, or a socket that was not closed properly, or one that is still in the process of closing. For server applications that need to bind multiple sockets to the same port number, consider using setsockopt (SO_REUSEADDR). Client applications usually need not call bind at all�connect chooses an unused port automatically. When bind is called with a wildcard address (involving ADDR_ANY), a WSAEADDRINUSE error could be delayed until the specific address is committed. This could happen with a call to another function later, including connect, listen, WSAConnect, or WSAJoinLeaf."
        Case WSAEADDRNOTAVAIL: WSAErrorDescriptionEx = "The requested address is not valid in its context. This normally results from an attempt to bind to an address that is not valid for the local computer. This can also result from connect, sendto, WSAConnect, WSAJoinLeaf, or WSASendTo when the remote address or port is not valid for a remote computer (for example, address or port 0)."
        Case WSAENETDOWN: WSAErrorDescriptionEx = "A socket operation encountered a dead network. This could indicate a serious failure of the network system (that is, the protocol stack that the Windows Sockets DLL runs over), the network interface, or the local network itself."
        Case WSAENETUNREACH: WSAErrorDescriptionEx = "A socket operation was attempted to an unreachable network. This usually means the local software knows no route to reach the remote host."
        Case WSAENETRESET: WSAErrorDescriptionEx = "The connection has been broken due to keep-alive activity detecting a failure while the operation was in progress. It can also be returned by setsockopt if an attempt is made to set SO_KEEPALIVE on a connection that has already failed."
        Case WSAECONNABORTED: WSAErrorDescriptionEx = "An established connection was aborted by the software in your host computer, possibly due to a data transmission time-out or protocol error."
        Case WSAECONNRESET: WSAErrorDescriptionEx = "An existing connection was forcibly closed by the remote host. This normally results if the peer application on the remote host is suddenly stopped, the host is rebooted, the host or remote network interface is disabled, or the remote host uses a hard close (see setsockopt for more information on the SO_LINGER option on the remote socket). This error may also result if a connection was broken due to keep-alive activity detecting a failure while one or more operations are in progress. Operations that were in progress fail with WSAENETRESET. Subsequent operations fail with WSAECONNRESET."
        Case WSAENOBUFS: WSAErrorDescriptionEx = "An operation on a socket could not be performed because the system lacked sufficient buffer space or because a queue was full."
        Case WSAEISCONN: WSAErrorDescriptionEx = "A connect request was made on an already-connected socket. Some implementations also return this error if sendto is called on a connected SOCK_DGRAM socket (for SOCK_STREAM sockets, the to parameter in sendto is ignored) although other implementations treat this as a legal occurrence."
        Case WSAENOTCONN: WSAErrorDescriptionEx = "A request to send or receive data was disallowed because the socket is not connected and (when sending on a datagram socket using sendto) no address was supplied. Any other type of operation might also return this error�for example, setsockopt setting SO_KEEPALIVE if the connection has been reset."
        Case WSAESHUTDOWN: WSAErrorDescriptionEx = "A request to send or receive data was disallowed because the socket had already been shut down in that direction with a previous shutdown call. By calling shutdown a partial close of a socket is requested, which is a signal that sending or receiving, or both have been discontinued."
        Case WSAETOOMANYREFS: WSAErrorDescriptionEx = "Too many references to some kernel object."
        Case WSAETIMEDOUT: WSAErrorDescriptionEx = "A connection attempt failed because the connected party did not properly respond after a period of time, or the established connection failed because the connected host has failed to respond."
        Case WSAECONNREFUSED: WSAErrorDescriptionEx = "No connection could be made because the target computer actively refused it. This usually results from trying to connect to a service that is inactive on the foreign host�that is, one with no server application running."
        Case WSAELOOP: WSAErrorDescriptionEx = "Cannot translate a name."
        Case WSAENAMETOOLONG: WSAErrorDescriptionEx = "A name component or a name was too long."
        Case WSAEHOSTDOWN: WSAErrorDescriptionEx = "A socket operation failed because the destination host is down. A socket operation encountered a dead host. Networking activity on the local host has not been initiated. These conditions are more likely to be indicated by the error WSAETIMEDOUT."
        Case WSAEHOSTUNREACH: WSAErrorDescriptionEx = "A socket operation was attempted to an unreachable host. See WSAENETUNREACH."
        Case WSAENOTEMPTY: WSAErrorDescriptionEx = "Cannot remove a directory that is not empty."
        Case WSAEPROCLIM: WSAErrorDescriptionEx = "A Windows Sockets implementation may have a limit on the number of applications that can use it simultaneously. WSAStartup may fail with this error if the limit has been reached."
        Case WSAEUSERS: WSAErrorDescriptionEx = "Ran out of user quota."
        Case WSAEDQUOT: WSAErrorDescriptionEx = "Ran out of disk quota."
        Case WSAESTALE: WSAErrorDescriptionEx = "The file handle reference is no longer available."
        Case WSAEREMOTE: WSAErrorDescriptionEx = "The item is not available locally."
        Case WSASYSNOTREADY: WSAErrorDescriptionEx = "This error is returned by WSAStartup if the Windows Sockets implementation cannot function at this time because the underlying system it uses to provide network services is currently unavailable. Users should check:" & vbCrLf & vbCrLf & _
            "* That the appropriate Windows Sockets DLL file is in the current path." & vbCrLf & _
            "* That they are not trying to use more than one Windows Sockets implementation simultaneously. If there is more than one Winsock DLL on your system, be sure the first one in the path is appropriate for the network subsystem currently loaded." & vbCrLf & _
            "* The Windows Sockets implementation documentation to be sure all necessary components are currently installed and configured correctly."
        Case WSAVERNOTSUPPORTED: WSAErrorDescriptionEx = "The current Windows Sockets implementation does not support the Windows Sockets specification version requested by the application. Check that no old Windows Sockets DLL files are being accessed."
        Case WSANOTINITIALISED: WSAErrorDescriptionEx = "Either the application has not called WSAStartup or WSAStartup failed. The application may be accessing a socket that the current active task does not own (that is, trying to share a socket between tasks), or WSACleanup has been called too many times."
        Case WSAEDISCON: WSAErrorDescriptionEx = "Returned by WSARecv and WSARecvFrom to indicate that the remote party has initiated a graceful shutdown sequence."
        Case WSAENOMORE: WSAErrorDescriptionEx = "No more results can be returned by the WSALookupServiceNext function."
        Case WSAECANCELLED: WSAErrorDescriptionEx = "A call to the WSALookupServiceEnd function was made while this call was still processing. The call has been canceled."
        Case WSAEINVALIDPROCTABLE: WSAErrorDescriptionEx = "The service provider procedure call table is invalid. A service provider returned a bogus procedure table to Ws2_32.dll. This is usually caused by one or more of the function pointers being NULL."
        Case WSAEINVALIDPROVIDER: WSAErrorDescriptionEx = "The requested service provider is invalid. This error is returned by the WSCGetProviderInfo and WSCGetProviderInfo32 functions if the protocol entry specified could not be found. This error is also returned if the service provider returned a version number other than 2.0."
        Case WSAEPROVIDERFAILEDINIT: WSAErrorDescriptionEx = "The requested service provider could not be loaded or initialized. This error is returned if either a service provider's DLL could not be loaded (LoadLibrary failed) or the provider's WSPStartup or NSPStartup function failed."
        Case WSASYSCALLFAILURE: WSAErrorDescriptionEx = "A system call that should never fail has failed. This is a generic error code, returned under various conditions." & vbCrLf & vbCrLf & _
            "Returned when a system call that should never fail does fail. For example, if a call to WaitForMultipleEvents fails or one of the registry functions fails trying to manipulate the protocol/namespace catalogs." & vbCrLf & vbCrLf & _
            "Returned when a provider does not return SUCCESS and does not provide an extended error code. Can indicate a service provider implementation error."
        Case WSASERVICE_NOT_FOUND: WSAErrorDescriptionEx = "No such service is known. The service cannot be found in the specified name space."
        Case WSATYPE_NOT_FOUND: WSAErrorDescriptionEx = "The specified class was not found."
        Case WSA_E_NO_MORE: WSAErrorDescriptionEx = "No more results can be returned by the WSALookupServiceNext function."
        Case WSA_E_CANCELLED: WSAErrorDescriptionEx = "A call to the WSALookupServiceEnd function was made while this call was still processing. The call has been canceled."
        Case WSAEREFUSED: WSAErrorDescriptionEx = "A database query failed because it was actively refused."
        Case WSAHOST_NOT_FOUND: WSAErrorDescriptionEx = "No such host is known. The name is not an official host name or alias, or it cannot be found in the database(s) being queried. This error may also be returned for protocol and service queries, and means that the specified name could not be found in the relevant database."
        Case WSATRY_AGAIN: WSAErrorDescriptionEx = "This is usually a temporary error during host name resolution and means that the local server did not receive a response from an authoritative server. A retry at some time later may be successful."
        Case WSANO_RECOVERY: WSAErrorDescriptionEx = "This indicates that some sort of nonrecoverable error occurred during a database lookup. This may be because the database files (for example, BSD-compatible HOSTS, SERVICES, or PROTOCOLS files) could not be found, or a DNS request was returned by the server with a severe error."
        Case WSANO_DATA: WSAErrorDescriptionEx = "The requested name is valid and was found in the database, but it does not have the correct associated data being resolved for. The usual example for this is a host name-to-address translation attempt (using gethostbyname or WSAAsyncGetHostByName) which uses the DNS (Domain Name Server). An MX record is returned but no A record�indicating the host itself exists, but is not directly reachable."
        Case WSA_QOS_RECEIVERS: WSAErrorDescriptionEx = "At least one QOS reserve has arrived."
        Case WSA_QOS_SENDERS: WSAErrorDescriptionEx = "At least one QOS send path has arrived."
        Case WSA_QOS_NO_SENDERS: WSAErrorDescriptionEx = "There are no QOS senders."
        Case WSA_QOS_NO_RECEIVERS: WSAErrorDescriptionEx = "There are no QOS receivers."
        Case WSA_QOS_REQUEST_CONFIRMED: WSAErrorDescriptionEx = "The QOS reserve request has been confirmed."
        Case WSA_QOS_ADMISSION_FAILURE: WSAErrorDescriptionEx = "A QOS error occurred due to lack of resources."
        Case WSA_QOS_POLICY_FAILURE: WSAErrorDescriptionEx = "The QOS request was rejected because the policy system couldn't allocate the requested resource within the existing policy."
        Case WSA_QOS_BAD_STYLE: WSAErrorDescriptionEx = "An unknown or conflicting QOS style was encountered."
        Case WSA_QOS_BAD_OBJECT: WSAErrorDescriptionEx = "A problem was encountered with some part of the filterspec or the provider-specific buffer in general."
        Case WSA_QOS_TRAFFIC_CTRL_ERROR: WSAErrorDescriptionEx = "An error with the underlying traffic control (TC) API as the generic QOS request was converted for local enforcement by the TC API. This could be due to an out of memory error or to an internal QOS provider error."
        Case WSA_QOS_GENERIC_ERROR: WSAErrorDescriptionEx = "A general QOS error."
        Case WSA_QOS_ESERVICETYPE: WSAErrorDescriptionEx = "An invalid or unrecognized service type was found in the QOS flowspec."
        Case WSA_QOS_EFLOWSPEC: WSAErrorDescriptionEx = "An invalid or inconsistent flowspec was found in the QOS structure."
        Case WSA_QOS_EPROVSPECBUF: WSAErrorDescriptionEx = "An invalid QOS provider-specific buffer."
        Case WSA_QOS_EFILTERSTYLE: WSAErrorDescriptionEx = "An invalid QOS filter style was used."
        Case WSA_QOS_EFILTERTYPE: WSAErrorDescriptionEx = "An invalid QOS filter type was used."
        Case WSA_QOS_EFILTERCOUNT: WSAErrorDescriptionEx = "An incorrect number of QOS FILTERSPECs were specified in the FLOWDESCRIPTOR."
        Case WSA_QOS_EOBJLENGTH: WSAErrorDescriptionEx = "An object with an invalid ObjectLength field was specified in the QOS provider-specific buffer."
        Case WSA_QOS_EFLOWCOUNT: WSAErrorDescriptionEx = "An incorrect number of flow descriptors was specified in the QOS structure."
        Case WSA_QOS_EUNKOWNPSOBJ: WSAErrorDescriptionEx = "An unrecognized object was found in the QOS provider-specific buffer."
        Case WSA_QOS_EPOLICYOBJ: WSAErrorDescriptionEx = "An invalid policy object was found in the QOS provider-specific buffer."
        Case WSA_QOS_EFLOWDESC: WSAErrorDescriptionEx = "An invalid QOS flow descriptor was found in the flow descriptor list."
        Case WSA_QOS_EPSFLOWSPEC: WSAErrorDescriptionEx = "An invalid or inconsistent flowspec was found in the QOS provider-specific buffer."
        Case WSA_QOS_EPSFILTERSPEC: WSAErrorDescriptionEx = "An invalid FILTERSPEC was found in the QOS provider-specific buffer."
        Case WSA_QOS_ESDMODEOBJ: WSAErrorDescriptionEx = "An invalid shape discard mode object was found in the QOS provider-specific buffer."
        Case WSA_QOS_ESHAPERATEOBJ: WSAErrorDescriptionEx = "An invalid shaping rate object was found in the QOS provider-specific buffer."
        Case WSA_QOS_RESERVED_PETYPE: WSAErrorDescriptionEx = "A reserved policy element was found in the QOS provider-specific buffer."
        Case Else: WSAErrorDescriptionEx = "Unknown error - " & errorCode
    End Select
End Function



