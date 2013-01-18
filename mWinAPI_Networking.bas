Attribute VB_Name = "mWinAPI_Networking"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011 David Briant - see https://github.com/DangerMouseB
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'************************************************************************************************************************************************
 
Option Explicit
Option Private Module

Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 6
Public Const MIB_IF_TYPE_TOKENRING = 9
Public Const MIB_IF_TYPE_FDDI = 15
Public Const MIB_IF_TYPE_PPP = 23
Public Const MIB_IF_TYPE_LOOPBACK = 24
Public Const MIB_IF_TYPE_SLIP = 28

Public Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Public Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Byte
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Public Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type

Declare Function apiGetNetworkParams Lib "IPHlpApi.dll" Alias "GetNetworkParams" (fixedInfo As Any, pOutBufLen As Long) As Long
Declare Function apiGetAdaptersInfo Lib "IPHlpApi.dll" Alias "GetAdaptersInfo" (ipAdapterInfo As Any, pOutBufLen As Long) As Long
