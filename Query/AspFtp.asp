<%
Dim obj
Set obj = CreateObject("sfFTPLib.FTPConnectionSTA")
 
WScript.Echo "Object created." & vbCrLf
 
Dim result
 
' Settings
obj.Host = "220.128.140.193"
obj.Username = "upduser"
obj.Password = "joly902f"
obj.Port = 21
obj.Protocol = 0 ' ftpProtocolNormal
obj.Passive = 1
obj.MLST = 1
' Enable logging
obj.LogFile.File = "Connect.log"
 
WScript.Echo "Port = " & obj.Port & vbCrLf
 
obj.Connect()
WScript.Echo "Connected" & vbCrLf
 
' Disconnect
obj.Disconnect()
WScript.Echo "Disconnected" & vbCrLf
%>