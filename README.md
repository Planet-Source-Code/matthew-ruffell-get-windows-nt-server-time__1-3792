<div align="center">

## Get Windows NT Server Time


</div>

### Description

Returns the time of day from a Windows NT workstation or server. Accounts for time zones. Requires Windows NT.
 
### More Info
 
ServerName [string] = name of target server.

Requires Windows NT.

Return the time of day.

Noen.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Ruffell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-ruffell.md)
**Level**          |Unknown
**User Rating**    |4.2 (75 globes from 18 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-ruffell-get-windows-nt-server-time__1-3792/archive/master.zip)

### API Declarations

```
Private Type TIME_OF_DAY
  t_elapsedt As Long
  t_msecs As Long
  t_hours As Long
  t_mins As Long
  t_secs As Long
  t_hunds As Long
  t_timezone As Long
  t_tinterval As Long
  t_day As Long
  t_month As Long
  t_year As Long
  t_weekday As Long
End Type
Private Declare Function NetRemoteTOD Lib "netapi32.dll" (ByVal server As String, buffer As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal Ptr As Long) As Long
```


### Source Code

```
Public Function Get_ServerTime(ByVal strServerName As String) As String
  Dim lngBuffer As Long
  Dim strServer As String
  Dim lngNet32ApiReturnCode As Long
  Dim days As Date
  Dim TOD As TIME_OF_DAY
  On Error Resume Next
  '// Get server time
  strServer = StrConv(strServerName, vbUnicode) '// Convert the server name to unicode
  lngNet32ApiReturnCode = NetRemoteTOD(strServer, lngBuffer)
  If lngNet32ApiReturnCode = 0 Then
    CopyMem TOD, ByVal lngBuffer, Len(TOD)
    days = DateSerial(70, 1, 1) + (TOD.t_elapsedt / 60 / 60 / 24) '// Convert the elapsed time since 1/1/70 to a date
    days = days - (TOD.t_timezone / 60 / 24) '// Adjust for TimeZone differences
    Get_ServerTime = days
  End If
  '// Free pointers from memory
  Call NetApiBufferFree(lngBuffer)
End Function
```

