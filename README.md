<div align="center">

## Time Conversion


</div>

### Description

To convert a given number into HH:MM:SS
 
### More Info
 
Number of seconds you are looking at converting, and bFixedLength (usually false)

We do alot of maintaining of time in the place I work, and most of the time it is in the format of large integers in number of seconds. This function will take that value and give you back the number of hours, minutes, and seconds in the format HH:MM:SS This is not something that is easily done with either VB formatting or SQL Server formatting...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Terry Hansen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/terry-hansen.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/terry-hansen-time-conversion__4-7642/archive/master.zip)





### Source Code

```
Public Function CTime(iSeconds, bFixedLength)
  Dim intHours As Integer
  Dim intMinutes As Integer
  Dim intDays As Integer
  Dim intSeconds As Integer
  Dim lngRemainder As Long
  Dim sHours As String
  Dim sMinutes As String
  Dim sSeconds As String
  If IsNull(iSeconds) Then iSeconds = 0
  '17 Jun 2002, TLH. This was added to take care of
  'negative values. This way the report will still run but
  'will show a zero value, and clue in the auditors that
  'there is a problem with that particular field...
  If iSeconds < 0 Then iSeconds = 0
  intHours = Int(iSeconds / 3600)
  lngRemainder = CLng(iSeconds Mod 3600)
  intMinutes = Int(lngRemainder / 60)
  intSeconds = CInt(lngRemainder Mod 60)
  sSeconds = String(2 - Len(CStr(intSeconds)), "0") & CStr(intSeconds)
  sMinutes = String(2 - Len(CStr(intMinutes)), "0") & CStr(intMinutes)
  'sHours = String(2-len(CStr(intHours)), "0") & CStr(intHours)
  sHours = CStr(intHours)
  CTime = sMinutes & ":" & sSeconds
  If intHours > 0 Or bFixedLength Then
    CTime = sHours & ":" & CTime
  End If
End Function
```

