<div align="center">

## GetMonDate


</div>

### Description

Get the previous week's Monday..."Kathleen A. Earley" <kearley@minnmutual.com>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-getmondate__1-517/archive/master.zip)





### Source Code

```
Function GetMonDate (CurrentDate)
  'checks to see if CurrentDate is a Date datatype
  If VarType(CurrentDate) <> 7 Then
    GetMonDate = Null
  Else
    Select Case Weekday(CurrentDate)
      Case 1   'Sunday
        GetMonDate = CurrentDate - 6
      Case 2   'Monday
        GetMonDate = CurrentDate
      Case 3 To 7 'Tuesday..Saturday
        GetMonDate = CurrentDate - Weekday(CurrentDate) + 2
    End Select
  End If
End Function
```

