<div align="center">

## A Non\-Repeating Random Number Generator


</div>

### Description

With this simple, and very fast, routine you can generate a series of non-repeating random numbers. You can select a series of 10 numbers, or a series of a million...It doesn't matter. Can be useful for image fades, deck shuffling, random tip of the day, etc. - It even tells you how long it took to generate the series.
 
### More Info
 
A popup message stating how many numbers had been mixed up and how long it took.

The larger the series of numbers the more RAM required. Uses arrays.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Lawrence](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-lawrence.md)
**Level**          |Unknown
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-lawrence-a-non-repeating-random-number-generator__1-892/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  '-------------------------------------------------------------
  ' Produces a series of X random numbers without repeating any
  '-------------------------------------------------------------
  'Results can be used by using array B(X)
  Dim A(10000) ' Sets the maximum number to pick
  Dim B(10000) ' Will be the list of new numbers (same as DIM above)
  Dim Message, Message_Style, Message_Title, Response
  'Set the original array
  MaxNumber = 10000 ' Must equal the DIM above
  For seq = 0 To MaxNumber
    A(seq) = seq
  Next seq
  'Main Loop (mix em all up)
  StartTime = Timer
  Randomize (Timer)
  For MainLoop = MaxNumber To 0 Step -1
    ChosenNumber = Int(MainLoop * Rnd)
    B(MaxNumber - MainLoop) = A(ChosenNumber)
    A(ChosenNumber) = A(MainLoop)
  Next MainLoop
  EndTime = Timer
  TotalTime = EndTime - StartTime
  Message = "The sequence of " + Format(MaxNumber, "#,###,###,###") + " numbers has been" + Chr$(10)
  Message = Message + "mixed up in a total of " + Format(TotalTime, "##.######") + " seconds!"
  Message_Style = vbInformationOnly + vbInformation + vbDefaultButton2
  Message_Title = "Sequence Generated"
  Response = MsgBox(Message, Message_Style, Message_Title)
End Sub
```

