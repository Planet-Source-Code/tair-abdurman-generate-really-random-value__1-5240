<div align="center">

## Generate Really Random Value


</div>

### Description

Seems VB generate predefined values when use functions RND and RANDOMIZE(6.0),

here is minimal improvement which generate You really reandom value...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tair Abdurman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tair-abdurman.md)
**Level**          |Beginner
**User Rating**    |2.6 (13 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tair-abdurman-generate-really-random-value__1-5240/archive/master.zip)





### Source Code

```
'generate random value between minVal and maxVal inclusive
'or return -1 if any error
Public Function GenerateRandom(minVal As Long, maxVal As Long) As Long
  intr = -1
  maxVal = maxVal + 1
  If maxVal > 0 Then
  If minVal >= maxVal Then
    minVal = 0
  End If
  Else
  minVal = 0
  maxVal = 10
  End If
  Randomize (DatePart("s", Now) + DatePart("m", Now))
  Do While (intr < minVal Or intr = maxVal)
   intr = CLng(Rnd() * maxVal)
  Loop
  GenerateRandom = intr
End Function
```

