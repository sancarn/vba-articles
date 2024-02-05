---
layout: post
title:  "Performance Tips for VBA"
published: true
authors:
  - "Sancarn"
  - "Excel-lent"
---

How to improve performance of VBA macros is a common question on the VBA reddit. In this section we will provide examples of code, and test their performance against each other. We believe it is best to work by example and hope this will give you ideas of how you can improve performance in your projects.

Performance of routines in this document are calculated using [`stdPerformance.cls`](https://github.com/sancarn/stdVBA/blob/master/src/stdPerformance.cls) which is part of the [`stdVBA` project](https://github.com/sancarn/stdVBA).

The performance test table is available [in Releases](https://github.com/Excel-lent/VBA-performance-tests/releases/download/v0.0.5-beta/Performance.tests.xlsm).

Notations:
- 2007/32: Excel 2007 32 bit
- 2010/32: Excel 2010 32 bit
- 2019/64: Excel 2019 64 bit
- 365/64: Excel 365 64 bit

<br>

> **FOREWORD**
> _This section does expect you to be familiar with VBA already, if you are still new to VBA check out the [resources](https://www.reddit.com/r/vba/wiki/resources) section_

<br>

### S1) Direct Value Setting

You don't need to select a cell to manipulate it. The following statements are equivalent

<table>
<tr>
<th>Select and set</th>
<th>Set directly</th>
</tr>
<tr>
<td>
  
```vb
cells(1, 1).select
selection.value = "hello"      
```

</td>
<td>

```vb
cells(1, 1).value = "hello"      


```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
With stdPerformance.Measure("#1 Select and set", C_MAX)
  For i = 1 To C_MAX
    cells(1, 1).select
    selection.value = "hello"
  Next
End With

With stdPerformance.Measure("#2 Set directly", C_MAX)
  For i = 1 To C_MAX
    cells(1, 1).value = "hello"
  Next
End With
```
</details>
<br>

**RESULTS**

Positive values indicate how much faster direct setting of values is compared to selecting and setting.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   54%   |   86%   |   192%  |   294%  |

<br>
<br>

### S2) Cut-Paste

Cutting is defined as the act of removing a value from one location and inserting those values into another location. In this test we will test 2 options, one using cut and paste and another using direct value assignment followed by clearing.

<table>
<tr>
<th>Cut and paste</th>
<th>Set directly + clear</th>
</tr>
<tr>
<td>
  
```vb
cells(1, 1).cut
cells(1, 2).select       
ActiveSheet.paste   
```

</td>
<td>

```vb
cells(1, 2).value = cells(1, 1).value      
cells(1, 1).Clear   

```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
const C_MAX as long = 1000
With stdPerformance.Measure("#1 Cut and paste", C_MAX)
  For i = 1 To C_MAX
    cells(1, 1).cut
    cells(1, 2).select
    ActiveSheet.paste
  Next
End With

With stdPerformance.Measure("#2 Set directly + clear", C_MAX)
  For i = 1 To C_MAX
    cells(1, 2).value = cells(1, 1).value
    cells(1, 1).Clear
  Next
End With
```
</details>
<br>

**RESULTS**

As you can see in the results, direct value assignment is significantly faster than cutting and pasting (and it doesn't alter the clipboard):

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   235%  |  -405%  |   -17%  |  1505%  |

<br>
<br>

### S3) Use arrays

Using arrays is a common bit of advice to any beginners to improve performance. This demonstrates this fact.

<table>
<tr>
<th>Looping through individual cells setting values</th>
<th>Exporting array in bulk, set values, Import array in bulk</th>
</tr>
<tr>
<td>
  
```vb
Cells(i, 1).Value2 = Rnd()

```

</td>
<td>

```vb
v(i, 1) = Rnd()
r.Value2 = v
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 50000
With stdPerformance.Measure("#1 Looping through individual cells setting values", C_MAX)
  For i = 1 To C_MAX
    Cells(i, 1).Value2 = Rnd()
  Next
End With

With stdPerformance.Measure("#2 Exporting array in bulk, set values, Import array in bulk", C_MAX)
  'GetRange
  Set r = ActiveSheet.Range("A1").Resize(C_MAX, 10)

  'Values of Range --> Array
  v = r.Value2
  
  'Modify array
  For i = 1 To C_MAX  'Using absolute just to be clear no extra work is done, but you'd usually use ubound(v,1)
    v(i, 1) = Rnd()
  Next

  'Values of array  -->  Range
  r.Value2 = v
End With
```
</details>
<br>

**RESULTS**

It can be concluded that to insert multiple values into a range it is always better to use `Range(...).value` (positive values in table below show increasing of relative performance).

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   196%  |   55%   |   269%  |   305%  |

<br>
<br>

### S4) Options

#### S4a) Option `ScreenUpdating`

Another common suggestion is that `ScreenUpdating` is turned off. `ScreenUpdating` will generally affect any code which could cause a visual change on the worksheet e.g. Setting a cell's value, changing the color of a cell, etc.

<table>
<tr>
<th>Looping through individual cells setting values</th>
<th>w/ ScreenUpdating</th>
</tr>
<tr>
<td>
  
```vb
cells(1, 1).value = Empty




```

</td>
<td>

```vb
Application.ScreenUpdating = False      
' Loop
cells(1, 1).value = Empty
' End of loop
Application.ScreenUpdating = True
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 1000
With stdPerformance.Measure("#1 Looping through individual cells setting values", C_MAX)
  For i = 1 To C_MAX
    cells(1, 1).value = Empty
  Next
End With
With stdPerformance.Measure("#2 w/ ScreenUpdating", C_MAX)
  Application.ScreenUpdating = False
    For i = 1 To C_MAX
      cells(1, 1).value = Empty
    Next
  Application.ScreenUpdating = True
End With
```
</details>
<br>

**RESULTS**

Positive values indicate an increase in performance if `ScreenUpdating` is disabled relative to `ScreenUpdating = true`.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   181%  |  3391%  |    2%   |   115%  |

<br>

It is important to note however that repeatedly setting cells is bad practice and if you were doing this using arrays toggling `ScreenUpdating` might have a negative impact as whenever `ScreenUpdating` is set to true it will forcefully update the screen:

<table>
<tr>
<th>w/ ScreenUpdating within loop</th>
</tr>
<tr>
<td>
  
```vb
' Loop
Application.ScreenUpdating= False     
cells(1, 1).value = Empty
Application.ScreenUpdating= True
' End of loop
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX as Long = 1000
With stdPerformance.Measure("#2 w/ ScreenUpdating within loop", C_MAX)
    For i = 1 To C_MAX
      Application.ScreenUpdating= False
        cells(1, 1).value = Empty
      Application.ScreenUpdating= True
    Next
End With
```
</details>
<br>

**RESULTS**
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  3541%  |  2560%  |    4%   |  3125%  |

<br>
<br>

#### S4b) Option `EnableEvents`

On many occasions people online will claim that setting `EnableEvents`to false will greatly speed up your code. But does it really?

<table>
<tr>
<th>Looping through individual cells setting values</th>
<th>w/ EnableEvents</th>
</tr>
<tr>
<td>
  
```vb
cells(1, 1).value = Empty




```

</td>
<td>

```vb
Application.EnableEvents = False      
' Loop
cells(1, 1).value = Empty
' End of loop
Application.EnableEvents = True
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
With stdPerformance.Measure("#1 Looping through individual cells setting values", C_MAX)
  For i = 1 To C_MAX
    cells(1, 1).value = Empty
  Next
End With

With stdPerformance.Measure("#2 w/ EnableEvents", C_MAX)
  Application.EnableEvents = False
    For i = 1 To C_MAX
      cells(1, 1).value = Empty
    Next
  Application.EnableEvents = True
End With
```
</details>
<br>

**RESULTS**

The results of relative performance due to setting `EnableEvents`to false are shown below: 
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |    1%   |    0%   |   12%   |   71%   |

It is important to note however that repeatedly setting cells is bad practice and if you were doing this virtually toggling enable events is only going to have negligible, if not negative impacts depending where the call is made:

<table>
<tr>
<th>w/ EnableEvents within loop</th>
</tr>
<tr>
<td>
  
```vb
' Loop
Application.EnableEvents= False     
cells(1, 1).value = Empty
Application.EnableEvents= True
' End of loop
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
With stdPerformance.Measure("#3 w/ EnableEvents within loop", C_MAX)
    For i = 1 To C_MAX
      Application.EnableEvents = False
        cells(1, 1).value = Empty
      Application.EnableEvents = True
    Next
End With
```
</details>
<br>

**RESULTS**
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   -1%   |    0%   |    1%   |  5569%  |

#### S4c) Option `Calculation`

Many sources suggest turning off calculation, and performing a manual calculation after edits is faster than continual calculation. In this test we'll test this hypothesis:

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 50000
Dim rCell As Range: Set rCell = wsTmp1.Range("A1")
wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
With stdPerformance.Measure("Calculation = xlCalculationAutomatic", C_MAX)
  For i = 1 To C_MAX
    rCell.Formula = "=1"
  Next
End With
With stdPerformance.Measure("Calculation = xlCalculationManual", C_MAX)
  Application.Calculation = xlCalculationManual
    For i = 1 To C_MAX
      rCell.Formula = "=1"
    Next
  Application.Calculation = xlCalculationAutomatic
  Application.Calculate
End With
```
</details>
<br>

**RESULTS**

It indeed appears to be the case that disabling calculation, and enabling it afterwards has a significant impact on performance of formula evaluation.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  1119%  |   298%  |   200%  |   203%  |

It should be noted however that this performance impact is only noticeable when changing cells on the worksheet. Either by setting the value directly (range.value) or by setting the formula (range.formula). If instead you are setting the value of an array, changing `Calculation` mode has a slightly negative impact on the code. It also should be noted that in `xlCalculationAutomatic` mode every cells with formulas will be recalculated on every change of the value in cell. Therefore the most performence gain will be reached if you have a lot of formulas on the sheet and apply `xlCalculationManual`. One should also note that if you applied `xlCalculationManual` and read some intermediate values from the sheet, than you can obtain wrong results since the formulas were not recalculated. 

<br>
<br>

#### S4d) Option `DisplayStatusBar`

Some sources suggest turning off the StatusBar will speed up macros, because Excel will no longer have to perform status-bar updates. In this test we'll test that hypothesis:

<table>
<tr>
<th>Don't change StatusBar</th>
<th>Change StatusBar setting</th>
</tr>
<tr>
<td>
  
```vb
v = ""




```

</td>
<td>

```vb
Application.DisplayStatusBar = False      
' Loop
v = ""
' End of loop
Application.DisplayStatusBar = True
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 5000000
With stdPerformance.Optimise   'Disable screen updating and application events
  Dim v
  With stdPerformance.Measure("Don't change StatusBar", C_MAX)
    For i = 1 To C_MAX
      v = ""
    Next
  End With
  With stdPerformance.Measure("Change StatusBar setting", C_MAX)
    Application.DisplayStatusBar = False
      For i = 1 To C_MAX
        v = ""
      Next
    Application.DisplayStatusBar = True
  End With
End With
```
</details>
<br>

**RESULTS**

As shown, changing the status bar will only increase the time the macro runs.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   -2%   |    0%   |    5%   |   -8%   |

One may suggest that this is only because the statusBar isn't being set, and/or because we are disabling screen updating and application events. So let's remove those and test again:

<table>
<tr>
<th>Change StatusBar setting</th>
<th>Don't change StatusBar</th>
</tr>
<tr>
<td>
  
```vb
Application.DisplayStatusBar = False      
' Loop
Application.StatusBar = i
' End of loop
Application.DisplayStatusBar = True
```

</td>
<td>

```vb
Application.StatusBar = i     




```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 5000
With stdPerformance.Measure("Change StatusBar setting", C_MAX)
  Application.DisplayStatusBar = False
    For i = 1 To C_MAX
      Application.StatusBar = i
    Next
  Application.DisplayStatusBar = True
End With
With stdPerformance.Measure("Don't change StatusBar", C_MAX)
  For i = 1 To C_MAX
    Application.StatusBar = i
  Next
End With
```
</details>
<br>

**RESULTS**

We still see changing the status bar display slows down the code.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   165%  |   89%   |  -11%   |   -1%   |

So the long and short of it is, it's best to avoid toggling the status bar unless you really want that feature.

<br>
<br>

#### S4e) Option `DisplayPageBreaks`

Some sources suggest that Option `DisplayPageBreaks` can be used to prevent Excel from calculating page breaks, thus speeding up macro execution.

Given that this example is much like others we'll only display results this time.

In a similar vein to `DisplayStatusBar`, displaying page breaks appears to have a negative impact on performance rather than a positive one.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   -15%  |   -8%   |  -19%   |   -6%   |

<br>
<br>

#### S4f) Option `EnableAnimations`

Some sources claim toggling `EnableAnimations` can speed up performance.

Given that this example is much like others we'll only display results this time.

Much like `DisplayStatusBar`, changing `EnableAnimations` appears to have a negative impact on performance rather than a positive one.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |    6%   |    1%   |   7%    |   -2%   |

<br>
<br>

#### S4g) Option `PrintCommunication`

Some sources claim toggling `PrintCommunication` can speed up performance.

Given that this example is much like others we'll only display results this time.

Much like `DisplayStatusBar`, changing `PrintCommunication` appears to have a negative impact on performance rather than a positive one.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |    -    |    0%   |   19%   |   -2%   |

<br>
<br>

### S5) Using `With` statements

It is often advised to use a `With` statement to increase performance, but how significantly does this affect performance? For this we will test 3 scenarios:

1. Searching an object directly.
2. Using a `With` statement.
3. Creating a helper variable to do the data lookup with.

The results vary wildly depending on how deep within an object you go, for this reason we have created a special helper class to perform this test:

<table>
<tr>
<th>D3, No `With` block</th>
<th>D3, `With` block</th>
<th>D3, Variable</th>
</tr>
<tr>
<td>
  
```vb
v = o.Self.Self.Data      


```

</td>
<td>

```vb
With o.Self.Self      
  v = .Data
End With
```

</td>
<td>

```vb
Set x = o.Self.Self      
v = x.Data

```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Self As oClass
Public Data As Boolean    

Private Sub Class_Initialize()
  Set Self = Me
  Data = True
End Sub
```
</details>
<br>

The above has a loop which allows us to continuously bury down into depths of the object to see how these affect our results.

Tests have been done up to 10 levels of depth. Here's the tests for the first 3 levels:

<details>
  <summary>Source code</summary>

```vb
Const C_MAX As Long = 10000000
Dim o As New oClass
Dim x As oClass

'Depth of 1
With stdPerformance.Measure("D1,No with block", C_MAX)
    For i = 1 To C_MAX
      v = o.Data
    Next
End With
With stdPerformance.Measure("D1,With block", C_MAX)
    With o
      For i = 1 To C_MAX
        v = .Data
      Next
    End With
End With
With stdPerformance.Measure("D1,Variable", C_MAX)
    Set x = o
    For i = 1 To C_MAX
      v = x.Data
    Next
End With

'Depth of 2    
With stdPerformance.Measure("D2,No with block", C_MAX)
    For i = 1 To C_MAX
      v = o.Self.Data
    Next
End With
With stdPerformance.Measure("D2,With block", C_MAX)
    With o.Self
      For i = 1 To C_MAX
        v = .Data
      Next
    End With
End With
With stdPerformance.Measure("D2,Variable", C_MAX)
    Set x = o.Self
    For i = 1 To C_MAX
      v = x.Data
    Next
End With

'Depth of 3
With stdPerformance.Measure("D3,No with block", C_MAX)
    For i = 1 To C_MAX
      v = o.Self.Self.Data
    Next
End With
With stdPerformance.Measure("D3,With block", C_MAX)
    With o.Self.Self
      For i = 1 To C_MAX
        v = .Data
      Next
    End With
End With
With stdPerformance.Measure("D3,Variable", C_MAX)
    Set x = o.Self.Self
    For i = 1 To C_MAX
      v = x.Data
    Next
End With
...
```
</details>
<br>

**RESULTS**

You'll note that `With` blocks behave exactly like use of variables. This is because internally `With` statements do set a special variable to the value assigned to them. In general you can see that both `With` statements and variables scale adequately to larger and larger depths. On the other hand if no with blocks or variables are used, the time taken scales up indefinitely:

<details>
  <summary>Measurement time in milliseconds for D1 till D10. The results of relative performance in % (below) are based on "D10,No with block" and "D10,With block".</summary>

    D1,No with block: 250 ms (0.025µs per operation)
    D1,With block: 250 ms (0.025µs per operation)
    D1,Variable: 234 ms (0.0234µs per operation)
    D2,No with block: 563 ms (0.0563µs per operation)
    D2,With block: 234 ms (0.0234µs per operation)
    D2,Variable: 234 ms (0.0234µs per operation)
    D3,No with block: 844 ms (0.0844µs per operation)
    D3,With block: 250 ms (0.025µs per operation)
    D3,Variable: 235 ms (0.0235µs per operation)
    D4,No with block: 1156 ms (0.1156µs per operation)
    D4,With block: 234 ms (0.0234µs per operation)
    D4,Variable: 250 ms (0.025µs per operation)
    D5,No with block: 1610 ms (0.161µs per operation)
    D5,With block: 250 ms (0.025µs per operation)
    D5,Variable: 234 ms (0.0234µs per operation)
    D6,No with block: 1812 ms (0.1812µs per operation)
    D6,With block: 250 ms (0.025µs per operation)
    D6,Variable: 235 ms (0.0235µs per operation)
    D7,No with block: 2109 ms (0.2109µs per operation)
    D7,With block: 234 ms (0.0234µs per operation)
    D7,Variable: 250 ms (0.025µs per operation)
    D8,No with block: 2391 ms (0.2391µs per operation)
    D8,With block: 234 ms (0.0234µs per operation)
    D8,Variable: 250 ms (0.025µs per operation)
    D9,No with block: 2703 ms (0.2703µs per operation)
    D9,With block: 235 ms (0.0235µs per operation)
    D9,Variable: 234 ms (0.0234µs per operation)
    D10,No with block: 3031 ms (0.3031µs per operation)
    D10,With block: 235 ms (0.0235µs per operation)
    D10,Variable: 250 ms (0.025µs per operation)

An chart view of this data can be found [here](https://i.imgur.com/Aid031a.png).
</details>

<br>

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  1215%  |  1437%  |  1479%  |  1190%  |

**A few important notes:**

1) We only get performance benefits when depth is greater than 2. However in this case there are no negative performance impacts either.
2) This demonstrates that use of variables can still obtain the performance benefits of `With` blocks, which can help get around some limitations with `With` blocks. E.G. imagine setting values in a sheet to values in an object:

<details>
  <summary>Source code</summary>

```vb
With ThisWorkbook.Sheets(1)
  With myObject.someData
    'This will syntax error because Range property is part of Worksheet object
    'not myObject.
    .Range("...").value = ,dataProperty
  End with
End with
```
</details>
<br>

Here we can use variables to work around the limitation, while still getting performance benefits:

<details>
  <summary>Source code</summary>

```vb
Dim ws as worksheet: set ws = ThisWorkbook.Sheets(1)
With myObject.someData
  ws.Range("...").value = ,dataProperty
End with
```
</details>
<br>
<br>

### S6) Late vs Early Binding

It is common knowledge that Late binding is slower than Early binding.

<table>
<tr>
<th>Late bound creation</th>
<th>Early bound creation</th>
<th>Late bound calls</th>
<th>Early bound calls</th>
</tr>
<tr>
<td>
  
```vb
Set r1 = CreateObject("VBScript.Regexp")      
```

</td>
<td>

```vb
Set r2 = New VBScript_RegExp_55.RegExp      
```

</td>
<td>

```vb
Set r1 = CreateObject("VBScript.Regexp")      
r1.pattern = "something"
```

</td>
<td>

```vb
Set r2 = New VBScript_RegExp_55.RegExp      
r2.pattern = "something"
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Dim r1 As Object, r2 As VBScript_RegExp_55.RegExp
With stdPerformance.Measure("#A-1 Late bound creation", C_MAX)
  For i = 1 To C_MAX
    Set r1 = CreateObject("VBScript.Regexp")
  Next
End With
With stdPerformance.Measure("#A-2 Early bound creation", C_MAX)
  For i = 1 To C_MAX
    Set r2 = New VBScript_RegExp_55.RegExp
  Next
End With

Set r1 = CreateObject("VBScript.Regexp")
With stdPerformance.Measure("#B-1 Late bound calls", C_MAX2)
  For i = 1 To C_MAX2
    r1.pattern = "something"
  Next
End With

Set r2 = New VBScript_RegExp_55.RegExp
With stdPerformance.Measure("#B-2 Early bound calls", C_MAX2)
  For i = 1 To C_MAX2
    r2.pattern = "something"
  Next
End With
```
</details>
<br>

**RESULTS**

Relative performance of late bound vs early bound creation:
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |    7%   |   -9%   |   -6%   |    3%   |


Relative performance of late bound vs early bound calls:
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   639%  |   609%  |  1018%  |   698%  |

Method calls on objects defined using `Dim ... as Object` (late-bound objects) are significantly slower than method calls on strictly typed objects e.g. `Dim ... as Regexp` or `Dim ... as Dictionary`.

Contrary to popular belief this has nothing todo with the use of `CreateObject(...)`. In fact we can do a test:

<details>
  <summary>Source code</summary>

```vb
Set r2 = CreateObject("VBScript.Regexp")
With stdPerformance.Measure("#B-3 Early bound calls via CreateObject", C_MAX2)
  For i = 1 To C_MAX2
    r2.pattern = "something"
  Next
End With
```
</details>
<br>

This will perform exactly the same as `#B-2`. 

It is anticipated that `Object` performs so much slower than strict types because `Object` uses the [`IDispatch`](https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch) interface of an object. I.E. whenever an object method is called `IDispatch::GetIDsOfNames` is called followed by `IDispatch::Invoke`. Due to the poor performance it is expected that the IDs are not cached for use within the function body, and are instead called each time a function is called. A faster approach would be to call `IDispatch::Invoke` ourselves with known `DispIDs` in scenarios where the object type is known. E.G.

<details>
  <summary>Source code</summary>

```vb
'Note `DispGetIDsOfNames` and `DispLetProp` are purely fictional functions. But would likely
'improve performance.
set r1 = CreateObject("VBScript.Regexp")
Dim ids() as long: ids = DispGetIDsOfNames(r1, Array("pattern"))
With stdPerformance.Measure("#B-3 Late bound fast dispatch", C_MAX2)
  For i = 1 To C_MAX2
    Call DispLetProp(r2, ids(0), "something")
  Next
End With
```
</details>
<br>

**CAVEATS**

Using direct references to types makes your code less portable than using `Object` type with `CreateObject`.

<br>
<br>

### S7) `ByRef` vs `ByVal`

<table>
<tr>
<th>`ByVal`</th>
<th>`ByRef`</th>
</tr>
<tr>
<td>
  
```vb
Sub testByVal(ByVal v)
    wsTmp1.Cells(1, 1).Value2 = Rnd()      
End Sub
Call testByVal(v)
```

</td>
<td>

```vb
Sub testByRef(ByRef v)
    wsTmp1.Cells(1, 1).Value2 = Rnd()      
End Sub
Call testByRef(v)
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
v = Split(Space(1000)," ")
With stdPerformance.measure("#1 `ByVal`", C_MAX)
  For i = 1 to C_MAX
    Call testByVal(v)
  Next
End With
With stdPerformance.measure("#2 `ByRef`", C_MAX)
  For i = 1 to C_MAX
    Call testByRef(v)
  Next
End With
```
</details>
<br>

**RESULTS**

Relative performance of `ByRef` vs `ByVal`. Positive values means `ByRef` is faster.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   19%   |   74%   |   218%  |    -    |

<br>
<br>

### S8) `Module` vs `Class`

Many people, myself included, use Classes extensively in our projects to augment data. However using Object-Oriented programming is only an option and the same sort of things can be done in a purely procedural-oriented way using `Type` and `Module` functions.

<table>
<tr>
<th>Class creation</th>
<th>Module creation</th>
<th>Class method calls</th>
<th>Module method calls</th>
</tr>
<tr>
<td>
  
```vb
Set c1 = Car1.Create("hello", 10, 2)      
```

</td>
<td>

```vb
c2 = Car2.Car_Create("hello", 10, 2)      
```

</td>
<td>

```vb
Call c1.Tick      
```

</td>
<td>

```vb
Call Car2.Car_Tick(c2)      
```

</td>

</tr>
</table>

<details>
  <summary>Source code</summary>

First we need to create our `Class` and `Module`:

**Car1 class**

<span style="color:red">B-1 and B-2 are removed because the performance is nearly the same.</span>

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Car1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public name As String
Public SeatCount As Long
Public DoorCount As Long
Public Distance As Long

Public Function Create(ByVal sName As String, iSeatCount As Long, iDoorCount As Long) As Car1
  Set Create = New Car1
  Call Create.protInit(sName, iSeatCount, iDoorCount)
End Function
Friend Sub protInit(ByVal sName As String, iSeatCount As Long, iDoorCount As Long)
  name = sName
  SeatCount = iSeatCount
  DoorCount = iDoorCount
End Sub

Public Sub Tick()
  Distance = Distance + 1
End Sub
```

**Car2 Module**

```vb
Attribute VB_Name = "Car2"
Public Type CarData
  name As String
  SeatCount As Long
  DoorCount As Long
  Distance As Long
End Type

Public Function Car_Create(ByVal sName As String, iSeatCount As Long, iDoorCount As Long) As CarData
  Car_Create.name = sName
  Car_Create.SeatCount = iSeatCount
  Car_Create.DoorCount = iDoorCount
End Function

Public Sub Car_Tick(ByRef data As CarData)
  data.Distance = data.Distance + 1
End Sub
```

Now we can test the performance of these 2 modules as follows:

```vb
Const C_MAX  As Long = 500000

Dim i As Long
Dim c1 As Car1
Dim c2 As Car2.CarData

With stdPerformance.Measure("A-#1 Object creation (Class)", C_MAX)
  For i = 1 To C_MAX
    Set c1 = Car1.Create("hello", 10, 2)
  Next
End With
With stdPerformance.Measure("A-#2 Object creation (Module)", C_MAX)
  For i = 1 To C_MAX
    c2 = Car2.Car_Create("hello", 10, 2)
  Next
End With

'Objects for instance tests
Set c1 = Car1.Create("hello", 10, 2)
c2 = Car2.Car_Create("hello", 10, 2)

'Test calling public methods speeds
With stdPerformance.Measure("B-#1 Object method calls (Class)", C_MAX)
  For i = 1 To C_MAX
    Call c1.Tick
  Next
End With
With stdPerformance.Measure("B-#2 Object method calls (Module)", C_MAX)
  For i = 1 To C_MAX
    Call Car2.Car_Tick(c2)
  Next
End With

'Test calling private method speeds
With stdPerformance.Measure("C-#1 Object private method calls (Class)", C_MAX)
  Call c1.TestPriv(C_MAX)
End With
With stdPerformance.Measure("C-#2 Object private method calls (Module)", C_MAX)
  Call Car2.TestPriv(c2, C_MAX)
End With
```
</details>
<br>

**RESULTS**

From these results we can see that class creation and initialisation is clearly significantly slower than module struct creation, however the calling of methods (and thus also setting of properties) are equally performant.

`Module` vs `Class` Object creation
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  540%   |   581%  |   670%  |   463%  |

`Module` vs `Class` Object method calls
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   45%   |   72%   |   66%   |    0%   |

<br>
<br>

### S9) `Variant` vs Typed Data

It is often suggested that you should use Typed data wherever possible for performance reasons.

<table>
<tr>
<th>Variant</th>
<th>Type</th>
</tr>
<tr>
<td>
  
```vb
Dim v() As Variant      
v(i) = i
```

</td>
<td>

```vb
Dim l() As Long      
l(i) = i
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX  As Long = 5000000

Dim i As Long
With stdPerformance.Measure("#1 - Variant", C_MAX)
  Dim v() As Variant
  ReDim v(1 To C_MAX)
  For i = 1 To C_MAX
    v(i) = i
  Next
End With
With stdPerformance.Measure("#2 - Type", C_MAX)
  Dim l() As Long
  ReDim l(1 To C_MAX)
  For i = 1 To C_MAX
    l(i) = i
  Next
```
</details>
<br>

**RESULTS**

In general this is only a real benefit when dealing with very large datasets (>5 million cells), but it definitely can happen so if you can, always type your data.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   16%   |   61%   |   81%   |   32%   |

<br>
<br>

### S10) Bulk range operations

It was anticipated that using `Application.Union()` to create a big range, and then calling delete on this single range would be faster than deleting each row individually. This turned out incorrect.

<table>
<tr>
<th>Delete rows 1 by 1</th>
<th>Delete all rows in a single operation less branching</th>
</tr>
<tr>
<td>
  
```vb
Rows(i).Delete      
```

</td>
<td>

```vb
rng.Delete      
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
Const C_MAX as long = 5000

Dim i As Long
Range("A1:X" & C_MAX).value = "Some cool data here"

With stdPerformance.Measure("#1 Delete rows 1 by 1", C_MAX)
  For i = C_MAX To 1 Step -1
    'Delete only even rows
    If i Mod 2 = 0 Then
      Rows(i).Delete
    End If
  Next
End With

With stdPerformance.Measure("#2 Delete all rows in a single operation less branching", C_MAX)
  Set rng = cells(Rows.count, 1)
  For i = 1 To C_MAX
    If i Mod 2 = 0 Then
      Set rng = Application.Union(rng, Rows(i))
    End If
  Next i
  Set rng = Application.Intersect(rng, Range("1:" & C_MAX))
  rng.Delete
End With
```
</details>
<br>

**RESULTS**

Positive values show how much faster deleting of all rows in a single operation than deleting of rows 1 by 1 is.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   -19%  |    6%   |    5%   |  -83%   |

<br>
<br>

### S11) When to use advanced filters

In many performance tutorials you will be told to use the Advanced filter to speed up filter operations on data. One might think that Advanced is so much faster than hand crafted techniques that it is always better to use it. 

In order to test this out we'll need some data. In this tutorial I'll use the following function to generate a section of data for us to filter on:

<details>
  <summary>Source code</summary>

```vb
'Obtain an array of data to C_MAX size.
'@param {Long} Max length of data
'@returns {Variant(nx2)} Returns an array of data 2 columns wide and n rows deep. 
Public Function getArray(C_MAX As Long) As Variant
  Dim arr() As Variant
  ReDim arr(1 To C_MAX, 1 To 2)
  
  arr(1, 1) = "ID"
  arr(1, 2) = "Key"
  Dim i As Long
  For i = 2 To C_MAX
    'ID
    arr(i, 1) = i
    Select Case True
      Case i Mod 17 = 0: arr(i, 2) = "A"
      Case i Mod 13 = 0: arr(i, 2) = "B"
      Case i Mod 11 = 0: arr(i, 2) = "C"
      Case i Mod 7 = 0: arr(i, 2) = "D"
      Case Else
        arr(i, 2) = "E"
    End Select
  Next
  getArray = arr
End Function
```
</details>
<br>

In this test, we will experiment with 3 different scenarios:

1. We have data in `Sheet1` and we want to obtain a filtered copy in `Sheet2`.
2. We already have an array in VBA and we want to filter it's contents.
3. We have data in `Sheet1` and we want to obtain a filtered copy as an array.

#### S11a) Sheet to Sheet

In scenario 1, we have some data in `Sheet1` and we'd like to obtain a copy of this data on `Sheet2`, filtered where the Key = "A".

<details>
  <summary>Source code</summary>

```vb
'In this test we are going to try the same thing but where the data is already on the sheet
Sub scenario1()
  'Initialisation. Initialise a sheet containing data and an output sheet containing headers.
  Dim arr As Variant: arr = getArray(1000000)
  Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
  Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
  Sheet1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
  Sheet2.UsedRange.Clear

  'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
  With stdPerformance.Measure("#1 Advanced filter and copy")
    'Choose headers
    Sheet2.Range("A1:B1").Value = Array("ID", "Key")
    
    'Choose filter
    HiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
    'filter and copy data
    With Sheet1.Range("A1").CurrentRegion
      Call .AdvancedFilter(xlFilterCopy, HiddenSheet.Range("A1:A2"), Sheet2.Range("A1:B1"))
    End With

    'Cleanup
    HiddenSheet.UsedRange.Clear
  End With
  
  With stdPerformance.Optimise()
    'Copy data from sheet into an array,
    'loop through rows, move filtered rows to top of array,
    'only return required size of array
    Sheet2.UsedRange.Clear
    With stdPerformance.Measure("#2 Use of array")
      Dim v: v = Sheet1.Range("A1").CurrentRegion.Value
      Dim iRowLen As Long: iRowLen = UBound(v, 1)
      Dim iColLen As Long: iColLen = UBound(v, 2)
      
      Dim i As Long, j As Long, iRet As Long
      iRet = 1
      For i = 2 To iRowLen
        If v(i, 2) = "A" Then
          iRet = iRet + 1
          If iRet < i Then
            For j = 1 To iColLen
              v(iRet, j) = v(i, j)
            Next
          End If
        End If
      Next
      
      Sheet2.Range("A1").Resize(iRet, iColLen).Value = v
    End With
    
  End With
End Sub
```
</details>
<br>

**RESULTS**

Positive values show how much faster advanced filter and copy than use of array is.
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   982%  |   962%  |   238%  |   30%   |

<br>
<br>

#### S11b) Array to Array

In scenario 2, we have some data in an array, and want a filtered set of IDs as an output. Is it worthwhile writing this data to the sheet and using AdvancedFilter? Or would it be better to filter the array in memory?

<details>
  <summary>Source code</summary>

```vb
Sub scenario2()
  'Obtain test data
  Dim arr As Variant: arr = getArray(100000)
  
  'Some of these tests may take some time, so optimise to ensure these don't have an impact
  With stdPerformance.Optimise()
    'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
    Dim vResult
    With stdPerformance.Measure("#1 Advanced filter and copy - array result")
      'Get data dimensions
      Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
      Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
      
      'Create filters data
      HiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
      With HiddenSheet.Range("A4").Resize(iNumRow, iNumCol)
        'Dump data to sheet
        .Value = arr
        
        'Call advanced filter
        Call .AdvancedFilter(xlFilterInPlace, HiddenSheet.Range("A1:A2"))
        
        'Get result
        .Resize(, 1).Copy HiddenSheet.Range("D4")
        vResult = HiddenSheet.Range("D4").CurrentRegion.Value
      End With
      
      'Cleanup
      HiddenSheet.ShowAllData
      HiddenSheet.UsedRange.Clear
    End With
    
    'Use advanced filter, extract results by looping over the range areas
    Dim vResult2() As Variant
    With stdPerformance.Measure("#2 Advanced filter and areas loop - array result")
      'get dimensions
      iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
      iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
      
      'Store capacity for at least iNumRow items in result
      ReDim vResult2(1 To iNumRow)
      
      'Create filters data
      HiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
      With HiddenSheet.Range("A4").Resize(iNumRow, iNumCol)
        'Set data and call filter
        .Value = arr
        Call .AdvancedFilter(xlFilterInPlace, HiddenSheet.Range("A1:A2"))
        
        'Loop over all visible cells and dump data to array
        Dim rArea As Range, vArea As Variant, i As Long, iRes As Long
        iRes = 0
        For Each rArea In .Resize(, 1).SpecialCells(xlCellTypeVisible).Areas
          vArea = rArea.Value
          If rArea.CountLarge = 1 Then
            iRes = iRes + 1
            vResult2(iRes) = vArea
          Else
            For i = 1 To UBound(vArea, 1)
              iRes = iRes + 1
              vResult2(iRes) = vArea(i, 1)
            Next
          End If
        Next
        
        'Trim size of array to total number of inserted elements
        ReDim Preserve vResult2(1 To iRes)
      End With
      
      'Cleanup
      HiddenSheet.ShowAllData
      HiddenSheet.UsedRange.Clear
    End With
    
    'Use a VBA filter
    Dim vResult3() As Variant: iRes = 0
    With stdPerformance.Measure("#3 Array filter - array result")
      'Get total row count
      iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
      
      'Make result at least the same number of rows as the base data (We can't return more data than rows in our source data)
      ReDim vResult3(1 To iNumRow)
      
      'Loop over rows, filter condition and assign to result
      For i = 1 To iNumRow
        If arr(i, 2) = "A" Then
          iRes = iRes + 1
          vResult3(iRes) = arr(i, 1)
        End If
      Next
      
      'Trim array to total result size
      ReDim Preserve vResult3(1 To iRes)
    End With
    
    'Use a VBA filter - return a collection
    'This algorithm is much the same as the above, however we simply add results to a collection instead of to an array. Collections are generally a fast way to have dynamic sizing data.
    Dim cResult As Collection
    With stdPerformance.Measure("#4 Array filter - collection result")
      Set cResult = New Collection
      iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
      For i = 1 To iNumRow
        If arr(i, 2) = "A" Then
          cResult.Add arr(i, 1)
        End If
      Next
    End With
  End With
End Sub
```
</details>
<br>

**RESULTS**

Positive values show how much faster advanced filter and areas loop than advanced filter and copy is.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |    12%  |  1610%  |   82%   |   66%   |

In reality the time taken for the array/collection methods isn't 0ms for either array or collection, but the calculation was so fast it is impossible to know how much faster it is on such a small amount of data. 

<br>
<br>

#### S11c) Sheet to Array

In our final scenario, we have some data in `Sheet1` and we'd like to obtain a copy of this data in an array which we require for further processing.

In this case what is faster? Using AdvancedFilter? Or using pure array logic? 

<details>
  <summary>Source code</summary>

```vb
Sub scenario3()
  'Initialisation
  Dim arr As Variant: arr = getArray(1000000)
  Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
  Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
  Sheet1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
  
  With stdPerformance.Optimise()
    'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
    Dim vResult1
    With stdPerformance.Measure("#1 Advanced filter and copy")
      'Choose output headers
      HiddenSheet.Range("A4:B4").Value = Array("ID", "Key")
      
      'Choose filters
      HiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
      
      'Call advanced filter
      With Sheet1.Range("A1").CurrentRegion
        Call .AdvancedFilter(xlFilterCopy, HiddenSheet.Range("A1:A2"), HiddenSheet.Range("A4:B4"))
      End With
      
      'Obtain results
      vResult1 = HiddenSheet.Range("A4").CurrentRegion.Value
      
      'Cleanup
      HiddenSheet.UsedRange.Clear
    End With
    
    'Array
    'Copy data from sheet into an array,
    'loop through rows, move filtered rows to top of array,
    'in this scenario remove value from any row which isn't required.
    Dim vResult2
    With stdPerformance.Measure("#2 Use of array")
      vResult2 = Sheet1.Range("A1").CurrentRegion.Value
      Dim iRowLen As Long: iRowLen = UBound(vResult2, 1)
      Dim iColLen As Long: iColLen = UBound(vResult2, 2)
      
      Dim i As Long, j As Long, iRet As Long
      iRet = 1
      For i = 2 To iRowLen
        If vResult2(i, 2) = "A" Then
          iRet = iRet + 1
          If iRet < i Then
            For j = 1 To iColLen
              vResult2(iRet, j) = vResult2(i, j)
              vResult2(i, j) = Empty
            Next
          End If
        Else
          vResult2(i, 1) = Empty
          vResult2(i, 2) = Empty
        End If
      Next
    End With
  End With
End Sub
```
</details>
<br>

**RESULTS**

Positive values show how much faster use of array than advanced filter and copy is.

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  -288%  |  -397%  |  -40%   |   13%   |

In summary, use of AdvancedFilters and arrays are fairly comparable. Use advanced filter if you are trying to keep all data within Excel and you don't need the data in VBA afterwards. However whether you've got the data in VBA to begin with, or you need the data in VBA afterwards, using an array based filter mechanism is preferable.

<br>
<br>

### S12) DLLs vs VBA Libraries

We often think of compiled C libraries from being faster alternatives to pure VBA code, and in some circumstances this is definitely the case. However in this test we are going to compare the following 2 subs:

```vb
Public Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)

Public Sub VariantCopyVBA(ByRef v1, ByVal v2)
  If isObject(v2) Then
    Set v1 = v2
  Else
    v1 = v2
  End If
End Sub
```

Generally speaking, these 2 subs perform the same task. They copy the data within one variant, to another variant. Let's compare their performance:

<details>
  <summary>Source code</summary>

```vb
Sub testVariantCopy()
  Const C_MAX As Long = 1000000
  Dim v1, v2
  With stdPerformance.Optimise
    With stdPerformance.Measure("#1 DLL", C_MAX)
      For i = 1 To C_MAX
        Call VariantCopy(v1, v2)
      Next
    End With
    With stdPerformance.Measure("#2 VBA", C_MAX)
      For i = 1 To C_MAX
        Call VariantCopyVBA(v1, v2)
      Next
    End With
  End With
End Sub
```
</details>
<br>

**RESULTS**

So this shows that there is likely some overhead involved when executing DLL functions, which causes them to be significantly slower than homebrew function calls. The results of relative performance of VBA and DLL are shown below. Positive performance indicates how much DLL is faster than VBA. 

> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   742%  |   245%  |  -1841% | -1363%  |

Given the slow speeds it is likely that this has something to do with either:

1. Lack of caching on DLL load
2. Lack of caching of function address (via GetProcAddress()).
3. Preparing the arguments of the function call.

<br>
<br>

### S13) Helper functions

Helper functions are often used to make code look clean, but do they have a performance impact? In order to test, the following 10 functions will be called:

<details>
  <summary>Source code</summary>

```vb
Function Help1() As Boolean
  Help1 = True
End Function
Function Help2() As Boolean
  Help2 = Help1()
End Function
Function Help3() As Boolean
  Help3 = Help2()
End Function
Function Help4() As Boolean
  Help4 = Help3()
End Function
Function Help5() As Boolean
  Help5 = Help4()
End Function
Function Help6() As Boolean
  Help6 = Help5()
End Function
Function Help7() As Boolean
  Help7 = Help6()
End Function
Function Help8() As Boolean
  Help8 = Help7()
End Function
Function Help9() As Boolean
  Help9 = Help8()
End Function
Function Help10() As Boolean
  Help10 = Help9()
End Function
```

The test itself looks like this:

```vb
Const C_MAX As Long = 1000000
Dim v As Boolean

With stdPerformance.Measure("Help0")
  For i = 1 To C_MAX
    v = True
  Next
End With
With stdPerformance.Measure("Help1", 1 * C_MAX)
  For i = 1 To C_MAX
    v = Help1()
  Next
End With
With stdPerformance.Measure("Help2", 2 * C_MAX)
  For i = 1 To C_MAX
    v = Help2()
  Next
End With
With stdPerformance.Measure("Help3", 3 * C_MAX)
  For i = 1 To C_MAX
    v = Help3()
  Next
End With
...
```
</details>
<br>

**RESULTS**

In this particular scenario it appears that 1-10 were effected by the same degree of delay, and each additional function call adds an additional ~0.03µs delay to the function. In general this isn't something to worry about, however if you are trying to make the most optimal code, it is something to consider.

    Help0:   15 ms (0.015µs per operation)
    Help1:   31 ms (0.031µs per operation)
    Help2:   63 ms (0.031µs per operation)
    Help3:   94 ms (0.031µs per operation)
    Help4:  141 ms (0.035µs per operation)
    Help5:  172 ms (0.034µs per operation)
    Help6:  188 ms (0.031µs per operation)
    Help7:  218 ms (0.031µs per operation)
    Help8:  266 ms (0.033µs per operation)
    Help9:  297 ms (0.033µs per operation)
    Help10: 312 ms (0.031µs per operation)

<br>
<br>

### D) Data structures

#### D1) Dictionary (Optimising data lookups)

Dictionaries are often used to optimise the lookup of data, as they provide an optimal means of doing these kinds of lookups. In this scenario we choose between 4 separate scenarios. Including using a dictionary.

<details>
  <summary>Source code</summary>

```vb
Sub dictionaryTest()
  'Arrays vs dictionary
  Const C_MAX As Long = 5000000
  Dim arr: arr = getArray(C_MAX)
  Dim arrLookup: arrLookup = getLookupArray()
  Dim dictLookup: Set dictLookup = getLookupDict()
  Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
  Dim i As Long, j As Long, iVal As Long
  
  With stdPerformance.Optimise
    With stdPerformance.Measure("#1 Lookup in array - Naieve approach", C_MAX)
      For i = 2 To C_MAX
        'Lookup key in arrLookup
        For j = 1 To 5
          If arr(i, 2) = arrLookup(j, 1) Then
            iVal = arrLookup(j, 2)
            Exit For
          End If
        Next
      Next
    End With
    
    With stdPerformance.Measure("#2 Lookup in dictionary - late binding", C_MAX)
      For i = 2 To C_MAX
        'Lookup key in dict
        iVal = dictLookup(arr(i, 2))
      Next
    End With
    
    With stdPerformance.Measure("#3 Lookup in dictionary - early binding", C_MAX)
      For i = 2 To C_MAX
        'Lookup key in dict
        iVal = dictLookup2(arr(i, 2))
      Next
    End With
    
    With stdPerformance.Measure("#4 Generate through logic", C_MAX)
      For i = 2 To C_MAX
        'Generate value from key
        iVal = getLookupFromCalc(arr(i, 2))
      Next
    End With
    
    With stdPerformance.Measure("#5 Generate through logic direct", C_MAX)
      For i = 2 To C_MAX
        'Generate value from key
        Dim iChar As Long: iChar = Asc(arr(i, 2)) - 64
        iVal = iChar * 10
        If iChar = 5 Then iVal = 99
      Next
    End With
  End With
End Sub

'Obtain an array of data to C_MAX size.
'@param {Long} Max length of data
'@returns {Variant(nx2)} Returns an array of data 2 columns wide and n rows deep.
Public Function getArray(C_MAX As Long) As Variant
  Dim arr() As Variant
  ReDim arr(1 To C_MAX, 1 To 2)

  arr(1, 1) = "ID"
  arr(1, 2) = "Key"
  Dim i As Long
  For i = 2 To C_MAX
    'ID
    arr(i, 1) = i
    Select Case True
      Case i Mod 17 = 0: arr(i, 2) = "A"
      Case i Mod 13 = 0: arr(i, 2) = "B"
      Case i Mod 11 = 0: arr(i, 2) = "C"
      Case i Mod 7 = 0: arr(i, 2) = "D"
      Case Else
        arr(i, 2) = "E"
    End Select
  Next
  getArray = arr
End Function

'Obtain dictionary to lookup key,value pairs
Public Function getLookupDict() As Object
  Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
  o("A") = 10: o("B") = 20: o("C") = 30: o("D") = 40: o("E") = 99
  Set getLookupDict = o
End Function

'Obtain array to lookup key,value pairs
Public Function getLookupArray() As Variant
  Dim arr()
  ReDim arr(1 To 5, 1 To 2)
  arr(1, 1) = "A": arr(1, 2) = 10
  arr(2, 1) = "B": arr(2, 2) = 20
  arr(3, 1) = "C": arr(3, 2) = 30
  arr(4, 1) = "D": arr(4, 2) = 40
  arr(5, 1) = "E": arr(5, 2) = 99
  getLookupArray = arr
End Function

'Obtain value from key
Public Function getLookupFromCalc(ByVal key As String) As Long
  Dim iChar As Long: iChar = Asc(key) - 64
  getLookupFromCalc = iChar * 10
  If iChar = 5 Then getLookupFromCalc = 99
End Function
```
</details>
<br>

**RESULTS**

From the below results we can see that the naieve approach of looping through the lookup array did not only produce more ugly code, but also produces significantly slower results. By comparison the dictionary was both clean and fast.

The other two tests were to show that actually generating the data directly using the available functions can often be faster than looking up data in a `Dictionary`. This may not be possible in a lot of situations however.

Finally #5 was included because it was surprising to me that including the function in the body of the loop would have that much of an impact on performance.

Lookup in array vs in dictionary (late binding):
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   287%  |   209%  |  210%   |   169%  |

Lookup in dictionary (late binding) vs in dictionary (early binding):
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   18%   |   266%  |  100%   |   191%  |

Generate through logic vs direct:
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   122%  |   126%  |  127%   |   116%  |

<br>
<br>

### A1) Advanced Array Manipulation

This section needs further work to include runnable examples instead of only explaining the concepts.

Further reading:

* [Arrays of Structs](https://www.vbforums.com/showthread.php?706851-RESOLVED-CopyMemory-Shift-Array-one-position).
* [Arrays of Variants](https://www.vbforums.com/showthread.php?848397-How-do-I-manipulate-data-in-variant-arrays-using-RtlMoveMemory-with-VB6-VBA).
* [VBSpeed](http://www.xbeat.net/vbspeed/) and [articles](http://www.xbeat.net/vbspeed/articles.htm).

#### A1a) Resizing and Transposing an array

If you're really keen to `ReDim` an array lightning fast, you can change the dims of the array using [`CopyMemory` and `SafeArray` struct](https://web.archive.org/web/20220121153221/https://bytecomb.com/vba-internals-array-variables-and-pointers-in-depth/)

This at least provides a `xd --> yd` translation.

Resizing or translating an array becomes tough. Let's say we wanted to resize the following array:

    a11,a12,a13
    a21,a22,a23
    a31,a32,a33

Internally this is represented as:

    a11,a21,a31,a12,a22,a32,a13,a23,a33

If we want to add a column we simply add 3 elements onto the end:

    a11,a21,a31,a12,a22,a32,a13,a23,a33,NEW14,NEW24,NEW34

this is why VBA is capable of `ReDim`-ing this bound, as it is really easy to do. If we want to add a row however we need to change the array as follows:

    a11,a21,a31,NEW41,a12,a22,a32,NEW42,a13,a23,a33,NEW43

    Pending implementation example...

In this case there is still some room for performance gain (we can use `CopyMemory` to copy `UBound(arr,1)` elements at a time into positions e.g. 1, 5, 9 in this case), but it's not so easy to achieve. This is likely one of the reasons why VBA doesn't support `redim`-ing of this bound.

If we want to Transpose the array we need to change it into a structure like:

    a11,a12,a13,a21,a22,a23,a31,a32,a33

Transposing isn't really possible to optimise in the programming language domain. For this we'd need to optimise byte level machine code.

The most common ways to perform this operation fast is using [`LAPACK`](https://netlib.sandia.gov/lapack/lapack-3.3.1.html) or [`CUDA SDK`](https://developer.nvidia.com/cuda-toolkit). P.S. this is what e.g. Python and FORTRAN use, and why they are so good at statistical programming.

<br>
<br>

### ST) String operations

The simple concatenation of two strings (`string3 = string1 & string2`) usually works quite fast, but only up to the point where the number of concatenations is small. For example, if you want to store an array of doubles with 150 x 150 elements, you will have to convert it to a string and use concatenation, which takes a lot of time. To speed up the conversion of the array to a string, there is an internal function `Join` that takes a 1-based 1D array and converts it to a string.

<table>
<tr>
<th>Concatenation</th>
<th>Join</th>
</tr>
<tr>
<td>
  
```vb
stringContent = stringContent & notBig2DArray(i, j) & ", "      
```

</td>
<td>

```vb
stringContent = Join(tmpStrings, vbCrLf)      
```

</td>
</tr>
</table>

<details>
  <summary>Source code</summary>

```vb
stringContent = ""
With stdPerformance.Measure("#1 Concatenation " & C_MAX, C_MAX ^ 2)
    For i = 1 To C_MAX
        For j = 1 To C_MAX
            stringContent = stringContent & notBig2DArray(i, j) & ", "
        Next j
        stringContent = stringContent & vbCrLf
    Next i
End With

stringContent = ""
Dim tmpStrings() As Variant
ReDim tmpStrings(1 To C_MAX)
Dim tmpDoubles() As Variant
ReDim tmpDoubles(1 To C_MAX)
With stdPerformance.Measure("#2 Join " & C_MAX, C_MAX ^ 2)
    For i = 1 To C_MAX
        For j = 1 To C_MAX
            tmpDoubles(j) = notBig2DArray(i, j)
        Next j
        tmpStrings(i) = Join(tmpDoubles, ", ")
    Next i
    stringContent = Join(tmpStrings, vbCrLf)
End With
```
</details>
<br>

It should also be mentioned that the concatenation function is a single Excel function with exponential growth of execution time, whereas the `Join` function exhibits almost linear growth in execution time with the size of the array:

![Comparison of measured `Join` and concatenation times as a function of square array](https://github.com/Excel-lent/VBA-performance-tests/raw/main/Pictures/Join%20vs%20concatenation.jpg)

**RESULTS**

Performance gain of `Join` relative to concatenation for conversion of 140*140 matrix to a string:
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |  6033%  |  6731%  |  2594%  |    -    |

<br>
<br>

### W) Writing to file

The majority of answers on StackOverflow regarding file writing in VBA recommend using FSO (`FileSystemObject`). In this chapter, we want to demonstrate that the internal VB functions `Print` and `Put` can accomplish the task more quickly. Each of test cases below stores a string with 2D matrix to a 44.6 Mb file. In all cases the files were written to slow SD card.

<table>
<tr>
<th>Standard writing function</th>
<th>FSO</th>
<th>Binary writing function</th>
</tr>
<tr>
<td>
  
```vb
Open Application.ActiveWorkbook.Path & "\" & fileName1 For Output As #iFile      
Print #iFile, strFileContent
```

</td>
<td>

```vb
Set oFile = fso.CreateTextFile(Application.ActiveWorkbook.Path & "\" & fileName2)      
oFile.WriteLine strFileContent
```

</td>
<td>

```vb
Open Application.ActiveWorkbook.Path & "\" & fileName2 For Binary As #iFile
Put #iFile, , strFileContent
```

</td>
</tr>
</table>

#### W1) `FileSystemObject` vs standard writing function (`Print`)

<details>
  <summary>Source code</summary>

```vb
With stdPerformance.Measure("#1 Standard writing function", C_MAX)
    iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\" & fileName1 For Output As #iFile
    Print #iFile, strFileContent
    Close #iFile
End With

With stdPerformance.Measure("#2 FSO", C_MAX)
    Set oFile = fso.CreateTextFile(Application.ActiveWorkbook.Path & "\" & fileName2)
    oFile.WriteLine strFileContent
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End With
```
</details>
<br>

**RESULTS**
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   34%   |   -1%   |   56%   |    -    |

<br>
<br>

#### W2) FSO vs binary writing

<details>
  <summary>Source code</summary>

```vb
With stdPerformance.Measure("#3 Binary writing function", C_MAX)
    iFile = FreeFile
    Open Application.ActiveWorkbook.Path & "\" & fileName2 For Binary As #iFile
    Put #iFile, , strFileContent
    Close #iFile
End With
```
</details>
<br>

**RESULTS**
> | 2007/32 | 2010/32 | 2019/64 |  365/64 |
> |:-------:|:-------:|:-------:|:-------:|
> |   34%   |   28%   |   31%   |    -    |

<br>
<br>

