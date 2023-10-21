---
layout: post
title:  "Why is VBA the most dreaded language?"
published: true
authors:
  - "Sancarn"
  - "TarVK"
---

In the 2020 StackOverflow survey, it is claimed that VBA is the most dreaded language:

![_](./assets/img/stackoverflow-2020-vba-most-dreaded.png)

In this article we will explore some of the reasons why this might be the case.

## Who uses VBA?

In order to answer this question we must first look at another question - who actually uses VBA in the first place? In 2021 I ran a poll on [/r/vba](http://reddit.com/r/vba) where I asked redditors why they code in VBA.

![_](./assets/img/reddit-2021-why-do-you-code-in-vba.png)

From these data we can clearly see that the majority of people who use VBA do so mainly because they have no other choice. Many organisations run their entire business processes with Excel and when a little bit of automation is required VBA is usually \#1 on the list because it's something that IT departments haven't locked down and haven't provided a better alternative for. In business culture IT rarely will allow its users to even create and query a database. This leads to even more data being trapped in Excel.

In the business I currently work for, in the engineering division, we have access to a variety of technologies:

* OnPrem - PowerShell (No access to `Install-Module`)
* OnPrem - Excel (VBA  / OfficeJS (limited access) / OfficeScripts / PowerQuery)
* OnPrem - PowerBI Desktop
* OnCloud - Power Platform (PowerApps, Power BI, PowerAutomate (non-premium only))
* OnCloud - Sharepoint
* SandboxedServer - ArcGIS (ArcPy)
* SandboxedServer - MapInfo (MapBasic)
* SandboxedServer - InfoWorks ICM (Ruby)
* SandboxedCloud - ArcGIS Online

Every request for a high level language to be installed across the team e.g. `Python` / `Ruby` etc. has been rejected by CyberSecurity in favour for technologies like `PowerAutomate`, `PowerApps` etc. It is supposedly "Against the strategic vision of the company". Why are these technologies unworkable? That's a bigger topic [for another day](/articles/Issues%20with%20PowerPlatform.html), but suffice to say sometimes the requirement is either `OnPrem`, or the task is so large a serverless PowerAutomate approach is too slow, or the algorithm so complex that a PowerAutomate solution would become infuriating to maintain and incomprehensible to even IT folks \(e.g. see [projection algorithms](https://www.movable-type.co.uk/scripts/latlong-os-gridref.html#source-code-osgridref)\).

Next, in 2022 I ran another poll on `/r/vba` where I asked redditors how they learned VBA. This was their responses:

![_](./assets/img/reddit-2022-how-did-you-learn-vba.png)

I echo the sentiment of most users of VBA here. I was also self-taught, but was fortunate enough to have learnt Lua before learning VBA, and have friends studying computer science, so I adopted many of their best practices. Many people who are self-taught are unlikely to know or have these best practices in mind. Looking at a [recent poll of mine](https://www.reddit.com/r/vba/comments/16ky8ja/do_you_know_or_write_code_in_other_programming), about 1/3 of respondents had not used other languages and therefore are unlikely to follow best practices.

## The state of VBA projects

Now that we've understood the users, we have to contemplate what state existing VBA projects are in. This can vary dramatically from project  to project, and largely dependent on the authors of the project in question.


### Poor indentation

Many people who write VBA code indent nothing.

```vb
Sub FindCombination()
Dim numArray() As Variant
Dim total As Double
Dim result As String
Dim i As Long, j As Long, k As Long, n As Long
numArray = Range("A1:A10").Value
total = 25
if range("A1") = Empty then total = 50
For i = 1 To 2 ^ UBound(numArray, 1)
result = ""
n = 0
For j = 0 To UBound(numArray, 1)
If i And 2 ^ j Then
result = result & numArray(j + 1, 1) & "+"
n = n + numArray(j + 1, 1)
End If
Next j
If n = total Then
Range("A1") = Left(result, Len(result) - 1)
End If
Next i
End Sub
```

Some only use 1 layer of indentation:

```vb
Sub FindCombination()
    Dim numArray() As Variant
    Dim total As Double
    Dim result As String
    Dim i As Long, j As Long, k As Long, n As Long
    numArray = Range("A1:A10").Value
    total = 25
    if range("A1") = Empty then total = 50
    For i = 1 To 2 ^ UBound(numArray, 1)
    result = ""
    n = 0
    For j = 0 To UBound(numArray, 1)
    If i And 2 ^ j Then
    result = result & numArray(j + 1, 1) & "+"
    n = n + numArray(j + 1, 1)
    End If
    Next j
    If n = total Then
    Range("A1") = Left(result, Len(result) - 1)
    End If
    Next i
End Sub
```

or sometimes the indenation is all over the shop...

```vb
Public Sub GradientCol(Ob As Object, AB As Single, R1 As Single, G1 As Single, B1 As Single, R2%, G2%, B2%)
On Error Resume Next 'just in case
Dim H%, rt As Single, Gt As Single, Bt As Single
Imagewait True
AB = AB / 10 'alpha blending
H = Ob.Height - 1
rt = (R2 - R1) / H
Gt = (G2 - G1) / H
Bt = (B2 - B1) / H
'Set the gradient
For xx = 0 To H
Ob.Line (0, xx)-(Ob.Width - 1, xx), RGB(R1, G1, B1)
R1 = R1 + rt
G1 = G1 + Gt
B1 = B1 + Bt
Next xx
'Read the gradient-colors, mix the with alpha-blend
'and put the new colors back.
For xx = 0 To Ob.Width - 1
For yy = 0 To Ob.Height - 1
    Color = GetPixel(Ob.hdc, xx, yy)
    R1 = Color Mod 256&
    G1 = ((Color And &HFF00) / 256&) Mod 256&
    B1 = (Color And &HFF0000) / 65536
'This is the actual alpha-blending
        R(xx, yy) = (R(xx, yy) * (1 - AB)) + (R1 * AB)
        G(xx, yy) = (G(xx, yy) * (1 - AB)) + (G1 * AB)
        B(xx, yy) = (B(xx, yy) * (1 - AB)) + (B1 * AB)
'put the new colors back
SetPixel Ob.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy, xx
Imagewait False
Ob.Refresh
End Sub
```

If you're lucky you'll have code which is indented well. But this is quite often a rarity among VBA developers. 

### Poor Commentary

Many many vba macro authors will write 0 commentary. This is quite alright if the code is clean, but most code written is not. And just like in many modern languages many programmers do not write adequate commentary on either how to use the code or how the internals work. This is definitely not isolated to a VBA problem, but the lack of skilled programmers writing VBA code often exacerbates this problem.

### Macro recorder junk

Because many VBA devs learn from the macro recorder they also often will leave recorded VBA in the subs. This VBA is notoriously awful, with many unrequired statements which is often totally unoptimised to the task at hand. For instance the below code:

```vb
Sub Macro1()
    Range("A1:A20").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=8
    Sheets.Add After:=ActiveSheet
    ActiveWindow.SmallScroll Down:=7
    ActiveSheet.Paste
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=18

    Columns("A:A").Select
    Selection.Style = "Percent"
End Sub
```

Should really be more like this:

```vb
Sub Macro1()
    With sheets.add()
        With .Range("A1:A20")
            .value = Sheet1.Range("A1:A20").value
            .Style = "Percent"
        End With
    End with
End sub
```

The reality is that the macro recorder makes use of the application state, across many different workbooks (Excel spreadsheets). This does have benefits like being able to record pretty much anything, but it does mean that you get these awful "Select the range, then apply to the selection" macro commands. Later macro recorders like OfficeScripts don't have the ability to work across multiple spreadsheets alleviating the need for a recorder which does this, but as a result OfficeScripts does become less powerful.

Another issue that vba projects which have used the macro recorder have is that quite often a mix of formulae and code will be used, and sometimes it's not clear why they used formulae, and why they used VBA.

```vb
Range("A2").Select
ActiveCell.Formula = _
    "=INDEX('Raw Data'!$A$1:$CC$67,ROWS($A$2:A2)+1,MATCH(A$1,'Raw Data'!$1:$1,0))"
Range("A2").Select
Selection.AutoFill Destination:=Range("A2:A67")
Range("B2").Select
ActiveCell.FormulaR1C1 = _
    "=IF((ISERROR(INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0)))),""-"",INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0)))"
'ActiveCell.FormulaR1C1 = _
'    "=INDEX('Raw Data'!R1C1:R67C81,MATCH(RC1,'Raw Data'!C1,0),MATCH('Dynamic Summary'!R1C,'Raw Data'!R1,0))"
Range("B2").Select
Selection.AutoFill Destination:=Range("B2:B67"), Type:=xlFillDefault
Range("B2:B67").Select
Selection.AutoFill Destination:=Range("B2:K67"), Type:=xlFillDefault
Range("A2:K67").Select
Selection.NumberFormat = "General"
Columns("E:F").Select
'Selection.Style = "Currency"
Columns("G:G").Select
Selection.Style = "Percent"
Columns("H:I").Select
Selection.Style = "Comma"
Columns("K:K").Select
Selection.Style = "Percent"

Sheets("Raw Data").Select
Range("CombinedRawData[[#Headers],[Project ID (PID)]]").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Application.CutCopyMode = False
Selection.Copy
Range("CombinedRawData[[#Headers],[Project ID (PID)]]").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Sheets("PMO Tracked CIAC Projects").Select
Application.CutCopyMode = False
'ActiveWindow.SelectedSheets.Delete
'Sheets("PIDHealthData").Select
'ActiveWindow.SelectedSheets.Delete
```

### Poor seperation of concerns

It is quite common for VBA applications to have monolithic subs instead of seperating code into many smaller functions or objects each with a seperate concern. This again comes mostly from the lack of programming knowledge and experience. Also it's very common that VBA developers will use global variables, or module level variables all the way throughout their code, sometimes not declaring them, all of which produces unexpected behaviour and difficult to maintain code. 

### Poor adherence to DRY

A best practice in programming is D.R.Y i.e. "Don't repeat yourself". However due to the naivity of many VBA developers, there is a lot of repeated code. Often the same piece of code will be copy and pasted all over the code base instead of modularising this code in a seperate function / library.

```vb
if Range("A1").value = "OK" then
  Debug.Print Range("B1")
  ...
end if
if Range("A2").value = "OK" then
  Debug.Print Range("B2")
  ...
end if
if Range("A3").value = "OK" then
  Debug.Print Range("B3")
  ...
end if
if Range("A4").value = "OK" then
  Debug.Print Range("B4")
  ...
end if
...
```

### Poor numerical reasoning

Developers typically have better mathematical skills than the average Joe. However most people who use VBA i.e. those who "make macros" are typically not the best when it comes to mathematics and numerical reasoning. This can lead to overcomplicated or unscalable solutions to problems. Below I've given a few examples of this which I've experienced in the past.

#### Finding generalised numerical shortcuts

One example I've come across is in finding the [AMP number](https://en.wikipedia.org/wiki/Asset_management_plan_period) for a given date - An asset management plan (AMP) period is a five-year time period used in the English and Welsh water industry. The first AMP, AMP1, was the period of time between April 1990 and 1995. The AMP is commonly used in databases as a means of aggregating data. In a lot of code I've seen engineers use something along the lines of this:

```vb
Function getAmp(ByVal dt as Date) as Long
  select case dt
    case Is < DateSerial(1990,3,1): getAmp = 0
    case Is < DateSerial(1995,3,1): getAmp = 1
    case Is < DateSerial(2000,3,1): getAmp = 2
    case Is < DateSerial(2005,3,1): getAmp = 3
    case Is < DateSerial(2010,3,1): getAmp = 4
    case Is < DateSerial(2015,3,1): getAmp = 5
    case Is < DateSerial(2020,3,1): getAmp = 6
    case Is < DateSerial(2025,3,1): getAmp = 7
  end select
End Function
```

And ultimately the code needs to be modified every 5 years to include an additional AMP number. Of course the better approach would be the following, which will continue working on into the future without modification:

```vb
Function getAmp(ByVal dt As Date) As Long
  getAmp = floor((Year(dt - 90) - 1990) / 5)
End Function
Function floor(ByVal x As Double) As Double
  floor = Int(x) - 1 * (Int(x) > x)
End Function
```

#### Use of numerical methods when generalised formula can be found

In a carbon assessment tool, there was a function which approximated how much carbon a tree would consume at an instance of time. Let's say `1/(x+1)` for the sake of argument. In order to calculate the total amount of carbon within a time frame the engineers had used an iterative approach as follows:

```vb
function getTotalCarbon(ByVal iNumYears as Long) as Double
    Dim sumCarbon as Double: sumCarbon = 0
    Dim i as long, j as long
    For i = 0 to iNumYears - 1
        For j = 1 to 100
            Dim x as Double: x = i + 1/100
            sumCarbon = sumCarbon + 1/(x+1) * 1/100
        next
    next
    getTotalCarbon = sumCarbon
end function
```

However this produces a poor estimate. Those with a little more mathematical know how will see this is simply integrating the expression, and thus the better approach would be to integrate `1/(x+1)` to `log(x+1)` and then plug values (also note `log(1)==0`):

```vb
function getTotalCarbon(ByVal iNumYears as Long) as Double
    getTotalCarbon = log(iNumYears)
end function
```

### Poor choice of algorithms

Finally a very common issue with many VBA projects is use of poor, slow algorithms or datastructures. For instance it is very common for people to use an array to store a list of items as follows:

```vb
Dim o() as object
Redim o(0 to 0)
For each row in rows
    if row("col") = condition then
        set o(ubound(o)) = row
        Redim Preserve o(ubound(o)+1)
    end if
next
```

The above, where we are inserting an element into an array is `O(n^2)`, where as a more optimal example would be using a collection, where adding all items is `O(n)`:

```vb
Dim o as collection
For each row in rows
    if row("col") = condition then
        o.add row
    end if
next
```

It is also extremely common for people to use funky algorithms due to the lack of available data structures / libraries in VBA. 

### A mix of VBA7 and VBA6 & 32-bit and 64-bit compatible code

In order to integrate with windows OS or use useful APIs not native to VBA, many VBA developers will use win32 APIs. Some examples of where you will need to use the win32 APIs are as follows:

* Making a userform resizable
* Integrating with the clipboard

If you are lucky a VBA developer will ensure their code works in both VBA7 and VBA6 as well as 32 and 64 bit versions:

```vb
#if Win64 then
    Private Const NULL_PTR as LongLong = 0^
#else
    Private Const NULL_PTR as Long = 0&
#end if

#if VBA7 then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (Optional ByVal hWnd As LongPtr = NULL_PTR) As Long
    '...
#else
    Private Declare Function OpenClipboard Lib "user32" (Optional ByVal hWnd As Long = NULL_PTR) As Long
    '...
#end if
```

However realistically most VBA developers who run into these problems will be inexperienced and be copy-paste coders, and these devs will [rarely have both](https://stackoverflow.com/a/2913670/6302131). This is not including the large VB6 community, who have created [thousands of examples](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md) of VBA6/32-bit only code, not because the VB6 community doesn't want 64-bit compatible applications, but because Microsoft never gave 64-bit compatibility to VB6... The upgraded office without upgrading VB6. As such, in reality there are a plethora of 32-bit examples out there which will work in 32-bit VBA, but VBA7/64-bit compatible code out there is scarce.


### Low level hacks and thunks

As a result of lack of development in the VBA language, devs who are extremely experienced and well versed in COM and the VBA runtime, will often resort to using low-level "tricks" and thunks to perform certain tasks. Ultimately many of these tricks or thunks come out of necessity but it doesn't help inexperienced VBA devs, even if they are full stack developers.

```vb
    hComctl32 = LoadLibrary(StrPtr("comctl32"))
    hMSVBVM60 = GetModuleHandle(StrPtr("msvbvm60"))
    
    If bIsInIDE Then
        hVBA6 = GetModuleHandle(StrPtr("vba6"))
    End If
    
    ' //
    ' // The code structure
    ' //
    ' // +--------+-----------------------------------+
    ' // | offset |           description             |
    ' // +--------+-----------------------------------+
    ' // |  0x00  | Reference counter                 |
    ' // |  0x04  | Mapping address of first instance |
    ' // |  0x08  | The asm code starts here          |
    ' // |  ....  |                                   |
    ' // +--------+-----------------------------------+
    ' //
    
    lCodeSize = (UBound(cOp) + 1) * 8 + 8
    
    cOp(0) = -172700121577779.7803@:  cOp(1) = 700803790855680.8192@:   cOp(2) = 32381283757.4281@:       cOp(3) = 8880706099704.8576@
    cOp(4) = 522461228120781.6192@:   cOp(5) = 508377703192199.1681@:   cOp(6) = -857365021913051.7204@:  cOp(7) = 443382.3494@
    cOp(8) = -441089100593967.8792@:  cOp(9) = 846995906805189.8372@:   cOp(10) = 620186570542432.3336@:  cOp(11) = 626874291154953.5999@
    cOp(12) = 20596.508@:             cOp(13) = -190689950864645.9042@: cOp(14) = 147676186889862.127@:   cOp(15) = -900719925218220.0576@
    cOp(16) = 15230408694.5918@:      cOp(17) = -6489317565906.8672@:   cOp(18) = 619989648405870.2934@:  cOp(19) = -147748414298680.988@
    cOp(20) = 849688782354884.2759@:  cOp(21) = 259535074382164.2883@:  cOp(22) = 583666403534889.1494@:  cOp(23) = -4953584075692.6977@
    cOp(24) = 850260259709754.0725@:  cOp(25) = 147436517215347.0992@:  cOp(26) = 491014245965020.6288@:  cOp(27) = 576460779796930.9778@
    cOp(28) = 82199524659143.1821@:   cOp(29) = -7165954328048.4288@:   cOp(30) = 783753342645325.5286@:  cOp(31) = -27796037319571.7501@
    cOp(32) = -842766526041578.1888@: cOp(33) = 178420980956634.426@:   cOp(34) = 620312248530782.8992@:  cOp(35) = 147436592801036.2344@
    cOp(36) = 850252832242876.1599@:  cOp(37) = -441089059744232.268@:  cOp(38) = 551550672246059.8296@:  cOp(39) = 166280520433.9748@
    cOp(40) = -843072447184994.304@:  cOp(41) = 465081571739304.8644@:  cOp(42) = 29843883910475.0084@:   cOp(43) = -169242102482309.7089@
    cOp(44) = 641312586460930.458@:   cOp(45) = 2043492.7755@:
        
    With m_tParams
        
        If bIsInIDE Then
            lCount = UBound(.pfnAPIs) + 1
        Else
            lCount = UBound(.pfnAPIs)
        End If
        
        For lIndex = 0 To lCount - 1
            
            Select Case lIndex
            
            Case 0:     sAPIName = "RemoveWindowSubclass"
            Case 1:     sAPIName = "DefSubclassProc"
            Case 2:     sAPIName = "__vbaRaiseEvent"
            Case 3:     sAPIName = "EbMode"
           
            End Select
            
            If lIndex < 2 Then
                .pfnAPIs(lIndex) = GetProcAddress(hComctl32, sAPIName)
            ElseIf lIndex < 3 Then
                .pfnAPIs(lIndex) = GetProcAddress(hMSVBVM60, sAPIName)
            Else
                .pfnAPIs(lIndex) = GetProcAddress(hVBA6, sAPIName)
            End If
            
            If .pfnAPIs(lIndex) = 0 Then
                Err.Raise 453, MODULE_NAME & "::" & FUNCTION_NAME
            End If
            
        Next
        
        .pHostObject = ObjPtr(Me)
        
    End With
    
    hMapping = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_EXECUTE_READWRITE, 0, lCodeSize, _
                                 StrPtr(MODULE_NAME & "#" & MODULE_VERSION & "_" & CStr(GetCurrentProcessId())))
    If hMapping = 0 Then
        Err.Raise 7, MODULE_NAME & "::" & FUNCTION_NAME, "CreateFileMapping failed"
    End If
    
    pCode = MapViewOfFile(hMapping, FILE_MAP_READ Or FILE_MAP_WRITE Or FILE_MAP_EXECUTE, 0, 0, 0)
    If pCode = 0 Then
        CloseHandle hMapping
        Err.Raise 7, MODULE_NAME & "::" & FUNCTION_NAME, "MapViewOfFile failed"
    End If
    
    ' // Increment ref counter
    If InterlockedIncrement(ByVal pCode) = 1 Then
        ' // Put the address for the first instance
        PutMem4 ByVal pCode + 4, pCode
    Else
        
        ' // There is already mapped region. Use previous mapping instead current one
        lTemp = pCode
        GetMem4 ByVal pCode + 4, pCode
        UnmapViewOfFile lTemp
        CloseHandle hMapping
        hMapping = 0
        
    End If
    
    ' // Copy asm code
    memcpy ByVal pCode + 8, cOp(0), lCodeSize - 4
    
    ' // Run initialization
    hr = DispCallFunc(ByVal 0&, pCode + 8, CC_STDCALL, vbLong, 1, vbLong, VarPtr(CVar(VarPtr(m_tParams))), vResult)
    
    If hr < 0 Or vResult = 0 Then
        
        CloseHandle hMapping
        
        If InterlockedDecrement(ByVal pCode) = 0 Then
            UnmapViewOfFile pCode
        End If
        
        If hr < 0 Then
            Err.Raise hr, MODULE_NAME & "::" & FUNCTION_NAME, "DispCallFunc failed"
        Else
            Err.Raise 5, MODULE_NAME & "::" & FUNCTION_NAME, "Initialization failed"
        End If
        
    End If
    
    m_pCode = pCode
    m_hMapping = hMapping
    m_bIsInit = True
```

The reality of the situation is this code is usually battle tested, and if an error occurs it will either be an edge case the author hadn't considered, or an edge case the caller didn't consider. Either way this is a developer's worst nightmare and even if you are a VBA developer with a significant amount of experience you are unlikely to know what is going on here.


## Why is VBA the most dreaded language?

The fact is that most modern developer's experience with VBA will likely be a declaration from the business that an old VBA tool needs to be re-written in a modern language as part of a business application. As such they will be given a project which likely has janky syntax, with awful indentation, no comments, and super old fashioned low level datastructures, that actually trying to make sense of the application in order to rebuild it is an utter nightmare. So it's not really that the language as a whole is dreaded, but that applications built in VBA are dreaded.

[VBA does have it's issues](./Issues%20with%20VBA.html) and these may also be contributing factors, but my personal opinion is that most developers who dread VBA mostly just have bad experiences understanding poorly written VBA applications. Unfortunately for them, sometimes devs will be required to understand them, for example in order to port them to other more modern platforms.

VBA in of itself is not an awful language, and it's basis in [COM](https://en.wikipedia.org/wiki/Component_Object_Model) is, dare I say, revolutionary! Most people do have a misunderstanding of COM, and many modern languages don't make working with COM particularly easy. It is quite often seen as a dark art! However COM, as a technology, is still used to this day by [modern Microsoft frameworks](https://en.wikipedia.org/wiki/Windows_Runtime) and is one of those technologies which people keep coming back to because it's so powerful.

Who's to blame for VBA's status? I think numerous parties are. End users are frequently blamed for creating automated business processes in VBA, but we have to remember that business users regularly have no power and no alternative apart from costly business projects. Ultimately I think the real people to blame comes down to:

* Microsoft:
    * for ceasing development and improvement of VBA, the macro recorder and the VBE.
        * and if they aren't going to develop the language the least they could do is OpenSource it...
    * for ceasing development of VB6, and not upgrading it (along with VBA) to 64-bit. In addition causing a schism between VB6 and VBA communities.
* System Administrators / Cyber Security in Businesses:
    * for gatekeeping modern programming environments and forcing end users into using legacy tools like VBA to automate business processes.
    * for continually suggesting the use of poor alternative technologies like `PowerPlatform`, instead of just suggesting a proper solution like on-prem Python.
* Businesses generically:
    * for not properly investing into the development of employee VBA skills. 

Of course, both parties are likely trying to do the right thing. Microsoft has shifted from a developer of "bespoke worse-than-average solutions" to a developer that prefers to stand on the shoulders of giants:

* Internet Explorer, a bespoke web engine, is being phased out for Chromium based Edge web browser.
* dotNET transitioned from the Windows-only dotNET framework to cross platform and open source dotNET core
* Visual Studio, a bespoke IDE, has been largely replaced by VS Code a HTML,CSS,JS Electron-based app.

So it's no wonder that Microsoft might want to phase out VBA, in benefit of languages like TypeScript. In a similar vein, System Admins and Cyber Security know that the safest system is one where the users have access to nothing. Users with access to programming languages not only might pose a cyber security risk themselves, but also become a new attack vector for malicious actors. As a result, Cyber Security and IT administrators prefer sandboxed systems like `OfficeScripts` or `Power Platform`.

But ultimately, this is only my theory as to why VBA is the most dreaded language. Here's hoping we get better fitting replacement technologies for VBA in the future.