---
layout: post
title:  "Issues with VBA"
published: true
---

This post is still under development but here's a dump of all issues. Hopefully I'll get around to finishing it some day 👀

## Issues with the language

1. Object creation and method call is slow when compared to modern languages (and compared to module only code).
2. Hidden features which are unimplementable - e.g. unable to implement IEnumVARIANT. I.E. No custom `For each ... in ... next` implementations (unless you delegate to a collection)
3. Some low level standard interfaces are forbidden in VBA (like IDispatch)
4. Inability to define `hidden` methods.
5. Inability to use `Evaluate` methods on custom classes unless you initially cast to Object (IDispatch) which also removes complete intellisense...
6. Lack of component based design for UserForms. - Modern UI frameworks are so much better at this by now!
7. Lack of standard libraries:
    a. Lack of a canvas component for UserForms - Fundamentally limits what you can do in a userform.
    b. No standard built-in implementation of HashMap - Dictionary is great, but has to be included externally and isn’t available on Mac. Libraries can also help here.
    c. Lack of standard libraries - there are community solutions for this e.g. stdVBA and vb core lib
8. Inconsistent setting of variables - the set keyword was added because Microsoft wanted to allow for use of default properties on COM objects. It would have been better if these were symbolised by illegal code e.g. something@()
9. Inconsistent call conventions for subs and functions 
10. ByRef is default where ByVal is more logical. Assume this was initially to keep code optimal, but generally leads to difficulty in learning. People often just learn “Use ByVal everywhere”, which isn’t correct either.
11. Inability to define collective types e.g. `Collection<Car>`. This leads to limitations in the type system.
12. No built in lambda syntax. Understandable due to VBA’s age, but modern languages use a lot of Lambda syntax to make code cleaner. To get around this, TarVK and I reinvented the wheel and created our own lambda syntax systems to get around this fundamental flaw in the language.
13. Inability to Multithread (or perform tasks asynchronously without re-writing the runtime)
14. Unable to compute on the GPU natively.
15. Lack of native libraries forces people to use APIs which only work on single platforms - Yes VBA can run on Mac but due to the lack of Windows APIs on Mac, what you can do with VBA is severely limited.
16. Lack of true inheritance (can sort of fake it with defaults)
17. No overloading - not a requirement, but a nice feature.
18. Interfacing - Lack of implicit cast to interfaces can make use somewhat clumsy: `Dim x as class: set x = new class: Dim y as IClass: set y = x: y.poop()`
19. Poor error reporting -  Lack of native stack traces, lack of line numbers in errors
20. Lack of reflection, metaprogramming and dynamic dependency injection.
21. Structs/Types appear to have been bodged on top of the runtime. No union types exist and recursive types don’t exist either.
22. PtrSafe keyword - Ridiculous that this only exists in some later versions of VBA…
23. Lack of Events on base types - e.g. `Collection::Add()`, `Collection::Remove()`
24. VBA Keywords - Print, Write, Debug, … - All of these are methods you cannot use! This wouldn’t be the case if VBA namespace wasn’t globally accessible. Why was it designed this way?!
25. Inability to pass structs BYVAL to low level functions
26. [Introduction of the pointless keyword `PtrSafe`](https://stackoverflow.com/a/77141128/6302131) - This keyword provides no function whatsoever nor any guarantee of safety.

## Other issues with VBAs Environment:

1. Not VBA's fault, but the Excel/Word/Powerpoint object libraries are a mess… Error handling in VBA looks bad because generally error handling in Excel/Word/Powerpoint APIs is awful.
2. Lack of Excel/Word/Powerpoint API events.
3. Not VBA's fault, but the Macro recorder produces garbage code. It's really useful for testing, and it's arguably very flexible, but that also means macros recorded with it (and based on current state) can really screw up other spreadsheets accidentally if something out of the ordinary happens.
4. Not VBA's fault, but the VBE - although it was great once - it's now unmaintained and frankly awful. Rubberduck is good if you can install addins etc. but if you aren't able to do these things then you're stuck with the rubbish editor which is in there currently…
5. Limited integration with new features such as PowerQuery, OfficeJS, etc.
6. VB7 (VBA) never released officially to VB6 users. This means many libraries don’t work. It also means VB7 cannot exist as a standalone application. It must always be hosted by some other application.
7. VBE doesn’t make it super easy to work outside of itself (lol).
8. Limitations in Office (e.g. Excel limited rows etc.)
9. Lack of type hinting
10. Cannot easily run VBA in the cloud. Cannot easily schedule VBA scripts either. Though, see [TwinBasic](https://twinbasic.com/).

## Benefits of VBA

1. VBA is on every computer with Microsoft Office. You don’t need permission to run it, everyone can use it 
2. VBA can access the local file system unlike OfficeJS / OfficeScript.
3. VBA has Excel APIs which are mostly intuitive
4. VBA has the ability to reference any other registered type library on the system.
5. VBA can use native APIs on Windows OS, and thus can be used to automate windows.
6. VBA is implemented on top of COM, which means VBA objects can be used by other languages,
7. VBA can inject machine code into memory and execute it (can use thunks)
8. Conditional compiling - Massive benefit to be able to “execute code” at compile time.
9. Default object methods can be useful for adding syntax sugar to code.
10. Ability to pass any datatype by reference can be very beneficial.
11. Ability to create raw array datatypes. Enables full user interaction with WinAPI
12. Implementation of `LongPtr` which can work between both 64-bit and 32 bit systems

## More reading:

https://rubberduckvba.wordpress.com/2019/04/10/whats-wrong-with-vba/
