---
layout: post
title:  "Issues with VBA"
published: true
authors:
  - "Sancarn"
---

This post is still under development but here's a dump of all issues. Hopefully I'll get around to finishing it some day 👀

## Fundamental issues with the language

1. Object creation and method call is slow when compared to modern languages (and compared to module only code).
2. Hidden features which are unimplementable - e.g. unable to implement IEnumVARIANT. I.E. No custom `For each ... in ... next` implementations (unless you delegate to a collection)
3. Some low level standard interfaces are forbidden in VBA (like IDispatch)
4. Inability to define `hidden` methods.
5. Inability to use `Evaluate` methods on custom classes unless you initially cast to Object (IDispatch) which also removes complete intellisense...
6. Inconsistent setting of variables - the set keyword was added because Microsoft wanted to allow for use of default properties on COM objects. It would have been better if calling of default members had a short-hand syntax like `something@()`. Having to write `set` everywhere for the benefit of unclear code like `v = Range("A1:B3")` is not a good trade-off.
7. `ByRef` is default where `ByVal` is more logical. Assume this was initially to keep code optimal, but generally leads to difficulty in learning. People often just learn "Use `ByVal` everywhere", which isn’t correct either.
8. Inability to define collective/generic types e.g. `Collection<Car>`. This leads to limitations in the type system.
9. Inability to define 1st class executable blocks of code
    * Means no native anonymous function syntax. Understandable due to VBA’s age, but modern languages use a lot of Lambda syntax to make code cleaner.
    * Means no native multithreading support.
    * Means no async/await support without creating a seperate runtime (see `stdFiber` for one such runtime)
    * Means no native GPU execution support
    * Worked around a bit with `stdVBA`'s `stdLambda` compiler/interpreter but far from ideal.
10. Lack of true inheritance (can sort of fake it with defaults)
11. No overloading - not a requirement, but a nice feature.
12. Interfacing - Lack of implicit cast to interfaces can make use somewhat clumsy: `Dim x as class: set x = new class: Dim y as IClass: set y = x: y.poop()`
13. Poor error reporting -  Lack of native stack traces, lack of line numbers in errors
14. Lack of reflection, metaprogramming and dynamic dependency injection.
15. Structs/Types appear to have been bodged on top of the runtime. No union types exist and recursive types don’t exist either.
16. [Introduction of the pointless keyword `PtrSafe`](https://stackoverflow.com/a/77141128/6302131) - This keyword provides no function whatsoever nor any guarantee of safety. Also ridiculous that this only exists in some later versions of VBA…
17. VBA Keywords - `Print`, `Write`, `Debug`, … - All of these are methods you cannot use! This wouldn’t be the case if VBA namespace wasn’t globally accessible. Why was it designed this way?!
18. Inability to pass structs `BYVAL` to low level functions
10. VBA Attributes are not editable post-import, classes with custom attributes have to be imported and can't just have attributes sitting in code. On many occasions users get confused when trying to use classes with Predeclared IDs as a result, because these must be imported, the code cannot be copy/pasted.
20. Lack of `static` functions (without also making the class pre-declared). VBA does have "static" keyword which can be assosciated with functions, but has a **completely different meaning** to that used by the rest of the programming world (and is rarely useful).
21. `Any` type exists for declares but cannot be used in user created subs and functions.
22. If an object wants to pass itself into a function call, i.e. `callback.run(me)` this will throw an error "Object doesn't support this property or method". Instead you have to create a reference of `Me` and pass this in instead `Dim oMe as Object: set oMe = me: callback.run(oMe)`
23. No block comments.
24. Only allowed 25 line continuations.
25. Application.Run only allows 30 params.
26. [This x64 bug](https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes)
27. Class Constructors and better intialisation `Dim x as New Y(a,b,c)`.
28. Difference between `Dim X as New Y` and `Dim X as Y: set X = new Y`.
29. Many functions and classes have been added to the [global namespace unnecessarily](https://rubberduckvba.blog/2024/08/14/understanding-libraries/comment-page-1/#respond). This has caused numerous issues  


## Issues which can be resolved with libraries

1. Lack of Events on base types - e.g. `Collection::Add()`, `Collection::Remove()`
2. Lack of component based design for UserForms. - Modern UI frameworks are so much better at this by now!
3. Lack of standard libraries:
    * No canvas component for UserForms - Fundamentally limits what you can do in a userform.
    * No standard built-in implementation of HashMap - Dictionary is great, but requires external DLL and isn’t available on Mac.
    * Lack of standard libraries - there are community solutions for this e.g. stdVBA and vb core lib
    * Lack of native libraries forces people to use Low level APIs which only work on single platforms - Yes VBA can run on Mac but due to the lack of Windows APIs on Mac, what you can do with VBA is severely limited.
4. Poor authentication with other Microsoft Services like Sharepoint Online. Excel can query from a Sharepoint Online list, so why is it so hard to do so with VBA?

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

--------

