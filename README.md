# MicroBeaut Visual Basic for Applications (VBA) Snippets

### Install snippets from the Marketplace
<br/>

<img src="./images/install-microbeaut-vba-snippets.gif" width="640px" height="480px">

<br/>

# Statements

## **``AppActivate``**
Activates an application window.

### **Prefix**
```vb
AppActivate 
```

### **Syntax**
```vb
AppActivate title, [ wait ]
```
<br/>


## **``Beep``**
Sounds a tone through the computer's speaker.

### **Prefix**
```vb
Beep 
```

### **Syntax**
```vb
Beep
```
<br/>

## **``Call``**
Transfers control to a Sub procedure, Function procedure, or dynamic-link library (DLL) procedure.

### **Prefix**
```vb
[ Call ] name [ argumentlist ]  
```

### **Syntax**
```vb
Call 
```
<br/>


## **``ChDir``**
Changes the current directory or folder.


### **Prefix**
```vb
ChDir
```

### **Syntax**
```vb
ChDir path
```
<br/>


## **``ChDrive``**
Changes the current drive.

### **Prefix**
```vb
ChDrive
```

### **Syntax**
```vb
ChDrive drive
```
<br/>


## **``Close ``**
Concludes input/output (I/O) to a file opened by using the Open statement.

### **Prefix**
```vb
Close 
```

### **Syntax**
```vb
Close [ filenumberlist ]
```
<br/>


## **``Const``**
Declares `constants` for use in place of literal values.

### **Prefix**
```vb
Const
```

### **Syntax**
```vb
[ Public | Private ] Const constname [ As type ] = expression
```
<br/>


## **``Date``**
Sets the current system date ``#mmmm d, yyyy#``.

### **Prefix**
```vb
Date
```

### **Syntax**
```vb
Date = date
```
<br/>


## **``DeleteSetting ``**
Deletes a section or key setting from an application's entry in the Windows ``registry`` or (on the Macintosh) information in the application's initialization file.

### **Prefix**
```vb
DeleteSetting 
```

### **Syntax**
```vb
DeleteSetting appname, section, key
```
<br/>


## **``Dim``**
Declares ``variables`` and allocates storage space.

### **Prefix**
```vb
Dim 
Dim WithEvents
```

### **Syntax**
```vb
Dim [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
```
<br/>


## **``Do...Loop``**
Repeats a block of ``statements`` while a condition is **True** or until a condition becomes **True**.

### **Prefix**
```vb
Do Until Loop
Do While Loop
```

### **Syntax**
```vb
Do [{ While | Until } condition ]
  [ statements ]
  [ Exit Do ]
  [ statements ]
Loop
```

Or,

### **Prefix**
```vb
Do Loop Until
Do Loop While
```

### **Syntax**
```vb
Do
  [ statements ]
  [ Exit Do ]
  [ statements ]
Loop [{ While | Until } condition ]
```
<br/>


## **``End``**
Ends a procedure or block.

### **Prefix**
```vb
End
End Function
End If
End Property
End Select
End Sub
End Type
End With
```

### **Syntax**
```vb
End
End Function
End If
End Property
End Select
End Sub
End Type
End With
```
<br/>


## **``Enum``**
Declares a type for an enumeration.

### **Prefix**
```vb
Enum
```

### **Syntax**
```vb
[ Public | Private ] Enum name
  membername [= constantexpression ]
  membername [= constantexpression ] . . .
End Enum
```
<br/>


## **``Erase``**
Reinitializes the elements of fixed-size ``arrays`` and releases dynamic-array storage space.

### **Prefix**
```vb
Erase
```

### **Syntax**
```vb
Erase arraylist
```
<br/>


## **``Error``**
Simulates the occurrence of an error.

### **Prefix**
```vb
Error 
```

### **Syntax**
```vb
Error errornumber
```
<br/>



## **``Event``**
Declares a user-defined event.

### **Prefix**
```vb
Event 
```

### **Syntax**
```vb
[ Public ] Event procedurename [ (arglist) ]
```
<br/>


## **``Exit``**
Exits a block of Do…Loop, For…Next, Function, Sub, or Property code.

### **Prefix**
```vb
Exit Do
Exit Fo
Exit Fu
Exit Pr
Exit Su
```

### **Syntax**
```vb
Exit Do
Exit For
Exit Function
Exit Property
Exit Sub
```
<br/>


## **``FileCopy``**
Copies a file.

### **Prefix**
```vb
FileCopy
```

### **Syntax**
```vb
FileCopy source, destination
```
<br/>

## **``For Each...Next``**
Repeats a group of ``statements`` for each element in an ``arrays`` or collection.

### **Prefix**
```vb
For Each
```

### **Syntax**
```vb
For Each element In group
  [ statements ]
  [ Exit For ]
  [ statements ]
Next [ element ]
```
<br/>


## **``For...Next``**
Repeats a group of ``statements`` a specified number of times.

### **Prefix**
```vb
For Next
```

### **Syntax**
```vb
For counter = start To end [ Step step ]
  [ statements ]
  [ Exit For ]
  [ statements ]
Next [ counter ]
```
<br/>


## **``Function``**
Declares the name, arguments, and code that form the body of a **Function** ``procedure``.

### **Prefix**
```vb
Function
Function Static
```

### **Syntax**
```vb
[Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
  [ statements ]
  [ name = expression ]
  [ Exit Function ]
  [ statements ]
  [ name = expression ]
End Function
```
<br/>


## **``Get``**
Reads data from an open disk file into a variable.

### **Prefix**
```vb
Get 
```

### **Syntax**
```vb
Get [ # ] filenumber, [ recnumber ], varname 
```
<br/>


## **``GoSub...Return``**
Branches to and returns from a subroutine within a procedure.

### **Prefix**
```vb
GoSub 
```

### **Syntax**
```vb
GoSub line
... line
line ...
Return
```
<br/>


## **``GoTo``**
Branches unconditionally to a specified line within a procedure.

### **Prefix**
```vb
GoTo line
```

### **Syntax**
```vb
GoTo line
```
<br/>

## **``If...Then...Else``**
Conditionally executes a group of ``statements``, depending on the value of an expression.

### **Prefix**
```vb
If 
```

### **Syntax**
```vb
If condition Then [ statements ] [ Else elsestatements ]

Or,

If condition Then
  [ statements ]
[ ElseIf condition-n Then
  [ elseifstatements ]]
[ Else
  [ elsestatements ]]
End If
```
<br/>

## **``Load``**
Loads an object but doesn't show it.

### **Prefix**
```vb
Load
```

### **Syntax**
```vb
Load object
```
<br/>


## **``Input #``**
Reads data from an open sequential file and assigns the data to variables.

### **Prefix**
```vb
Input 
```

### **Syntax**
```vb
Input #filenumber, varlist
```
<br/>


## **``Kill``**
Deletes files from a disk.

### **Prefix**
```vb
Kill 
```

### **Syntax**
```vb
Kill pathname
```
<br/>


## **``Let``**
Assigns the value of an expression to a variable or property.

### **Prefix**
```vb
Let 
```

### **Syntax**
```vb
[ Let ] varname = expression
```
<br/>

## **``Line Input #``**
Reads a single line from an open sequential file and assigns it to a String variable.

### **Prefix**
```vb
Line Input
```

### **Syntax**
```vb
Line Input #
```
<br/>


## **``Lock, Unlock``**
Controls access by other processes to all or part of a file opened by using the Open statement.

### **Prefix**
```vb
Lock 
Unlock 
```

### **Syntax**
```vb
Lock [ # ] filenumber, [ recordrange ]
Unlock [ # ] filenumber, [ recordrange ]
```
<br/>

## **``LSet``**
Left aligns a string within a string variable, or copies a variable of one ``user-defined type`` to another variable of a different user-defined type.

### **Prefix**
```vb
LSet
```

### **Syntax**
```vb
LSet stringvar = string
LSet varname1 = varname2
```
<br/>


## **``Mid``**
Replaces a specified number of characters in a **Variant (String)** variable with characters from another string.

### **Prefix**
```vb
Mid
```

### **Syntax**
```vb
Mid(stringvar, start, [ length ] ) = string
```
<br/>

## **``MkDir``**
Creates a new directory or folder.

### **Prefix**
```vb
MkDir
```

### **Syntax**
```vb
MkDir path
```
<br/>


## **``Name``**
Renames a disk file, directory, or folder.

### **Prefix**
```vb
Name 
```

### **Syntax**
```vb
Name oldpathname As newpathname
```
<br/>


## **``On Error``**
Enables an error-handling routine and specifies the location of the routine within a procedure; can also be used to disable an error-handling routine.

### **Prefix**
```vb
On Error Go
On Error Re
```

### **Syntax**
```vb
On Error GoTo 0
On Error GoTo line
On Error Resume Next
```
<br/>


## **``On...GoSub, On...GoTo``**
Branch to one of several specified lines, depending on the value of an expression.

### **Prefix**
```vb
On GoSub
On GoTo
```

### **Syntax**
```vb
On expression GoSub destinationlist
On expression GoTo destinationlist
```
<br/>

## **``Open``**
Enables input/output (I/O) to a file.

### **Prefix**
```vb
Open 
```

### **Syntax**
```vb
Open pathname For mode [ Access access ] [ lock ] As [ # ] filenumber [ Len = reclength ]
```
<br/>


## **``Option Base``**
Used at the module level to declare the default lower bound for array subscripts.

### **Prefix**
```vb
Option Ba
```

### **Syntax**
```vb
Option Base { 0 | 1 }
```
<br/>


## **``Option Compare``**
Used at the module level to declare the default comparison method to use when string data is compared.

### **Prefix**
```vb
Option Co
```

### **Syntax**
```vb
Option Compare { Binary | Text | Database }
```
<br/>


## **``Option Explicit``**
Used at the module level to force explicit declaration of all variables in that module.

### **Prefix**
```vb
Option Ex
```

### **Syntax**
```vb
Option Explicit
```
<br/>

## **``Option Private Module``**
When used in host applications that allow references across multiple projects, **Option Private Module** prevents a module's contents from being referenced outside its project. In host applications that don't permit such references, for example, standalone versions of Visual Basic, **Option Private** has no effect.

### **Prefix**
```vb
Option Pr
```

### **Syntax**
```vb
Option Private Module
```
<br/>


## **``Print #``**
Writes display-formatted data to a sequential file.

### **Prefix**
```vb
Print #
```

### **Syntax**
```vb
Print #filenumber, [ outputlist ]
```
<br/>


## **``Private``**
Used at the module level to declare private variables and allocate storage space.

### **Prefix**
```vb
Private
Private With
```

### **Syntax**
```vb
Private [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
```
<br/>



## **``Property Get``**
Declares the name, arguments, and code that form the body of a **Property** procedure, which gets the value of a property.

### **Prefix**
```vb
Property Get
Property Get Static
```

### **Syntax**
```vb
[ Public | Private | Friend ] [ Static ] Property Get name [ (arglist) ] [ As type ]
  [ statements ]
  [ name = expression ]
[ Exit Property ]
  [ statements ]
  [ name = expression ]
End Property
```
<br/>


## **``Property Let``**
Declares the name, arguments, and code that form the body of a **Property** procedure, which assigns a value to a property.

### **Prefix**
```vb
Property Let
Property Let Static
```

### **Syntax**
```vb
[ Public | Private | Friend ] [ Static ] Property Let name ( [ arglist ], value )
  [ statements ]
  [ Exit Property ]
  [ statements ]
End Property
```
<br/>

## **``Property Set``**
Declares the name, arguments, and code that form the body of a **Property** procedure, which sets a reference to an object.

### **Prefix**
```vb
Property Set
Property Set Static
```

### **Syntax**
```vb
[ Public | Private | Friend ] [ Static ] Property Set name ( [ arglist ], reference )
  [ statements ]
  [ Exit Property ]
  [ statements ]
End Property
```
<br/>


## **``Public``**
Used at the module level to declare public variables and allocate storage space.

### **Prefix**
```vb
Public
Public With
```

### **Syntax**
```vb
Public [ WithEvents ] varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
```
<br/>

## **``Put``**
Writes data from a variable to a disk file.

### **Prefix**
```vb
Put 
```

### **Syntax**
```vb
Put [ # ] filenumber, [ recnumber ], varname
```
<br/>

## **``RaiseEvent``**
Fires an event declared at the module level within a class, form, or document.

### **Prefix**
```vb
RaiseEvent 
```

### **Syntax**
```vb
RaiseEvent eventname [ ( argumentlist ) ]
```
<br/>

## **``Randomize``**
Initializes the random-number generator.

### **Prefix**
```vb
Randomize 
```

### **Syntax**
```vb
Randomize [ number ]
```
<br/>


## **``ReDim``**
Used at the procedure level to reallocate storage space for dynamic array variables.

### **Prefix**
```vb
ReDim 
```

### **Syntax**
```vb
ReDim [ Preserve ] varname ( subscripts )
```
<br/>

## **``Rem``**
Used to include explanatory remarks in a program.

### **Prefix**
```vb
Rem  
```

### **Syntax**
```vb
Rem comment
```
<br/>

## **``Reset``**
Closes all disk files opened by using the Open statement.

### **Prefix**
```vb
Reset  
```

### **Syntax**
```vb
Reset
```
<br/>

## **``Resume``**
Removes an existing directory or folder.

### **Prefix**
```vb
Resume 
Resume Ne
Resume Li
```

### **Syntax**
```vb
Resume [ 0 ]
Resume Next
Resume line
```
<br/>


## **``RmDir``**
Removes an existing directory or folder.

### **Prefix**
```vb
RmDir
```

### **Syntax**
```vb
RmDir path
```
<br/>


## **``RSet``**
Right aligns a string within a string variable, or copies a variable of one user-defined type to another variable of a different user-defined type.

### **Prefix**
```vb
RSet
```

### **Syntax**
```vb
RSet stringvar = string
RSet varname1 = varname2
```
<br/>


## **``SaveSetting``**
Saves or creates an application entry in the application's entry in the Windows registry or (on the Macintosh) information in the application's initialization file.

### **Prefix**
```vb
SaveSetting
```

### **Syntax**
```vb
SaveSetting appname, section, key, setting
```

### **Remarks**
The root of these registry settings is: <code>Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings</code>.
<br/>


## **``Seek``**
Sets the position for the next read/write operation within a file opened by using the Open statement.

### **Prefix**
```vb
Seek 
```

### **Syntax**
```vb
Seek [ # ] filenumber, position
```
<br/>


## **``Select Case``**
Executes one of several groups of statements, depending on the value of an expression.

### **Prefix**
```vb
Select Case
```

### **Syntax**
```vb
Select Case testexpression
[ Case expressionlist-n
  [ statements-n ]]
[ Case Else 
  [ elsestatements ]]
End Select
```
<br/>

## **``SendKeys``**
Sends one or more keystrokes to the active window as if typed at the keyboard.

### **Prefix**
```vb
SendKeys
```

### **Syntax**
```vb
SendKeys [ # ] filenumber, position
```
<br/>


## **``Set``**
Assigns an object reference to a variable or property.

### **Prefix**
```vb
Set 
```

### **Syntax**
```vb
Set objectvar = {[ New ] objectexpression | Nothing }
```
<br/>

## **``SetAttr``**
Sets attribute information for a file.

### **Prefix**
```vb
SetAttr 
```

### **Syntax**
```vb
SetAttr pathname, attributes
```
<br/>


## **``Static``**
Used at the procedure level to declare variables and allocate storage space. Variables declared with the **Static** statement retain their values as long as the code is running.

### **Prefix**
```vb
Static
```

### **Syntax**
```vb
Static varname [ ( [ subscripts ] ) ] [ As [ New ] type ]
```
<br/>

## **``Stop``**
Suspends execution.

### **Prefix**
```vb
Stop
```

### **Syntax**
```vb
Stop
```
<br/>


## **``Sub``**
Declares the name, arguments, and code that form the body of a **Sub** procedure.

### **Prefix**
```vb
Sub
Sub Static
```

### **Syntax**
```vb
[ Private | Public | Friend ] [ Static ] Sub name [ ( arglist ) ]
  [ statements ]
  [ Exit Sub ]
  [ statements ]
End Sub
```
<br/>


## **``Time``**
Sets the system time. ``#hh:mm:ss AM/PM#``

### **Prefix**
```vb
Time 
```

### **Syntax**
```vb
Time = time
```
<br/>


## **``Type``**
Used at the module level to define a user-defined data type containing one or more elements.

### **Prefix**
```vb
Type 
```

### **Syntax**
```vb
[ Private | Public ] Type varname
  elementname [ ( [ subscripts ] ) ] As type
  [ elementname [ ( [ subscripts ] ) ] As type ] . . .
End Type
```
<br/>

## **``Unload``**
Removes an object from memory.

### **Prefix**
```vb
Unload 
```

### **Syntax**
```vb
Unload object
```
<br/>


## **``While...Wend``**
Removes an object from memory.

### **Prefix**
```vb
While Wend
```

### **Syntax**
```vb
While condition
  [ statements ]
Wend
```
<br/>


## **``Width #``**
Assigns an output line width to a file opened by using the Open statement.

### **Prefix**
```vb
Width 
```

### **Syntax**
```vb
Width #filenumber, width
```
<br/>


## **``With``**
Executes a series of statements on a single object or a user-defined type.

### **Prefix**
```vb
With
```

### **Syntax**
```vb
With object
  [ statements ]
End With
```
<br/>

## **``Write #``**
Writes data to a sequential file.

### **Prefix**
```vb
Write
```

### **Syntax**
```vb
Write #filenumber, [ outputlist ]
```
<br/>

# Conversion functions

## **``Asc``**
Returns an Integer representing the character code corresponding to the first letter in a string.

### **Prefix**
```vb
Asc
```

### **Syntax**
```vb
Asc(string)
```
<br/>

## **``Chr``**
Returns a String containing the character associated with the specified character code.

### **Prefix**
```vb
Chr
```

### **Syntax**
```vb
Chr(charcode)
```
<br/>

## **``Format``**
Returns a Variant (String) containing an expression formatted according to instructions contained in a format expression.

### **Prefix**
```vb
Format
```

### **Syntax**
```vb
Format(Expression, [ Format ], [ FirstDayOfWeek ], [ FirstWeekOfYear ])
```
<br/>

## **``Hex``**
Returns a Variant (String) containing an expression formatted according to instructions contained in a format expression.

### **Prefix**
```vb
Hex
```

### **Syntax**
```vb
Hex(number)
```
<br/>


## **``Oct``**
Returns a Variant (String) representing the octal value of a number.

### **Prefix**
```vb
Oct
```

### **Syntax**
```vb
Oct(number)
```
<br/>

## **``Str``**
Returns a Variant (String) representation of a number.

### **Prefix**
```vb
Str
```

### **Syntax**
```vb
Str(number)
```
<br/>

## **``Val``**
Returns the numbers contained in a string as a numeric value of appropriate type.

### **Prefix**
```vb
Val
```

### **Syntax**
```vb
Val(string)
```
<br/>


# Math functions

## **``Abs``**
Returns a value of the same type that is passed to it specifying the absolute value of a number.

### **Prefix**
```vb
Abs
```

### **Syntax**
```vb
Abs(number)
```
<br/>


## **``Atn``**
Returns a Double specifying the arctangent of a number.

### **Prefix**
```vb
Atn
```

### **Syntax**
```vb
Atn(number)
```
<br/>


## **``Cos``**
Returns a Double specifying the cosine of an angle.

### **Prefix**
```vb
Cos
```

### **Syntax**
```vb
Cos(number)
```
<br/>


## **``Exp``**
Returns a Double specifying e (the base of natural logarithms) raised to a power.

### **Prefix**
```vb
Exp
```

### **Syntax**
```vb
Exp(number)
```
<br/>


## **``Int, Fix``**
Returns a Double specifying e (the base of natural logarithms) raised to a power.

### **Prefix**
```vb
Int
Fix
```

### **Syntax**
```vb
Int(number)
Fix(number)
```
<br/>

## **``Log``**
Returns a Double specifying the natural logarithm of a number.

### **Prefix**
```vb
Log
```

### **Syntax**
```vb
Log(number)
```
<br/>

## **``Rnd``**
Returns a Single containing a pseudo-random number.

### **Prefix**
```vb
Rnd
```

### **Syntax**
```vb
Rnd [ (Number) ]
```
<br/>

## **``Sgn``**
Returns a Variant (Integer) indicating the sign of a number.

### **Prefix**
```vb
Sgn
```

### **Syntax**
```vb
Sgn(number)
```
<br/>


## **``Sin``**
Returns a Double specifying the cosine of an angle.

### **Prefix**
```vb
Sin
```

### **Syntax**
```vb
Sin(number)
```
<br/>

## **``Sqr``**
Returns a Double specifying the sine of an angle.

### **Prefix**
```vb
Sqr
```

### **Syntax**
```vb
Sqr(number)
```
<br/>

## **``Tan ``**
Returns a Double specifying the tangent of an angle.

### **Prefix**
```vb
Tan 
```

### **Syntax**
```vb
Tan (number)
```
<br/>



# Reference

### [Language reference for Visual Basic for Applications (VBA)](https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference)
<br/>


# Release Notes

### [0.0.1]
- Initial release of MicroBeaut VBA Snippets

### [0.0.2]
- Changed package description
- Revised statements
- Added new statements

### [0.0.3]
- Changed the prefix for,
    - ```Dim WithEvent```
    - ``` On Error *```
    - ```Private WithEvent```
    - ```Public WithEvent```
<br/>


### [0.0.4]
- Removed statements ``If...Then..Exit`` inside,
    - ``Do...Loop``
    - ``For Each...Next``
    - ``For...Next``
    - ``Function``
    - ``If...Then...Else``
- Updated descriptions
- Added new statements
- Added Conversion functions
- Added Math functions

### [0.0.5]
- Fixed package not updated.
<br/>

# License

MIT License

Copyright &copy; 2023 MicroBeaut