# MicroBeaut VBA Snippets
MicroBeaut Visual Basic for Applications (VBA) Snippets

# Language Support
Visual Basic for Applications (VBA)
 
# Statements

## **Dim statement**
Declares variables

### ***Prefix***
```vb
Dim
```

### ***Syntax***
```vb
Dim varname [ ( [ subscripts ] ) ] [ As type ]
```


## **Do...Loop statement**
Repeats a block of statements while a condition is True or until a condition becomes True.

### ***Prefix***
```vb
Do Until Loop

Or,

Do While Loop
```

### ***Syntax***
```vb
Do [{ While | Until } condition ]
  [ statements ]
  [ Exit Do ]
  [ statements ]
Loop
```

Or,

### ***Prefix***
```vb
Do Loop Until

or,

Do Loop While
```

### ***Syntax***
```vb
Do
  [ statements ]
  [ Exit Do ]
  [ statements ]
Loop [{ While | Until } condition ]
```


## **For Each...Next statement**
Repeats a group of statements for each element in an array or collection.

### ***Prefix***
```vb
For Each
```

### ***Syntax***
```vb
For Each element In group
  [ statements ]
  [ Exit For ]
  [ statements ]
Next [ element ]
```


## **For...Next statement**
Repeats a group of statements a specified number of times.

### ***Prefix***
```vb
For Next
```

### ***Syntax***
```vb
For counter = start To end [ Step step ]
  [ statements ]
  [ Exit For ]
  [ statements ]
Next [ counter ]
```


## **Function statement**
Declares the name, arguments, and code that form the body of a Function procedure.

### ***Prefix***
```vb
Function
```

### ***Syntax***
```vb
[Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
  [ statements ]
  [ name = expression ]
  [ Exit Function ]
  [ statements ]
  [ name = expression ]
End Function
```


## **If...Then...Else statement**
Conditionally executes a group of statements, depending on the value of an expression.

### ***Prefix***
```vb
If
```

### ***Syntax***
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


## **Property Get statement**
Declares the name, arguments, and code that form the body of a Property procedure, which gets the value of a property.

### ***Prefix***
```vb
Property Get
```

### ***Syntax***
```vb
[ Public | Private ] Property Get name [ (arglist) ] [ As type ]
  [ statements ]
  [ name = expression ]
  [ Exit Property ]
  [ statements ]
  [ name = expression ]
End Property
```


## **Property Let statement**
Declares the name, arguments, and code that form the body of a Property procedure, which assigns a value to a property.

### ***Prefix***
```vb
Property Let
```

### ***Syntax***
```vb
[ Public | Private ] Property Let name ( [ arglist ], value )
  [ statements ]
  [ Exit Property ]
  [ statements ]
End Property
```


## **Property Set statement**
Declares the name, arguments, and code that form the body of a Property procedure, which sets a reference to an object.

### ***Prefix***
```vb
Property Set
```

### ***Syntax***
```vb
[ Public | Private] Property Set name ( [ arglist ], reference )
  [ statements ]
  [ Exit Property ]
  [ statements ]
End Property
```


## **ReDim statement**
Used at the procedure level to reallocate storage space for dynamic array variables.

### ***Prefix***
```vb
ReDim
```

### ***Syntax***
```vb
ReDim [ Preserve ] varname ( subscripts )
```


## **Select Case statement**
Executes one of several groups of statements, depending on the value of an expression.

### ***Prefix***
```vb
Select Case
```

### ***Syntax***
```vb
Select Case testexpression
[ Case expressionlist-n
  [ statements-n ]]
[ Case Else 
  [ elsestatements ]]
End Select
```

# Release Notes

### 0.0.1

- Initial release of MicroBeaut VBA Snippets


# License

Copyright (c) 2023 MicroBeaut

Licensed under the MIT License.