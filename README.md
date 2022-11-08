# ivvbscript
A VBScript transpiler made in Java for a custom language I made

## Documentation

### Print

Display a dialogue box message.

`.<message>`

`message` Message to be displayed.

Example:

```
."Hello, world!";

.7 + 4;
'Displays 11;
```

### Input

Prompt for user input.

`,<variable>`

`,<variable>:<prompt>`

`variable` The variable to be assigned.

`prompt` The dialogue box prompt text.

Example:

```
,input : "Enter your message";
.input;
```

### Comments

Add a comment.

`'<comment>;`

`comment` An explanatory remark not included in the execution of the program.

Example:

```
,input : "Enter your message";
'This is a comment;
.input;
```

### If Then

Conditionally execute a block of statements.

```
? <condition>:
    <statements>
:? <condition>:
    <statements>
:
    <statements>
?;
```

`condition` An expression that evaluates to True or False.

`statements` Program code to be executed if `condition` is True.

Example:

```
,input : "What's 9 + 10?";

? input = 21:
    ."You stupid";
:? input = 19:
    ."Correct";
:
    ."Wrong";
?;
```

### Functions

Define a function procedure.

```
> <name>(<parameters>) {
    <statements>
    <<expression>;
}
```

`name` The name of the function.

`parameters` Argument variabless passed to the function, comma separated.

`statements` Program code.

`expression` The value to return.

Example:

```
> add(val1, val2) {
    < val1 + val2;
}

.add(4, 7);
'Displays 11;
```

### Increment/Decrement

Increases/decreases the value of a variable by 1.

`+<variable>;`

`-<variable>;`

`variable` The variable to be incremented/decremented.

Example:

```
x = 5;
+x;
.x;
'Displays 6;

y = 5;
-y;
.y;
'Displays 4;
```

### Loops

Repeat a block of statements.

```
[
    <statements>
    /;
]
```

`statements` Program code to be repeated until `/;` is called.

### Objects

Assign an object reference.

`# <variable> : <object>`

`variable` The variable to be assigned.

`object` The name of the Windows Scripting Host (WSH) automation object.

Example:

```
# objShell : "wscript.shell";
objShell.Run("notepad.exe");
```

### Eval

Evaluate an expression.

`=<expression>;`

`expression` Any expression that returns a string or a number.

Example:

```
x = "7 + 4";
.=x;
'Displays 11;
```

### Exit

Stop the execution.

`!;`

Example:

```
x = "Hello, world!";
!;

'Unreachable code;
.x;
```

## Loop techniques

### While

Conditionally repeat a block of statements.

```
[
    ? <condition>:
    :;
        /;
    ?;
    
    <statements>
]
```

### For

Repeat a block of statements a given number of times.

```
<counter> = initialValue - 1;

[
    +<counter>;
    
    ? <counter> > <maxValue>:
        /;
    ?;
    
    <statement>
]
```
