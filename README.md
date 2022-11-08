# ivvbscript
A VBScript transpiler made in Java for a custom language I made

## Usage

### Command-line

`cd "c:\<projectDirectory>\" && javac Program.java && java Program`

### Arguments

`-i <filePath>` Sets the path of the source code to be transpiled.

`-o <filePath>` Sets the path where the transpiled code will be written to. If not declared, the program will run directly from its source code.

`-p` Prints the entire transpiled code to the terminal.

If only one argument is given, it will be treated as the path of the source code to be transpiled.

If no argument is given, it will display the list of arguments that can be used in this program.

### Running the source code

`cd "c:\<projectDirectory>\" && javac Program.java && java Program "c:\fileDirectory\program.ivvbs"`

## Documentation

### Print

Display a dialogue box message.

`.<message>;`

`message` Message to be displayed.

Example:

```
."Hello, world!";

.7 + 4;
'Displays 11;
```

### Input

Prompt for user input.

`,<variable>;`

`,<variable> : <prompt>;`

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
    < <expression>;
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

## Changelog

### v1.1.0

- Added a help section that displays when no argument is entered.
- Added an output argument `-o` that saves the transpiled code to a file path.
- Changed the VBScript statement used for the print syntax `.<variable>;` from `wscript.echo(<variable>)` to `msgbox(<variable>)`
- Changed a phrase in the transpilation error message from `Compilation error in line:` to `Transpilation error in statement:`

### v1.0.0

- Created this project as a Java practice material for CCE101 and a potential submission for my final exam.
- Released to GitHub.
