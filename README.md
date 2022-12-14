# ivvbscript
A custom language to VBScript transpiler written in a crappy language full of boilerplates and unnecessarily long syntaxes called Java for a crappy custom language full of unintuitive and unnecessarily short syntaxes that I made.

## Usage

### Command-line

`cd "c:\<projectDirectory>\" && javac Program.java && java Program`

### Arguments

`-i <filePath>` Sets the path of the source code to be transpiled.

`-o <filePath>` Sets the path where the transpiled code will be written to. If not declared, the program will run directly from the source code.

`-p` Prints the entire transpiled code to the terminal.

`-c` Transpiles the source code as cscript for console.

If only one argument is given, it will be treated as the path of the source code to be transpiled.

If no argument is given, it will display the list of arguments that can be used in this program.

### Running the source code

`cd "c:\<projectDirectory>\" && javac Program.java && java Program "c:\fileDirectory\program.ivvbs"`

### Exporting the transpiled code.

`cd "c:\<projectDirectory>\" && javac Program.java && java Program -i "c:\fileDirectory\program.ivvbs" -o "c:\fileDirectory\program.vbs"`

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

### Comment

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

### Dim

Declare a new variable or array variable within a scope.

`> <variables>`

`variables` The variables to be declared.

Example:

```
> var1, var2;
var1 = 7;
var2 = 4;

> add() {
    > var3;
    var3 = var1 + var2;
}
```

### Function

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

### Loop

Repeat a block of statements.

```
[
    <statements>
    /;
]
```

`statements` Program code to be repeated until `/;` is called.

### Object

Assign an object reference.

`# <variable> : <object>;`

`variable` The variable to be assigned.

`object` The name of the Windows Scripting Host (WSH) automation object.

Example:

```
# objShell : "wscript.shell";
objShell.Run("notepad.exe");
```

### Eval

Evaluate an expression.

`= <expression>;`

`expression` A string expression that returns a value.

Example:

```
x = "7 + 4";
.=x;
'Displays 11;
```

### Assign

Assign a value to a variable.

`= <variable> : <value>;`

`variable` The variable to be assigned.

`value` The value that is assigned to the variable.

Example:

```
= x : "Hello, world!";
.x;
'Displays "Hello, world!";
```

This is usually used with the eval statement.

```
var1 = 7;
var2 = 4;
var3 = "var1 + var2";
= result : = var3;
.result;
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

### Breakpoint

Stop the execution and print the value of the expression.

`!<expression>`

`expression` A statement that returns a value to be printed.

Example:

```
x = "Hello, world!";
!x;
'Stops the program and displays "Hello, world!";
```

### Options

`@e;` Option explicit. 	Forces all variables to be declared.

`@c;` Transpiles the source code as cscript for terminal.

`@p;` Prints the entire transpiled code to the terminal.

```
@e;
x = 1;
.x;
'Displays an error because x has not been declared yet before using;

@e;
> x;
x = 1;
.x;
'Displays 1;
```

### Template literal

Easily use double quotes and interpolate variables and expressions into a string.

`` `<text>{<expression>}` ``

`text` Text written into the string.

`expression` Variables or expressions to be interpolated into the string.

If there are characters included in the template literal that affects the template literal statement such as `` ` `` `\` `$`, you can add an escape character `\` before it to treat it as a character.

The escape symbol `\n` inserts a newline within the template literal.

It is required to escape the characters `;` `:` in strings and template literals when using it as a character.

Example:

```
var1 = "fox";
var2 = "dog";
.`The quick brown ${var1} "jumps" over the lazy ${var2}`;
```

## Loop techniques

### While

Conditionally repeat a block of statements.

```
[
    ? <condition>:;:
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

### 2022-11-22 v1.3.0

- Added an escape symbol `\n` to add a newline within a template literal.
- Adding an expression next to the exit statement `!;` turns it into a breakpoint statement `!<expression>;` and allows it to stop the program and print the value of the expression.
- It is now possible to use statement terminators `;` without having to pair it with a newline so it is now also possible to have multiple statements within a single line. However it is now required to escape semicolons `;` with the escape character `\` when using it in strings and template literals.
- Interpolating variables into template literals now use `${}` instead of only `{}`. This is so it is no longer required to escape the characters `{` and `}` when using it in template literals.
- It is now possible to run this program on older versions of Java where the method `.strip()` is not a thing yet.
- A complete rearrangement of the code where switch cases are used instead of if else statements for the `transpile()` method, and switched the switch case layering between `currentState` and `element` in the `scan()` method to make the code easier to read. However because of this, it is now required to escape the escape character `\\` when using it in template literals.
- Fixed a bug regarding colons where using colons within a string in a syntax where colons are optional are treated as part of the syntax. However it is now required to escape colons `:` with the escape character `\` when using it in strings and template literals.
- Fixed a bug where statements with only 2 characters are treated as constant methods that only require 2 characters and are not parsed.

### 2022-11-13 v1.2.0

- Added an argument command `-c` for transpiling the source code as cscript.
- Changed the default output filename of the transpiled code from `compiled.vbs` to the filename of the source code. This may overwrite and delete any existing `.vbs` files in the directory with the same filename.
- Added a statement `= <variables> : <parameters>;` as an alternative for assigning variables which supports the eval statement `= <expression>` within the parameters.
- Added a statement `> <variables>;` without `{ }` as an alternative for the statement `dim` when declaring a new variable.
- Added a statement `@e` as an alternative for `option explicit` statement when forcing all variables to be declared.
- Added a statement `@c` as an alternative for adding `-c` in the arguments when transpiling the source code as cscript.
- Added a statement `@p` as an alternative for adding `-p` in the arguments when printing the entire transpiled code to the terminal.
- Added a new literal `` ` ` `` which has a similar function as template literals in JavaScript.
- Added `run.vbs` in the sample folder to easily transpile and run a file within the directory without having to type the command in the terminal.
- Made some small optimizations.

### 2022-11-08 v1.1.0

- Added a help section that displays when no argument is entered.
- Added an output argument `-o` that saves the transpiled code to a file path.
- Changed the VBScript statement used for the print statement `.<variable>;` from `wscript.echo(<variable>)` to `msgbox(<variable>)`
- Changed a phrase in the transpilation error message from `Compilation error in line:` to `Transpilation error in statement:`

### 2022-11-05 v1.0.0

- Created this project as my first Java program and a practice material for CCE101, and a potential submission for my final exam.
- Released to GitHub.
