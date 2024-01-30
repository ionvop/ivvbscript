# ivvbscript (IonVop VBScript)
A custom language to VBScript transpiler written in VBScript that focuses on hacky syntax at the cost of readability and intuitiveness.

## Usage

`ivvbs "path\to\file.ivvbs"`

## Documentation

### Variable declaration

Declare a new variable or array variable within a scope.

`> <variables>`

`variables` The variables to be declared.

Example:

```
> var1, var2;
```

### Comment

Add a comment.

`'<comment>;`

`comment` An explanatory remark not included in the execution of the program.

Example:

```
'This is a comment;
```

### Print

Display a dialogue box message.

`.<value>;`

`message` Message to be displayed.

Example:

```
."Hello, world!";
'Outputs Hello, world!;

.7 + 4;
'Outputs 11;
```

### Input

Prompt for user input.

`,<variable>;`

`,<variable> : <prompt>;`

`variable` The variable to be assigned.

`prompt` The prompt text.

Example:

```
> input;
,input: "Enter your message";
.input;
```

### If Then

Conditionally execute a block of statements.

```
? <condition>:
    <statements>
:? <elseif_condition>:
    <elseif_statements>
:
    <else_statements>
?;
```

`condition` An expression that evaluates to True or False.

`statements` Program code to be executed if `condition` is True.

Example:

```
> input;
,input: "What's 9 + 10?";

? input = 21:
    ."You stupid";
:? input = 19:
    ."Correct";
:
    ."Wrong";
?;
```

### Function

Define a function procedure.

```
> <name>(<parameters>)
    <statements>
< <value>;
```

`name` The name of the function.

`parameters` Argument variabless passed to the function, comma separated.

`statements` Program code.

`value` The value to return.

Example:

```
> add(num1, num2)
    > result;
    result = num1 + num2;
< result;

.add(4, 7);
'Outputs 11;
```

### Operators

Adds/subtracts/multiplies/divides/concatenates the value of a variable by a given value.

`+<variable>: <value>;`

`-<variable>: <value>;`

`*<variable>: <value>;`

`/<variable>: <value>;`

`&<variable>: <value>;`

`variable` The variable to be incremented/decremented.

If a value is not given, the variable is added/subtracted by 1, multiplied/divided by 2, or concatenated with a newline by default.

Example:

```
> num;
num = 7;

+num: 4; 'num is now 11;
-num: 3; 'num is now 8;
*num: 4: 'num is now 32;
/num: 8; 'num is now 4;
&num: "hello"; 'num is now "4hello";

num = 4;

+num; 'num is now 5;
-num; 'num is now 4;
*num; 'num is now 8;
/num; 'num is now 4;
&num; 'num is now "4\n";
```

### Loop

Repeat a block of statements.

```
[
    <statements>
    /
]
```

`statements` Program code to be repeated until `/` is called.

### Exit

Stop the execution and print a message.

`!<message>`

`message` Message to be displayed.

Example:

```
> x;
x = "Hello, world!";
!x;
'Stops the program and outputs "Hello, world!";
```

### Template literal

Easily use double quotes and interpolate variables and expressions into a string.

`` `<text>{<expression>}` ``

`text` Text written into the string.

`expression` Variables or expressions to be interpolated into the string.

If there are characters included in the template literal that affects the template literal statement such as `` ` `` `\` `$`, you can add an escape character `\` before it to treat it as a character.

The escape symbol `\n` inserts a newline within the template literal.

Example:

```
> var1, var2;
var1 = "fox";
var2 = "dog";
.`The quick brown ${var1} "jumps" over the lazy ${var2}`;
```

### Special objects

Commonly used ActiveX objects in VBScript projects can now be accessed using a special character.

`$` WScript.Shell

`#` Scripting.FileSystemObject

`%` MSXML2.XMLHTTP.6.0

Example:

```
> content;
content = #opentextfile("test.txt").readall();
.content;

$run("mspaint");
```

## Loop techniques

### While

Conditionally repeat a block of statements.

```
[
    ? <condition>::
        /
    ?;
    
    <statements>
]
```

### For

Repeat a block of statements a given number of times.

```
> <counter>;
<counter> = initialValue;

[
    ? <condition>::
        /;
    ?;
    
    <statements>
    +<counter>;
]
```

## Changelog

### 2024-01-28 v2.0.0

- A complete rewrite of the source code from scratch.
- Rewritten in VBScript because it funny and improved maintainability of the code.
- It is no longer required to escape semicolons `;` within strings and template literals.
- Scripts can be executed using the `ivvbs` command.

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
