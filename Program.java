import java.io.*;
import java.nio.file.Paths;

public class Program {
    public static void main(String[] args) throws FileNotFoundException, IOException, InterruptedException {
        String directory = System.getProperty("user.dir");
        String input = "";
        String outputPath = "";

        switch (args.length) {
            case 0:
                p.print("ivvbscript v1.1.0 by Ionvop");
                p.print("-i <filePath> //Sets the path of the source code to be transpiled.");
                p.print("-o <filePath> //Sets the path where the transpiled code will be written to. If not declared, the program will run directly from its source code.");
                p.print("-p //Prints the entire transpiled code.");
                p.print("-c //Transpiles the source code as cscript for console.");
                return;
            case 1:
                input = args[0];
                break;
            default:
                input = getParameter(args, "-i");
                outputPath = getParameter(args, "-o");

                if (findArray(args, "-c") != -1) {
                    isConsole = true;
                }

                if (findArray(args, "-p") != -1) {
                    isPrinted = true;
                }
        }

        String inputFileName = Paths.get(input).getFileName().toString();
        inputFileName = inputFileName.substring(0, inputFileName.lastIndexOf("."));
        String inputFolder = Paths.get(input).getParent().toString();
        String content = fso.openFile(input);
        content += "\r\n";
        content = content.replaceAll("\\\\\r\n", "");
        content = content.replaceAll("\\\\;", "177013escapedsemicolon");
        content = content.replaceAll("\\\\:", "177013escapedcolon");
        String[] untranspiledCode = content.split(";");
        String transpiledCode = "";

        for (String element : untranspiledCode) {
            transpiledCode += transpile(element) + "\r\n";
        }

        transpiledCode = transpiledCode.replaceAll("177013escapedsemicolon", ";");
        transpiledCode = transpiledCode.replaceAll("177013escapedcolon", ":");

        if (isPrinted) {
            p.print(transpiledCode);
        }
    
        if (findArray(args, "-o") != -1) {
            fso.overwriteFile(outputPath, transpiledCode);
            p.print("Successfully transpiled to: " + outputPath);
        } else {
            fso.overwriteFile(inputFolder + "\\" + inputFileName + ".vbs", transpiledCode);

            if (isConsole) {
                fso.overwriteFile(directory + "\\run.vbs", "CreateObject(\"WScript.Shell\").run(\"%comspec% /k cscript \"\"" + inputFolder + "\\" + inputFileName + ".vbs\"\"\")");
                Runtime.getRuntime().exec(new String[] {"C:\\Windows\\System32\\wscript.exe ", "\"" + directory + "\\run.vbs\""});
                Thread.sleep(1000);
                fso.deleteFile(directory + "\\run.vbs");
            } else {
                Runtime.getRuntime().exec(new String[] {"C:\\Windows\\System32\\wscript.exe ", "\"" + inputFolder + "\\" + inputFileName + ".vbs\""});
            }

            Thread.sleep(1000);
            fso.deleteFile(inputFolder + "\\" + inputFileName + ".vbs");
        }
    }

    private static Print p = new Print();

    private static ionFSO fso = new ionFSO();

    private static Boolean isConsole = false;

    private static Boolean isPrinted = false;

    private static String getParameter(String[] input, String command) {
        int temp = findArray(input, command);

        if (temp == -1) {
            return "";
        } else {
            return input[temp + 1];
        }
    }

    private static int findArray(String[] input, String toFind) {
        int i = 0;

        for (String element : input) {
            if (element.equals(toFind)) {
                return i;
            }

            i++;
        }

        return -1;
    }

    private static String currentFunction = "";

    private static String transpile(String input) {
        String parameter;
        String variable;
        String condition;
        String statement;
        input = input.strip();

        try {
            input = input.strip();
        } catch (Exception e) {
            input = input.trim();
        }

        try {
            switch (input.length()) {
                case 0:
                    return input;
                case 1:
                    switch (input) {
                        case "?":
                            return "end if";
                        case "}":
                            return "end function";
                        case "]":
                            return "loop";
                        case "/":
                            return "exit do";
                        case "!":
                            return "wscript.quit";
                    }
                case 2:
                    switch (input) {
                        case "@c":
                            isConsole = true;
                            return "'Set as a console script";
                        case "@e":
                            return "option explicit";
                        case "@p":
                            isPrinted = true;
                            return "'Is printed to terminal";
                    }
                default:
                    if (input.length() > 2) {
                        switch (input.substring(0, 2)) {
                            case ":?":
                                condition = input.substring(2, input.indexOf(":", 2));
                                condition = transpile(condition);
                                statement = input.substring(input.indexOf(":", 2) + 1);
                                statement = transpile(statement);
                                return "elseif " + condition + " then\r\n" + statement;
                        }
                    }

                    switch (input.substring(0, 1)) {
                        case ".":
                            parameter = input.substring(1);
                            parameter = transpile(parameter);

                            if (isConsole) {
                                return "wscript.echo(" + parameter + ")";
                            }

                            return "msgbox(" + parameter + ")";
                        case ",":
                            if (input.contains(":")) {
                                variable = input.substring(1, input.indexOf(":"));
                                variable = transpile(variable);
                                parameter = input.substring(input.indexOf(":") + 1);
                                parameter = transpile(parameter);

                                if (isConsole) {
                                    return "wscript.echo(" + parameter + ")\r\n" + variable + " = WScript.StdIn.ReadLine()";
                                }

                                return variable + " = inputbox(" + parameter + ")";
                            }

                            variable = input.substring(1);
                            variable = transpile(variable);

                            if (isConsole) {
                                return variable + " = WScript.StdIn.ReadLine()";
                            }

                            return variable + " = inputbox(\"\")";
                        case "?":
                            condition = input.substring(1, input.indexOf(":"));
                            condition = transpile(condition);
                            statement = input.substring(input.indexOf(":") + 1);
                            statement = transpile(statement);
                            return "if " + condition + " then\r\n" + statement;
                        case ":":
                            statement = input.substring(1);
                            statement = transpile(statement);
                            return "else\r\n" + statement;
                        case "+":
                            variable = input.substring(1);
                            variable = transpile(variable);
                            return variable + " = " + variable + " + 1";
                        case "-":
                            variable = input.substring(1);
                            variable = transpile(variable);
                            return variable + " = " + variable + " - 1";
                        case ">":
                            if (input.contains("{")) {
                                parameter = input.substring(1, input.indexOf("{"));
                                parameter = transpile(parameter);
                                currentFunction = parameter.substring(0, parameter.indexOf("("));
                                currentFunction = transpile(currentFunction);
                                statement = input.substring(input.indexOf("{") + 1);
                                statement = transpile(statement);
                                return "function " + parameter + "\r\n" + statement;
                            }
                            
                            parameter = input.substring(1);
                            parameter = transpile(parameter);
                            return "dim " + parameter;
                        case "<":
                            parameter = input.substring(1);
                            parameter = transpile(parameter);
                            return currentFunction + " = " + parameter + "\r\nexit function";
                        case "}":
                            statement = input.substring(1);
                            statement = transpile(statement);
                            return "end function\r\n" + statement;
                        case "[":
                            statement = input.substring(1);
                            statement = transpile(statement);
                            return "do\r\n" + statement;
                        case "]":
                            statement = input.substring(1);
                            statement = transpile(statement);
                            return "loop\r\n" + statement;
                        case "#":
                            variable = input.substring(1, input.indexOf(":"));
                            variable = transpile(variable);
                            parameter = input.substring(input.indexOf(":") + 1);
                            parameter = transpile(parameter);
                            return "set " + variable + " = CreateObject(" + parameter + ")";
                        case "=":
                            if (input.contains(":")) {
                                variable = input.substring(1, input.indexOf(":"));
                                variable = transpile(variable);
                                parameter = input.substring(input.indexOf(":") + 1);
                                parameter = transpile(parameter);
                                return variable + " = " + parameter;
                            }

                            parameter = input.substring(1);
                            parameter = transpile(parameter);
                            return "eval(" + parameter + ")";
                        case "!":
                            parameter = input.substring(1);
                            parameter = transpile(parameter);
                            return "wscript.echo(" + parameter + ")\r\nwscript.quit";
                    }

                    if (input.contains("`")) {
                        return scan(input);
                    }
            }
        } catch (Exception e) {
            p.print("Transpilation error in statement: \"" + input + ";\"");
        }

        return input;
    }

    private static String scan(String input) {
        String currentState = "normal";
        char element = ' ';
        String scannedInput = "";

        try {
            for (int i = 0; i < input.length(); i++) {
                element = input.charAt(i);

                switch (currentState) {
                    case "normal":
                        switch (element) {
                            case '\"':
                                currentState = "string";
                                scannedInput += "\"";
                                break;
                            case '`':
                                currentState = "template";
                                scannedInput += "\"";
                                break;
                            default:
                                scannedInput += element;
                        }

                        break;
                    case "string":
                        switch (element) {
                            case '"':
                                currentState = "normal";
                                scannedInput += "\"";
                                break;
                            default:
                                scannedInput += element;
                        }

                        break;
                    case "template":
                        switch (element) {
                            case '`':
                                currentState = "normal";
                                scannedInput += "\"";
                                break;
                            case '\"':
                                scannedInput += "\"\"";
                                break;
                            case '\\':
                                currentState = "escape";
                                break;
                            case '$':
                                if (input.charAt(i + 1) == '{') {
                                    currentState = "interpolate";
                                } else {
                                    scannedInput += "$";
                                }
                                
                                break;
                            default:
                                scannedInput += element;
                        }

                        break;
                    case "escape":
                        switch (element) {
                            case 'n':
                                scannedInput += "\" & vbcrlf & \"";
                                break;
                            default:
                                scannedInput += element;
                        }

                        currentState = "template";
                        break;
                    case "interpolate":
                        switch (element) {
                            case '{':
                                scannedInput += "\" & ";
                                break;
                            case '}':
                                currentState = "template";
                                scannedInput += " & \"";
                                break;
                            default:
                                scannedInput += element;
                                break;
                        }
                }
            }
        } catch (Exception e) {
            p.print("Transpilation error in statement: \"" + input + ";\"");
        }

        return scannedInput;
    }
}
