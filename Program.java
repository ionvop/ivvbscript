import java.io.*;
import java.nio.file.Paths;

public class Program {
    public static void main(String[] args) throws FileNotFoundException, IOException, InterruptedException {
        //String directory = System.getProperty("user.dir");
        String input = "";
        String outputPath = "";

        switch (args.length) {
            case 0:
                p.print("ivvbscript v1.1.0 by Ionvop");
                p.print("-i <filePath> //Sets the path of the source code to be transpiled.");
                p.print("-o <filePath> //Sets the path where the transpiled code will be written to. If not declared, the program will run directly from its source code.");
                p.print("-p //Prints the entire transpiled code.");
                return;
            case 1:
                input = args[0];
                break;
            default:
                input = getCommand(args, "-i");
                outputPath = getCommand(args, "-o");
        }

        String inputFolder = Paths.get(input).getParent().toString();
        ionFSO fso = new ionFSO();
        String content = fso.openFile(input);
        content += "\r\n";
        content = content.replaceAll(";\r", ";177013");
        content = content.replaceAll(";\n", ";177013");
        String[] untranspiledCode = content.split("\\;177013");
        String transpiledCode = "";

        for (String element : untranspiledCode) {
            transpiledCode = transpiledCode + transpile(element) + "\r\n";
        }

        if (findArray(args, "-p") != -1) {
            p.print(transpiledCode);
        }
    
        if (findArray(args, "-o") != -1) {
            fso.overwriteFile(outputPath, transpiledCode);
            p.print("Successfully transpiled to: " + outputPath);
        } else {
            fso.overwriteFile(inputFolder + "\\compiled.vbs", transpiledCode);
            Runtime.getRuntime().exec(new String[] {"C:\\Windows\\System32\\wscript.exe", "\"" + inputFolder + "\\compiled.vbs\""});
            Thread.sleep(1000);
            fso.deleteFile(inputFolder + "\\compiled.vbs");
        }
    }

    private static Print p = new Print();

    private static String getCommand(String[] input, String command) {
        var temp = findArray(input, command);

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

                    break;
                default:
                    if (input.substring(0, 1).equals(".")) {
                        parameter = input.substring(1);
                        parameter = transpile(parameter);
                        return "msgbox(" + parameter + ")";
                    } else if (input.substring(0, 1).equals(",")) {
                        if (input.indexOf(":") == -1) {
                            variable = input.substring(1);
                            variable = transpile(variable);
                            return variable + " = inputbox(\"\")";
                        }

                        variable = input.substring(1, input.indexOf(":"));
                        variable = transpile(variable);
                        parameter = input.substring(input.indexOf(":") + 1);
                        parameter = transpile(parameter);
                        return variable + " = inputbox(" + parameter + ")";
                    } else if (input.substring(0, 2).equals(":?")) {
                        condition = input.substring(2, input.indexOf(":", 2));
                        condition = transpile(condition);
                        statement = input.substring(input.indexOf(":", 2) + 1);
                        statement = transpile(statement);
                        return "elseif " + condition + " then\r\n" + statement;
                    } else if (input.substring(0, 1).equals("?")) {
                        condition = input.substring(1, input.indexOf(":"));
                        condition = transpile(condition);
                        statement = input.substring(input.indexOf(":") + 1);
                        statement = transpile(statement);
                        return "if " + condition + " then\r\n" + statement;
                    } else if (input.substring(0, 1).equals(":")) {
                        statement = input.substring(1);
                        statement = transpile(statement);
                        return "else\r\n" + statement;
                    } else if (input.substring(0, 1).equals("+")) {
                        variable = input.substring(1);
                        variable = transpile(variable);
                        return variable + " = " + variable + " + 1";
                    } else if (input.substring(0, 1).equals("-")) {
                        variable = input.substring(1);
                        variable = transpile(variable);
                        return variable + " = " + variable + " - 1";
                    } else if (input.substring(0, 1).equals(">")) {
                        parameter = input.substring(1, input.indexOf("{"));
                        parameter = transpile(parameter);
                        currentFunction = parameter.substring(0, parameter.indexOf("("));
                        currentFunction = transpile(currentFunction);
                        statement = input.substring(input.indexOf("{") + 1);
                        statement = transpile(statement);
                        return "function " + parameter + "\r\n" + statement;
                    } else if (input.substring(0, 1).equals("<")) {
                        parameter = input.substring(1);
                        parameter = transpile(parameter);
                        return currentFunction + " = " + parameter + "\r\nexit function";
                    } else if (input.substring(0, 1).equals("}")) {
                        statement = input.substring(1);
                        statement = transpile(statement);
                        return "end function\r\n" + statement;
                    } else if (input.substring(0, 1).equals("[")) {
                        statement = input.substring(1);
                        statement = transpile(statement);
                        return "do\r\n" + statement;
                    } else if (input.substring(0, 1).equals("]")) {
                        statement = input.substring(1);
                        statement = transpile(statement);
                        return "loop\r\n" + statement;
                    } else if (input.substring(0, 1).equals("#")) {
                        variable = input.substring(1, input.indexOf(":"));
                        variable = transpile(variable);
                        parameter = input.substring(input.indexOf(":") + 1);
                        parameter = transpile(parameter);
                        return "set " + variable + " = CreateObject(" + parameter + ")";
                    } else if (input.substring(0, 1).equals("=")) {
                        parameter = input.substring(1);
                        parameter = transpile(parameter);
                        return "eval(" + parameter + ")";
                    }
            }
        } catch(Exception e) {
            p.print("Transpilation error in statement: \"" + input + ";\"");
        }

        return input;
    }
}
