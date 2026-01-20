package com.logicsolutions;

import com.logicsolutions.commands.*;

import java.util.Arrays;

public class Main {

    private static final String VERSION = "1.0.0";

    public static void main(String[] args) {
        if (args.length < 1) {
            printHelp();
            return;
        }

        String command = args[0].toLowerCase();
        String[] commandArgs = Arrays.copyOfRange(args, 1, args.length);

        switch (command) {
            case "cuadros":
                new CuadrosTextoCommand().execute(commandArgs);
                break;

            case "encabezados":
                new EncabezadosCommand().execute(commandArgs);
                break;

            case "pies":
                new PiesCommand().execute(commandArgs);
                break;

            case "todos":
                new TodosCommand().execute(commandArgs);
                break;

            case "predeterminado":
                new PredeterminadoCommand().execute(commandArgs);
                break;
            case "loop":
                new LoopCommand().execute(commandArgs);
                break;
            case "help":
            case "-h":
            case "--help":
                printHelp();
                break;

            case "version":
            case "-v":
            case "--version":
                System.out.println("docx-tool version " + VERSION);
                break;

            default:
                System.err.println("Comando desconocido: " + command);
                System.err.println("Use 'java -jar docx-tool.jar help' para ver los comandos disponibles.");
                System.exit(1);
        }
    }

    private static void printHelp() {
        System.out.println("╔══════════════════════════════════════════════════════════════════╗");
        System.out.println("║                    DOCX-TOOL v" + VERSION + "                              ║");
        System.out.println("║     Herramienta para reemplazo de tags en documentos Word        ║");
        System.out.println("╚══════════════════════════════════════════════════════════════════╝");
        System.out.println();
        System.out.println("USO: java -jar docx-tool.jar <comando> [argumentos]");
        System.out.println();
        System.out.println("COMANDOS DISPONIBLES:");
        System.out.println();
        System.out.println("  cuadros       Reemplaza tags SOLO en cuadros de texto");
        System.out.println("                Uso: cuadros <origen> <destino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("                Ej:  cuadros doc.docx out.docx \"<<NOMBRE>>\" \"Juan\"");
        System.out.println();
        System.out.println("  encabezados   Reemplaza tags SOLO en encabezados (headers)");
        System.out.println("                Uso: encabezados <origen> <destino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("                Ej:  encabezados doc.docx out.docx \"<<TITULO>>\" \"Mi Doc\"");
        System.out.println();
        System.out.println("  pies          Reemplaza tags SOLO en pies de página (footers)");
        System.out.println("                Uso: pies <origen> <destino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("                Ej:  pies doc.docx out.docx \"<<PIE>>\" \"Confidencial\"");
        System.out.println();
        System.out.println("  todos         Reemplaza tags en TODO el documento EXCEPTO cuadros de texto");
        System.out.println("                Soporta saltos de línea con ||BR||");
        System.out.println("                Uso: todos <origen> <destino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("                Ej:  todos doc.docx out.docx \"<<DIR>>\" \"Calle 1||BR||Ciudad\"");
        System.out.println();
        System.out.println("  predeterminado  Reemplaza TODAS las etiquetas <<...>> con un valor único");
        System.out.println("                  Uso: predeterminado <origen> <destino> <textoReemplazo>");
        System.out.println("                  Ej:  predeterminado doc.docx out.docx \"[PENDIENTE]\"");
        System.out.println();
        System.out.println("  help          Muestra esta ayuda");
        System.out.println("  version       Muestra la versión del programa");
        System.out.println();
        System.out.println("NOTAS:");
        System.out.println("  - Los tags deben estar en formato <<TAG>>");
        System.out.println("  - Use comillas para valores con espacios");
        System.out.println("  - El comando 'todos' ignora cuadros de texto (use 'cuadros' para esos)");
    }
}
