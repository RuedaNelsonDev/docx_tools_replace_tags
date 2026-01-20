package com.logicsolutions.commands;

import com.logicsolutions.util.DocxUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.transform.Transformer;
import java.io.*;
import java.util.*;
import java.util.zip.*;

public class CuadrosTextoCommand {

    private Map<String, Integer> reemplazosPorTag = new HashMap<>();
    private Map<String, Integer> reemplazosPorArchivo = new HashMap<>();

    public void execute(String[] args) {
        if (args.length < 4 || (args.length - 2) % 2 != 0) {
            printUsage();
            return;
        }

        String inputFilePath = args[0];
        String outputFilePath = args[1];

        Map<String, String> reemplazos = new LinkedHashMap<>();
        for (int i = 2; i < args.length; i += 2) {
            reemplazos.put(args[i], args[i + 1]);
        }

        System.out.println("========================================");
        System.out.println("COMANDO: cuadros");
        System.out.println("Archivo origen: " + inputFilePath);
        System.out.println("Archivo destino: " + outputFilePath);
        System.out.println("Reemplazos a realizar:");
        for (Map.Entry<String, String> entry : reemplazos.entrySet()) {
            System.out.println("  - \"" + entry.getKey() + "\" -> \"" + entry.getValue() + "\"");
        }
        System.out.println("========================================");

        try {
            File tempFile = File.createTempFile("docx_temp", ".zip");
            tempFile.deleteOnExit();
            copyFile(new File(inputFilePath), tempFile);
            procesarArchivosXML(tempFile, outputFilePath, reemplazos);
            mostrarResumen();
            tempFile.delete();
        } catch (Exception e) {
            System.err.println("Error al procesar el archivo: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void printUsage() {
        System.out.println("Uso: cuadros <archivoOrigen> <archivoDestino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("Ejemplo: cuadros doc.docx doc_mod.docx \"<<NOMBRE>>\" \"Juan\" \"<<FECHA>>\" \"2024\"");
    }

    private void copyFile(File source, File dest) throws IOException {
        try (InputStream is = new FileInputStream(source);
             OutputStream os = new FileOutputStream(dest)) {
            byte[] buffer = new byte[4096];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
        }
    }

    private void procesarArchivosXML(File zipFile, String outputPath, Map<String, String> reemplazos) throws Exception {
        Map<String, byte[]> modifiedFiles = new HashMap<>();
        List<String> archivosAProcesar = new ArrayList<>(Arrays.asList(
                "word/document.xml",
                "word/header1.xml", "word/header2.xml", "word/header3.xml",
                "word/footer1.xml", "word/footer2.xml", "word/footer3.xml"
        ));

        // Buscar headers/footers adicionales
        try (ZipFile zip = new ZipFile(zipFile)) {
            Enumeration<? extends ZipEntry> entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();
                if ((name.startsWith("word/header") || name.startsWith("word/footer")) 
                        && name.endsWith(".xml") && !archivosAProcesar.contains(name)) {
                    archivosAProcesar.add(name);
                }
            }
        }

        DocumentBuilder builder = DocxUtils.createDocumentBuilder();
        Transformer transformer = DocxUtils.createTransformer();

        try (ZipFile zip = new ZipFile(zipFile)) {
            for (String archivoXML : archivosAProcesar) {
                ZipEntry entry = zip.getEntry(archivoXML);
                if (entry != null) {
                    System.out.println("\nProcesando: " + archivoXML);
                    try (InputStream is = zip.getInputStream(entry)) {
                        Document document = builder.parse(is);
                        int reemplazosEnArchivo = 0;

                        // Procesar diferentes tipos de cuadros de texto
                        reemplazosEnArchivo += procesarNodos(document, "w:txbxContent", reemplazos);
                        reemplazosEnArchivo += procesarNodos(document, "v:textbox", reemplazos);
                        reemplazosEnArchivo += procesarNodos(document, "wps:txbx", reemplazos);

                        if (reemplazosEnArchivo > 0) {
                            modifiedFiles.put(archivoXML, DocxUtils.documentToBytes(document, transformer));
                            reemplazosPorArchivo.merge(archivoXML, reemplazosEnArchivo, Integer::sum);
                        }
                    } catch (Exception e) {
                        System.err.println("  Error al procesar " + archivoXML + ": " + e.getMessage());
                    }
                }
            }
        }

        crearArchivoModificado(zipFile, outputPath, modifiedFiles);
    }

    private int procesarNodos(Document document, String tagName, Map<String, String> reemplazos) {
        int totalReemplazos = 0;
        NodeList nodes = document.getElementsByTagName(tagName);

        if (nodes.getLength() > 0) {
            System.out.println("  Encontrados " + nodes.getLength() + " elementos <" + tagName + ">");
        }

        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            totalReemplazos += procesarTextoEnNodo(node, reemplazos);
        }
        return totalReemplazos;
    }

    private int procesarTextoEnNodo(Node node, Map<String, String> reemplazos) {
        int totalReemplazos = 0;
        totalReemplazos += procesarTextoSimple(node, reemplazos);
        totalReemplazos += procesarTextoFragmentado(node, reemplazos);
        return totalReemplazos;
    }

    private int procesarTextoSimple(Node node, Map<String, String> reemplazos) {
        int totalReemplazos = 0;
        List<Node> textNodes = DocxUtils.buscarNodosRecursivamente(node, "w:t");

        for (Node textNode : textNodes) {
            String textoActual = textNode.getTextContent();
            String textoNuevo = textoActual;
            boolean modificado = false;

            for (Map.Entry<String, String> entry : reemplazos.entrySet()) {
                if (textoNuevo.contains(entry.getKey())) {
                    textoNuevo = textoNuevo.replace(entry.getKey(), entry.getValue());
                    modificado = true;
                    reemplazosPorTag.merge(entry.getKey(), 1, Integer::sum);
                }
            }

            if (modificado) {
                textNode.setTextContent(textoNuevo);
                totalReemplazos++;
                System.out.println("    Reemplazo simple: \"" + textoActual + "\" -> \"" + textoNuevo + "\"");
            }
        }
        return totalReemplazos;
    }

    private int procesarTextoFragmentado(Node node, Map<String, String> reemplazos) {
        int totalReemplazos = 0;
        List<Node> paragraphs = DocxUtils.buscarNodosRecursivamente(node, "w:p");

        for (Node paragraph : paragraphs) {
            List<Node> runs = DocxUtils.buscarNodosRecursivamente(paragraph, "w:r");
            List<Node> textNodes = new ArrayList<>();
            StringBuilder textoCompleto = new StringBuilder();

            for (Node run : runs) {
                List<Node> texts = DocxUtils.buscarNodosRecursivamente(run, "w:t");
                for (Node textNode : texts) {
                    textNodes.add(textNode);
                    textoCompleto.append(textNode.getTextContent());
                }
            }

            String textoOriginal = textoCompleto.toString();
            String textoNuevo = textoOriginal;
            boolean modificado = false;

            for (Map.Entry<String, String> entry : reemplazos.entrySet()) {
                if (textoNuevo.contains(entry.getKey())) {
                    textoNuevo = textoNuevo.replace(entry.getKey(), entry.getValue());
                    modificado = true;
                    reemplazosPorTag.merge(entry.getKey(), 1, Integer::sum);
                }
            }

            if (modificado && !textNodes.isEmpty()) {
                textNodes.get(0).setTextContent(textoNuevo);
                for (int n = 1; n < textNodes.size(); n++) {
                    textNodes.get(n).setTextContent("");
                }
                totalReemplazos++;
                System.out.println("    Reemplazo fragmentado: \"" + textoOriginal + "\" -> \"" + textoNuevo + "\"");
            }
        }
        return totalReemplazos;
    }

    private void mostrarResumen() {
        System.out.println("\n========================================");
        System.out.println("RESUMEN DE REEMPLAZOS:");
        int totalGeneral = 0;

        System.out.println("\nPor tag:");
        for (Map.Entry<String, Integer> entry : reemplazosPorTag.entrySet()) {
            System.out.println("  - \"" + entry.getKey() + "\": " + entry.getValue() + " reemplazos");
            totalGeneral += entry.getValue();
        }

        System.out.println("\nPor archivo:");
        for (Map.Entry<String, Integer> entry : reemplazosPorArchivo.entrySet()) {
            if (entry.getValue() > 0) {
                System.out.println("  - " + entry.getKey() + ": " + entry.getValue() + " reemplazos");
            }
        }

        System.out.println("\nTotal general: " + totalGeneral + " reemplazos");
        System.out.println("========================================");
    }

    private void crearArchivoModificado(File archivoOriginal, String outputPath, Map<String, byte[]> modifiedFiles) throws IOException {
        try (ZipFile zipOriginal = new ZipFile(archivoOriginal);
             FileOutputStream fos = new FileOutputStream(outputPath);
             ZipOutputStream zipOut = new ZipOutputStream(fos)) {

            Enumeration<? extends ZipEntry> entries = zipOriginal.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String entryName = entry.getName();

                zipOut.putNextEntry(new ZipEntry(entryName));
                if (modifiedFiles.containsKey(entryName)) {
                    zipOut.write(modifiedFiles.get(entryName));
                } else {
                    try (InputStream is = zipOriginal.getInputStream(entry)) {
                        byte[] buffer = new byte[4096];
                        int length;
                        while ((length = is.read(buffer)) > 0) {
                            zipOut.write(buffer, 0, length);
                        }
                    }
                }
                zipOut.closeEntry();
            }
        }
        System.out.println("\nArchivo modificado guardado en: " + outputPath);
    }
}
