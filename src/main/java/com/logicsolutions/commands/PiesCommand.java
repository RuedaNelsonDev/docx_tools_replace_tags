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

public class PiesCommand {

    public void execute(String[] args) {
        if (args.length < 4 || (args.length - 2) % 2 != 0) {
            printUsage();
            return;
        }

        String inputFilePath = args[0];
        String outputFilePath = args[1];

        // Construir mapa de reemplazos (múltiples pares tag/valor)
        Map<String, String> reemplazos = new LinkedHashMap<>();
        for (int i = 2; i < args.length; i += 2) {
            reemplazos.put(args[i], args[i + 1]);
        }

        System.out.println("========================================");
        System.out.println("COMANDO: pies");
        System.out.println("Archivo origen: " + inputFilePath);
        System.out.println("Archivo destino: " + outputFilePath);
        System.out.println("Reemplazos a realizar:");
        for (Map.Entry<String, String> entry : reemplazos.entrySet()) {
            System.out.println("  - \"" + entry.getKey() + "\" -> \"" + entry.getValue() + "\"");
        }
        System.out.println("========================================");

        try {
            processFooters(inputFilePath, outputFilePath, reemplazos);
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void printUsage() {
        System.out.println("Uso: pies <archivoOrigen> <archivoDestino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("Ejemplo: pies doc.docx doc_mod.docx \"<<PIE>>\" \"Página confidencial\"");
        System.out.println("Ejemplo: pies doc.docx doc_mod.docx \"<<PIE>>\" \"Confidencial\" \"<<PAGINA>>\" \"1\"");
    }

    private void processFooters(String inputPath, String outputPath, Map<String, String> reemplazos) throws Exception {
        DocumentBuilder builder = DocxUtils.createDocumentBuilder();
        Transformer transformer = DocxUtils.createTransformer();
        int totalReemplazos = 0;

        try (ZipFile zipFile = new ZipFile(inputPath);
             FileOutputStream fos = new FileOutputStream(outputPath);
             ZipOutputStream zipOut = new ZipOutputStream(fos)) {

            var entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();

                if (DocxUtils.isFooterXml(name)) {
                    System.out.println("\nProcesando pie de página: " + name);

                    try (InputStream is = zipFile.getInputStream(entry)) {
                        Document document = builder.parse(is);
                        int reemplazosEnArchivo = replaceTagsInDocument(document, reemplazos);
                        totalReemplazos += reemplazosEnArchivo;

                        zipOut.putNextEntry(new ZipEntry(name));
                        zipOut.write(DocxUtils.documentToBytes(document, transformer));
                        zipOut.closeEntry();
                    }
                } else {
                    DocxUtils.copyEntry(zipFile, entry, zipOut);
                }
            }
        }

        System.out.println("\n========================================");
        System.out.println("Total de reemplazos en pies de página: " + totalReemplazos);
        System.out.println("Archivo guardado en: " + outputPath);
        System.out.println("========================================");
    }

    private int replaceTagsInDocument(Document document, Map<String, String> reemplazos) {
        int count = 0;
        NodeList textNodes = document.getElementsByTagName("w:t");

        for (int i = 0; i < textNodes.getLength(); i++) {
            Node textNode = textNodes.item(i);
            String textoActual = textNode.getTextContent();
            String textoNuevo = textoActual;
            boolean modificado = false;

            for (Map.Entry<String, String> entry : reemplazos.entrySet()) {
                if (textoNuevo.contains(entry.getKey())) {
                    textoNuevo = textoNuevo.replace(entry.getKey(), entry.getValue());
                    modificado = true;
                }
            }

            if (modificado) {
                System.out.println("  Reemplazando: \"" + textoActual + "\" -> \"" + textoNuevo + "\"");
                textNode.setTextContent(textoNuevo);
                count++;
            }
        }
        return count;
    }
}
