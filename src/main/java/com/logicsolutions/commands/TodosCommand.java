package com.logicsolutions.commands;

import com.logicsolutions.util.DocxUtils;
import com.logicsolutions.util.DocxUtils.RunInfo;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.transform.Transformer;
import java.io.*;
import java.util.*;
import java.util.zip.*;

public class TodosCommand {

    public void execute(String[] args) {
        if (args.length < 4 || (args.length - 2) % 2 != 0) {
            printUsage();
            return;
        }

        String inputDocx = args[0];
        String outputDocx = args[1];

        Map<String, String> reemplazos = new LinkedHashMap<>();
        for (int i = 2; i < args.length; i += 2) {
            reemplazos.put(args[i], args[i + 1]);
        }

        System.out.println("========================================");
        System.out.println("COMANDO: todos");
        System.out.println("Archivo origen: " + inputDocx);
        System.out.println("Archivo destino: " + outputDocx);
        System.out.println("Reemplazos a realizar:");
        reemplazos.forEach((k, v) -> System.out.println("  - \"" + k + "\" -> \"" + v + "\""));
        System.out.println();
        System.out.println("NOTA: Los cuadros de texto NO serán modificados.");
        System.out.println("NOTA: El delimitador ||BR|| será convertido a salto de línea.");
        System.out.println("========================================");

        try {
            processDocument(inputDocx, outputDocx, reemplazos);
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void printUsage() {
        System.out.println("Uso: todos <archivoOrigen> <archivoDestino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("Ejemplo: todos doc.docx doc_mod.docx \"<<NOMBRE>>\" \"Juan\" \"<<FECHA>>\" \"2024\"");
        System.out.println();
        System.out.println("Para insertar saltos de línea, use ||BR|| en el valor de reemplazo:");
        System.out.println("  todos doc.docx doc_mod.docx \"<<DIRECCION>>\" \"Calle 1||BR||Ciudad||BR||País\"");
    }

    private void processDocument(String inputDocx, String outputDocx, Map<String, String> reemplazos) throws Exception {
        DocumentBuilder builder = DocxUtils.createDocumentBuilder();
        Transformer transformer = DocxUtils.createTransformer();

        try (ZipFile zipOrig = new ZipFile(inputDocx);
             FileOutputStream fos = new FileOutputStream(outputDocx);
             ZipOutputStream zipOut = new ZipOutputStream(fos)) {

            var entries = zipOrig.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String name = entry.getName();

                boolean esDocument = DocxUtils.isDocumentXml(name);
                boolean esHeader = DocxUtils.isHeaderXml(name);
                boolean esFooter = DocxUtils.isFooterXml(name);

                if (esDocument || esHeader || esFooter) {
                    System.out.println("\nProcesando: " + name);

                    try (InputStream is = zipOrig.getInputStream(entry)) {
                        Document doc = builder.parse(is);
                        doc.getDocumentElement().normalize();

                        marcarElementosEnTextBox(doc);
                        reemplazarEnDocumento(doc, reemplazos);
                        DocxUtils.limpiarMarcas(doc);

                        zipOut.putNextEntry(new ZipEntry(name));
                        zipOut.write(DocxUtils.documentToBytes(doc, transformer));
                        zipOut.closeEntry();
                    }
                } else {
                    DocxUtils.copyEntry(zipOrig, entry, zipOut);
                }
            }
        }

        System.out.println("\n========================================");
        System.out.println("Documento guardado en: " + outputDocx);
        System.out.println("========================================");
    }

    private void marcarElementosEnTextBox(Document doc) {
        NodeList textBoxes = doc.getElementsByTagName("w:txbxContent");
        for (int i = 0; i < textBoxes.getLength(); i++) {
            DocxUtils.marcarDescendientes(textBoxes.item(i));
        }
    }

    private void reemplazarEnDocumento(Document document, Map<String, String> reemplazos) {
        NodeList nodosTexto = document.getElementsByTagName("w:t");
        int elementosProcesados = 0;
        int elementosIgnorados = 0;

        List<Element> elementosAProcesar = new ArrayList<>();
        for (int i = 0; i < nodosTexto.getLength(); i++) {
            Element elemento = (Element) nodosTexto.item(i);
            if (elemento.hasAttribute("ignorar")) {
                elementosIgnorados++;
            } else {
                elementosProcesados++;
                elementosAProcesar.add(elemento);
            }
        }

        for (Element elemento : elementosAProcesar) {
            String textoOriginal = elemento.getTextContent();
            String textoModificado = textoOriginal;
            boolean huboReemplazo = false;

            for (Map.Entry<String, String> entrada : reemplazos.entrySet()) {
                if (textoModificado.contains(entrada.getKey())) {
                    textoModificado = textoModificado.replace(entrada.getKey(), entrada.getValue());
                    System.out.println("  Reemplazado: " + entrada.getKey() + " -> " + entrada.getValue());
                    huboReemplazo = true;
                }
            }

            if (huboReemplazo && !textoModificado.equals(textoOriginal)) {
                if (textoModificado.contains("||BR||")) {
                    DocxUtils.insertarTextoConSaltos(document, elemento, textoModificado);
                } else {
                    elemento.setTextContent(textoModificado);
                }
            }
        }

        System.out.println("  Elementos procesados: " + elementosProcesados);
        System.out.println("  Elementos en cuadros de texto ignorados: " + elementosIgnorados);

        reemplazarEnParagrafos(document, reemplazos);
    }

    private void reemplazarEnParagrafos(Document document, Map<String, String> reemplazos) {
        NodeList paragraphs = document.getElementsByTagName("w:p");

        for (int i = 0; i < paragraphs.getLength(); i++) {
            Element paragraph = (Element) paragraphs.item(i);

            // Verificar si está en un cuadro de texto
            Node parent = paragraph.getParentNode();
            boolean enTextBox = false;
            while (parent != null) {
                if (parent.getNodeName().equals("w:txbxContent")) {
                    enTextBox = true;
                    break;
                }
                parent = parent.getParentNode();
            }

            if (enTextBox) continue;

            // Recolectar runs y textos
            NodeList runs = paragraph.getElementsByTagName("w:r");
            List<RunInfo> runInfos = new ArrayList<>();

            for (int j = 0; j < runs.getLength(); j++) {
                Element run = (Element) runs.item(j);
                NodeList texts = run.getElementsByTagName("w:t");
                for (int k = 0; k < texts.getLength(); k++) {
                    Element textElement = (Element) texts.item(k);
                    if (!textElement.hasAttribute("ignorar")) {
                        runInfos.add(new RunInfo(textElement, textElement.getTextContent()));
                    }
                }
            }

            if (runInfos.isEmpty()) continue;

            // Concatenar texto completo
            StringBuilder fullText = new StringBuilder();
            for (RunInfo info : runInfos) {
                fullText.append(info.text);
            }

            String originalFullText = fullText.toString();
            String modifiedFullText = originalFullText;
            boolean huboReemplazos = false;

            for (Map.Entry<String, String> entrada : reemplazos.entrySet()) {
                if (modifiedFullText.contains(entrada.getKey())) {
                    modifiedFullText = modifiedFullText.replace(entrada.getKey(), entrada.getValue());
                    System.out.println("  Reemplazado (fragmentado): " + entrada.getKey() + " -> " + entrada.getValue());
                    huboReemplazos = true;
                }
            }

            if (huboReemplazos && !modifiedFullText.equals(originalFullText)) {
                if (modifiedFullText.contains("||BR||")) {
                    DocxUtils.insertarTextoConSaltos(document, runInfos.get(0).element, modifiedFullText);
                    for (int idx = 1; idx < runInfos.size(); idx++) {
                        runInfos.get(idx).element.setTextContent("");
                    }
                } else {
                    runInfos.get(0).element.setTextContent(modifiedFullText);
                    for (int idx = 1; idx < runInfos.size(); idx++) {
                        runInfos.get(idx).element.setTextContent("");
                    }
                }
            }
        }
    }
}
