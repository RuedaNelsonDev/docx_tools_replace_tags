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

/**
 * Comando para reemplazo de tags en modo LOOP.
 * - Procesa cada párrafo como unidad
 * - Reemplaza la PRIMERA ocurrencia del tag en cada párrafo
 * - Soporta ||BR|| para saltos de línea
 * - Ignora cuadros de texto
 * - Maneja tags fragmentados (divididos entre múltiples w:r)
 * - PRESERVA saltos de línea existentes
 */
public class LoopCommand {

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
        System.out.println("COMANDO: loop");
        System.out.println("Archivo origen: " + inputDocx);
        System.out.println("Archivo destino: " + outputDocx);
        System.out.println("Reemplazos a realizar:");
        reemplazos.forEach((k, v) -> System.out.println("  - \"" + k + "\" -> \"" + v + "\""));
        System.out.println();
        System.out.println("NOTA: Los cuadros de texto NO serán modificados.");
        System.out.println("NOTA: El delimitador ||BR|| será convertido a salto de línea.");
        System.out.println("NOTA: Solo se reemplaza la PRIMERA ocurrencia por párrafo.");
        System.out.println("========================================");

        try {
            processDocument(inputDocx, outputDocx, reemplazos);
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void printUsage() {
        System.out.println("Uso: loop <archivoOrigen> <archivoDestino> <tag1> <valor1> [<tag2> <valor2> ...]");
        System.out.println("Ejemplo: loop doc.docx doc_mod.docx \"<<NOMBRE>>\" \"Juan\"");
        System.out.println();
        System.out.println("Para insertar saltos de línea, use ||BR|| en el valor de reemplazo:");
        System.out.println("  loop doc.docx doc_mod.docx \"<<ITEMS>>\" \"Item1||BR||    Item2||BR||    Item3\"");
        System.out.println();
        System.out.println("Este comando está diseñado para procesamiento iterativo de secciones.");
        System.out.println("Reemplaza la PRIMERA ocurrencia de cada tag en cada párrafo.");
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
                        int totalReemplazos = reemplazarEnParagrafos(doc, reemplazos);
                        DocxUtils.limpiarMarcas(doc);

                        System.out.println("  Total reemplazos: " + totalReemplazos);

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

    private int reemplazarEnParagrafos(Document document, Map<String, String> reemplazos) {
        int totalReemplazos = 0;
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

            // Recolectar runs y textos del párrafo
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

            // PRIMERO: Intentar reemplazar en elementos individuales (sin concatenar)
            // Esto preserva los saltos de línea existentes
            boolean reemplazoIndividual = false;
            for (Map.Entry<String, String> entrada : reemplazos.entrySet()) {
                String tag = entrada.getKey();
                String valor = entrada.getValue();

                for (RunInfo info : runInfos) {
                    if (info.text.contains(tag)) {
                        String nuevoTexto = info.text.replace(tag, valor);
                        System.out.println("  Reemplazado en párrafo " + i + ": " + tag);

                        if (nuevoTexto.contains("||BR||")) {
                            DocxUtils.insertarTextoConSaltos(document, info.element, nuevoTexto);
                        } else {
                            info.element.setTextContent(nuevoTexto);
                        }

                        reemplazoIndividual = true;
                        totalReemplazos++;
                        break; // Solo primera ocurrencia por párrafo
                    }
                }

                if (reemplazoIndividual) break;
            }

            // Si no se encontró en elementos individuales, buscar tags fragmentados
            if (!reemplazoIndividual) {
                // Concatenar texto completo del párrafo
                StringBuilder fullText = new StringBuilder();
                for (RunInfo info : runInfos) {
                    fullText.append(info.text);
                }

                String originalFullText = fullText.toString();
                String modifiedFullText = originalFullText;
                boolean huboReemplazos = false;

                // Reemplazar tags fragmentados
                for (Map.Entry<String, String> entrada : reemplazos.entrySet()) {
                    String tag = entrada.getKey();
                    String valor = entrada.getValue();

                    int idx = modifiedFullText.indexOf(tag);
                    if (idx >= 0) {
                        modifiedFullText = modifiedFullText.substring(0, idx) + valor +
                                modifiedFullText.substring(idx + tag.length());
                        System.out.println("  Reemplazado (fragmentado) en párrafo " + i + ": " + tag);
                        huboReemplazos = true;
                        totalReemplazos++;
                        break; // Solo primera ocurrencia
                    }
                }

                if (huboReemplazos && !modifiedFullText.equals(originalFullText)) {
                    if (modifiedFullText.contains("||BR||")) {
                        DocxUtils.insertarTextoConSaltos(document, runInfos.get(0).element, modifiedFullText);
                    } else {
                        runInfos.get(0).element.setTextContent(modifiedFullText);
                    }

                    // Limpiar los demás elementos del párrafo
                    for (int idx = 1; idx < runInfos.size(); idx++) {
                        runInfos.get(idx).element.setTextContent("");
                    }
                }
            }
        }

        return totalReemplazos;
    }
}