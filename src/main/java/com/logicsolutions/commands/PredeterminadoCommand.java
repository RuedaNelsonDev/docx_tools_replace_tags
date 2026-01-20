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
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.*;

public class PredeterminadoCommand {

    private static final Pattern TAG_PATTERN = Pattern.compile("<<[^>]+>>");

    public void execute(String[] args) {
        if (args.length != 3) {
            printUsage();
            return;
        }

        String inputDocx = args[0];
        String outputDocx = args[1];
        String textoReemplazo = args[2];

        System.out.println("========================================");
        System.out.println("COMANDO: predeterminado");
        System.out.println("Archivo origen: " + inputDocx);
        System.out.println("Archivo destino: " + outputDocx);
        System.out.println("Texto de reemplazo: " + textoReemplazo);
        System.out.println("Se reemplazarán TODAS las etiquetas del formato <<...>>");
        System.out.println("NOTA: Los cuadros de texto y campos NO serán modificados.");
        System.out.println("========================================");

        try {
            processDocument(inputDocx, outputDocx, textoReemplazo);
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void printUsage() {
        System.out.println("Uso: predeterminado <archivoOrigen> <archivoDestino> <textoReemplazo>");
        System.out.println("Ejemplo: predeterminado doc.docx doc_mod.docx \"[PENDIENTE]\"");
        System.out.println();
        System.out.println("Este comando reemplaza TODAS las etiquetas <<...>> con el texto especificado.");
    }

    private void processDocument(String inputDocx, String outputDocx, String textoReemplazo) throws Exception {
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

                        marcarElementosProtegidos(doc);
                        int totalReemplazos = reemplazarEnDocumento(doc, textoReemplazo);
                        System.out.println("  Total de etiquetas reemplazadas: " + totalReemplazos);
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

    private void marcarElementosProtegidos(Document doc) {
        // Marcar contenido de controles de contenido (sdtContent)
        NodeList sdtContents = doc.getElementsByTagName("w:sdtContent");
        for (int i = 0; i < sdtContents.getLength(); i++) {
            DocxUtils.marcarDescendientes(sdtContents.item(i));
        }

        // Marcar cuadros de texto
        NodeList textBoxes = doc.getElementsByTagName("w:txbxContent");
        for (int i = 0; i < textBoxes.getLength(); i++) {
            DocxUtils.marcarDescendientes(textBoxes.item(i));
        }

        // Marcar campos simples
        NodeList fieldSimples = doc.getElementsByTagName("w:fldSimple");
        for (int i = 0; i < fieldSimples.getLength(); i++) {
            DocxUtils.marcarDescendientes(fieldSimples.item(i));
        }

        // Marcar campos complejos (fldChar begin...end)
        NodeList fieldChars = doc.getElementsByTagName("w:fldChar");
        for (int i = 0; i < fieldChars.getLength(); i++) {
            Element fld = (Element) fieldChars.item(i);
            if ("begin".equals(fld.getAttribute("w:fldCharType"))) {
                Node sibling = fld.getParentNode();
                boolean inField = true;
                while (sibling != null && inField) {
                    DocxUtils.marcarDescendientes(sibling);
                    if (sibling.getNodeType() == Node.ELEMENT_NODE) {
                        NodeList endChars = ((Element) sibling).getElementsByTagName("w:fldChar");
                        for (int j = 0; j < endChars.getLength(); j++) {
                            Element endChar = (Element) endChars.item(j);
                            if ("end".equals(endChar.getAttribute("w:fldCharType"))) {
                                inField = false;
                                break;
                            }
                        }
                    }
                    sibling = sibling.getNextSibling();
                }
            }
        }

        // Marcar instrucciones de campo
        NodeList instrTexts = doc.getElementsByTagName("w:instrText");
        for (int i = 0; i < instrTexts.getLength(); i++) {
            ((Element) instrTexts.item(i)).setAttribute("ignorar", "true");
        }
    }

    private int reemplazarEnDocumento(Document document, String textoReemplazo) {
        int total = 0;
        total += reemplazarEnElementosIndividuales(document, textoReemplazo);
        total += reemplazarEnParagrafos(document, textoReemplazo);
        return total;
    }

    private int reemplazarEnElementosIndividuales(Document document, String textoReemplazo) {
        int count = 0;
        NodeList nodosTexto = document.getElementsByTagName("w:t");
        int elementosProcesados = 0;
        int elementosIgnorados = 0;

        for (int i = 0; i < nodosTexto.getLength(); i++) {
            Element elemento = (Element) nodosTexto.item(i);

            if (elemento.hasAttribute("ignorar") || DocxUtils.estaEnEstiloTOC(elemento)) {
                elementosIgnorados++;
                continue;
            }

            elementosProcesados++;
            String textoOriginal = elemento.getTextContent();
            Matcher matcher = TAG_PATTERN.matcher(textoOriginal);
            StringBuffer sb = new StringBuffer();
            boolean encontrado = false;

            while (matcher.find()) {
                System.out.println("    Encontrado: " + matcher.group() + " -> " + textoReemplazo);
                matcher.appendReplacement(sb, Matcher.quoteReplacement(textoReemplazo));
                count++;
                encontrado = true;
            }
            matcher.appendTail(sb);

            if (encontrado) {
                elemento.setTextContent(sb.toString());
                if (elemento.hasAttribute("xml:space")) {
                    elemento.setAttribute("xml:space", "preserve");
                }
            }
        }

        System.out.println("  Elementos procesados: " + elementosProcesados);
        System.out.println("  Elementos protegidos ignorados: " + elementosIgnorados);
        return count;
    }

    private int reemplazarEnParagrafos(Document document, String textoReemplazo) {
        int count = 0;
        NodeList paragraphs = document.getElementsByTagName("w:p");

        for (int i = 0; i < paragraphs.getLength(); i++) {
            Element paragraph = (Element) paragraphs.item(i);

            // Verificar si está en elemento protegido
            Node parent = paragraph.getParentNode();
            boolean enElementoProtegido = false;
            while (parent != null) {
                String name = parent.getNodeName();
                if ("w:txbxContent".equals(name) || "w:fldSimple".equals(name)) {
                    enElementoProtegido = true;
                    break;
                }
                parent = parent.getParentNode();
            }
            if (enElementoProtegido) continue;

            // Verificar estilo TOC
            NodeList pPrList = paragraph.getElementsByTagName("w:pPr");
            if (pPrList.getLength() > 0) {
                Element pPr = (Element) pPrList.item(0);
                NodeList pStyleList = pPr.getElementsByTagName("w:pStyle");
                if (pStyleList.getLength() > 0) {
                    Element pStyle = (Element) pStyleList.item(0);
                    String styleId = pStyle.getAttribute("w:val");
                    if (styleId != null && (styleId.startsWith("TOC") ||
                            styleId.startsWith("Toc") ||
                            styleId.contains("TableofContents") ||
                            styleId.contains("ndice"))) {
                        continue;
                    }
                }
            }

            // Recolectar runs
            NodeList runs = paragraph.getElementsByTagName("w:r");
            List<RunInfo> runInfos = new ArrayList<>();

            for (int j = 0; j < runs.getLength(); j++) {
                Element run = (Element) runs.item(j);
                NodeList texts = run.getElementsByTagName("w:t");
                for (int k = 0; k < texts.getLength(); k++) {
                    Element textEl = (Element) texts.item(k);
                    if (!textEl.hasAttribute("ignorar")) {
                        runInfos.add(new RunInfo(textEl, textEl.getTextContent()));
                    }
                }
            }

            if (runInfos.isEmpty()) continue;

            // Concatenar texto
            StringBuilder fullText = new StringBuilder();
            for (RunInfo ri : runInfos) {
                fullText.append(ri.text);
            }

            Matcher matcher = TAG_PATTERN.matcher(fullText.toString());
            StringBuffer sb = new StringBuffer();
            boolean encontrado = false;

            while (matcher.find()) {
                System.out.println("    Encontrado (fragmentado): " + matcher.group() + " -> " + textoReemplazo);
                matcher.appendReplacement(sb, Matcher.quoteReplacement(textoReemplazo));
                count++;
                encontrado = true;
            }
            matcher.appendTail(sb);

            if (encontrado) {
                String modified = sb.toString();
                runInfos.get(0).element.setTextContent(modified);
                for (int idx = 1; idx < runInfos.size(); idx++) {
                    runInfos.get(idx).element.setTextContent("");
                }
            }
        }

        return count;
    }
}
