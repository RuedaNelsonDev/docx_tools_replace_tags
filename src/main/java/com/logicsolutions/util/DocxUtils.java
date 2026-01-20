package com.logicsolutions.util;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class DocxUtils {

    public static final String WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static DocumentBuilder createDocumentBuilder() throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        return factory.newDocumentBuilder();
    }

    public static Transformer createTransformer() throws Exception {
        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer transformer = factory.newTransformer();
        transformer.setOutputProperty("indent", "no");
        transformer.setOutputProperty("method", "xml");
        transformer.setOutputProperty("encoding", "UTF-8");
        transformer.setOutputProperty("standalone", "yes");
        return transformer;
    }

    public static byte[] documentToBytes(Document doc, Transformer transformer) throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        transformer.transform(new DOMSource(doc), new StreamResult(baos));
        byte[] result = baos.toByteArray();
        baos.close();
        return result;
    }

    public static void copyEntry(ZipFile zipFile, ZipEntry entry, ZipOutputStream zipOut) throws Exception {
        zipOut.putNextEntry(new ZipEntry(entry.getName()));
        try (InputStream is = zipFile.getInputStream(entry)) {
            byte[] buffer = new byte[4096];
            int len;
            while ((len = is.read(buffer)) > 0) {
                zipOut.write(buffer, 0, len);
            }
        }
        zipOut.closeEntry();
    }

    public static boolean isDocumentXml(String name) {
        return "word/document.xml".equals(name);
    }

    public static boolean isHeaderXml(String name) {
        return name.startsWith("word/header") && name.endsWith(".xml");
    }

    public static boolean isFooterXml(String name) {
        return name.startsWith("word/footer") && name.endsWith(".xml");
    }

    /**
     * Marca todos los elementos w:t descendientes de un nodo con el atributo "ignorar"
     */
    public static void marcarDescendientes(Node node) {
        if (node.getNodeName().equals("w:t") && node.getNodeType() == Node.ELEMENT_NODE) {
            ((Element) node).setAttribute("ignorar", "true");
        }
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            marcarDescendientes(children.item(i));
        }
    }

    /**
     * Limpia el atributo "ignorar" de todos los elementos w:t
     */
    public static void limpiarMarcas(Document doc) {
        NodeList nodosTexto = doc.getElementsByTagName("w:t");
        for (int i = 0; i < nodosTexto.getLength(); i++) {
            Element element = (Element) nodosTexto.item(i);
            if (element.hasAttribute("ignorar")) {
                element.removeAttribute("ignorar");
            }
        }
    }

    /**
     * Busca nodos recursivamente por nombre de tag
     */
    public static List<Node> buscarNodosRecursivamente(Node parent, String tagName) {
        List<Node> nodes = new ArrayList<>();
        buscarNodosRecursivamenteHelper(parent, tagName, nodes);
        return nodes;
    }

    private static void buscarNodosRecursivamenteHelper(Node node, String tagName, List<Node> foundNodes) {
        if (node.getNodeName().equals(tagName)) {
            foundNodes.add(node);
        }
        NodeList children = node.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            buscarNodosRecursivamenteHelper(children.item(i), tagName, foundNodes);
        }
    }

    /**
     * Verifica si un elemento está dentro de un estilo TOC (tabla de contenidos)
     */
    public static boolean estaEnEstiloTOC(Element elemento) {
        Node parent = elemento.getParentNode();
        while (parent != null && !parent.getNodeName().equals("w:p")) {
            parent = parent.getParentNode();
        }
        if (parent != null && parent.getNodeType() == Node.ELEMENT_NODE) {
            NodeList pPrList = ((Element) parent).getElementsByTagName("w:pPr");
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
                        return true;
                    }
                }
            }
        }
        return false;
    }

    /**
     * Inserta texto con saltos de línea (||BR|| -> w:br)
     */
    public static void insertarTextoConSaltos(Document document, Element elementoTexto, String texto) {
        // Buscar el w:r padre (puede no ser el padre directo)
        Node runNode = elementoTexto.getParentNode();
        while (runNode != null && !runNode.getNodeName().equals("w:r")) {
            runNode = runNode.getParentNode();
        }

        if (runNode == null) {
            // No se encontró w:r, intentar crear estructura básica
            System.out.println("  ADVERTENCIA: No se encontró w:r padre, insertando texto sin formato de salto");
            elementoTexto.setTextContent(texto.replace("||BR||", "\n"));
            return;
        }

        Element run = (Element) runNode;
        Node parentOfRun = run.getParentNode();

        if (parentOfRun == null) {
            elementoTexto.setTextContent(texto.replace("||BR||", "\n"));
            return;
        }

        String[] lineas = texto.split("\\|\\|BR\\|\\|", -1);

        if (lineas.length <= 1) {
            elementoTexto.setTextContent(texto);
            return;
        }

        System.out.println("  Insertando " + (lineas.length - 1) + " saltos de línea");

        // Obtener formato original
        Element rPrOriginal = null;
        NodeList rPrList = run.getElementsByTagName("w:rPr");
        if (rPrList.getLength() > 0) {
            rPrOriginal = (Element) rPrList.item(0);
        }

        // Primera línea en el elemento original
        elementoTexto.setTextContent(lineas[0]);

        Node insertAfter = run;
        for (int i = 1; i < lineas.length; i++) {
            // Crear run con salto de línea
            Element nuevoRunBr = document.createElementNS(WORD_NS, "w:r");
            if (rPrOriginal != null) {
                nuevoRunBr.appendChild(rPrOriginal.cloneNode(true));
            }
            Element br = document.createElementNS(WORD_NS, "w:br");
            nuevoRunBr.appendChild(br);

            if (insertAfter.getNextSibling() != null) {
                parentOfRun.insertBefore(nuevoRunBr, insertAfter.getNextSibling());
            } else {
                parentOfRun.appendChild(nuevoRunBr);
            }
            insertAfter = nuevoRunBr;

            // Crear run con texto (incluso si está vacío, para mantener estructura)
            Element nuevoRunTexto = document.createElementNS(WORD_NS, "w:r");
            if (rPrOriginal != null) {
                nuevoRunTexto.appendChild(rPrOriginal.cloneNode(true));
            }
            Element nuevoT = document.createElementNS(WORD_NS, "w:t");
            if (lineas[i].startsWith(" ") || lineas[i].endsWith(" ") || lineas[i].isEmpty()) {
                nuevoT.setAttribute("xml:space", "preserve");
            }
            nuevoT.setTextContent(lineas[i]);
            nuevoRunTexto.appendChild(nuevoT);

            if (insertAfter.getNextSibling() != null) {
                parentOfRun.insertBefore(nuevoRunTexto, insertAfter.getNextSibling());
            } else {
                parentOfRun.appendChild(nuevoRunTexto);
            }
            insertAfter = nuevoRunTexto;
        }
    }

    /**
     * Clase auxiliar para almacenar info de runs
     */
    public static class RunInfo {
        public Element element;
        public String text;

        public RunInfo(Element element, String text) {
            this.element = element;
            this.text = text;
        }
    }
}