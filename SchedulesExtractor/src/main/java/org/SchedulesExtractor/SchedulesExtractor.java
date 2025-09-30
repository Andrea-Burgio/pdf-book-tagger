package org.SchedulesExtractor;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.SchedulesExtractor.LOCSubject.getLOCSubjectFromChar;

public class SchedulesExtractor {

    /**
     * Helper class that manages multiple BufferedWriter instances for different LOC subjects.
     * Implements AutoCloseable to work with try-with-resources statements.
     * Creates one output file per LOC subject in the Resources directory.
     */
    private static class BufferedWriterManager implements AutoCloseable {
        private final Map<LOCSubject, BufferedWriter> writers = new HashMap<>();

        /**
         * Constructs a BufferedWriterManager and initializes BufferedWriter instances for all LOC subjects.
         * The Resource directory was previously created by PowerShell.
         * Creates one txt output file per LOC subject (A.txt, B.txt, etc.).
         *
         * @throws IOException if any of the output files cannot be created or opened
         */
        public BufferedWriterManager() throws IOException {
            for (LOCSubject locSubject : LOCSubject.values()) {
                String filePath = "Resources/" + locSubject.name() + ".txt";
                writers.put(locSubject, new BufferedWriter(new FileWriter(filePath), 1024 * 1024)); // 1MB buffer
            }
        }

        /**
         * Returns the map of LOC subjects to their corresponding BufferedWriter instances.
         *
         * @return a map where keys are LOCSubject enums and values are BufferedWriter instances
         */
        public Map<LOCSubject, BufferedWriter> getWriters() {
            return writers;
        }

        @Override
        public void close() throws IOException {
            for (BufferedWriter writer : writers.values()) {
                writer.close();
            }
        }
    }


    /**
     * Processes all the subfields of a datafield element with tag "153" from the XML stream.
     * Extracts the classification code from subfields "z" and "a", applies optional ranges from "c",
     * and appends hierarchy ("h") and caption ("j") information. The formatted record is then written
     * to the appropriate output file based on the first character of the classification code.
     *
     * Special handling rules:
     * <ul>
     *   <li>If "z" is present, it sets the classification prefix. "a" is appended to "z".</li>
     *   <li>If "a" begins with ".", it continues from "z". If no "z" exists, "a" alone is used.</li>
     *   <li>If "c" defines a collapsed range (e.g., ".A-Z"), it is merged into "a".</li>
     *   <li>
     *       If "c" defines a numeric range (e.g., KBM2474–KBM2478), the record is skipped entirely,
     *       since the following records will contain the individual codes covered by the range.
     *   </li>
     *   <li>All "h" subfields are concatenated with "/", then "j" subfields are appended after another "/".</li>
     * </ul>
     * Some cases that can occur:
     * <pre>{@code
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="z">P-PZ20</subfield>
     *      <subfield code="a">176.5.H57</subfield>
     *      <subfield code="h">Table for literature (194 nos.)</subfield>
     *      ...
     *
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="z">BX7</subfield>
     *      <subfield code="a">.x8</subfield>
     *      ...
     *
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="z">G6</subfield>
     *      <subfield code="a">70.A4</subfield>
     *      <subfield code="c">70.Z</subfield>
     *      <subfield code="h">Table of geographical subdivisions (96 numbers)</subfield>
     *      ...
     *
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBR15.A</subfield>
     *      <subfield code="c">KBR15.Z</subfield>
     *      ...
     *  <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBR15.5.A</subfield>
     *      <subfield code="c">KBR15.5.Z</subfield>
     *    ...
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBM2474</subfield>
     *      <subfield code="c">KBM2478</subfield>
     *      <subfield code="h">Jewish law. Halakhah. הלכה</subfield>
     *       ...
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBM2474</subfield>
     *      <subfield code="h">Jewish law. Halakhah. הלכה</subfield>
     *      ...
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBM2476</subfield>
     *      <subfield code="h">Jewish law. Halakhah. הלכה</subfield>
     *      ...
     * <datafield tag="153" ind1=" " ind2=" ">
     *      <subfield code="a">KBM2478</subfield>
     *      <subfield code="h">Jewish law. Halakhah. הלכה</subfield>
     *
     * </pre>
     *}
     * @param reader the XMLStreamReader positioned at the start of a datafield element
     * @param lccToWriterMap map of LOC subjects to their corresponding BufferedWriter instances
     * @throws XMLStreamException if an error occurs while reading the XML stream
     * @throws IOException if an error occurs while writing to the output file
     */
    private static void processSubfields(XMLStreamReader reader, Map<LOCSubject, BufferedWriter> lccToWriterMap)
            throws XMLStreamException, IOException {

        StringBuilder buffer = new StringBuilder(512);
        BufferedWriter bufferedWriter = null;

        boolean zFound = false;
        String lastACode = null;
        LOCSubject locSubject = null;

        while (reader.hasNext()) {
            int event = reader.next();

            if (event == XMLStreamConstants.START_ELEMENT && reader.getLocalName().equals("subfield")) {

                String code = reader.getAttributeValue(null, "code");
                String tagContent = reader.getElementText();

                switch (code) {

                    /*
                     * Classification code determines output file.
                     * Process
                     * <subfield code="a">...</subfield>
                     * or
                     * <subfield code="z">...</subfield>
                     * <subfield code="a">...</subfield>
                     */
                    case "z": { //next subfield is going to be 'a'
                        zFound = true;
                        char prefix = tagContent.charAt(Character.isLetter(tagContent.charAt(0)) ? 0 : 1);
                        locSubject = getLOCSubjectFromChar(prefix);  // also sets locSubject for 'a'

                        bufferedWriter = lccToWriterMap.get(locSubject); //'z' determines the output file
                        buffer.append(tagContent);

                        break;
                    }

                    case "a": {
                        lastACode = tagContent; // Used when 'c' tag is present.
                        // If was found, append 'a' content to the 'z' content
                        if (zFound) {   // Previous subfield was 'z'. locSubject is already determined.
                            String SubjectDescription = locSubject.getDescription();
                            if (tagContent.charAt(0) == '.') {
                                buffer.append(tagContent).append(SubjectDescription);
                            }
                            else {
                                buffer.append(".").append(tagContent).append(" - ").append(SubjectDescription);
                            }
                            zFound = false;
                        }
                        // If 'z' was not found, use 'a' content as prefix.
                        else {
                            char prefix = tagContent.charAt(Character.isLetter(tagContent.charAt(0)) ? 0 : 1);
                            locSubject = getLOCSubjectFromChar(prefix);
                            String SubjectDescription = locSubject.getDescription();

                            bufferedWriter = lccToWriterMap.get(locSubject); // 'a' determines the output file
                            buffer.append(tagContent).append(" - ").append(SubjectDescription);
                        }
                        break;
                    }

                    case "c": {
                        buffer.setLength(0); // Delete stored 'a' content first
                        String subjectDescription = locSubject.getDescription();

                        if (lastACode != null && !tagContent.endsWith("Z")) { // if it's a pure range (KBM2474 -> KBM2478)
                            Matcher matcher = Pattern.compile("\\d").matcher(tagContent);
                            if (matcher.find()) {
                                buffer.append(lastACode)
                                        .append("-")
                                        .append(tagContent.substring(matcher.start()))
                                        .append(" - ")
                                        .append(subjectDescription);
                            }
                        }
                        else { // merge "a" and "c" -> e.g. QK584.6.A-Z
                            // take lastACode (e.g. QK584.6.A) and merge with trailing "Z"
                            buffer.append(lastACode)
                                  .append("-")
                                  .append(tagContent.substring(tagContent.lastIndexOf('.') + 1))
                                  .append(" - ")
                                  .append(subjectDescription);
                        }
                        break;
                    }

                    case "h":   // Heading information.
                        buffer.append("/").append(tagContent);
                        break;

                    case "j":   // Caption information
                        buffer.append("/").append(tagContent);
                        break;
                }
            }
            else if (event == XMLStreamConstants.END_ELEMENT && reader.getLocalName().equals("datafield")) {
                // Append content only if something has to be written. This is not always true for the 'c' case
                if (bufferedWriter != null) {
                    bufferedWriter.append(buffer).append('\n');
                    buffer.setLength(0);
                }
                break;  // Exit loop. Finished processing datafield with tag 153.
            }
        }
    }

    /**
     * <p>Extracts Library of Congress classification data from <i>Classification.d170424.xml</i> (data from
     * <a href="https://www.loc.gov/cds/products/MDSConnect-classification.html">MDSConnect-classification</a> (2016))
     * and distributes records into separate output files (A.xml, B.xml, C.xml, ...) based on the first letter of the
     * classification code. Output files are created in the Resources directory.</p>
     * <p>Only records with datafield tag "153" are processed. These records contain:
     * <ul>
     *   <li>Subfield "z":<br>Classification code</li>
     *   <li>
     *       Subfield "a":<br>Main classification code (e.g., "KBR39.2"). If no "z" is present, "a" determines the
     *       classification and the output file. May start with "." to continue from "z".
     *   </li>
     *   <li>
     *       Subfield "c":<br>End of a classification range. If it represents a collapsed range (e.g., "A-Z"), it is merged
     *       with "a". If it defines a numeric span (e.g., KBM2474–KBM2478), the record is skipped and nothing is written.
     *   </li>
     *   <li>Subfield "h":<br>Heading information</li>
     *   <li>Subfield "j":<br>Caption information</li>
     * </ul>
     *
     * <p>Example of a record from the <i>Classification.d170424.xml</i><br>
     * <pre>{@code
     * <record>
     *   <leader>00364nw  a2200121n  4500</leader>
     *   <controlfield tag="001">CF 00433935</controlfield>
     *   <controlfield tag="003">DLC</controlfield>
     *   <controlfield tag="005">20030306082331.0</controlfield>
     *   <controlfield tag="008">000119acaaaaaa</controlfield>
     *   <datafield tag="010" ind1=" " ind2=" ">
     *     <subfield code="a">CF 00433935</subfield>
     *   </datafield>
     *   <datafield tag="040" ind1=" " ind2=" ">
     *     <subfield code="a">DLC</subfield>
     *     <subfield code="c">DLC</subfield>
     *   </datafield>
     *   <datafield tag="084" ind1="0" ind2=" ">
     *     <subfield code="a">lcc</subfield>
     *   </datafield>
     *   <datafield tag="153" ind1=" " ind2=" ">
     *     <subfield code="a">KBR39.2</subfield>
     *     <subfield code="c">KBR39.22</subfield>
     *     <subfield code="h">History of canon law</subfield>
     *     <subfield code="h">Official acts of the Holy See</subfield>
     *     <subfield code="h">Decrees and decisions of the Curia Romana</subfield>
     *     <subfield code="j">Signatura Gratiae. Signatura of Grace</subfield>
     *   </datafield>
     * </record>
     * }</pre>
     * </p>
     */
    public static void extractSchedules(String XMlClassificationPath) {
        XMLInputFactory factory = XMLInputFactory.newInstance();
        factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);

        try (FileInputStream fileInputStream = new FileInputStream(XMlClassificationPath);
             BufferedWriterManager bufferedWriterManager = new BufferedWriterManager()
        ) {
            Map<LOCSubject, BufferedWriter> lccToWriter = bufferedWriterManager.getWriters();

            XMLStreamReader reader = factory.createXMLStreamReader(fileInputStream);

            while (reader.hasNext()) {
                int event = reader.next();

                // Process one record
                if (event == XMLStreamConstants.START_ELEMENT &&
                    reader.getLocalName().equals("datafield") &&
                    reader.getAttributeValue(null, "tag").equals("153")
                ) {
                    processSubfields(reader, lccToWriter);
                }
            }

            reader.close(); // Not in try-with-resources
        } catch (Exception e) {
            System.err.println(e.getMessage());
            System.exit(1);
        }
    }
}