package org.example;


import org.apache.poi.ss.usermodel.*;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.Namespace;
import org.jdom2.output.XMLOutputter;

import javax.swing.*;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.Year;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;


public class Main {
    private static final String COMPANY_EORI = "RO29631670";
    private static final String LOCATION_AUTHORIZATION = "ROTSTRODRVTM00030-2016";
    private static final Namespace NS_COMMON =
            Namespace.getNamespace("urn:uccv:xsd:imp:1.0.0_2.00:CCommon");
    private static final Namespace NS_CC415 =
            Namespace.getNamespace("ns2", "urn:uccv:xsd:imp:1.0.0_2.00:CC415");
    private static final Namespace NS_DECLARATION =
            Namespace.getNamespace("ns3", "urn:uccv:xsd:imp:1.0.0_2.00:Declaration");
    private static JTextArea logArea;
    public static void main(String[] args) {

        try{
            final JFrame frame = new JFrame("Generare XMLs");

            final JFileChooser fc = new JFileChooser();
            fc.setMultiSelectionEnabled(true);
            fc.setCurrentDirectory(new File("C:\\tmp"));

            JButton btn2 = new JButton("Choose excel file:");
            btn2.addActionListener(e -> {
                int retVal = fc.showOpenDialog(frame);
                if (retVal == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fc.getSelectedFile();

                    try {
                        processFile(selectedFile, frame);
                    } catch (SQLException | ParserConfigurationException | TransformerException | IOException |
                             URISyntaxException ex) {
                        logMessage(ex.getCause().toString());
                    } catch (Exception ex) {
                        logMessage(ex.toString());
                        logMessage(Arrays.toString(ex.getStackTrace()));
                        throw new RuntimeException(ex);
                    }

                }

            });
            logArea = new JTextArea();
            logArea.setEditable(false);
            JScrollPane scrollPane = new JScrollPane(logArea);

            Container pane = frame.getContentPane();
            pane.setLayout(new GridLayout(3, 1, 10, 10));
            pane.add(btn2);
            pane.add(scrollPane, BorderLayout.CENTER);

            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(300, 200);
            frame.setVisible(true);
        }catch (Exception e) {
            logMessage(e.toString());
        }

    }

    private static void logMessage(String message) {
        SwingUtilities.invokeLater(() -> logArea.append(message + "\n"));
    }

    private static void processFile(File selectedFile, JFrame frame)
            throws Exception {
        try {
            String dbPath = selectedFile.getAbsolutePath();
            logMessage("File path: " +dbPath);
            File file = new File(dbPath);
            logMessage("File exists. Continuing ..");

            FileInputStream fileInputStream = new FileInputStream(file);
            Workbook workbook = WorkbookFactory.create(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);
            logMessage("Sheet found for reading: " + sheet.getSheetName());
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();
            rowIterator.next();
            List<CargoManifest> cargoManifestList = new ArrayList<>();
            logMessage("Iterating rows in excel");
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                CargoManifest cargoManifest = new CargoManifest();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    DataFormatter formatter = new DataFormatter();
                    switch (cell.getColumnIndex()) {
                        case 0:
                            cargoManifest.setCodColet(cell.getStringCellValue());
                            break;
                        case 1:
                            cargoManifest.setNrContainer(formatter.formatCellValue(cell));
                            break;
                        case 2:
                            cargoManifest.setPreviousDocumentReference(formatter.formatCellValue(cell));
                            break;
                        case 4:
                            cargoManifest.setNrDeclaratieVamala(formatter.formatCellValue(cell));
                            break;
                        case 5:
                            cargoManifest.setNrT1(formatter.formatCellValue(cell));
                            break;
                        case 6:
                            String importerName = formatter.formatCellValue(cell).trim();
                            int importerSplitIndex = importerName.indexOf(" ");
                            if (importerSplitIndex > 0) {
                                cargoManifest.setNumeDest(importerName.substring(0, importerSplitIndex));
                                cargoManifest.setPrenumeDest(importerName.substring(importerSplitIndex + 1));
                            } else {
                                cargoManifest.setNumeDest(importerName);
                                cargoManifest.setPrenumeDest("");
                            }
                            break;
                        case 7:
                            String addressImp = formatter.formatCellValue(cell);
                            if (addressImp.length() > 34)
                                addressImp = addressImp.substring(0, 34);
                            cargoManifest.setAdresaDest(addressImp);
                            break;
                        case 8:
                            cargoManifest.setLocalitate(formatter.formatCellValue(cell));
                            break;
                        case 9:
                            cargoManifest.setCodPostalImportator(formatter.formatCellValue(cell));
                            break;
                        case 10:
                            cargoManifest.setTaraImportator(formatter.formatCellValue(cell));
                            break;
                        case 13:
                            String exporterName = formatter.formatCellValue(cell).trim();
                            int exporterSplitIndex = exporterName.indexOf(" ");
                            if (exporterSplitIndex > 0) {
                                cargoManifest.setPrenumeExp(exporterName.substring(0, exporterSplitIndex));
                                cargoManifest.setNumeExp(exporterName.substring(exporterSplitIndex + 1));
                            } else {
                                cargoManifest.setPrenumeExp(exporterName);
                                cargoManifest.setNumeExp("");
                            }
                            break;
                        case 14:
                            String address = formatter.formatCellValue(cell);
                            if (address.length() > 34)
                                address = address.substring(0, 34);
                            cargoManifest.setAdresaExp(address);
                            break;
                        case 15:
                            cargoManifest.setOrasDest(formatter.formatCellValue(cell));
                            break;
                        case 16:
                            //cargoManifest.setCodPostalExpeditor(cell.getCellType().equals(CellType.NUMERIC) || cell.getCachedFormulaResultType().equals(CellType.NUMERIC) ? String.valueOf(Double.valueOf(cell.getNumericCellValue())) : cell.getStringCellValue());
                            cargoManifest.setCodPostalExpeditor(formatter.formatCellValue(cell));
                            break;
                        case 17:
                            cargoManifest.setTara(formatter.formatCellValue(cell));
                            break;
                        case 18:
                            cargoManifest.setTypeOfLocation(formatter.formatCellValue(cell));
                            break;
                        case 19:
                            cargoManifest.setQualifierOfIdentification(formatter.formatCellValue(cell));
                            break;
                        case 20:
                            cargoManifest.setCustomsOffice(formatter.formatCellValue(cell));
                            break;
                        case 25:
                            cargoManifest.setHarmonizedCode(formatter.formatCellValue(cell));
                            break;
                        case 26:
                            cargoManifest.setGreutateKg(formatter.formatCellValue(cell).replaceAll(",", "."));
                            break;
                        case 27:
                            cargoManifest.setValEur(formatter.formatCellValue(cell));
                            break;
                        case 28:
                            cargoManifest.setCurrency(formatter.formatCellValue(cell));
                            break;
                        case 29:
                            cargoManifest.setRegimSuplimentar(cell.getCellType().equals(CellType.NUMERIC) ? String.valueOf(Double.valueOf(cell.getNumericCellValue())) : cell.getStringCellValue());
                            break;
                        case 31:
                            String descriere = formatter.formatCellValue(cell);
                            if (descriere.length() > 209)
                                descriere = descriere.substring(0, 209);
                            cargoManifest.setObservatii(descriere);
                            break;
                        case 32:
                            cargoManifest.setNrCol(formatter.formatCellValue(cell));
                            break;
                        case 33:
                            cargoManifest.setSupportingDocumentType(formatter.formatCellValue(cell));
                            break;
                        case 34:
                            cargoManifest.setSupportingDocumentReference(formatter.formatCellValue(cell));
                            break;
                        case 35:
                            try {
                                cargoManifest.setDataDocument(cell.getDateCellValue().toInstant()
                                        .atZone(ZoneId.systemDefault()).toLocalDate().toString().replaceAll("-", ""));
                            } catch (Exception ignored) {
                                cargoManifest.setDataDocument(formatter.formatCellValue(cell));
                            }
                            break;
                        case 37:
                            cargoManifest.setDeposit(formatter.formatCellValue(cell));
                            break;
                        case 48:
                            cargoManifest.setTsdNr(formatter.formatCellValue(cell));
                            break;
                        case 49:
                            cargoManifest.setMarfuriMultiple(formatter.formatCellValue(cell));
                            break;
                        case 50:
                            cargoManifest.setTipTranzit(formatter.formatCellValue(cell));
                            break;
                    }
                }
                cargoManifestList.add(cargoManifest);
            }

            for (CargoManifest cargoManifest : cargoManifestList) {
                generateXmlFile(cargoManifest);
            }
            logMessage("Finished generating.");

        } catch (Exception e) {
            logMessage(e.toString());
        }

    }

    private static void generateXmlFile(CargoManifest cargoManifest) throws IOException {

        Document document = new Document();
        String lrn = String.format("%ty", Year.now()) + COMPANY_EORI + safe(cargoManifest.getNrDeclaratieVamala());

        Element root = new Element("CC415", NS_CC415);
        root.addNamespaceDeclaration(NS_COMMON);
        root.addNamespaceDeclaration(NS_DECLARATION);
        document.setRootElement(root);

        Element message = new Element("MESSAGE", NS_CC415);
        root.addContent(message);
        addMessage(message);

        Element importOperation = new Element("ImportOperation", NS_CC415);
        root.addContent(importOperation);
        addImportOperation(importOperation, lrn);

        Element nationalDeclaration = new Element("NationalDeclaration", NS_CC415);
        root.addContent(nationalDeclaration);
        addNationalDeclaration(nationalDeclaration);

        Element customsOffice = new Element("CustomsOffice", NS_CC415);
        root.addContent(customsOffice);
        addCustomsOffice(customsOffice, cargoManifest);

        Element importer = new Element("Importer", NS_CC415);
        root.addContent(importer);
        addImporter(importer, cargoManifest);

        Element declarant = new Element("Declarant", NS_CC415);
        root.addContent(declarant);
        addDeclarant(declarant);

        Element representative = new Element("Representative", NS_CC415);
        root.addContent(representative);
        addRepresentative(representative);

        Element goodsShipment = new Element("GoodsShipment", NS_CC415);
        root.addContent(goodsShipment);
        addGoodsShipment(goodsShipment, cargoManifest);

        LocalDate timestamp = LocalDateTime.now().toLocalDate();
        XMLOutputter xmlOutputter = new XMLOutputter();

        String path = System.getProperty("user.home") + "/Desktop/xmls/" + timestamp + "//";
        File filepath = new File(path);

        Files.createDirectories(Paths.get(filepath.toURI()));

        File file = new File(path, lrn + ".xml");

        try (FileOutputStream fileOutputStream =
                     new FileOutputStream(file)) {
            xmlOutputter.output(document, fileOutputStream);
        }
        logMessage("Done creating XML File");
    }

    private static void addMessage(Element message) {
        addTextElement(message, "MessageSender", NS_COMMON, COMPANY_EORI);
        addTextElement(message, "MessageRecipient", NS_COMMON, "IMP.RO");
        addTextElement(message, "PreparationDateAndTime", NS_COMMON,
                LocalDateTime.now().truncatedTo(ChronoUnit.SECONDS).toString());
        addTextElement(message, "MessageIdentification", NS_COMMON, UUID.randomUUID().toString());
        addTextElement(message, "MessageType", NS_COMMON, "CC415");
    }

    private static void addImportOperation(Element importOperation, String lrn) {
        addTextElement(importOperation, "LRN", NS_DECLARATION, lrn);
        addTextElement(importOperation, "declarationType", NS_DECLARATION, "IM");
        addTextElement(importOperation, "additionalDeclarationType", NS_DECLARATION, "D");
    }

    private static void addNationalDeclaration(Element nationalDeclaration) {
        addTextElement(nationalDeclaration, "CustomType", NS_DECLARATION, "standard");
        addTextElement(nationalDeclaration, "DeclarationCategory", NS_DECLARATION, "H7");
        addTextElement(nationalDeclaration, "SafetySecurityFeatures", NS_DECLARATION, "false");
    }

    private static void addCustomsOffice(Element customsOffice, CargoManifest cargoManifest) {
        addTextElement(customsOffice, "referenceNumber", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getCustomsOffice(), "ROTM0200"));
    }

    private static void addImporter(Element importer, CargoManifest cargoManifest) {
        addTextElement(importer, "name", NS_DECLARATION,
                fullName(cargoManifest.getNumeDest(), cargoManifest.getPrenumeDest()));

        Element address = new Element("Address", NS_DECLARATION);
        importer.addContent(address);
        addTextElement(address, "streetAndNumber", NS_DECLARATION, cargoManifest.getAdresaDest());
        addTextElement(address, "city", NS_DECLARATION, cargoManifest.getLocalitate());
        addTextElement(address, "postcode", NS_DECLARATION, cargoManifest.getCodPostalImportator());
        addTextElement(address, "country", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getTaraImportator(), "RO"));
    }

    private static void addDeclarant(Element declarant) {
        addTextElement(declarant, "identificationNumber", NS_DECLARATION, COMPANY_EORI);
    }

    private static void addRepresentative(Element representative) {
        addTextElement(representative, "status", NS_DECLARATION, "3");
    }

    private static void addGoodsShipment(Element goodsShipment, CargoManifest cargoManifest) {
        addTextElement(goodsShipment, "sequenceNumber", NS_DECLARATION, "1");

        Element supportingDocument = new Element("SupportingDocument", NS_DECLARATION);
        goodsShipment.addContent(supportingDocument);
        addTextElement(supportingDocument, "sequenceNumber", NS_DECLARATION, "1");
        addTextElement(supportingDocument, "type", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getSupportingDocumentType(), "C665"));
        addTextElement(supportingDocument, "referenceNumber", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getSupportingDocumentReference(), "DA"));

        Element supportingDocument2 = new Element("SupportingDocument", NS_DECLARATION);
        goodsShipment.addContent(supportingDocument2);
        addTextElement(supportingDocument2, "sequenceNumber", NS_DECLARATION, "2");
        addTextElement(supportingDocument2, "type", NS_DECLARATION,
                defaultIfBlank("1094", "1094"));
        addTextElement(supportingDocument2, "referenceNumber", NS_DECLARATION,
                defaultIfBlank("APROBARE VAMA NR. 690/09.03.2026", "APROBARE VAMA NR. 690/09.03.2026"));

        Element exporter = new Element("Exporter", NS_DECLARATION);
        goodsShipment.addContent(exporter);
        addTextElement(exporter, "name", NS_DECLARATION,
                fullName(cargoManifest.getNumeExp(), cargoManifest.getPrenumeExp()));
        Element exporterAddress = new Element("Address", NS_DECLARATION);
        exporter.addContent(exporterAddress);
        addTextElement(exporterAddress, "streetAndNumber", NS_DECLARATION, cargoManifest.getAdresaExp());
        addTextElement(exporterAddress, "city", NS_DECLARATION, cargoManifest.getOrasDest());
        addTextElement(exporterAddress, "postcode", NS_DECLARATION, cargoManifest.getCodPostalExpeditor());
        addTextElement(exporterAddress, "country", NS_DECLARATION, cargoManifest.getTara());

        Element consignment = new Element("Consignment", NS_DECLARATION);
        goodsShipment.addContent(consignment);

        Element locationOfGoods = new Element("LocationOfGoods", NS_DECLARATION);
        consignment.addContent(locationOfGoods);
        addTextElement(locationOfGoods, "typeOfLocation", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getTypeOfLocation(), "A"));
        addTextElement(locationOfGoods, "qualifierOfIdentification", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getQualifierOfIdentification(), "Y"));
        Element locationCustomsOffice = new Element("CustomsOffice", NS_DECLARATION);
        locationOfGoods.addContent(locationCustomsOffice);
        addCustomsOffice(locationCustomsOffice, cargoManifest);

        Element transportDocument = new Element("TransportDocument", NS_DECLARATION);
        consignment.addContent(transportDocument);
        addTextElement(transportDocument, "sequenceNumber", NS_DECLARATION, "1");
        addTextElement(transportDocument, "referenceNumber", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getNrContainer(), "CMR"));
        addTextElement(transportDocument, "type", NS_DECLARATION, "N730");

        Element goodsShipmentItem = new Element("GoodsShipmentItem", NS_DECLARATION);
        goodsShipment.addContent(goodsShipmentItem);
        addTextElement(goodsShipmentItem, "declarationGoodsItemNumber", NS_DECLARATION, "1");

        Element procedure = new Element("Procedure", NS_DECLARATION);
        goodsShipmentItem.addContent(procedure);
        addTextElement(procedure, "requestedProcedure", NS_DECLARATION, "40");
        addTextElement(procedure, "previousProcedure", NS_DECLARATION, "00");
        addAdditionalProcedure(procedure, 1, "000");
        addAdditionalProcedure(procedure, 2, defaultIfBlank(cargoManifest.getRegimSuplimentar(), "C08"));

        Element previousDocument = new Element("PreviousDocument", NS_DECLARATION);
        goodsShipmentItem.addContent(previousDocument);
        addTextElement(previousDocument, "sequenceNumber", NS_DECLARATION, "1");
        addTextElement(previousDocument, "referenceNumber", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getTsdNr(),
                        defaultIfBlank(cargoManifest.getPreviousDocumentReference(), "1")));
        addTextElement(previousDocument, "type", NS_DECLARATION, cargoManifest.getTipTranzit());

        Element commodity = new Element("Commodity", NS_DECLARATION);
        goodsShipmentItem.addContent(commodity);
        addTextElement(commodity, "descriptionOfGoods", NS_DECLARATION, cargoManifest.getObservatii() +
                (cargoManifest.getMarfuriMultiple().equalsIgnoreCase("DA") ? " - Marfa declarata CF Art. 177 din Reg. UE 952/2013 in urma aprobarii Vamii" : ""));

        Element commodityCode = new Element("CommodityCode", NS_DECLARATION);
        commodity.addContent(commodityCode);
        addTextElement(commodityCode, "harmonizedSystemSubheadingCode", NS_DECLARATION, cargoManifest.getHarmonizedCode());

        Element goodsMeasure = new Element("GoodsMeasure", NS_DECLARATION);
        commodity.addContent(goodsMeasure);
        addTextElement(goodsMeasure, "grossMass", NS_DECLARATION, cargoManifest.getGreutateKg());

        Element packaging = new Element("Packaging", NS_DECLARATION);
        goodsShipmentItem.addContent(packaging);
        addTextElement(packaging, "sequenceNumber", NS_DECLARATION, "1");
        addTextElement(packaging, "numberOfPackages", NS_DECLARATION, cargoManifest.getNrCol());

        Element intrinsicValue = new Element("IntrinsicValue", NS_DECLARATION);
        goodsShipmentItem.addContent(intrinsicValue);
        addTextElement(intrinsicValue, "intrinsicValueCurrency", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getCurrency(), "USD"));
        addTextElement(intrinsicValue, "intrinsicValueAmount", NS_DECLARATION,
                defaultIfBlank(cargoManifest.getValEur(), "0"));
    }

    private static void addAdditionalProcedure(Element procedure, int sequenceNumber, String additionalProcedureCode) {
        Element additionalProcedure = new Element("AdditionalProcedure", NS_DECLARATION);
        procedure.addContent(additionalProcedure);
        addTextElement(additionalProcedure, "sequenceNumber", NS_DECLARATION, String.valueOf(sequenceNumber));
        addTextElement(additionalProcedure, "additionalProcedure", NS_DECLARATION, additionalProcedureCode);
    }

    private static void addTextElement(Element parent, String name, Namespace namespace, String value) {
        parent.addContent(new Element(name, namespace).setText(safe(value)));
    }

    private static String defaultIfBlank(String value, String fallback) {
        return safe(value).isEmpty() ? fallback : safe(value);
    }

    private static String fullName(String firstName, String lastName) {
        String first = safe(firstName);
        String last = safe(lastName);
        if (first.isEmpty()) {
            return last;
        }
        if (last.isEmpty()) {
            return first;
        }
        return first + " " + last;
    }

    private static String safe(String value) {
        return value == null ? "" : value.trim();
    }

    private static CargoManifest mapRow(ResultSet rs) throws SQLException {
        return CargoManifest.builder()
                .nrDeclaratieVamala(rs.getString(1))
                .codColet(rs.getString("COD COLET"))
                .numeExp(rs.getString("NUME EXP"))
                .adresaExp(rs.getString("ADRESA EXP"))
                .prenumeExp(rs.getString("PRENUME"))
                .tara(rs.getString("TARA"))
                .numeDest(rs.getString("NUME DEST"))
                .prenumeDest(rs.getString("PRENUME DEST"))
                .telefon(rs.getString("TELEFON"))
                .judet(rs.getString("JUDET"))
                .localitate(rs.getString("LOCALITATE"))
                .adresaDest(rs.getString("ADRESA"))
                .nrCol(rs.getString("NR COL"))
                .greutateKg(rs.getString("G(KG)"))
                .valEur(rs.getString("VAL EURO"))
                .observatii(rs.getString("OBSERVATII"))
                .codPostalImportator(rs.getString("Cod Postal Importator"))
                .codPostalExpeditor(rs.getString("Cod Postal Exportator"))
                .nrContainer(rs.getString("Nr Container"))
                .nrT1(rs.getString("Nr T1"))
                .dataDocument(rs.getString("Data Documente"))
                .regimSuplimentar(rs.getString("Regim Suplimentar"))
                .build();
    }
}
