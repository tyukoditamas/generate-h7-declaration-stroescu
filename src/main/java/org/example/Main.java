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


public class Main {
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
                        case 4:
                            cargoManifest.setNrDeclaratieVamala(cell.getStringCellValue());
                            break;
                        case 5:
                            cargoManifest.setNrT1(formatter.formatCellValue(cell));
                            break;
                        case 6:
                            cargoManifest.setPrenumeDest(cell.getStringCellValue().substring(cell.getStringCellValue().indexOf(" ") + 1));
                            cargoManifest.setNumeDest(cell.getStringCellValue().substring(0, cell.getStringCellValue().indexOf(" ")));
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
                        case 13:
                            cargoManifest.setPrenumeExp(cell.getStringCellValue().substring(0, cell.getStringCellValue().indexOf(" ")));
                            cargoManifest.setNumeExp(cell.getStringCellValue().substring(cell.getStringCellValue().indexOf(" ") + 1));
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
                        case 35:
                            cargoManifest.setDataDocument(cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate().toString().replaceAll("-", ""));
                            break;
                        case 37:
                            cargoManifest.setDeposit(formatter.formatCellValue(cell));
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
        String lrn = String.format("%ty", Year.now()) + "-33147998-" + cargoManifest.getNrDeclaratieVamala();

        Namespace sNS = Namespace.getNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");

        Element root = new Element("SRD415");
        root.addNamespaceDeclaration(sNS);
        document.addContent(root);

        Element messages = new Element("MESSAGE");
        Element declaration = new Element("Declaration");
        Element goodsShipment = new Element("GoodsShipment");
        root.addContent(messages);
        root.addContent(declaration);
        root.addContent(goodsShipment);

        addMessages(messages);
        addDeclaration(declaration, lrn, cargoManifest);
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

    private static void addGoodsShipment(Element goodsShipment, CargoManifest cargoManifest) {
        Element gsExporter = new Element("GSExporter");
        goodsShipment.addContent(gsExporter);
        addGSExporter(gsExporter, cargoManifest);

        Element gSPreviousDocumentsX337 = new Element("GSPreviousDocuments");
        goodsShipment.addContent(gSPreviousDocumentsX337);
        addGSPreviousDocumentsX337(gSPreviousDocumentsX337, cargoManifest);

        Element gSSupportingDocumentsN271 = new Element("GSSupportingDocuments");
        goodsShipment.addContent(gSSupportingDocumentsN271);
        addGSSupportingDocumentsN271(gSSupportingDocumentsN271, cargoManifest);

        Element gSSupportingDocumentsN730 = new Element("GSSupportingDocuments");
        goodsShipment.addContent(gSSupportingDocumentsN730);
        addGSSupportingDocumentsN730(gSSupportingDocumentsN730, cargoManifest);

//        Element gSSupportingDocumentsN821 = new Element("GSSupportingDocuments");
//        goodsShipment.addContent(gSSupportingDocumentsN821);
//        addGSSupportingDocumentsN821(gSSupportingDocumentsN821, cargoManifest);

        Element locationOfGoods = new Element("LocationOfGoods");
        goodsShipment.addContent(locationOfGoods);
        addLocationOfGoods(locationOfGoods, cargoManifest);

        Element goodsItem = new Element("GoodsItem");
        goodsShipment.addContent(goodsItem);
        addGoodsItem(goodsItem, cargoManifest);
    }

    private static void addGSPreviousDocumentsX337(Element gSPreviousDocumentsX337, CargoManifest cargoManifest) {
        gSPreviousDocumentsX337.addContent(new Element("PreviousDocumentCategory").setText("X"));
        gSPreviousDocumentsX337.addContent(new Element("PreviousDocumentType").setText("337"));
        gSPreviousDocumentsX337.addContent(new Element("PreviousDocumentReferenceNumber").setText("NU ESTE CAZUL"));
        gSPreviousDocumentsX337.addContent(new Element("PreviousDocumentDate").setText(cargoManifest.getDataDocument()));
    }

    private static void addGoodsItem(Element goodsItem, CargoManifest cargoManifest) {
        goodsItem.addContent(new Element("GoodsItemNumber").setText("1"));

        Element goodsInformation = new Element("GoodsInformation");
        goodsItem.addContent(goodsInformation);
        addGoodsInformation(goodsInformation, cargoManifest);

        Element intrinsicValue = new Element("IntrinsicValue");
        goodsItem.addContent(intrinsicValue);
        addIntrinsicValue(intrinsicValue, cargoManifest);

        Element procedures = new Element("Procedures");
        goodsItem.addContent(procedures);
        addProcedures(procedures, cargoManifest);

        goodsItem.addContent(new Element("SIGrossMass").setText(cargoManifest.getGreutateKg()));
    }

    private static void addProcedures(Element procedures, CargoManifest cargoManifest) {
        procedures.addContent(new Element("AdditionalProcedure").setText(cargoManifest.getRegimSuplimentar()));
    }

    private static void addIntrinsicValue(Element intrinsicValue, CargoManifest cargoManifest) {
        intrinsicValue.addContent(new Element("ValueAmount").setText(cargoManifest.getValEur()));
        intrinsicValue.addContent(new Element("ValueCurrency").setText(cargoManifest.getCurrency()));
    }

    private static void addGoodsInformation(Element goodsInformation, CargoManifest cargoManifest) {
        Element descriptionOfGoods = new Element("DescriptionOfGoods");
        goodsInformation.addContent(descriptionOfGoods);
        addDescriptionOfGoods(descriptionOfGoods, cargoManifest);

        goodsInformation.addContent(new Element("PackageNumber").setText(cargoManifest.getNrCol()));

        Element commodityCode = new Element("CommodityCode");
        goodsInformation.addContent(commodityCode);
        addCommodityCode(commodityCode, cargoManifest);
    }

    private static void addCommodityCode(Element commodityCode, CargoManifest cargoManifest) {
        commodityCode.addContent(new Element("CommodityCodeHarmonizedSystemSubHeadingCode").setText(cargoManifest.getHarmonizedCode()));
    }

    private static void addDescriptionOfGoods(Element descriptionOfGoods, CargoManifest cargoManifest) {
        descriptionOfGoods.addContent(new Element("DescriptionOfGood").setText(cargoManifest.getObservatii()));
    }

    private static void addLocationOfGoods(Element locationOfGoods, CargoManifest cargoManifest) {
        locationOfGoods.addContent(new Element("LocationOfGoodsTypeOfLocation").setText("B"));
        locationOfGoods.addContent(new Element("LocationOfGoodsQualifierOfIdentification").setText("V"));
        locationOfGoods.addContent(new Element("LocationOfGoodsCustomsOffice").setText("ROTM0200"));
    }

    private static void addGSSupportingDocumentsN821(Element gSSupportingDocumentsN821, CargoManifest cargoManifest) {
        gSSupportingDocumentsN821.addContent(new Element("SupportingDocumentType").setText("N821"));
        gSSupportingDocumentsN821.addContent(new Element("SupportingDocumentReferenceNumber").setText(cargoManifest.getNrT1()));
        gSSupportingDocumentsN821.addContent(new Element("SupportingDocumentDate").setText(cargoManifest.getDataDocument()));
    }

    private static void addGSSupportingDocumentsN730(Element gSSupportingDocumentsN730, CargoManifest cargoManifest) {
        gSSupportingDocumentsN730.addContent(new Element("SupportingDocumentType").setText("N730"));
        gSSupportingDocumentsN730.addContent(new Element("SupportingDocumentReferenceNumber").setText(cargoManifest.getNrContainer()));
        gSSupportingDocumentsN730.addContent(new Element("SupportingDocumentDate").setText(cargoManifest.getDataDocument()));
    }

    private static void addGSSupportingDocumentsN271(Element gSSupportingDocumentsN271, CargoManifest cargoManifest) {
        gSSupportingDocumentsN271.addContent(new Element("SupportingDocumentType").setText("C665"));
        gSSupportingDocumentsN271.addContent(new Element("SupportingDocumentReferenceNumber").setText("CN23"));
        gSSupportingDocumentsN271.addContent(new Element("SupportingDocumentDate").setText(cargoManifest.getDataDocument()));
    }

    private static void addGSExporter(Element gsExporter, CargoManifest cargoManifest) {
        Element exporter = new Element("Exporter");
        gsExporter.addContent(exporter);
        addExporter(exporter, cargoManifest);

        Element additionalFiscalReference = new Element("AdditionalFiscalReference");
        gsExporter.addContent(additionalFiscalReference);
        addAdditionalFiscalReference(additionalFiscalReference);
    }

    private static void addAdditionalFiscalReference(Element additionalFiscalReference) {
        additionalFiscalReference.addContent(new Element("AdditionalFiscalReferenceRole").setText("FR1"));
    }

    private static void addExporter(Element exporter, CargoManifest cargoManifest) {
        exporter.addContent(new Element("ExporterName").setText(cargoManifest.getNumeExp() + " " + cargoManifest.getPrenumeExp()));
        Element exporterAddress = new Element("ExporterAddress");
        exporter.addContent(exporterAddress);
        addExporterAddress(exporterAddress, cargoManifest);
    }

    private static void addExporterAddress(Element exporterAddress, CargoManifest cargoManifest) {
        exporterAddress.addContent(new Element("ExporterAddressCity").setText(cargoManifest.getOrasDest()));
        exporterAddress.addContent(new Element("ExporterAddressCountry").setText(cargoManifest.getTara()));
        exporterAddress.addContent(new Element("ExporterAddressStreetAndNumber").setText(cargoManifest.getAdresaExp()));
        exporterAddress.addContent(new Element("ExporterAddressPostCode").setText(cargoManifest.getCodPostalExpeditor()));
    }

    private static void addDeclaration(Element declaration, String lrn, CargoManifest cargoManifest) {
        declaration.addContent(new Element("LRN").setText(lrn));
        declaration.addContent(new Element("AdditionalDeclarationType").setText("A"));

        Element customsOffices = new Element("CustomsOffices");
        declaration.addContent(customsOffices);
        addCustomsOffices(customsOffices);

        Element importer = new Element("Importer");
        declaration.addContent(importer);
        addImporter(importer, cargoManifest);

        Element declarant = new Element("Declarant");
        declaration.addContent(declarant);
        addDeclarant(declarant);

        Element representative = new Element("Representative");
        declaration.addContent(representative);
        addRepresentative(representative);

        if (cargoManifest.getDeposit() != null && !cargoManifest.getDeposit().isEmpty()){
            Element depositAdvancePayment = new Element("DepositAdvancePayment").setText(cargoManifest.getDeposit());
            declaration.addContent(depositAdvancePayment);
            // addDepositAdvancePayment(depositAdvancePayment);
        }
    }



    private static void addRepresentative(Element representative) {
        representative.addContent(new Element("RepresentativeIdentificationNumber").setText("RO33147998"));
        representative.addContent(new Element("RepresentativeStatus").setText("3"));
    }

    private static void addDeclarant(Element declarant) {
        declarant.addContent(new Element("DeclarantIdentificationNumber").setText("RO33147998"));
    }

    private static void addImporter(Element importer, CargoManifest cargoManifest) {
        importer.addContent(new Element("ImporterName").setText(cargoManifest.getNumeDest()
                + " " + cargoManifest.getPrenumeDest()));

        Element importerAddress = new Element("ImporterAddress");
        importer.addContent(importerAddress);
        addImporterAddress(importerAddress, cargoManifest);

    }

    private static void addImporterAddress(Element importerAddress, CargoManifest cargoManifest) {
        importerAddress.addContent(new Element("ImporterAddressCity").setText(cargoManifest.getLocalitate()));
        importerAddress.addContent(new Element("ImporterAddressCountry").setText("RO"));
        importerAddress.addContent(new Element("ImporterAddressStreetAndNumber").setText(cargoManifest.getAdresaDest()));
        importerAddress.addContent(new Element("ImporterAddressPostCode").setText(cargoManifest.getCodPostalImportator()));

    }

    private static void addCustomsOffices(Element customsOffices) {
        customsOffices.addContent(new Element("CustomsOffice").setText("ROTM0200"));
    }

    private static void addMessages(Element messages) {
        messages.addContent(new Element("MessageSender").setText("RO33147998"));
        messages.addContent(new Element("MessageRecipient").setText("MESSRD"));
        messages.addContent(new Element("PreparationDateAndTime").setText(LocalDateTime.now().truncatedTo(ChronoUnit.SECONDS).toString()));
        messages.addContent(new Element("MessageIdentification").setText("RO33147998-000000001"));
        messages.addContent(new Element("MessageType").setText("SRD415"));
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
