package org.example;

import lombok.*;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CargoManifest {
    private String nrDeclaratieVamala;
    private String codColet;
    private String tsdNr;
    private String previousDocumentReference;
    private String numeExp;
    private String prenumeExp;
    private String adresaExp;
    private String tara;
    private String taraImportator;
    private String orasDest;
    private String numeDest;
    private String prenumeDest;
    private String judet;
    private String telefon;
    private String localitate;
    private String adresaDest;
    private String nrCol;
    private String greutateKg;
    private String valEur;
    private String observatii;
    private String codPostalImportator;
    private String codPostalExpeditor;
    private String nrContainer;
    private String nrT1;
    private String dataDocument;
    private String regimSuplimentar;
    private String harmonizedCode;
    private String currency;
    private String deposit;
    private String typeOfLocation;
    private String qualifierOfIdentification;
    private String customsOffice;
    private String supportingDocumentType;
    private String supportingDocumentReference;
    private String marfuriMultiple;
    private String tipTranzit;
}
