package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.*;

public class LabelService {
    private static final String TAG_NAME_LABELS = "labels";
    private static final String TAG_NAME_FULL_NAME = "fullName";
    private static final String SHEET_NAME_FIELD_VALUES = "Filed Values";
    private static final String CELL_VALUE_CUSTOM_LABEL_NAME = "Custom Label Name";
    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(new InputStreamReader(System.in));
        System.out.println("Please enter the full path and name of the XML file with labels from Org:");
        String xmlFileNameWithLabelsFromOrg = scanner.nextLine();
        System.out.println("Please enter the full path and name of the Excel file with translation labels:");
        String excelFileNameWithTranslationLabels = scanner.nextLine();

        List<String> allLabelsFromQaSubstrings = getAllLabelsFromOrg(xmlFileNameWithLabelsFromOrg);
        excelFileProcessing(excelFileNameWithTranslationLabels, allLabelsFromQaSubstrings);
    }

    private static void excelFileProcessing(String excelFileNameWithTranslationLabels, List<String> allLabelsFromQaSubstrings) throws IOException {
        FileInputStream file = new FileInputStream(excelFileNameWithTranslationLabels);
        Workbook workbook = new XSSFWorkbook(file);
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            System.out.println("Sheet name is: " + sheetName);

            if (SHEET_NAME_FIELD_VALUES.equals(sheetName)) {
                System.out.println("Not processed yet.");
            } else {
                int customLabelNameColumnIndex = getCustomLabelNameColumnIndex(sheet);
                List<String> customLabelsFromTranslationExcelFile = new ArrayList<>();
                for (Row row : sheet) {
                    int rowNum = row.getRowNum();
                    for (Cell cell : row) {
                        if (CellType.STRING.equals(cell.getCellType()) && cell.getColumnIndex() == customLabelNameColumnIndex && rowNum > 0) {
                            customLabelsFromTranslationExcelFile.add(cell.getStringCellValue());
                        }
                    }
                }

                if (customLabelsFromTranslationExcelFile.isEmpty()) {
                    System.out.println("___IS_EMPTY___");
                } else {
                    List<String> existingCommonLabelsFromQA = getExistingLabels(customLabelsFromTranslationExcelFile, allLabelsFromQaSubstrings);
                    checkLabels(existingCommonLabelsFromQA, customLabelsFromTranslationExcelFile);
                }
            }
            System.out.println("========================================");
        }
    }

    private static int getCustomLabelNameColumnIndex(Sheet sheet) {
        int customLabelNameColumnIndex = 0;
        Row firstRow = sheet.getRow(0);
        for (int j = 0; j < firstRow.getLastCellNum(); j++) {
            Cell cell = firstRow.getCell(j);
            if (CELL_VALUE_CUSTOM_LABEL_NAME.equalsIgnoreCase(cell.getStringCellValue())) {
                customLabelNameColumnIndex = cell.getColumnIndex();
            }
        }
        return customLabelNameColumnIndex;
    }

    private static List<String> getAllLabelsFromOrg(String xmlFileNameWithLabelsFromOrg) throws ParserConfigurationException, SAXException, IOException {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document document = builder.parse(new File(xmlFileNameWithLabelsFromOrg));
        document.getDocumentElement().normalize();

        NodeList nList = document.getElementsByTagName(TAG_NAME_LABELS);
        List<String> fullNameOfLabels = new ArrayList<>();
        for (int temp = 0; temp < nList.getLength(); temp++) {
            Node node = nList.item(temp);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element eElement = (Element) node;
                fullNameOfLabels.add(eElement.getElementsByTagName(TAG_NAME_FULL_NAME).item(0).getTextContent());
            }
        }
        return fullNameOfLabels;
    }

    private static void checkLabels(List<String> existingSpecificLabelsFromQA, List<String> specificLabelsFromTranslationTableSubstrings) {
        List<String> specificLabelsToCreate = new ArrayList<>();
        List<String> specificLabelsToWriteToTranslationTable = new ArrayList<>();
        if (existingSpecificLabelsFromQA.size() == specificLabelsFromTranslationTableSubstrings.size()) {
            System.out.println("___GOOD___");
        } else if (existingSpecificLabelsFromQA.size() <= specificLabelsFromTranslationTableSubstrings.size()) {
            for (String specificLabelFromTranslationTable : specificLabelsFromTranslationTableSubstrings) {
                if (!existingSpecificLabelsFromQA.contains(specificLabelFromTranslationTable)) {
                    specificLabelsToCreate.add(specificLabelFromTranslationTable);
                }
            }
            System.out.println("Labels to create is: " + specificLabelsToCreate);
        } else {
            for (String existingSpecificLabelFromQA : existingSpecificLabelsFromQA) {
                if (!specificLabelsFromTranslationTableSubstrings.contains(existingSpecificLabelFromQA)) {
                    specificLabelsToWriteToTranslationTable.add(existingSpecificLabelFromQA);
                }
            }
            System.out.println("Labels to write to Translation table is: " + specificLabelsToWriteToTranslationTable);
        }
    }

    private static List<String> getExistingLabels(List<String> labelsFromTranslationTableSubstrings, List<String> allLabelsFromQaSubstrings) {
        List<String> existingLabels = new ArrayList<>();
        for (String labelFromQA : allLabelsFromQaSubstrings) {
            for (String labelFromTranslationTable : labelsFromTranslationTableSubstrings) {
                if (Objects.equals(labelFromQA, labelFromTranslationTable)) {
                    existingLabels.add(labelFromQA);
                }
            }
        }
        return existingLabels;
    }
}
