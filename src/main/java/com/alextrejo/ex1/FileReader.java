package com.alextrejo.ex1;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.util.*;
import java.util.List;



public class FileReader {
    private static final List<String> ALLOWED_FILE_EXTENSIONS = Arrays.asList("csv", "xlsx", "xls", "txt");
    public static void readFile(){
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.println("Please input the directoy of the file you wish to read without quotes: ");
            String userInput = scanner.nextLine();
            openFile(userInput);
        }
            catch(IllegalStateException | NoSuchElementException e) {
            System.out.println("System.in was closed; exiting");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void openFile (String filePath) throws IOException {
        try {

            File file = new File(filePath);
            if(!file.exists()){
                System.err.println("File specified at: "+filePath + " was not found.");
            }
            System.out.println("File at: "+filePath + " successfully found.");
            String fileExtension = getExtensionByApacheCommonLib(file.getName());
            validateFileExtension(fileExtension);
            if(fileExtension.equals("xlsx")){
                createNewXLSXFileFromSpreadsheet(filePath, file.getName());
            }
            if(fileExtension.equals("xls")){
                createNewXLSFileFromSpreadsheet(filePath, file.getName());
            }
            } catch (RuntimeException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }
    private static void validateFileExtension(String fileExtension){

        if(!ALLOWED_FILE_EXTENSIONS.contains(fileExtension)){
            throw new RuntimeException("File extension of type: "+ fileExtension + " is not supported.");
        }
        System.out.println("File found. Has valid extension: "+fileExtension);
    }

    public static String getExtensionByApacheCommonLib(String filename) {
        return FilenameUtils.getExtension(filename);
    }

    private static void createNewXLSXFileFromSpreadsheet(String filePath, String fileName) throws IOException {
        FileInputStream file = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(file);

        createxls(workbook, fileName);
    }

    private static void createNewXLSFileFromSpreadsheet(String filePath, String filename) throws InvalidFormatException, IOException {
        Workbook wb = WorkbookFactory.create(new File(filePath));

        createxls(wb, filename);

    }
    public static String reverseString(String str){
        StringBuilder sb=new StringBuilder(str);
        sb.reverse();
        return sb.toString();
    }

    private static void createxls(Workbook wb, String filename) throws IOException {
        Sheet sheet = wb.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING: cell.setCellValue(reverseString(cell.getStringCellValue())); break;
                    case NUMERIC :{
                        if(DateUtil.isCellDateFormatted(cell)){
                            cell.setCellValue(String.valueOf(LocalDateTime.from(cell.getDateCellValue().toInstant()).toLocalDate().plusDays(1)));
                        }else{
                            cell.setCellValue(cell.getNumericCellValue() + 1500);
                        }
                        break;
                    }
                    case BOOLEAN: cell.setCellValue(!cell.getBooleanCellValue()); break;
                    default: break;
                }
            }
            i++;
        }
        String pathName = ".";
        String fileLocation = "";
        Scanner scanner = new Scanner(System.in);
        System.out.println("Want to save the new file in a specific location: y/n");
        String userChoice = scanner.nextLine();
        if(userChoice.equals("y")) {
            scanner = new Scanner(System.in);
            System.out.println("Please input the desired location without quotes: ");
            String userInput = scanner.nextLine();
            pathName = userInput;
            File currDir = new File(pathName);
            String path = currDir.getAbsolutePath();
            fileLocation = path + "\\EDITED_" + filename;
        }
        else{
            File currDir = new File(pathName);
            String path = currDir.getAbsolutePath();
            fileLocation = path.substring(0, path.length() - 1) +"EDITED_" + filename;
        }
        scanner.close();
        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        wb.write(outputStream);
        wb.close();
        System.out.println("File saved at: " + fileLocation);
    }
}


