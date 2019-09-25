package Horizon.DailyReport;

import java.io.InputStream;
import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Date;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.File;
import java.util.Properties;
import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.Logger;

public class ReadExcel
{
    static Logger log;
    
    static {
        ReadExcel.log = Logger.getLogger(ReadExcel.class.getName());
    }
    
    public static void main(final String[] args) {
        final Properties prop = getProperties();
        PropertyConfigurator.configure(prop.getProperty("log4jConfPath"));
        final String tableData = getDateTable(prop.getProperty("EXECUATION_DATE"));
        final String devData = readDevRelease(prop.getProperty("DEV_XLSX_FILE_PATH"), prop.getProperty("EXECUATION_DATE"));
        final String prodData = readProdRelease(prop.getProperty("PROD_XLSX_FILE_PATH"), prop.getProperty("EXECUATION_DATE"));
        final String supportData = readIssueRelease(prop.getProperty("DEV_ISSUE_XLSX_FILE_PATH"), prop.getProperty("PROD_ISSUE_XLSX_FILE_PATH"), prop.getProperty("EXECUATION_DATE"));
        final String execuationDate = getFinalExecuationDate(prop.getProperty("EXECUATION_DATE"));
        final String[] reportBody = { execuationDate, tableData, devData, prodData, " <table ><tr ><td cellpadding =\"-1\" width=100%><font size=\"2\" color=\"black\" face=\"Calibri\">1. Daily ENV Checks</td></tr><tr><td><font size=\"2\" color=\"black\" face=\"Calibri\">2. Daily Sanity Test</td></tr></table> ", supportData };
        ReadExcel.log.info((Object)reportBody);
        SendEmail.sendMail(reportBody);
    }
    
    public static String readDevRelease(final String DEV_XLSX_FILE_PATH, final String EXECUATION_DATE) {
        final String currentDate = getFinalExecuationDate(EXECUATION_DATE);
        String devDeployments = " <table >";
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(DEV_XLSX_FILE_PATH));
        }
        catch (EncryptedDocumentException e) {
            e.printStackTrace();
        }
        catch (InvalidFormatException e2) {
            e2.printStackTrace();
        }
        catch (IOException e3) {
            e3.printStackTrace();
        }
        final Iterator<Sheet> sheetIterator = (Iterator<Sheet>)workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            final Sheet sheet2 = sheetIterator.next();
        }
        int counter = 0;
        for (int tabCounter = 0; tabCounter <= 1; ++tabCounter) {
            final Sheet sheet = workbook.getSheetAt(tabCounter);
            final DataFormatter dataFormatter = new DataFormatter();
            final Iterator<Row> rowIterator = (Iterator<Row>)sheet.rowIterator();
            while (rowIterator.hasNext()) {
                final Row row = rowIterator.next();
                final String cellValue = dataFormatter.formatCellValue(row.getCell(8));
                final String verification_Status = dataFormatter.formatCellValue(row.getCell(10));
                if (currentDate.equals(cellValue) && (verification_Status.trim().equalsIgnoreCase("PASS") || verification_Status.trim().equalsIgnoreCase("FAIL"))) {
                    ++counter;
                    if (verification_Status.trim().equalsIgnoreCase("PASS")) {
                        devDeployments = String.valueOf(devDeployments) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(3)) + " deployed on " + dataFormatter.formatCellValue(row.getCell(2)) + "</td></tr>";
                    }
                    else {
                        devDeployments = String.valueOf(devDeployments) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(3)) + " rejected/failed on " + dataFormatter.formatCellValue(row.getCell(2)) + "</td></tr>";
                    }
                }
            }
        }
        if (counter == 0) {
            devDeployments = String.valueOf(devDeployments) + "<tr><td><font size=\"2\" color=\"black\" face=\"Calibri\">1. None </td></tr>";
        }
        devDeployments = String.valueOf(devDeployments) + "</font></table> ";
        return devDeployments;
    }
    
    public static String readProdRelease(final String PROD_XLSX_FILE_PATH, final String EXECUATION_DATE) {
        final String currentDate = getFinalExecuationDate(EXECUATION_DATE);
        String prodDeployments = " <table >";
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(PROD_XLSX_FILE_PATH));
        }
        catch (EncryptedDocumentException e) {
            e.printStackTrace();
        }
        catch (InvalidFormatException e2) {
            e2.printStackTrace();
        }
        catch (IOException e3) {
            e3.printStackTrace();
        }
        final Iterator<Sheet> sheetIterator = (Iterator<Sheet>)workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            final Sheet sheet2 = sheetIterator.next();
        }
        final Sheet sheet = workbook.getSheetAt(0);
        final DataFormatter dataFormatter = new DataFormatter();
        final Iterator<Row> rowIterator = (Iterator<Row>)sheet.rowIterator();
        int counter = 0;
        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            final String cellValue = dataFormatter.formatCellValue(row.getCell(10));
            final String verification_Status = dataFormatter.formatCellValue(row.getCell(11));
            if (currentDate.equals(cellValue)) {
                ++counter;
                prodDeployments = String.valueOf(prodDeployments) + "<tr><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(1)) + "-" + dataFormatter.formatCellValue(row.getCell(4)) + " deployed on " + dataFormatter.formatCellValue(row.getCell(2)) + "</td></tr>";
            }
        }
        if (counter == 0) {
            prodDeployments = String.valueOf(prodDeployments) + "<tr><td><font size=\"2\" color=\"black\" face=\"Calibri\">1. None </td></tr>";
        }
        prodDeployments = String.valueOf(prodDeployments) + "</table> ";
        return prodDeployments;
    }
    
    public static String readIssueRelease(final String DEV_ISSUE_XLSX_FILE_PATH, final String PROD_ISSUE_XLSX_FILE_PATH, final String EXECUATION_DATE) {
        final String finalexecuationDate = getFinalExecuationDate(EXECUATION_DATE);
        String supportTasks = " <table>";
        supportTasks = String.valueOf(supportTasks) + "<!--<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">1. Jenkins support.</td></tr>-->";
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(DEV_ISSUE_XLSX_FILE_PATH));
        }
        catch (EncryptedDocumentException e) {
            e.printStackTrace();
        }
        catch (InvalidFormatException e2) {
            e2.printStackTrace();
        }
        catch (IOException e3) {
            e3.printStackTrace();
        }
        final Iterator<Sheet> sheetIterator = (Iterator<Sheet>)workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            final Sheet sheet2 = sheetIterator.next();
        }
        final Sheet sheet = workbook.getSheetAt(0);
        final DataFormatter dataFormatter = new DataFormatter();
        final Iterator<Row> rowIterator = (Iterator<Row>)sheet.rowIterator();
        int counter = 0;
        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            final String cellValue = dataFormatter.formatCellValue(row.getCell(6));
            final String Issue_Status = dataFormatter.formatCellValue(row.getCell(9));
            if (Issue_Status.trim().equalsIgnoreCase("OPEN")) {
                ++counter;
                supportTasks = String.valueOf(supportTasks) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(1)) + " On " + dataFormatter.formatCellValue(row.getCell(3)) + " for " + dataFormatter.formatCellValue(row.getCell(2)) + " : " + dataFormatter.formatCellValue(row.getCell(4)) + " investigated/supported. Status : " + Issue_Status.trim().toUpperCase() + "</td></tr>";
            }
            else {
                if (!finalexecuationDate.equals(cellValue)) {
                    continue;
                }
                ++counter;
                supportTasks = String.valueOf(supportTasks) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(1)) + " On " + dataFormatter.formatCellValue(row.getCell(3)) + " for " + dataFormatter.formatCellValue(row.getCell(2)) + " : " + dataFormatter.formatCellValue(row.getCell(4)) + " investigated/supported. Status : " + Issue_Status.trim().toUpperCase() + "</td></tr>";
            }
        }
        supportTasks = String.valueOf(supportTasks) + readProdIssueRelease(PROD_ISSUE_XLSX_FILE_PATH, EXECUATION_DATE, counter);
        supportTasks = String.valueOf(supportTasks) + "</font></table> ";
        return supportTasks;
    }
    
    public static String readProdIssueRelease(final String PROD_ISSUE_XLSX_FILE_PATH, final String EXECUATION_DATE, final int devLastcounter) {
        final String finalexecuationDate = getFinalExecuationDate(EXECUATION_DATE);
        String supportTasks = "";
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(new File(PROD_ISSUE_XLSX_FILE_PATH));
        }
        catch (EncryptedDocumentException e) {
            e.printStackTrace();
        }
        catch (InvalidFormatException e2) {
            e2.printStackTrace();
        }
        catch (IOException e3) {
            e3.printStackTrace();
        }
        final Iterator<Sheet> sheetIterator = (Iterator<Sheet>)workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            final Sheet sheet2 = sheetIterator.next();
        }
        final Sheet sheet = workbook.getSheetAt(0);
        final DataFormatter dataFormatter = new DataFormatter();
        final Iterator<Row> rowIterator = (Iterator<Row>)sheet.rowIterator();
        int counter = devLastcounter;
        while (rowIterator.hasNext()) {
            final Row row = rowIterator.next();
            final String cellValue = dataFormatter.formatCellValue(row.getCell(6));
            final String Issue_Status = dataFormatter.formatCellValue(row.getCell(9));
            if (Issue_Status.trim().equalsIgnoreCase("OPEN")) {
                ++counter;
                supportTasks = String.valueOf(supportTasks) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(1)) + " On " + dataFormatter.formatCellValue(row.getCell(3)) + " for " + dataFormatter.formatCellValue(row.getCell(2)) + " : " + dataFormatter.formatCellValue(row.getCell(4)) + " investigated/supported. Status : " + Issue_Status.trim().toUpperCase() + "</td></tr>";
            }
            else {
                if (!finalexecuationDate.equals(cellValue)) {
                    continue;
                }
                ++counter;
                supportTasks = String.valueOf(supportTasks) + "<tr ><td><font size=\"2\" color=\"black\" face=\"Calibri\">" + counter + ". " + dataFormatter.formatCellValue(row.getCell(1)) + " On " + dataFormatter.formatCellValue(row.getCell(3)) + " for " + dataFormatter.formatCellValue(row.getCell(2)) + " : " + dataFormatter.formatCellValue(row.getCell(4)) + " investigated/supported. Status : " + Issue_Status.trim().toUpperCase() + "</td></tr>";
            }
        }
        return supportTasks;
    }
    
    public static String getDateTable(final String execuationDate) {
        final String dateTable = " <table><tr><td><font color=\"black\" face=\"Calibri\" size=\"2\" >" + getFinalExecuationDate(execuationDate) + "</td></tr>" + "</table> ";
        return dateTable;
    }
    
    public static String getFinalExecuationDate(final String execuationDate) {
        final DateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy");
        String finalexecuationDate = "";
        if ("".equals(execuationDate.toString().trim())) {
            final Date todayDate = new Date();
            finalexecuationDate = dateFormat.format(todayDate);
            final Calendar calendar = Calendar.getInstance();
            calendar.setTime(todayDate);
            calendar.add(6, -1);
            final Date previousDate = calendar.getTime();
            finalexecuationDate = dateFormat.format(previousDate);
        }
        else {
            try {
                final Date date = new SimpleDateFormat("dd-MM-yy").parse(execuationDate);
                finalexecuationDate = dateFormat.format(date);
            }
            catch (ParseException e) {
                e.printStackTrace();
            }
        }
        if (finalexecuationDate.startsWith("0")) {
            finalexecuationDate = finalexecuationDate.replaceFirst("0", "");
        }
        return finalexecuationDate;
    }
    
    public static Properties getProperties() {
        final Properties prop = new Properties();
        InputStream input = null;
        try {
        	input = new FileInputStream("D:\\Test\\config\\config.properties");
           // input = new FileInputStream("./config/config.properties");
            
            prop.load(input);
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return prop;
    }
}
