/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.wmk;

import com.spire.xls.ExcelVersion;
import com.spire.xls.FileFormat;
import com.spire.xls.HyperLink;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;

/**
 *
 * @author HP
 */
public class Kwft_document_extraction {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws SQLException, IOException {
        // TODO code application logic here
        
        Connection DBConnection = null;
        String strDBConnectionUrl = "jdbc:sqlserver://127.0.0.1:1433;databaseName=ILAB_DEMO_DB";
        String strUser = "sa";
        String strPassword = "slyfox37337";
        String strQuery = "select distinct b.documentid, a.EMPLOYEE_NAME, b.LOCATION, b.VERSIONNUMBER,b.FILENAME, replace(replace(b.LOCATION, '/', '\\'), '\\sonora', 'c:\\sonora') as proper_filename from FS_STAFF_DOCS a inner join SONORASTRGELEMENTS b on a.DOCUMENTID = b.DOCUMENTID";
        Statement DBStatement;
        ResultSet myResults;
        String strColumnValue;
        String strOutputDirectory = "C:\\test_kwft_docs";
        String strExcelPath = "C:\\test_kwft_docs\\csop.xlsx";
        DBConnection = DriverManager.getConnection(strDBConnectionUrl, strUser, strPassword);
        if (DBConnection == null)
        {
            System.out.println("Could not connect to DB server");
        }
        
        DBStatement = DBConnection.createStatement();
        myResults = DBStatement.executeQuery(strQuery);
        
        ResultSetMetaData myResultSetMetaData = myResults.getMetaData();
        
        //create excel document
        
        Workbook myExcelWorkBook = new Workbook();
        
        //load the sample excel file
        
        //myExcelWorkBook.loadFromFile(strExcelPath);
        
        //Get the first worksheet
        
        Worksheet myExcelWorkSheet = myExcelWorkBook.getWorksheets().get(0);
        
        
        
        
       
        
        if (myResults != null)
        {
            //get the number of columns and get column names
                
            int intNoOfColumns = myResultSetMetaData.getColumnCount();
            
            
            while (myResults.next())
            {
                String strDocumentId = myResults.getString("documentid");
                String strEmployeeName = myResults.getString("EMPLOYEE_NAME");
                String strLocation = myResults.getString("LOCATION");
                String strFileName = myResults.getString("FILENAME");
                String strVersionNumber = myResults.getString("VERSIONNUMBER");
                String strProperFilename = myResults.getString("proper_filename");
                String strProperFileFullPath = strOutputDirectory + "\\" + strFileName;
                
                int intDocumentId = Integer.parseInt(strDocumentId);
                int intVersionNumber = Integer.parseInt(strVersionNumber);
                String strDocFileName = strFileName;
                String strTheTargetFileName = strFileName;
                String strTargetFileNameWithoutExtension = strFileName.substring(0, strFileName.indexOf('.'));
                
                for (int i = 1; i < intNoOfColumns + 1 ; i++) //add 1 for the extra column for the for loop since the counter (i) is starting from 1
                {
                    String strColumnName;
                    int intCurrentRow = myResults.getRow()+1;
                    String strExcelCellName;
                    String strCurrentRow = Integer.toString(intCurrentRow);
                    if (i == 1)
                    {
                        strColumnName = "A";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            
                            //Header for excel document
                            
                            strColumnValue = "DOCUMENTID";
                            
                            System.out.println("DEBUG ============> " + strColumnValue);
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strDocumentId);
                        }
                        
                    }
                    if (i == 2)
                    {
                        strColumnName = "B";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            
                            //Header for excel document
                            
                            strColumnValue = "EMPLOYEE_NAME";
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strEmployeeName);
                        }
                        
                    }
                    if (i == 3)
                    {
                        strColumnName = "C";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            
                            //Header for excel document
                            
                            strColumnValue = "VERSION_NUMBER";
                            
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strVersionNumber);
                        }
                        
                    }
                    if (i == 4)
                    {
                        strColumnName = "D";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            
                            //Header for excel document
                            
                            strColumnValue = "FILE_NAME";
                            
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strFileName);
                        }
                        
                    }
                    if (i == 5)
                    {
                        strColumnName = "E";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            //Header for excel document
                            
                            strColumnValue = "PROPER_FILE_NAME";
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strProperFilename);
                        }
                        
                    }
                    if (i == 6)
                    {
                        
                        strColumnName = "F";
                        //if(intCurrentRow ==1)
                        if(myResults.getRow() ==1)
                        {
                            //strColumnValue = myResults.getString(myResultSetMetaData.getColumnName(i));
                            
                            //Header for excel document
                            
                            strColumnValue = "LINKED_DOCUMENT";
                            
                            strExcelCellName = strColumnName + strCurrentRow;
                            myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strColumnValue);
                        }
                        else
                        {
                            strExcelCellName = strColumnName + strCurrentRow;
                        
                            //add a link on excel that opens the document
                        
                            HyperLink fileLink = myExcelWorkSheet.getHyperLinks().add(myExcelWorkSheet.getCellRange(strExcelCellName));
                            fileLink.setTextToDisplay(strFileName);
                            fileLink.setAddress(strProperFileFullPath);
                        
                        
                            //myExcelWorkSheet.getCellRange(strExcelCellName).setValue(strProperFileFullPath);
                        }
                        
                    }
                }
               
                
                //display database results
                
                System.out.println("||" + strDocumentId + "||" + strEmployeeName + "||" + strLocation + "||" + strFileName + "||" + strVersionNumber +"||" + strProperFilename + "||" + strProperFileFullPath + "||");
                
                //insert function for extracting documents
                
                extractStorageControllerDocuments(strLocation, intDocumentId, intVersionNumber, strDocFileName, strOutputDirectory, strTargetFileNameWithoutExtension);
                
                //insert function for writing path to excel document.
                
            }
            
        }
        
        //save the excel file
        
        myExcelWorkBook.saveToFile(strExcelPath, ExcelVersion.Version2013);
        
        myResults.close();
        DBStatement.close();
        DBConnection.close();
        
        
    }
    
    public static void extractStorageControllerDocuments(String strStorageControllerPath, int intDocumentID, int intVersionNumber, String strDocFileName, String strOutputDirectory, String strTargetFileNameWithoutExtension) throws IOException
    {
        String strFullDocumentPath = null;
        String strFileNameWithoutExtension = strDocFileName.substring(0, strDocFileName.indexOf('.'));
        String strExtension = strDocFileName.substring(strDocFileName.indexOf('.'), strDocFileName.length());
        String strTargetFileName = strTargetFileNameWithoutExtension + strExtension;
        String strTargetFileNameWithFullPath = strOutputDirectory + "\\" + strTargetFileName;
        
        
        
        if ((intDocumentID > 0) && (intDocumentID < 10))
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            //char chrHundredsCharacter = chrArrDocumentID[0];
            
            char chrOnesCharacter = chrArrDocumentID[0];
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\000\\0" + strOnes + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 9) && (intDocumentID < 100))
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            //char chrHundredsCharacter = chrArrDocumentID[0];
            char chrTensCharacter = chrArrDocumentID[0];
            char chrOnesCharacter = chrArrDocumentID[1];
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\000\\" + strTens + strOnes + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 99) && (intDocumentID < 1000)) //234
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrHundredsCharacter = chrArrDocumentID[0];
            char chrTensCharacter = chrArrDocumentID[1];
            char chrOnesCharacter = chrArrDocumentID[2];
            
            //convert the character to string
            
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\00" + strHundreds + "\\" + strTens + strOnes + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 999) && (intDocumentID < 10000)) //2345
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrThousandCharacter = chrArrDocumentID[0];
            char chrHundredsCharacter = chrArrDocumentID[1];
            char chrTensCharacter = chrArrDocumentID[2];
            char chrOnesCharacter = chrArrDocumentID[3];
            
            //convert the character to string
            
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\0" + strThousands + strHundreds + "\\"  + strTens + strOnes + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 9999) && (intDocumentID < 100000)) //23456
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrTenThousandCharacter = chrArrDocumentID[0];
            char chrThousandCharacter = chrArrDocumentID[1];
            char chrHundredsCharacter = chrArrDocumentID[2];
            char chrTensCharacter = chrArrDocumentID[3];
            char chrOnesCharacter = chrArrDocumentID[4];
            
            //convert the character to string
            
            String strTenThousands = String.valueOf(chrTenThousandCharacter);
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            //strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\" +strTenThousands + strThousands + strHundreds + "\\" + Integer.toString(intDocumentID) + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\000\\" +strTenThousands + strThousands + strHundreds + "\\" + strTens + strOnes + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 99999) && (intDocumentID < 1000000)) //234567
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrHundredThousandCharacter = chrArrDocumentID[0];
            char chrTenThousandCharacter = chrArrDocumentID[1];
            char chrThousandCharacter = chrArrDocumentID[2];
            char chrHundredsCharacter = chrArrDocumentID[3];
            char chrTensCharacter = chrArrDocumentID[4];
            char chrOnesCharacter = chrArrDocumentID[5];
            
            //convert the character to string
            
            String strHundredThousands = String.valueOf(chrHundredThousandCharacter);
            String strTenThousands = String.valueOf(chrTenThousandCharacter);
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\00" + strHundredThousands + "\\" +strTenThousands + strThousands + strHundreds + "\\" + Integer.toString(intDocumentID) + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 999999) && (intDocumentID < 10000000)) //2345678
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrMillionCharacter = chrArrDocumentID[0];
            char chrHundredThousandCharacter = chrArrDocumentID[1];
            char chrTenThousandCharacter = chrArrDocumentID[2];
            char chrThousandCharacter = chrArrDocumentID[3];
            char chrHundredsCharacter = chrArrDocumentID[4];
            char chrTensCharacter = chrArrDocumentID[5];
            char chrOnesCharacter = chrArrDocumentID[6];
            
            //convert the character to string
            
            String strMillions = String.valueOf(chrMillionCharacter);
            String strHundredThousands = String.valueOf(chrHundredThousandCharacter);
            String strTenThousands = String.valueOf(chrTenThousandCharacter);
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\0" + strMillions + strHundredThousands + "\\" +strTenThousands + strThousands + strHundreds + "\\" + Integer.toString(intDocumentID) + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 9999999) && (intDocumentID < 100000000)) //23456789
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrTenMillionCharacter = chrArrDocumentID[0];
            char chrMillionCharacter = chrArrDocumentID[1];
            char chrHundredThousandCharacter = chrArrDocumentID[2];
            char chrTenThousandCharacter = chrArrDocumentID[3];
            char chrThousandCharacter = chrArrDocumentID[4];
            char chrHundredsCharacter = chrArrDocumentID[5];
            char chrTensCharacter = chrArrDocumentID[6];
            char chrOnesCharacter = chrArrDocumentID[7];
            
            //convert the character to string
            
            String strTenMillions = String.valueOf(chrTenMillionCharacter);
            String strMillions = String.valueOf(chrMillionCharacter);
            String strHundredThousands = String.valueOf(chrHundredThousandCharacter);
            String strTenThousands = String.valueOf(chrTenThousandCharacter);
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\000\\"+ strTenMillions + strMillions + strHundredThousands + "\\" +strTenThousands + strThousands + strHundreds + "\\" + Integer.toString(intDocumentID) + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
        }
        if ((intDocumentID > 99999999) && (intDocumentID < 1000000000)) //234567891
        {
            char[] chrArrDocumentID;
            chrArrDocumentID = new char[255];
            
            //convert intDocumentID to String and then convert the String to an array of characters
            
            chrArrDocumentID = Integer.toString(intDocumentID).toCharArray();
            
            char chrHundredMillionCharacter = chrArrDocumentID[0];
            char chrTenMillionCharacter = chrArrDocumentID[1];
            char chrMillionCharacter = chrArrDocumentID[2];
            char chrHundredThousandCharacter = chrArrDocumentID[3];
            char chrTenThousandCharacter = chrArrDocumentID[4];
            char chrThousandCharacter = chrArrDocumentID[5];
            char chrHundredsCharacter = chrArrDocumentID[6];
            char chrTensCharacter = chrArrDocumentID[7];
            char chrOnesCharacter = chrArrDocumentID[8];
            
            //convert the character to string
            
            String strHundredMillions = String.valueOf(chrHundredMillionCharacter);
            String strTenMillions = String.valueOf(chrTenMillionCharacter);
            String strMillions = String.valueOf(chrMillionCharacter);
            String strHundredThousands = String.valueOf(chrHundredThousandCharacter);
            String strTenThousands = String.valueOf(chrTenThousandCharacter);
            String strThousands = String.valueOf(chrThousandCharacter);
            String strHundreds = String.valueOf(chrHundredsCharacter);
            String strTens = String.valueOf(chrTensCharacter);
            String strOnes = String.valueOf(chrOnesCharacter);
            strFullDocumentPath = strStorageControllerPath + "\\" + "0000000\\00" + strHundredMillions + "\\"+ strTenMillions + strMillions + strHundredThousands + "\\" +strTenThousands + strThousands + strHundreds + "\\" + Integer.toString(intDocumentID) + "0000" + Integer.toString(intVersionNumber)+ "-" + strDocFileName;
            
        }
        CopyFileToTargetFolder(strFullDocumentPath, strTargetFileNameWithFullPath);
        
        
    }
    
    public static void CopyFileToTargetFolder(String strSourceFile, String strDestinationFile) throws FileNotFoundException, IOException
    {
         FileInputStream ins = null;
         FileOutputStream outs = null;
         File infile = new File(strSourceFile);
         File outfile = new File(strDestinationFile);
         ins = new FileInputStream(infile);
         outs = new FileOutputStream(outfile);
         byte[] buffer = new byte[1024];
         int length;
         
         while ((length = ins.read(buffer)) > 0) {
            outs.write(buffer, 0, length);
         } 
         ins.close();
         outs.close();
    }
    
}
