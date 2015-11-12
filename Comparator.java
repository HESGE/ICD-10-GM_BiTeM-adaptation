/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package comparator;

import java.io.* ;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;


/**
 *
 * @author mottinl
 */
public class Comparator {

    /**
     * Main function
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String[] args) throws IOException {
        delta_MVC_MTC();
        translation();
        transcoding_Map_HUG();
    }


    public static void transcoding_Map_HUG() throws IOException {
        //Get the input files
        FileInputStream mvcFile = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\Informal_epSOS-MVC_V2_0_(DRAFT)_03.xlsx"));
        //Get the workbook instance for XLS file 
        XSSFWorkbook mvcWorkbook = new XSSFWorkbook(mvcFile);
        XSSFSheet mvcSheet;
        Iterator<Row> mvcRowIterator;
        String mvcSheetName;
        int mvcCol;
        boolean mvcColFound;
        Row mvcRow;
        Row mvcRow2;
        Iterator<Cell> mvcCellIterator;
        boolean statusOK = false;

        //OUTPUT
        String code_src;
        String code_dest;
        String name_dest = "";
        String value_set_name_dest = "";
        String status = "none";
        String value_set_name_source = "";
        String value_set_oid_dest = "";
        String parent_system_code_dest = "";
        String parent_system_oid_dest = "";
        String comment = "";
        String map_level = "0";
        String review = "0";
        String version = "";

        //Prepare the output file
        Writer csvW = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\Map_HUG\\map_hug_to_mvc_2.0.csv"), "UTF-8"));
        csvW.write('\ufeff');
        csvW.write("code_src;code_dest;name_dest;value_set_name_dest;status;value_set_name_source;value_set_oid_dest;parent_system_code_dest;parent_system_oid_dest;comment;map_level;review;version;");
        csvW.write("\n");

        //Read csv map
        String map = "\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\Map_HUG\\map_hug_to_mvc_1_9.csv";
        try {
            BufferedReader br = new BufferedReader(new FileReader(map));
            String line = "";
            String csvSplitBy = ";";
            String[] maLigne;

            //jump over the first line
            br.readLine();
            //pour chaque ligne de la map
            while ((line = br.readLine()) != null) {
                statusOK = false;

                maLigne = line.split(csvSplitBy);
                code_src = maLigne[0];
                code_dest = maLigne[1];

                //Get the sheet from the MTC workbook
                for(int i=0; i<mvcWorkbook.getNumberOfSheets(); i++) {
                    mvcSheet = mvcWorkbook.getSheetAt(i);

                    //Get iterator to all the rows in current MTC sheet
                    mvcRowIterator = mvcSheet.iterator();

                    //Get the name of MTTC sheet, compare them MAP entries
                    //MVC data files are called "epSOSsheetName"
                    mvcSheetName = mvcSheet.getSheetName();

                    //And process the file matching to find the good sheet
                    if(mvcSheetName.equals(maLigne[3])) {
                        value_set_name_dest = mvcSheetName;
                        value_set_name_source = maLigne[5];

                        mvcCol = 0;
                        mvcColFound = false;

                        while(mvcRowIterator.hasNext()) {
                            mvcRow = mvcRowIterator.next();
                            mvcRow2 = mvcRow;

                            if(mvcColFound == false) {
                                mvcCellIterator = mvcRow.cellIterator();

                                while(mvcCellIterator.hasNext()) {
                                    Cell mvcCell = mvcCellIterator.next();

                                    if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("Parent Code System:"))) {
                                        mvcCol = mvcCell.getColumnIndex()+1;
                                        mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                        parent_system_code_dest = mvcRow.getCell(mvcCol).getStringCellValue().trim();
                                    }
                                    if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("OID Parent Code System:"))) {
                                        mvcCol = mvcCell.getColumnIndex()+1;
                                        mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                        parent_system_oid_dest = mvcRow.getCell(mvcCol).getStringCellValue().trim();
                                    }
                                    if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("epSOS OID:"))) {
                                        mvcCol = mvcCell.getColumnIndex()+1;
                                        mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                        value_set_oid_dest = mvcRow.getCell(mvcCol).getStringCellValue().trim();
                                    }
                                    if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("version:"))) {
                                        mvcCol = mvcCell.getColumnIndex()+1;
                                        mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                        version = mvcRow.getCell(mvcCol).getStringCellValue().trim();
                                    }

                                    if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("epSOS Code") || mvcCell.getStringCellValue().equals("Code"))) {
                                        mvcCol = mvcCell.getColumnIndex();
                                        mvcColFound = true;
                                        break;
                                    }
                                }
                            }
                            else {
                                mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                if (mvcRow.getCell(mvcCol).getStringCellValue().trim().equals(code_dest)) {
                                    statusOK = true;
                                    mvcRow2.getCell(mvcCol+1, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                                    name_dest = mvcRow2.getCell(mvcCol+1).getStringCellValue().trim();
                                    break;
                                }

                            }
                        }
                        if(statusOK==true) {
                            break;
                        }
                        else {
                            parent_system_code_dest = "";
                            parent_system_oid_dest = "";
                            value_set_oid_dest = "";
                            version = "";
                        }
                    }
                }

                if(statusOK!=true) {
                    //TO CHECK MANUALY
                    status = "manual";
                    name_dest = maLigne[2];
                    comment = "mvc2.0 no hug code";
                }

                //Write the mapping
                csvW.write(code_src+";"+code_dest+";"+name_dest+";"+value_set_name_dest+";"+status+";"+value_set_name_source+";"+value_set_oid_dest+";"+parent_system_code_dest+";"+parent_system_oid_dest+";"+comment+";"+map_level+";"+review+";"+version+";");
                csvW.write("\n");
                //reset status
                status = "none";
                comment = "";

            }

            br.close();

        } catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}

        csvW.flush();
        csvW.close();

    }


    public static void delta_MVC_MTC() throws IOException {
        //Get the input files
        //FileInputStream mtcFile = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\Catalogues\\workingMTC.xlsx"));
        //FileInputStream mvcFile = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\Catalogues\\Informal_epSOS-MVC_V1_9.xlsx"));
        FileInputStream mtcFile = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\MTC_2.0.xlsx"));
        FileInputStream mtcFile2 = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\MTC_2.0.xlsx"));
        FileInputStream mvcFile = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\Informal_epSOS-MVC_V2_0_(DRAFT)_03.xlsx"));

        //Prepare the output file
        //Writer csvW = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\Catalogues\\delta_Mtc-Mvc.csv"), "UTF-8"));
        Writer csvW = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\delta_Mtc-Mvc2.1.csv"), "UTF-8"));
        csvW.write('\ufeff');

        csvW.write("Expand Project;");
        csvW.write("\n\n");

        //Get the workbook instance for XLS file 
        XSSFWorkbook mtcWorkbook = new XSSFWorkbook(mtcFile);
        XSSFWorkbook mtcWorkbook2 = new XSSFWorkbook(mtcFile2);
        XSSFWorkbook mvcWorkbook = new XSSFWorkbook(mvcFile);

        //Output
        csvW.write("One MTC sheet is missing in MVC : VS16_epSOSErrorCodes;");
        csvW.write("\n");
        csvW.write("********************;");
        csvW.write("\n");
        csvW.write("Set name;");
        csvW.write("\n");
        csvW.write("MTC mismatches;List of the codes missing in MVC");
        csvW.write("\n");
        csvW.write("MVC mismatches;List of the codes missing in MTC");
        csvW.write("\n");
        csvW.write("********************;");

        XSSFSheet mtcSheet;
        XSSFSheet mtcSheet2;
        Iterator<Row> mtcRowIterator;
        Iterator<Row> mtcRowIterator2;
        Iterator<Row> mvcRowIterator;
        Iterator<Cell> mtcCellIterator;
        Iterator<Cell> mvcCellIterator;
        int mtcCol;
        int mvcCol;
        boolean mtcColFound;
        boolean mvcColFound;
        ArrayList mtcCodes;
        ArrayList mvcCodes;
        ArrayList mtcEnglishNames;
        ArrayList mvcEnglishNames;
        ArrayList englishNamesdifferences;
        Row mtcRow;
        Row mtcRow2;
        Row mvcRow;
        Row mvcRow2;
        Row newRow;
        Cell newCell;
        CellStyle myStyle;
        String mtcSplit[];
        String mvcSplit[];
        String mtcSheetName;
        String mvcSheetName;

        //Get the sheet from the MTC workbook
        for(int i=0; i<mtcWorkbook.getNumberOfSheets(); i++) {
            mtcSheet = mtcWorkbook.getSheetAt(i);
            mtcSheet2 = mtcWorkbook2.getSheetAt(i);

            //Get iterator to all the rows in current MTC sheet
            mtcRowIterator = mtcSheet.iterator();
            mtcRowIterator2 = mtcSheet2.iterator();

            //Get the sheet from the MVC workbook
            for(int j=0; j<mvcWorkbook.getNumberOfSheets(); j++) {
                XSSFSheet mvcSheet = mvcWorkbook.getSheetAt(j);

                //Get iterator to all the rows in current MVC sheet
                mvcRowIterator = mvcSheet.iterator();

                //Get the name of MTC sheet and MVC sheet, compare them if they contain data
                //MTC data files are called "VSX_sheetName"
                //MVC data files are called "epSOSsheetName"
                mtcSplit = mtcSheet.getSheetName().split("_");
                mvcSplit = mvcSheet.getSheetName().split("SOS");
                mtcSheetName = mtcSplit[mtcSplit.length-1];
                mvcSheetName = mvcSplit[mvcSplit.length-1];

                //And process the file matching or throw out the file that has no equivalent
                if(mtcSheetName.equals(mvcSheetName)) {

                    mtcCol = 0;
                    mvcCol = 0;
                    mtcColFound = false;
                    mvcColFound = false;
                    mtcCodes = new ArrayList();
                    mvcCodes = new ArrayList();
                    mtcEnglishNames = new ArrayList();
                    mvcEnglishNames = new ArrayList();
                    englishNamesdifferences = new ArrayList();

                    //For each row, iterate through each columns
                    //Get iterator to all cells of current row
                    //In MTC
                    while(mtcRowIterator.hasNext()) {
                        mtcRow = mtcRowIterator.next();
                        mtcRow2 = mtcRow;

                        if(mtcColFound == false) {
                            mtcCellIterator = mtcRow.cellIterator();

                            while(mtcCellIterator.hasNext()) {
                                Cell mtcCell = mtcCellIterator.next();
                                if(mtcCell.getCellType() == 1 && (mtcCell.getStringCellValue().equals("Code") || mtcCell.getStringCellValue().equals("epSOS Code"))) {
                                    mtcCol = mtcCell.getColumnIndex();
                                    mtcColFound = true;
                                    break;
                                }
                            }
                        }
                        else{
                            mtcRow.getCell(mtcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                            mtcRow2.getCell(mtcCol+1, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                            mtcCodes.add(mtcRow.getCell(mtcCol).getStringCellValue().trim());
                            mtcEnglishNames.add(mtcRow2.getCell(mtcCol+1).getStringCellValue().trim());
                        }
                    }

                    //In MVC
                    while(mvcRowIterator.hasNext()) {
                        mvcRow = mvcRowIterator.next();
                        mvcRow2 = mvcRow;
                        if(mvcColFound == false) {
                            mvcCellIterator = mvcRow.cellIterator();

                            while(mvcCellIterator.hasNext()) {
                                Cell mvcCell = mvcCellIterator.next();

                                if(mvcCell.getCellType() == 1 && (mvcCell.getStringCellValue().equals("epSOS Code") || mvcCell.getStringCellValue().equals("Code"))) {
                                    mvcCol = mvcCell.getColumnIndex();
                                    mvcColFound = true;
                                    break;
                                }
                            }
                        }
                        else {
                            mvcRow.getCell(mvcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                            mvcRow2.getCell(mvcCol+1, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                            mvcCodes.add(mvcRow.getCell(mvcCol).getStringCellValue().trim());
                            mvcEnglishNames.add(mvcRow2.getCell(mvcCol+1).getStringCellValue().trim());
                        }
                    }

                    //Processing
                    colCompare(mtcCodes, mvcCodes, mvcEnglishNames, mtcEnglishNames, englishNamesdifferences);

                    //Output
                    //if((!mtcCodes.isEmpty()) || (!mvcCodes.isEmpty())) {}
                    csvW.write("\n\n");
                    csvW.write(mtcSheetName + ";");
                    csvW.write("\n");
                    csvW.write("MTC mismatches;");
                    for(int a=0; a<mtcCodes.size(); a++) {
                        csvW.write(mtcCodes.get(a) + ";");
                    }
                    csvW.write("\n");
                    csvW.write("MVC mismatches\n");
                    for(int b=0; b<mvcCodes.size(); b++) {
                        csvW.write(mvcCodes.get(b) + ";"+mvcEnglishNames.get(b) + "\n");
                    }

                    csvW.write("english names differences\n");
                    if (!englishNamesdifferences.isEmpty()) {
                        csvW.write("code;MTC 2.0;MVC 2.0.1\n");
                        for(int c=0; c<englishNamesdifferences.size(); c=c+3) {
                            csvW.write(englishNamesdifferences.get(c) + ";" + englishNamesdifferences.get(c+1) + ";" + englishNamesdifferences.get(c+2) + "\n");
                        }
                    }

                    /* work on currents MTC2.0 sheet */
                    mtcColFound = false;
                    mtcCol = 0;
                    List<Integer> delRows = new ArrayList();

                    //recreate iterator to all the rows in current MTC sheet
                    while(mtcRowIterator2.hasNext()) {
                        mtcRow = mtcRowIterator2.next();
                        mtcRow2 = mtcRow;
                        if(mtcColFound == false) {
                            mtcCellIterator = mtcRow.cellIterator();

                            while(mtcCellIterator.hasNext()) {
                                Cell mtcCell = mtcCellIterator.next();
                                if(mtcCell.getCellType() == 1 && (mtcCell.getStringCellValue().equals("Code") || mtcCell.getStringCellValue().equals("epSOS Code"))) {
                                    mtcCol = mtcCell.getColumnIndex();
                                    mtcColFound = true;
                                    break;
                                }
                            }
                        }
                        else{
                            mtcRow.getCell(mtcCol, Row.RETURN_NULL_AND_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                            mtcRow2.getCell(mvcCol+1, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);

                            for(int a=0; a<mtcCodes.size(); a++) {
                                if (mtcRow.getCell(mtcCol).getStringCellValue().trim().equals(mtcCodes.get(a))) {
                                    // delete row corresponding to useless code
                                    delRows.add(mtcRow.getRowNum());
                                    break;
                                }
                            }

                            if (!englishNamesdifferences.isEmpty()) {
                                for(int c=0; c<englishNamesdifferences.size(); c=c+3) {
                                    if (mtcRow2.getCell(mtcCol+1).getStringCellValue().trim().equals(englishNamesdifferences.get(c+1))) {
                                        mtcRow2.getCell(mtcCol+1).setCellValue(englishNamesdifferences.get(c+2).toString());
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    for (int d=delRows.size()-1;d>=0;d--) {
                        mtcSheet2.shiftRows(delRows.get(d)+1, mtcSheet2.getLastRowNum()+1, -1);
                    }
                    myStyle = mtcSheet2.getRow(0).getCell(0).getCellStyle();
                    for(int b=0; b<mvcCodes.size(); b++) {
                        newRow = mtcSheet2.createRow(mtcSheet2.getLastRowNum()+1);
                        for(int bb=0; bb<mtcSheet2.getRow(0).getLastCellNum(); bb++) {
                            newCell = newRow.createCell(bb);
                            newCell.setCellStyle(myStyle);
                            if (bb==mtcCol) {
                                newCell.setCellValue(mvcCodes.get(b).toString());
                            }
                            else if (bb==mtcCol+1) {
                                newCell.setCellValue(mvcEnglishNames.get(b).toString());
                            }
                        }
                    }
                }
            }
        }
        //close InputStream
        mtcFile.close();
        mtcFile2.close();
        mvcFile.close();
        //close OutputStream
        csvW.close();

        //Open FileOutputStream to write updates
        FileOutputStream output_file =new FileOutputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\MTC_2.0_new.xlsx"));
        //write changes
        mtcWorkbook2.write(output_file);
        //close the stream
        output_file.close(); 
    }


    /**
    * Array comparison function
    * @param mtcCodes
    * @param mvcCodes
    * @param mvcEnglishNames
    * @param mtcEnglishNames
    * @param englishNamesdifferences
    */
    public static void colCompare(ArrayList mtcCodes, ArrayList mvcCodes, ArrayList mvcEnglishNames, ArrayList mtcEnglishNames, ArrayList englishNamesdifferences) {
        //If mvcCodes is bigger than mtcCodes, we'll match all elements from mtc in mvc
        //Remove the matching ones
        //And keep the others for the output
        if(mtcCodes.size() < mvcCodes.size()) {
            for(int a=0; a<mtcCodes.size(); a++) {
                for(int b=0; b<mvcCodes.size(); b++) {
                    if(mtcCodes.isEmpty()) {
                        break;
                    }
                    if(a == -1) {
                        a=0;
                    }
                    if(mtcCodes.get(a).equals(mvcCodes.get(b))) {
                        if(!mtcEnglishNames.get(a).equals(mvcEnglishNames.get(b))) {
                            englishNamesdifferences.add(mtcCodes.get(a));
                            englishNamesdifferences.add(mtcEnglishNames.get(a));
                            englishNamesdifferences.add(mvcEnglishNames.get(b));
                        }
                        mtcCodes.remove(a);
                        mvcCodes.remove(b);
                        mvcEnglishNames.remove(b);
                        mtcEnglishNames.remove(a);
                        a--;
                        b=-1;
                    }
                }
            }
        }
        //The opposite
        else if(mtcCodes.size() > mvcCodes.size()) {
            for(int b=0; b<mvcCodes.size(); b++) {
                for(int a=0; a<mtcCodes.size(); a++) {
                    if(mvcCodes.isEmpty()) {
                        break;
                    }
                    if(b == -1) {
                        b=0;
                    }
                    if(mtcCodes.get(a).equals(mvcCodes.get(b))) {
                        if(!mtcEnglishNames.get(a).equals(mvcEnglishNames.get(b))) {
                            englishNamesdifferences.add(mtcCodes.get(a));
                            englishNamesdifferences.add(mtcEnglishNames.get(a));
                            englishNamesdifferences.add(mvcEnglishNames.get(b));
                        }
                        mtcCodes.remove(a);
                        mvcCodes.remove(b);
                        mvcEnglishNames.remove(b);
                        mtcEnglishNames.remove(a);
                        b--;
                        a=-1;
                    }
                }
            }
        }
        //And the other cases when MTC and MVC have the same number of codes
        else if(mtcCodes.size() == mvcCodes.size()){
            for(int a=0; a<mtcCodes.size(); a++) {
                for(int b=0; b<mvcCodes.size(); b++) {
                    if(mtcCodes.isEmpty()) {
                        break;
                    }
                    if(a == -1) {
                        a=0;
                    }
                    if(mtcCodes.get(a).equals(mvcCodes.get(b))) {
                        if(!mtcEnglishNames.get(a).equals(mvcEnglishNames.get(b))) {
                            englishNamesdifferences.add(mtcCodes.get(a));
                            englishNamesdifferences.add(mtcEnglishNames.get(a));
                            englishNamesdifferences.add(mvcEnglishNames.get(b));
                        }
                        mtcCodes.remove(a);
                        mvcCodes.remove(b);
                        mvcEnglishNames.remove(b);
                        mtcEnglishNames.remove(a);
                        a=-1;
                        b=-1;
                    }
                }
            }
        }
    }


    /*
        !!!!!!!!!!
        TAKE CARE : ATC codes file (belgium) not reliable
        can use : http://stoppstart.free.fr/atc.php?cl=A
        or : https://www.vidal.fr/classifications/atc/
        !!!!!!!!!!
    */
    //MTC filling
    public static void translation() throws IOException {

        //Get the input files
        FileInputStream newMTC = new FileInputStream(new File("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\MTC_2.0.xlsx"));

        String icdCodes = "\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\CIM-10\\CIM10GM2014_S_FR_ClaML_2014.10.31.xml";
        String atcCodes = "\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\ATCcodes\\ATCDPP.CSV";

        //Prepare the output file
        Writer csvW = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\CIM10-treated.csv"), "UTF-8"));
        csvW.write('\ufeff');
        Writer csvW2 = new BufferedWriter(new OutputStreamWriter(new FileOutputStream("\\\\hes-nas-drize.hes.adhes.hesge.ch\\home\\luc.mottin\\Documents\\Expand\\MVC2.0\\ATC-treated.csv"), "UTF-8"));
        csvW2.write('\ufeff');
        List<String> translationList = new ArrayList();
        Map<String, String> translatList = new HashMap();
        Map<String, String> translatAtcList = new HashMap();
        Map<String, String> translatAtcList2 = new HashMap();
        String codeTemp = "";
        boolean prefered = false;

        InputStream ips = new FileInputStream(icdCodes); 
        //Cp1252 --> ANSI
        InputStreamReader ipsr = new InputStreamReader(ips, "UTF-8");
        BufferedReader br = new BufferedReader(ipsr);
        String ligne;
        Pattern p1 = Pattern.compile("<Class code=\"(.+?)\"");
        Pattern p2 = Pattern.compile("xml:space=\"default\">(.+?)<");
        Pattern p3 = Pattern.compile("(.+?)\\..");
        Pattern pActiveIngredient = Pattern.compile("(?:.*;){8}\"(.+?)\";(?:.*;)\"(.+?)\";(?:.*;){5}.*");
        Pattern pActiveIngredient2 = Pattern.compile("(?:.*;){4}\"(.+?)\";(?:.*;){5}\"(.+?)\";(?:.*;){5}.*");
        Matcher m1;
        Matcher m2;
        Matcher m3;
        Matcher mActiveIngredient;
        Matcher mActiveIngredient2;

        while((ligne=br.readLine())!= null) {
            m1 = p1.matcher(ligne);
            m2 = p2.matcher(ligne);

            if (ligne.matches("</Class>")) {
                prefered = false;
                codeTemp = "";
            }

            if (m1.find()){
                codeTemp = m1.group(1);
            }

            if (ligne.matches("(.*)kind=\"preferred\"(.*)")){
                prefered = true;
            }

            if (m2.find() && prefered==true){
                translatList.put(codeTemp, m2.group(1));
                prefered = false;
            }

            //si traduction français ET anglais
            if (ligne.matches(".*<FR_OMS>.*</FR_OMS>.*") && ligne.matches(".*<EN_OMS>.*</EN_OMS>.*")) {
                    translationList.add(ligne.replace("\u00A0", " "));
            }
        }
        br.close(); 

        ips = new FileInputStream(atcCodes); 
        //Cp1252 --> ANSI
        ipsr = new InputStreamReader(ips, "UTF-8");
        br = new BufferedReader(ipsr);

        while((ligne=br.readLine())!= null) {
            mActiveIngredient = pActiveIngredient.matcher(ligne);
            mActiveIngredient2 = pActiveIngredient2.matcher(ligne);
            if (mActiveIngredient.find()){
                translatAtcList.put(mActiveIngredient.group(1), mActiveIngredient.group(2));
            }
            if (mActiveIngredient2.find()){
                translatAtcList2.put(mActiveIngredient.group(1), mActiveIngredient.group(2));
            }
        }
        br.close();

        //Get the workbook instance for XLS file 
        XSSFWorkbook newMtcWorkbook = new XSSFWorkbook(newMTC);
        XSSFSheet newMtcSheet;
        Iterator<Row> newMtcRowIterator;
        Iterator<Cell> newMtcCellIterator;
        int newMtcCol;
        boolean newMtcColFound;
        ArrayList newMtcCodes;
        ArrayList newMtcCodes2;
        Row newMtcRow;
        Row newMtcRow2;
        Cell newMtcCell;

        //Get the sheet from the MTC workbook
        for(int i=0; i<newMtcWorkbook.getNumberOfSheets(); i++) {
            newMtcSheet = newMtcWorkbook.getSheetAt(i);

            //Get iterator to all the rows in current MTC sheet
            newMtcRowIterator = newMtcSheet.iterator();
            
            //And process the file matching or throw out the file that has no equivalent
            if(newMtcSheet.getSheetName().equals("VS21_IllnessesandDisorders")) {
                newMtcCol = 0;
                newMtcColFound = false;
                newMtcCodes = new ArrayList();

                //For each row, iterate through each columns
                //Get iterator to all cells of current row
                //In MTC
                while(newMtcRowIterator.hasNext()) {
                    newMtcRow = newMtcRowIterator.next();

                    if(newMtcColFound == false) {
                        newMtcCellIterator = newMtcRow.cellIterator();

                        while(newMtcCellIterator.hasNext()) {
                            newMtcCell = newMtcCellIterator.next();
                            if(newMtcCell.getCellType() == 1 && newMtcCell.getStringCellValue().equals("Code")){
                                newMtcCol = newMtcCell.getColumnIndex();
                                newMtcColFound = true;
                                break;
                            }
                        }
                    }
                    else {
                        newMtcRow.getCell(newMtcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                        newMtcCodes.add(newMtcRow.getCell(newMtcCol).getStringCellValue().trim());
                    }
                }

                for (int j=0;j<newMtcCodes.size();j++) {
                    csvW.write(newMtcCodes.get(j) + ";");

                    if (translatList.containsKey(newMtcCodes.get(j))) {
                        csvW.write(translatList.get(newMtcCodes.get(j)));
                    }
                    else {
                        m3 = p3.matcher((String)newMtcCodes.get(j));
                        if (m3.find() && translatList.containsKey(m3.group(1))) {
                            csvW.write(translatList.get(m3.group(1)));
                        }
                    }


                    /*for (int k=0; k<translationList.size(); k++) {
                        String frTrad = "";

                        if (translationList.get(k).trim().contains("<EN_OMS>"+newMtcCodes.get(j)+"</EN_OMS>")) {
                            Pattern p = Pattern.compile("<FR_OMS>(.+?)</FR_OMS>");
                            Matcher m = p.matcher(translationList.get(k).trim());
                            if (m.find()){
                                frTrad = m.group(1);
                                translationList.remove(k);
                            }
                            csvW.write(StringUtils.capitalize(frTrad));
                        }
                    }*/
                    csvW.write("\n");
                }
            }
            else if (newMtcSheet.getSheetName().equals("VS3_ActiveIngredient")) {
                newMtcCol = 0;
                newMtcColFound = false;
                newMtcCodes = new ArrayList();
                newMtcCodes2 = new ArrayList();

                //For each row, iterate through each columns
                //Get iterator to all cells of current row
                //In MTC
                while(newMtcRowIterator.hasNext()) {
                    newMtcRow = newMtcRowIterator.next();
                    newMtcRow2 = newMtcRow;

                    if(newMtcColFound == false) {
                        newMtcCellIterator = newMtcRow.cellIterator();

                        while(newMtcCellIterator.hasNext()) {
                            newMtcCell = newMtcCellIterator.next();
                            if(newMtcCell.getCellType() == 1 && newMtcCell.getStringCellValue().equals("English Display Name")){
                                newMtcCol = newMtcCell.getColumnIndex();
                                newMtcColFound = true;
                                break;
                            }
                        }
                    }
                    else {
                        newMtcRow.getCell(newMtcCol, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                        newMtcCodes.add(newMtcRow.getCell(newMtcCol).getStringCellValue().trim());
                        newMtcRow2.getCell(newMtcCol-1, Row.CREATE_NULL_AS_BLANK).setCellType(Cell.CELL_TYPE_STRING);
                        newMtcCodes2.add(newMtcRow.getCell(newMtcCol-1).getStringCellValue().trim());
                    }
                }

                for (int j=0;j<newMtcCodes.size();j++) {
                    csvW2.write(newMtcCodes2.get(j) + ";");
                    csvW2.write(newMtcCodes.get(j) + ";");

                    if (translatAtcList.containsKey(newMtcCodes.get(j))) {
                        csvW2.write(translatAtcList.get(newMtcCodes.get(j)));
                    }
                    else if (translatAtcList2.containsKey(newMtcCodes2.get(j))) {
                        csvW2.write(translatAtcList.get(newMtcCodes.get(j)));
                    }
                    else {
                        System.out.println(newMtcCodes.get(j));
                    }

                    csvW2.write("\n");
                }
            }
        }

        csvW.close();
        csvW2.close();
        newMTC.close();

    }
}
