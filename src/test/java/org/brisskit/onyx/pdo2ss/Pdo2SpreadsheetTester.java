package org.brisskit.onyx.pdo2ss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Pdo2SpreadsheetTester {
	
	public static final String SPREADSHEET_TEMPLATE_PATH = "/home/jeff/ws-brisskit/pdo2spreadsheet/src/test/resources/appts-template.xls" ;
	public static final String WS_PDO_PATH = "/home/jeff/ws-brisskit/pdo2spreadsheet/src/test/resources/onyx-4-sql-pdo.xml" ;
	
	public static final String EXISTING_LIST_PATH = "/home/jeff/ws-brisskit/pdo2spreadsheet/target/new-appointments-01.xls" ;
	public static final String TARGET_PATH = "/home/jeff/ws-brisskit/pdo2spreadsheet/target/new-appointments-01.xls" ;
	
	public static void main( String[] argv )  throws IOException {
		
		Pdo2SpreadsheetTester.augmentedListTest() ;
		
	}
	
	public Pdo2SpreadsheetTester() {}
	
	public static void freshListTest() throws IOException {
		try {

			FileInputStream templateFileStream = new FileInputStream( SPREADSHEET_TEMPLATE_PATH );
			HSSFWorkbook templateWB = new HSSFWorkbook( templateFileStream );
			templateFileStream.close();

			File pdoFile = new File( WS_PDO_PATH ) ;

			FreshListProcessor flp = new FreshListProcessor( templateWB, pdoFile, "brisskit001" ) ;

			HSSFWorkbook target = flp.processList() ;

			FileOutputStream fileOut = new FileOutputStream( TARGET_PATH ) ;
			target.write(fileOut);
			fileOut.close();
			System.out.println( "Done!" ) ;
		}
		catch ( Pdo2Spreadsheet.P2SException ex ) {
			ex.printStackTrace() ;
		}
	}
	
	public static void augmentedListTest() throws IOException {
		try {

			FileInputStream existingFileStream = new FileInputStream( EXISTING_LIST_PATH );
			HSSFWorkbook existingWB = new HSSFWorkbook( existingFileStream );
			existingFileStream.close();

			File pdoFile = new File( WS_PDO_PATH ) ;

			AugmentedListProcessor alp = new AugmentedListProcessor( existingWB, pdoFile, "brisskit001" ) ;

			HSSFWorkbook target = alp.processList() ;

			FileOutputStream fileOut = new FileOutputStream( TARGET_PATH ) ;
			target.write(fileOut);
			fileOut.close();
			System.out.println( "Done!" ) ;
		}
		catch ( Pdo2Spreadsheet.P2SException ex ) {
			ex.printStackTrace() ;
		}
	}

}
