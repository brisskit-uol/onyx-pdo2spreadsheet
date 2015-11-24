package org.brisskit.onyx.pdo2ss;

import java.io.File;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.brisskit.onyx.pdo2ss.Pdo2Spreadsheet.P2SException;
import org.brisskit.pdo.beans.PatientType;

public class FreshListProcessor extends Pdo2Spreadsheet {
	
	private static Log log = LogFactory.getLog( FreshListProcessor.class ) ;

	public FreshListProcessor( HSSFWorkbook template, File pdoFile ) throws P2SException {
		super( template, pdoFile ) ;
		init() ;
	} 
	
	public FreshListProcessor( HSSFWorkbook template, File pdoFile, String collectionCenterId ) throws P2SException {
		super( template, pdoFile, collectionCenterId ) ; 
		init() ;
	}
	
	private void init() throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "FreshListProcessor.init()" ) ; } 
		int numberDataRows = workbook.getSheetAt( DATA_SHEET_INDEX ).getLastRowNum() - FIRST_DATA_ROW_INDEX + 1;
		if( numberDataRows == 0) {
			throw new P2SException( "The template workbook is missing the template data row required." ) ;
		}
		else if( numberDataRows > 1 ) {
			throw new P2SException( "The template workbook contains more than the one template data row required." ) ;
		}
		acquireCellStyles() ;
		this.numberPreviousAppts = 0 ;
		if( log.isTraceEnabled() ) { exitTrace( "FreshListProcessor.init()" ) ; }
	}
	
	@Override
	public HSSFWorkbook processList() throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "FreshListProcessor.processList()" ) ; } 
		HSSFSheet dataSheet = this.workbook.getSheetAt( DATA_SHEET_INDEX ) ;
		PatientType[] pta = this.pdo.getPatientSet().getPatientArray() ;
		//
		// Add an appointment for each new participant...
		for( int i=0; i<this.numberNewAppts; i++ ) {
			HSSFRow apptRow ;
			//
			// If this is a fresh sheet, there is a template row already in existence which
			// we have used to acquire the relevant cell styles.
			// We re-use this template row for the first data row, and create all the rest...
			if( i == 0 ) {
				apptRow = dataSheet.getRow( this.numberPreviousAppts + FIRST_DATA_ROW_INDEX + i ) ;
			}
			else {
				apptRow = dataSheet.createRow( this.numberPreviousAppts + FIRST_DATA_ROW_INDEX + i ) ;
			}
			formatRow( apptRow ) ;
			acquirePatientData( pta[i] ) ;
			addDataToRow( apptRow ) ;
		}		
		if( log.isTraceEnabled() ) { exitTrace( "FreshListProcessor.processList()" ) ; } 
		return this.workbook ;
	}

}
