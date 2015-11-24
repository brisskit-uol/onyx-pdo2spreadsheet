package org.brisskit.onyx.pdo2ss;

import java.io.File;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.brisskit.pdo.beans.PatientType;

public class AugmentedListProcessor extends Pdo2Spreadsheet {
	
	private static Log log = LogFactory.getLog( AugmentedListProcessor.class ) ;
	
	public AugmentedListProcessor( HSSFWorkbook workbook, File pdoFile ) throws P2SException {
		super( workbook, pdoFile ) ;
		init() ;
	}
	
	public AugmentedListProcessor( HSSFWorkbook workbook, File pdoFile, String collectionCenterId ) throws P2SException {
		super( workbook, pdoFile, collectionCenterId ) ;
		init() ;
	}
	
	private void init() throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "AugmentedListProcessor.init()" ) ; }
		this.numberPreviousAppts = workbook.getSheetAt( DATA_SHEET_INDEX ).getLastRowNum() - FIRST_DATA_ROW_INDEX + 1 ;
		if( this.numberPreviousAppts == 0) {
			throw new P2SException( "The previous list contains no appointments." ) ;
		}
		acquireCellStyles() ;
		if( log.isTraceEnabled() ) { exitTrace( "AugmentedListProcessor.init()" ) ; }
	}

	@Override
	public HSSFWorkbook processList() throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "AugmentedListProcessor.processList()" ) ; } 
		HSSFSheet dataSheet = this.workbook.getSheetAt( DATA_SHEET_INDEX ) ;
		PatientType[] pta = this.pdo.getPatientSet().getPatientArray() ;
		//
		// Add an appointment for each new participant...
		for( int i=0; i<this.numberNewAppts; i++ ) {
			HSSFRow apptRow = dataSheet.createRow( this.numberPreviousAppts + FIRST_DATA_ROW_INDEX + i ) ;
			formatRow( apptRow ) ;
			acquirePatientData( pta[i] ) ;
			addDataToRow( apptRow ) ;
		}		
		if( log.isTraceEnabled() ) { exitTrace( "AugmentedListProcessor.processList()" ) ; } 
		return this.workbook ;
	}
	


}
