package org.brisskit.onyx.pdo2ss;

import java.io.File;
import java.util.Calendar;

import org.apache.commons.logging.Log;   
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.xmlbeans.XmlCalendar;

import org.brisskit.pdo.beans.ParamType;
import org.brisskit.pdo.beans.PatientDataDocument;
import org.brisskit.pdo.beans.PatientDataType;
import org.brisskit.pdo.beans.PatientType;
import org.brisskit.pdo.beans.PidSetDocument.PidSet;
import org.brisskit.pdo.beans.PidType;

public abstract class Pdo2Spreadsheet {
	
	public static final int APPOINTMENT_COLUMN_INDEX = 0 ;
	public static final int CENTER_ID_COLUMN_INDEX = 1 ;
	public static final int ENROLLMENT_ID_COLUMN_INDEX = 2 ;
	public static final int LAST_NAME_COLUMN_INDEX = 3 ;
	public static final int FIRST_NAME_COLUMN_INDEX = 4 ;
	public static final int GENDER_COLUMN_INDEX = 5 ;
	public static final int BIRTHDATE_COLUMN_INDEX = 6 ;
	
	public static final int DATA_SHEET_INDEX = 0 ;
	
	public static final int FIRST_DATA_ROW_INDEX = 2 ;
	
	public static final String ENROLLMENT_ID = "enrollment id" ;
	public static final String CENTER_ID = "center id" ;
	public static final String APPOINTMENT_DATE = "appointment date" ;
	public static final String LAST_NAME = "last name" ;
	public static final String FIRST_NAME = "first name" ;
	public static final String GENDER = "sex" ;
	public static final String BIRTHDATE = "birthdate" ;
	
	private static Log log = LogFactory.getLog( Pdo2Spreadsheet.class ) ;
	
	private static StringBuffer logIndent = null ;
	
	
	
	protected HSSFWorkbook workbook ;
	protected PatientDataType pdo ;
	protected String collectionCenterId ;
	protected int numberNewAppts ;
	protected int numberPreviousAppts ;
	protected HSSFCellStyle apptCellStyle ;
	protected HSSFCellStyle birthdateCellStyle ;
	
	protected String enrollmentId ;
	protected Calendar appointmentDate ;
	protected String centerId ;
	protected String lastName ;
	protected String firstName ;
	protected String gender ;
	protected Calendar birthdate ;
	
	@SuppressWarnings("unused")
	private Pdo2Spreadsheet() {}
	
	protected Pdo2Spreadsheet( HSSFWorkbook workbook, File pdoFile ) throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet(HSSFWorkbook,File)" ) ; } 
		this.workbook = workbook ;
		//
		// Acquire the PDO...
		this.pdo = acquirePDO( pdoFile ) ;
		//
		// Calculate the number of new participants;
		// ie; the number of new appointments we wish to add to the list...
		PidSet pids = this.pdo.getPidSet() ;
		this.numberNewAppts = pids.sizeOfPidArray() ;
		if( numberNewAppts == 0 ) {
			throw new P2SException( "No participant identifiers found within the PDO file." ) ;
		}
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet(HSSFWorkbook,File)" ) ; } 
			
	}
	
	protected Pdo2Spreadsheet( HSSFWorkbook workbook, File pdoFile, String collectionCenterId ) throws P2SException {
		this( workbook, pdoFile) ;
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet(HSSFWorkbook,File,String)" ) ; } 
		this.collectionCenterId = collectionCenterId ;
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet(HSSFWorkbook,File,String)" ) ; } 
			
	}
	
	public abstract HSSFWorkbook processList() throws P2SException ;
	
	private PatientDataType acquirePDO( File pdoFile ) throws P2SException {
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet.acquirePDO()" ) ; } 
		try {
			PatientDataDocument pdo = PatientDataDocument.Factory.parse( pdoFile ) ;
			return pdo.getPatientData() ;
		}
		catch( Exception ex ) {
			throw new P2SException( "Problem reading PDO file: " + pdoFile.getAbsolutePath(), ex ) ;
		}
		finally {
			if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet.acquirePDO()" ) ; }
		} 
	}
	
	protected void acquireCellStyles() {
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet.acquireCellStyles()" ) ; } 
		HSSFSheet dataSheet = this.workbook.getSheetAt( DATA_SHEET_INDEX ) ;		
		HSSFRow dataRow1 = dataSheet.getRow( FIRST_DATA_ROW_INDEX ) ;		
		this.apptCellStyle = dataRow1.getCell( APPOINTMENT_COLUMN_INDEX ).getCellStyle() ;
		this.birthdateCellStyle = dataRow1.getCell( BIRTHDATE_COLUMN_INDEX ).getCellStyle() ; 
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet.acquireCellStyles()" ) ; }
	}
	
	protected void acquirePatientData( PatientType patientDimension ) {
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet.acquirePatientData()" ) ; } 
		//
		// Initialize the transient patient data...
		this.enrollmentId = null ;
		this.appointmentDate = null ;
		this.centerId = null ;
		this.lastName = null ;
		this.firstName = null ;
		this.gender = null ;
		this.birthdate = null ;
		//
		// Get the patient dimension array...
		ParamType[] pta = patientDimension.getParamArray() ;
		//
		// Acquire the data...
		for( int i=0; i<pta.length; i++ ) {
			String name = pta[i].getName() ;
			this.enrollmentId = getEnrollmentId( patientDimension ) ;
			this.centerId = this.collectionCenterId ;
			if( name.equalsIgnoreCase( APPOINTMENT_DATE ) ) {
				this.appointmentDate = new XmlCalendar( pta[i].getStringValue() ) ;
			}
			else if( name.equalsIgnoreCase( LAST_NAME ) ) {
				this.lastName = pta[i].getStringValue() ;
			}
			else if( name.equalsIgnoreCase( FIRST_NAME ) ) {
				this.firstName = pta[i].getStringValue() ;
			}
			else if( name.equalsIgnoreCase( GENDER ) ) {
				this.gender = pta[i].getStringValue() ;
				if( this.gender.equalsIgnoreCase( "MALE" ) ) {
					this.gender = "M" ;
				}
				else if( this.gender.equalsIgnoreCase( "FEMALE" ) ) {
					this.gender = "F" ;
				}
			}
			else if( name.equalsIgnoreCase( BIRTHDATE ) ) {
				this.birthdate = new XmlCalendar( pta[i].getStringValue() ) ;
			}
		}
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet.acquirePatientData()" ) ; } 
	}
	
	protected void formatRow( HSSFRow apptRow ) {		
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet.formatRow()" ) ; }
		//
		// Create sufficient cells to hold an appointment...
		apptRow.createCell( 0 ) ;
		apptRow.createCell( 1 ) ;
		apptRow.createCell( 2 ) ;
		apptRow.createCell( 3 ) ;
		apptRow.createCell( 4 ) ;
		apptRow.createCell( 5 ) ;
		apptRow.createCell( 6 ) ;
		//
		// Set the style for the date columns. All the rest can accept the default style...
		apptRow.getCell( Pdo2Spreadsheet.APPOINTMENT_COLUMN_INDEX ).setCellStyle( this.apptCellStyle  ) ;
		apptRow.getCell( Pdo2Spreadsheet.BIRTHDATE_COLUMN_INDEX ).setCellStyle( this.birthdateCellStyle ) ;
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet.formatRow()" ) ; } 
	}
	
	protected void addDataToRow( HSSFRow apptRow ) {
		if( log.isTraceEnabled() ) { enterTrace( "Pdo2Spreadsheet.addDataToRow()" ) ; } 
		apptRow.getCell( Pdo2Spreadsheet.ENROLLMENT_ID_COLUMN_INDEX ).setCellValue( this.enrollmentId ) ;
		apptRow.getCell( Pdo2Spreadsheet.APPOINTMENT_COLUMN_INDEX ).setCellValue( this.appointmentDate ) ;
		apptRow.getCell( Pdo2Spreadsheet.CENTER_ID_COLUMN_INDEX ).setCellValue( this.centerId ) ;
		apptRow.getCell( Pdo2Spreadsheet.LAST_NAME_COLUMN_INDEX ).setCellValue( this.lastName ) ;
		apptRow.getCell( Pdo2Spreadsheet.FIRST_NAME_COLUMN_INDEX ).setCellValue( this.firstName ) ;
		apptRow.getCell( Pdo2Spreadsheet.GENDER_COLUMN_INDEX ).setCellValue( this.gender ) ;
		apptRow.getCell( Pdo2Spreadsheet.BIRTHDATE_COLUMN_INDEX ).setCellValue( this.birthdate ) ;
		if( log.isTraceEnabled() ) { exitTrace( "Pdo2Spreadsheet.addDataToRow()" ) ; } 
	}
	
	/**
	 * There are two circumstances to cover:<br/>
	 * (1) A web service styled PDO<br/>
	 * (2) An SQL direct styled PDO.<br/>
	 * The patient mapping within these two styles is different. For a web service style, 
	 * everything is plain and the patient id is given as the brisskit id. But for the 
	 * SQL direct style, the internal HIVE mapping is already present and we must navigate
	 * back to the source id to find the brisskit identifier.
	 * 
	 * @param patientDimension
	 * @return the brisskit identifier as the enrollment id.
	 */
	private String getEnrollmentId( PatientType patientDimension ) {
		String enrollmentId = null ;
		String source = patientDimension.getPatientId().getSource() ;
		//
		// If the source is not the internal HIVE, we can assume
		// the brisskit id is the patient id given within the patient dimension...
		if( !source.equalsIgnoreCase( "HIVE" ) ) {
			enrollmentId = patientDimension.getPatientId().getStringValue() ;
		}
		//
		// If the source is the internal HIVE, then we need to address 
		// the mapping within the PID set...
		else {
			String patientId = patientDimension.getPatientId().getStringValue() ;
			PidType[] pta = this.pdo.getPidSet().getPidArray() ;
			for( int i=0; i<pta.length; i++ ) {
				//
				// First find the matching patient id...
				if( pta[i].getPatientId().getStringValue().equals( patientId ) ) {
					//
					// Then take the first source mapping (there should not be more than one!)...
					enrollmentId = pta[i].getPatientMapIdArray(0).getStringValue() ;
					break ;
				}
			}
		}		
		return enrollmentId ;
	}
		
	/**
	 * Utility routine to enter a structured message in the trace log that the given method 
	 * has been entered. 
	 * 
	 * @param entry: the name of the method entered
	 */
	public static void enterTrace( String entry ) {
		log.trace( getIndent().toString() + "enter: " + entry ) ;
		indentPlus() ;
	}

    /**
     * Utility routine to enter a structured message in the trace log that the given method 
	 * has been exited. 
	 * 
     * @param entry: the name of the method exited
     */
    public static void exitTrace( String entry ) {
    	indentMinus() ;
		log.trace( getIndent().toString() + "exit : " + entry ) ;
	}
	
    /**
     * Utility method used to maintain the structured trace log.
     */
    public static void indentPlus() {
		getIndent().append( ' ' ) ;
	}
	
    /**
     * Utility method used to maintain the structured trace log.
     */
    public static void indentMinus() {
        if( logIndent.length() > 0 ) {
            getIndent().deleteCharAt( logIndent.length()-1 ) ;
        }
	}
	
    /**
     * Utility method used for indenting the structured trace log.
     */
    public static StringBuffer getIndent() {
	    if( logIndent == null ) {
	       logIndent = new StringBuffer() ;	
	    }
	    return logIndent ;	
	}
    
    @SuppressWarnings("unused")
	private static void resetIndent() {
        if( logIndent != null ) { 
            if( logIndent.length() > 0 ) {
               logIndent.delete( 0, logIndent.length() )  ;
            }
        }   
    }
	
    /**
     * Utility class covering all exceptions possibly thrown whilst producing the Spreadsheet list.
     * 
     * @author jl99
     *
     */
    public static class P2SException extends Exception {
    	
    	/**
		 * 
		 */
		private static final long serialVersionUID = 1L;

		public P2SException( String message ) {
    		super( message ) ;
    	}
    	
    	public P2SException( String message, Throwable cause ) {
    		super( message, cause ) ;
    	}
    	
    }

}
