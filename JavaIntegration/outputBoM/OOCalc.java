//***************************************************************************
// comment: Step 1: get the remote component context from the office
//          Step 2: open an empty calc document
//          Step 3: create cell styles
//          Step 4: get the sheet an insert some data
//          Step 5: apply the created cell syles
//          Step 6: insert a 3D Chart
//***************************************************************************

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;

import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;

import com.sun.star.frame.XComponentLoader;

import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XMultiComponentFactory;

import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XInterface;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.sheet.XSpreadsheetDocument;

import com.sun.star.style.XStyleFamiliesSupplier;

import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;

public class OOCalc {
	//oooooooooooooooooooooooooooStep 1oooooooooooooooooooooooooooooooooooooooooo
	// call UNO bootstrap method and get the remote component context form
	// the a running office (office will be started if necessary)
	//***************************************************************************
	public static XComponentContext bootstrapCalc()
	{
		XComponentContext xContext = null;
        
		// get the remote office component context
		try {
			xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
			System.out.println("Connected to a running office ...");
		} catch( Exception e) {
			e.printStackTrace(System.err);
			System.exit(1);
		}
		return xContext;
	}

    //oooooooooooooooooooooooooooStep 2oooooooooooooooooooooooooooooooooooooooooo
    // open an empty document. In this case it's a calc document.
    // For this purpose an instance of com.sun.star.frame.Desktop
    // is created. The desktop provides the XComponentLoader interface,
    // which is used to open the document via loadComponentFromURL
    //***************************************************************************
    public static XSpreadsheetDocument openCalc(XComponentContext xContext)
    {    
        //define variables
        XMultiComponentFactory xMCF = null;
        XComponentLoader xCLoader;
        XSpreadsheetDocument xSpreadSheetDoc = null;
        XComponent xComp = null;
        
        try {
            // get the servie manager rom the office
            xMCF = xContext.getServiceManager();

            // create a new instance of the the desktop
            Object oDesktop = xMCF.createInstanceWithContext(
                "com.sun.star.frame.Desktop", xContext );

            // query the desktop object for the XComponentLoader
            xCLoader = ( XComponentLoader ) UnoRuntime.queryInterface(
                XComponentLoader.class, oDesktop );
            
            PropertyValue [] szEmptyArgs = new PropertyValue [0];
            String strDoc = "private:factory/scalc";

            xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0, szEmptyArgs );
            xSpreadSheetDoc = (XSpreadsheetDocument) UnoRuntime.queryInterface(
                XSpreadsheetDocument.class, xComp);
            
        } catch(Exception e){            
            System.err.println(" Exception " + e);
            e.printStackTrace(System.err);
        }        
        
        return xSpreadSheetDoc;
    }
    
    //oooooooooooooooooooooooooooStep 3oooooooooooooooooooooooooooooooooooooooooo
    // create cell styles.
    // For this purpose get the StyleFamiliesSupplier and the the familiy
    // CellStyle. Create an instance of com.sun.star.style.CellStyle and
    // add it to the family. Now change some properties
    //***************************************************************************
    public static void createStyles(XSpreadsheetDocument xDoc)
    {    
        try {
        	XStyleFamiliesSupplier xSFS = (XStyleFamiliesSupplier)
            	UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, xDoc);
        	XNameAccess xSF = (XNameAccess) xSFS.getStyleFamilies();
        	XNameAccess xCS = (XNameAccess) UnoRuntime.queryInterface(
        			XNameAccess.class, xSF.getByName("CellStyles"));
        	XMultiServiceFactory oDocMSF = (XMultiServiceFactory)
        		UnoRuntime.queryInterface(XMultiServiceFactory.class, xDoc );
	        XNameContainer oStyleFamilyNameContainer = (XNameContainer)
	            UnoRuntime.queryInterface(
	            XNameContainer.class, xCS);
	        XInterface oInt1 = (XInterface) oDocMSF.createInstance(
	            "com.sun.star.style.CellStyle");
	        oStyleFamilyNameContainer.insertByName("My Style", oInt1);
	        XPropertySet oCPS1 = (XPropertySet)UnoRuntime.queryInterface(
	            XPropertySet.class, oInt1 );
	        oCPS1.setPropertyValue("IsCellBackgroundTransparent", new Boolean(false));
	        oCPS1.setPropertyValue("CellBackColor",new Integer(6710932));
	        oCPS1.setPropertyValue("CharColor",new Integer(16777215));
	        XInterface oInt2 = (XInterface) oDocMSF.createInstance(
	            "com.sun.star.style.CellStyle");
	        oStyleFamilyNameContainer.insertByName("My Style2", oInt2);
	        XPropertySet oCPS2 = (XPropertySet)UnoRuntime.queryInterface(
	            XPropertySet.class, oInt2 );
	        oCPS2.setPropertyValue("IsCellBackgroundTransparent", new Boolean(false));
	        oCPS2.setPropertyValue("CellBackColor",new Integer(13421823));
	    } catch (Exception e) {
	        e.printStackTrace(System.err);
	    }
    }
    
    //oooooooooooooooooooooooooooStep 4oooooooooooooooooooooooooooooooooooooooooo
    // get the sheet an insert some data.
    // Get the sheets from the document and then the first from this container.
    // Now some data can be inserted. For this purpose get a Cell via
    // getCellByPosition and insert into this cell via setValue() (for floats)
    // or setFormula() for formulas and Strings
    //***************************************************************************
    public static XSpreadsheet getSheet(XSpreadsheetDocument xDoc)
    {
	    XSpreadsheet xSheet=null;
	    
	    try {
	        System.out.println("Getting spreadsheet") ;
	        XSpreadsheets xSheets = xDoc.getSheets() ;
	        XIndexAccess oIndexSheets = (XIndexAccess) UnoRuntime.queryInterface(
	            XIndexAccess.class, xSheets);
	        xSheet = (XSpreadsheet) UnoRuntime.queryInterface(
	            XSpreadsheet.class, oIndexSheets.getByIndex(0));
	        
	    } catch (Exception e) {
	        System.out.println("Couldn't get Sheet " +e);
	        e.printStackTrace(System.err);
	    }
	    return xSheet;
    }

    public static void insertIntoCell(int CellX, int CellY, String theValue,
                                      XSpreadsheet TT1, String flag)
    {    
        XCell xCell = null;
        
        try {
            xCell = TT1.getCellByPosition(CellX, CellY);
        } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
            System.err.println("Could not get Cell");
            ex.printStackTrace(System.err);
        }

        if (flag.equals("V")) {
            xCell.setValue((new Float(theValue)).floatValue());
        } else {
            xCell.setFormula(theValue);
        }
        
    }
    
    public static void chgbColor( int x1, int y1, int x2, int y2,
                                  String template, XSpreadsheet TT )
    {
        XCellRange xCR = null;
        try {
            xCR = TT.getCellRangeByPosition(x1,y1,x2,y2);
        } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
            System.err.println("Could not get CellRange");
            ex.printStackTrace(System.err);
        }
        
        XPropertySet xCPS = (XPropertySet)UnoRuntime.queryInterface(
            XPropertySet.class, xCR );
        
        try {
            xCPS.setPropertyValue("CellStyle", template);
        } catch (Exception e) {
            System.err.println("Can't change colors chgbColor" + e);
            e.printStackTrace(System.err);
        }
    }
    
    public static void closeDocument(XSpreadsheetDocument xDocument)
    { 
        //check this supports xmodifyable 
        XCloseable xcloseable = 
           (XCloseable) UnoRuntime.queryInterface(XCloseable.class, xDocument); 
        if (xcloseable != null) { 
           try { 
              xcloseable.close(true); 
           } catch (CloseVetoException ex) { 
              System.out.println("close thrown a close veto exception"); 
              ex.printStackTrace(); 
           } 
           System.out.println(" ...closed successfully"); 
        } 
     } 
}
