import com.jacob.com.*;
//import com.sun.star.sheet.XSpreadsheet;
//import com.sun.star.sheet.XSpreadsheetDocument;
//import com.sun.star.uno.XComponentContext;

import MGCPCB.*;
import MGCPCBAutomationLicensing.*;
import java.io.*;

public class OOBoM {

	/**
	 * @param args
	 * @throws FileNotFoundException 
	 */
	public static void main(String[] args) throws FileNotFoundException {
		System.runFinalizersOnExit(true);
		// open the PCB document
		ExpeditionPCBApplication pcbapp = new ExpeditionPCBApplication();
		pcbapp.IMGCPCBApplication_setVisible(true);
		System.out.println(pcbapp.IMGCPCBApplication_getFullName());
		String curDir = System.getProperty("user.dir");
        //		IMGCPCBDocument pcbdoc = pcbapp.IMGCPCBApplication_OpenDocument("/users/dube/projects/Vidar_WG/Pcb/Vidar.pcb");
		IMGCPCBDocument pcbdoc = pcbapp.IMGCPCBApplication_OpenDocument(curDir+"\\Demonstrations\\Vidar_WG\\Pcb\\Vidar.pcb");
		PrintWriter fw = new PrintWriter(curDir+"\\Demonstrations\\Vidar_WG\\Pcb\\output\\Vidar_BOM.txt");



		
		
		
		System.out.println(pcbdoc.IMGCPCBDocument_getFullName());
			
		int key = pcbdoc.IMGCPCBDocument_Validate(0);
		MGCPCBAutomationLicensing.Application pLicenseServer = new MGCPCBAutomationLicensing.Application();
		int licenseToken = pLicenseServer.IApplication_GetToken(key);
		pcbdoc.IMGCPCBDocument_Validate(licenseToken);
		pcbapp.IMGCPCBApplication_LockServer(false);
		
		// open the spreadsheet document
//		XComponentContext xContext = OOCalc.bootstrapCalc();
//		XSpreadsheetDocument xDoc = OOCalc.openCalc(xContext);
//		XSpreadsheet xSheet = OOCalc.getSheet(xDoc);
//		
//		// create the header
//		OOCalc.insertIntoCell(0, 0, "REFDES", xSheet, "");
//		OOCalc.insertIntoCell(1, 0, "PartNum", xSheet, "");
//		OOCalc.insertIntoCell(2, 0, "X", xSheet, "");
//		OOCalc.insertIntoCell(3, 0, "Y", xSheet, "");
		
		try {
		    IMGCPCBComponents pComps = pcbdoc.IMGCPCBDocument_getComponents(mwEPcbSelectionType.epcbSelectAll, mwEPcbComponentType.epcbCompAll, mwEPcbCelltype.epcbCelltypeAll, "*");
		    pComps.IMGCPCBComponents_Sort();
		    int count = pComps.IMGCPCBComponents_getCount();
		    String str="  "+"REFDES"+"  "+"PartNum"+"  "+"X"+"   "+"Y"+"\r\n";
		    fw.write(str);
				
		    for (int i = 1; i <= count; i++) {
				IMGCPCBComponent pComp = pComps.IMGCPCBComponents_getItem(new Variant(i));
				str=i+"  "+pComp.IMGCPCBComponent_getName()+"  "+pComp.IMGCPCBComponent_getPartName()+"  "+Double.toString(pComp.IMGCPCBComponent_getPositionX(mwEPcbUnit.epcbUnitMils))+"   "+Double.toString(pComp.IMGCPCBComponent_getPositionY(mwEPcbUnit.epcbUnitMils))+"\r\n";
			    fw.write(str);
		}} catch (Exception e) {
			System.out.println("Couldn't get component data; " +e);
			e.printStackTrace(System.err);
		}
	    fw.close();
		
		pcbapp.IMGCPCBApplication_UnlockServer(false);
		pcbdoc.IMGCPCBDocument_Close(false);
		pcbdoc = null;
		pcbapp.IMGCPCBApplication_Quit();
		pcbapp = null;
		
        System.exit(0);
	}
}
