import com.jacob.com.*;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.uno.XComponentContext;

import MGCPCB.*;
import MGCPCBAutomationLicensing.*;

public class OOBoM {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.runFinalizersOnExit(true);
		// open the PCB document
		ExpeditionPCBApplication pcbapp = new ExpeditionPCBApplication();
		pcbapp.IMGCPCBApplication_setVisible(true);
		System.out.println(pcbapp.IMGCPCBApplication_getFullName());
		
//		IMGCPCBDocument pcbdoc = pcbapp.IMGCPCBApplication_OpenDocument("/users/dube/projects/Vidar_WG/Pcb/Vidar.pcb");
		IMGCPCBDocument pcbdoc = pcbapp.IMGCPCBApplication_OpenDocument("c:\\Demonstrations\\Vidar_WG\\Pcb\\Vidar.pcb");
		
		System.out.println(pcbdoc.IMGCPCBDocument_getFullName());
			
		int key = pcbdoc.IMGCPCBDocument_Validate(0);
		MGCPCBAutomationLicensing.Application pLicenseServer = new MGCPCBAutomationLicensing.Application();
		int licenseToken = pLicenseServer.IApplication_GetToken(key);
		pcbdoc.IMGCPCBDocument_Validate(licenseToken);
		pcbapp.IMGCPCBApplication_LockServer(false);
		
		// open the spreadsheet document
		XComponentContext xContext = OOCalc.bootstrapCalc();
		XSpreadsheetDocument xDoc = OOCalc.openCalc(xContext);
		XSpreadsheet xSheet = OOCalc.getSheet(xDoc);
		
		// create the header
		OOCalc.insertIntoCell(0, 0, "REFDES", xSheet, "");
		OOCalc.insertIntoCell(1, 0, "PartNum", xSheet, "");
		OOCalc.insertIntoCell(2, 0, "X", xSheet, "");
		OOCalc.insertIntoCell(3, 0, "Y", xSheet, "");
		
		try {
		    IMGCPCBComponents pComps = pcbdoc.IMGCPCBDocument_getComponents(mwEPcbSelectionType.epcbSelectAll, mwEPcbComponentType.epcbCompAll, mwEPcbCelltype.epcbCelltypeAll, "*");
		    pComps.IMGCPCBComponents_Sort();
		    int count = pComps.IMGCPCBComponents_getCount();
				
		    for (int i = 1; i <= count; i++) {
				IMGCPCBComponent pComp = pComps.IMGCPCBComponents_getItem(new Variant(i));
				OOCalc.insertIntoCell(0, i, pComp.IMGCPCBComponent_getName(), xSheet, "");
				OOCalc.insertIntoCell(1, i, pComp.IMGCPCBComponent_getPartName(), xSheet, "");
				OOCalc.insertIntoCell(2, i, Double.toString(pComp.IMGCPCBComponent_getPositionX(mwEPcbUnit.epcbUnitMils)), xSheet, "V");
				OOCalc.insertIntoCell(3, i, Double.toString(pComp.IMGCPCBComponent_getPositionY(mwEPcbUnit.epcbUnitMils)), xSheet, "V");
			}
		} catch (Exception e) {
			System.out.println("Couldn't get component data; " +e);
			e.printStackTrace(System.err);
		}
		
		pcbapp.IMGCPCBApplication_UnlockServer(false);
		pcbdoc.IMGCPCBDocument_Close(false);
		pcbdoc = null;
		pcbapp.IMGCPCBApplication_Quit();
		pcbapp = null;
		
        System.exit(0);
	}
}
