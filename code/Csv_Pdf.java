package email.code;

import java.io.File;

import com.aspose.cells.CellArea;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.PageOrientationType;
import com.aspose.cells.PrintingPageType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.SheetRender;
import com.aspose.cells.TxtLoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class Csv_Pdf {

	static File csv(String str4) throws Exception {
		TxtLoadOptions loadOptions = new TxtLoadOptions(com.aspose.cells.LoadFormat.CSV);
		loadOptions.setCheckExcelRestriction(false);
		Workbook workbook = new Workbook(str4, loadOptions);

		Worksheet worksheet = workbook.getWorksheets().get(0);
		worksheet.autoFitColumns();
		worksheet.getPageSetup().setPrintGridlines(true);
		worksheet.getPageSetup().setOrientation(PageOrientationType.LANDSCAPE);
		ImageOrPrintOptions printoption = new ImageOrPrintOptions();
		printoption.setPrintingPage(PrintingPageType.DEFAULT);
		SheetRender sr = new SheetRender(worksheet, printoption);
		int pageCount = sr.getPageCount();
		System.out.println(pageCount);
		CellArea[] area = worksheet.getPrintingPageBreaks(printoption);
		System.out.println(area.length);
		int strow = 0;
		int stcol = 0;
		for (int i = 0; i < area.length; i++) {
			// Get each page range/area.
			strow = area[i].StartRow;
			stcol = area[i].StartColumn;

			// Add Watermark
			com.aspose.cells.Shape wordart = worksheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1, "",
					"Arial Black", 50, false, true, strow + 18, 8, stcol + 1, 0, 130, 800);
			// Get the fill format of the word art
			FillFormat wordArtFormat = wordart.getFill();
			// Set the color
			wordArtFormat.setOneColorGradient(com.aspose.cells.Color.getRed(), 0.2, GradientStyleType.HORIZONTAL, 2);
			// Set the transparency
			wordArtFormat.setTransparency(0.9);

			// Make the line invisible
			LineFormat lineFormat = wordart.getLine();
			lineFormat.setWeight(0.0);
		}

		workbook.save(str4.replace("csv", "") + "pdf", SaveFormat.PDF);
		File f = new File(str4.replace("csv", "") + "pdf");
		System.out.println(f.getAbsolutePath());
		System.out.println("Done");
		return f;

	}

}
