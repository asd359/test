package ExRW;

import java.io.File;

import edu.npu.fastexcel.BIFFSetting;
import edu.npu.fastexcel.ExcelException;
import edu.npu.fastexcel.FastExcel;
import edu.npu.fastexcel.Sheet;
import edu.npu.fastexcel.Workbook;
public class AllRW {

	
	public void AllRw(){
		
		try {
			testDump();
		} catch (ExcelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void testWrite() throws ExcelException{
		File f=new File("./test2.xls");
		Workbook wb=FastExcel.createWriteableWorkbook(f);
		wb.open();
		Sheet sheet=wb.addSheet("sheetA");
		sheet.setCell(1, 2, "some string");
		
		for (int i = 0; i < 50000; i++) {
//			System.out.print(i+"#\n"+s.getRow(i));
			for (int j = 0; j < 20; j++) {
				sheet.setCell(i, j, "678676");
			}
//			System.out.println();
			}
		
		
		wb.close();
	}
	
	
	public void testDump() throws ExcelException {
		Workbook workBook;
		workBook = FastExcel.createReadableWorkbook(new File("test.xlsx"));
		workBook.setSSTType(BIFFSetting.SST_TYPE_DEFAULT);//memory storage
		workBook.open();
		Sheet s;
		s = workBook.getSheet(0);
		System.out.println("SHEET:"+s);
		for (int i = s.getFirstRow(); i < s.getLastRow(); i++) {
		System.out.print(i+"#\n"+s.getRow(i));
		for (int j = s.getFirstColumn(); j < s.getLastColumn(); j++) {
			System.out.print(","+s.getCell(i, j));
		}
//		System.out.println();
		}
		workBook.close();
	}
		
	
	
	
	public void testStreamWrite() throws Exception{
		File f=new File("write1.xls");
		Workbook wb=FastExcel.createWriteableWorkbook(f);
		wb.open();
		Sheet sheet=wb.addStreamSheet("SheetA");
		sheet.addRow(new String[]{"aaa","bbb","aaa","bbb",
				"aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb","aaa","bbb"});
		
		for (int i = 0; i < 50000; i++) {
			sheet.addRow(new String[]{"aaa","bbb"});
		}
		
		
		wb.close();
	}	
	
}
