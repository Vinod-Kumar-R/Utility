package Excel.Reader;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class ExcelTest  
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
	
	@Test
	public void ReadrowExcel() {
		
		try {
			ExcelReader er = new ExcelReader("D:\\Vinod\\ExcelTest.xlsx");
			System.out.println(er.Rowcout(0));
			System.out.println(er.ColumnCout(3));
			
			/*for(int i = 0;i<er.Rowcout(0);i++) {
				for(int j= 0;j<er.ColumnCout(i);j++) {
					System.out.println(er.GetCellData(i,j));
				}
			}*/
			
			for(int i = 1;i<er.Rowcout(0);i++) {
				Map<Integer,String> RowData = er.GetRowData(i);
				
				for(int j= 0;j<RowData.size();j++) {
					System.out.println(RowData.get(j));
				}
			}
			
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
    
}
