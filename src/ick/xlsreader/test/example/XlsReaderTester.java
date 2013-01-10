package ick.xlsreader.test.example;
import ick.xlsreader.test.util.ExcelReader;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class XlsReaderTester {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		ExcelReader reader = new ExcelReader("test.xls");

		reader.start();

		// 특정 시트, 특정 행의 모든 값 추출
		List<String> list1 = reader.getRowCellsAt(1, 9);
		System.out.println("1번 시트, 10행의 모든 값");
		for (int i = 0; i < list1.size(); i++) {
			System.out.println(list1.get(i));
		}

		// 특정 시트, 특정 열의 모든 값 추출
		List<String> list2 = reader.getColCellsAt(1, 3);
		System.out.println("1번 시트, 3열의 모든 값");
		for (int i = 0; i < list2.size(); i++) {
			if (!(list2.get(i).equals("")))
				System.out.println(list2.get(i));
		}

		// 특정 시트, 행, 열에서 추출
		String result1 = reader.getRowColCellsAt(1, 10, 3);
		String result2 = reader.getRowColCellsAt(2, 11, 3);
		String result3 = reader.getRowColCellsAt(3, 12, 3);
		String result4 = reader.getRowColCellsAt(4, 13, 3);
		String result5 = reader.getRowColCellsAt(5, 14, 3);
		
		System.out.println("1번 시트, 11행 3열의 값 : " + result1);
		System.out.println("2번 시트, 12행 3열의 값 : " + result2);
		System.out.println("3번 시트, 13행 3열의 값 : " + result3);
		System.out.println("4번 시트, 14행 3열의 값 : " + result4);
		System.out.println("5번 시트, 15행 3열의 값 : " + result5);
	}

}
