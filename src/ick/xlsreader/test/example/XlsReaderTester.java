package ick.xlsreader.test.example;
import ick.xlsreader.test.util.ExcelReader;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class XlsReaderTester {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		ExcelReader reader = new ExcelReader("test.xls");

		reader.start();

		// Ư�� ��Ʈ, Ư�� ���� ��� �� ����
		List<String> list1 = reader.getRowCellsAt(1, 9);
		System.out.println("1�� ��Ʈ, 10���� ��� ��");
		for (int i = 0; i < list1.size(); i++) {
			System.out.println(list1.get(i));
		}

		// Ư�� ��Ʈ, Ư�� ���� ��� �� ����
		List<String> list2 = reader.getColCellsAt(1, 3);
		System.out.println("1�� ��Ʈ, 3���� ��� ��");
		for (int i = 0; i < list2.size(); i++) {
			if (!(list2.get(i).equals("")))
				System.out.println(list2.get(i));
		}

		// Ư�� ��Ʈ, ��, ������ ����
		String result1 = reader.getRowColCellsAt(1, 10, 3);
		String result2 = reader.getRowColCellsAt(2, 11, 3);
		String result3 = reader.getRowColCellsAt(3, 12, 3);
		String result4 = reader.getRowColCellsAt(4, 13, 3);
		String result5 = reader.getRowColCellsAt(5, 14, 3);
		
		System.out.println("1�� ��Ʈ, 11�� 3���� �� : " + result1);
		System.out.println("2�� ��Ʈ, 12�� 3���� �� : " + result2);
		System.out.println("3�� ��Ʈ, 13�� 3���� �� : " + result3);
		System.out.println("4�� ��Ʈ, 14�� 3���� �� : " + result4);
		System.out.println("5�� ��Ʈ, 15�� 3���� �� : " + result5);
	}

}
