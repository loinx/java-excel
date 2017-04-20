/**
 * 
 */
package com.examle.java.excel;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.examle.java.excel.model.Country;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.util.JSONPObject;

/**
 * @author loinx
 *
 */
public class ReadFile {
	/**
	 * Read excel file and return the list of the Object
	 * 
	 * @param file
	 * @param sheetName
	 * @return
	 */
	public static List<Country> parseObject(final InputStream is, String... sheetName) {
		if (Objects.isNull(is)) {
			throw new IllegalArgumentException("The excel file must not be null.");
		}
		Workbook workbook = null;
		try {
			List<Country> list = new ArrayList<>();
			workbook = new XSSFWorkbook(is);
			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = firstSheet.iterator();

			while (iterator.hasNext()) {
				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();
				if (nextRow.getRowNum() == 0) {
					continue;
				}
				Country country = new Country();
				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();
					int columnIndex = nextCell.getColumnIndex();
					Object object = getCellValue(nextCell);
					String cellValue = Objects.isNull(object) ? null : object.toString();
					if (Objects.isNull(cellValue) || "".equals(cellValue)) {
						System.err.println("Null");
						continue;
					}
					switch (columnIndex) {
					case 0:
						country.setCountryCode(cellValue);
						break;
					case 1:
						country.setCountryName(cellValue);
						break;
					default:
						break;
					}

				}
				list.add(country);
			}
			return list;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(workbook);
			IOUtils.closeQuietly(is);

		}

		return Collections.emptyList();
	}

	private static Object getCellValue(Cell cell) {
		final CellType cellType = cell.getCellTypeEnum();
		switch (cellType) {
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case NUMERIC:
			return cell.getNumericCellValue();
		default:
			// unknown the cell type
			return null;
		}
	}

	public static void main(String[] args) throws JsonProcessingException {
		List<Country> countries = ReadFile.parseObject(Thread.currentThread().getContextClassLoader()
				.getResourceAsStream("ISO 3166-1 alpha 3 Country Codes.xlsx"), "sheet1");

		System.err.println(countries);
		ObjectMapper mapper = new ObjectMapper();
		String json = mapper.writeValueAsString(countries);
		System.err.println(json);

	}
}
