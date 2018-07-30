import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dbc.DataBaseConnection;

public class ImportUser {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook hw;
		try {
			hw = new XSSFWorkbook(new FileInputStream("user1.xlsx"));
			XSSFSheet hsheet = hw.getSheetAt(0);
			importTotalData(hsheet);

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void importTotalData(XSSFSheet hsheet) {

		for (int i = 5; i < 17; i++) {
			XSSFRow hrow = hsheet.getRow(i);
			String month = "";
			if(i < 14) {
				month = "20170" + (i - 4);
			} else {
				month = "2017" + (i - 4);
			}
			
			
			XSSFCell cell = hrow.getCell(1);
			String zysr = getCellValue(cell);
			cell = hrow.getCell(2);
			String zysrHB = getCellValue(cell);

			cell = hrow.getCell(5);
			String ydcz = getCellValue(cell);
			cell = hrow.getCell(6);
			String ydczHB = getCellValue(cell);

			cell = hrow.getCell(9);
			String dyjz = getCellValue(cell);
			cell = hrow.getCell(10);
			String dyjzHB = getCellValue(cell);

			cell = hrow.getCell(13);
			String ljjz = getCellValue(cell);
			cell = hrow.getCell(14);
			String ljjzHB = getCellValue(cell);

			cell = hrow.getCell(17);
			String xfz = getCellValue(cell);
			cell = hrow.getCell(18);
			String xfzHB = getCellValue(cell);

			cell = hrow.getCell(21);
			String czls = getCellValue(cell);
			cell = hrow.getCell(22);
			String czlsHB = getCellValue(cell);

			cell = hrow.getCell(25);
			String arpu = getCellValue(cell);
			cell = hrow.getCell(26);
			String arpuHB = getCellValue(cell);
			importdata(zysr, zysrHB, ydcz, ydczHB, dyjz, dyjzHB, ljjz, ljjzHB, xfz, xfzHB, czls, czlsHB, arpu, arpuHB,
					month);

			if (i <= 9) {
				month = "2018" + (i - 4);
				cell = hrow.getCell(3);
				zysr = getCellValue(cell);
				cell = hrow.getCell(4);
				zysrHB = getCellValue(cell);

				cell = hrow.getCell(7);
				ydcz = getCellValue(cell);
				cell = hrow.getCell(8);
				ydczHB = getCellValue(cell);

				cell = hrow.getCell(11);
				dyjz = getCellValue(cell);
				cell = hrow.getCell(12);
				dyjzHB = getCellValue(cell);

				cell = hrow.getCell(15);
				ljjz = getCellValue(cell);
				cell = hrow.getCell(16);
				ljjzHB = getCellValue(cell);

				cell = hrow.getCell(19);
				xfz = getCellValue(cell);
				cell = hrow.getCell(20);
				xfzHB = getCellValue(cell);

				cell = hrow.getCell(23);
				czls = getCellValue(cell);
				cell = hrow.getCell(24);
				czlsHB = getCellValue(cell);

				cell = hrow.getCell(27);
				arpu = getCellValue(cell);
				cell = hrow.getCell(28);
				arpuHB = getCellValue(cell);
				importdata(zysr, zysrHB, ydcz, ydczHB, dyjz, dyjzHB, ljjz, ljjzHB, xfz, xfzHB, czls, czlsHB, arpu,
						arpuHB, month);
			}
		}

	}

	public static void importdata(String zysr, String zysrHB, String ydcz, String ydczHB, String dyjz, String dyjzHB,
			String ljjz, String ljjzHB, String xfz, String xfzHB, String czls, String czlsHB, String arpu,
			String arpuHB, String date) {
		try {
			double zysrDouble = Double.parseDouble(zysr);
			zysr = String.format("%.1f", zysrDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double zysrHBDouble = Double.parseDouble(zysrHB) * 100;
			zysrHB = String.format("%.1f", zysrHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			double ydczDouble = Double.parseDouble(ydcz);
			ydcz = String.format("%.1f", ydczDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double ydczHBDouble = Double.parseDouble(ydczHB) * 100;
			ydczHB = String.format("%.1f", ydczHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}


		try {
			double dyjzDouble = Double.parseDouble(dyjz);
			dyjz = String.format("%.1f", dyjzDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double dyjzHBDouble = Double.parseDouble(dyjzHB) * 100;
			dyjzHB = String.format("%.1f", dyjzHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			double ljjzDouble = Double.parseDouble(ljjz);
			ljjz = String.format("%.1f", ljjzDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double ljjzHBDouble = Double.parseDouble(ljjzHB) * 100;
			ljjzHB = String.format("%.1f", ljjzHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			double xfzDouble = Double.parseDouble(xfz);
			xfz = String.format("%.1f", xfzDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double xfzHBDouble = Double.parseDouble(xfzHB) * 100;
			xfzHB = String.format("%.1f", xfzHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			double czlsDouble = Double.parseDouble(czls);
			czls = String.format("%.1f", czlsDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double czlsHBDouble = Double.parseDouble(czlsHB) * 100;
			czlsHB = String.format("%.1f", czlsHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			double arpuDouble = Double.parseDouble(arpu);
			arpu = String.format("%.1f", arpuDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			double arpuHBDouble = Double.parseDouble(arpuHB) * 100;
			arpuHB = String.format("%.1f", arpuHBDouble);
		} catch (Exception e) {
			e.printStackTrace();
		}

		String sql = "insert into `provinceUserData`(zysr,zysrHB,ydcz,ydczHB,dyjz,dyjzHB,ljjz,ljjzHB,xfz,xfzHB,czls,czlsHB,arpu,arpuHB,month) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
		System.out.println(sql);
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1, zysr);
			prep.setString(2, zysrHB);
			prep.setString(3, ydcz);
			prep.setString(4, ydczHB);
			prep.setString(5, dyjz);
			prep.setString(6, dyjzHB);
			prep.setString(7, ljjz);
			prep.setString(8, ljjzHB);
			prep.setString(9, xfz);
			prep.setString(10, xfzHB);
			prep.setString(11, czls);
			prep.setString(12, czlsHB);
			prep.setString(13, arpu);
			prep.setString(14, arpuHB);
			prep.setString(15, date);
			prep.executeUpdate();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			try {
				if (prep != null)
					prep.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}

			DataBaseConnection.returnConnection(conn);

		}
	}

	public static String getCellValue(XSSFCell cell) {
		String value = "0.0";
		try {
			if (cell != null) {
				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_FORMULA:
					try {
						value = String.valueOf(cell.getNumericCellValue());
					} catch (IllegalStateException e) {
						try {
							value = String.valueOf(cell.getRichStringCellValue());
						} catch (Exception ex) {
							value = "error";
						}
					}
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					value = String.valueOf(cell.getNumericCellValue());
					break;
				case XSSFCell.CELL_TYPE_STRING:
					value = String.valueOf(cell.getRichStringCellValue());
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			value = "0.0";
		}
		if (value == null)
			value = "0.0";
		return value;
	}

}
