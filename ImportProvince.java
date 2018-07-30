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

public class ImportProvince {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook hw;
		try {
			hw = new XSSFWorkbook(new FileInputStream("province.xlsx"));
			
			XSSFSheet hsheet = hw.getSheetAt(1);
			importTotalData(hsheet);
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void importTotalData(XSSFSheet hsheet) {
		int rows = hsheet.getPhysicalNumberOfRows();
		int flag = 0;
		int province = 0;
		for (int i = 3; i < 38; i++) {
			province = i - 3;
			XSSFRow hrow = hsheet.getRow(i);
			for (int j = 2; j <= 85; j++) {
				String date = "2017";
				int month = j / 7 + 1;
				if(month < 10) {
					date = date + "0" + month;
				} else {
					date = date + month;
				}
				
				XSSFCell zyszbcell = hrow.getCell(j);
				String zyszb = getCellValue(zyszbcell);
				j++;
				XSSFCell ARPUcell = hrow.getCell(j);
				String ARPU = getCellValue(ARPUcell);
				j++;
				XSSFCell cxywzdwcell = hrow.getCell(j);
				String cxywzdw = getCellValue(cxywzdwcell);
				j++;
				XSSFCell cxywzjkcell = hrow.getCell(j);
				String cxywzjk = getCellValue(cxywzjkcell);
				j++;
				XSSFCell ydjczjkcell = hrow.getCell(j);
				String ydjczjk = getCellValue(ydjczjkcell);
				j++;
				XSSFCell gwjczjkcell = hrow.getCell(j);
				String gwjczjk = getCellValue(gwjczjkcell);
				j++;
				XSSFCell zlsrzdwcell = hrow.getCell(j);
				String zlsrzdw = getCellValue(zlsrzdwcell);
				System.out.println(zyszb+ "  " +ARPU +"  " +cxywzdw + "  " + cxywzjk + "  " + ydjczjk+ "  " + gwjczjk+ "  " + zlsrzdw + "  " + date);
				importdata(zyszb,ARPU ,cxywzdw,cxywzjk,ydjczjk,gwjczjk,zlsrzdw,date,province + "");
				
			}
		}
	}
	
	public static void importdata(String zysrzb,String ARPU,String cxywzdw,String cxywzjk,String ydjczjk,String gwjczjk,String zlsrzdw,String date,String province) {
		try {
			double zysrzbDouble = Double.parseDouble(zysrzb) * 100;
			zysrzb = String.format("%.1f", zysrzbDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double cxywzdwDouble = Double.parseDouble(cxywzdw) * 100;
			cxywzdw = String.format("%.1f", cxywzdwDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double cxywzjkDouble = Double.parseDouble(cxywzjk) * 100;
			cxywzjk = String.format("%.1f", cxywzjkDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double ydjczjkDouble = Double.parseDouble(ydjczjk) * 100;
			ydjczjk = String.format("%.1f", ydjczjkDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double gwjczjkDouble = Double.parseDouble(gwjczjk) * 100;
			gwjczjk = String.format("%.1f", gwjczjkDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double zlsrzdwDouble = Double.parseDouble(zlsrzdw) * 100;
			zlsrzdw = String.format("%.1f", zlsrzdwDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double ARPUDouble = Double.parseDouble(ARPU);
			ARPU = String.format("%.1f", ARPUDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		String sql = "update `provincesubjecttotal_copy` set zysrzb = ?,ARPU = ?,cxywzdw=?,cxywzjk=?,ydjczjk=?,gwjczjk=?,zlsrzdw=? where subject = 0 and province = ? and month = ?";
		System.out.println(sql);
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1,zysrzb);
			prep.setString(2, ARPU);
			prep.setString(3, cxywzdw);
			prep.setString(4, cxywzjk);
			prep.setString(5, ydjczjk);
			prep.setString(6, gwjczjk);
			prep.setString(7, zlsrzdw);
			prep.setString(8, province);
			prep.setString(9, date);
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
		if(value == null)
			value = "0.0";
		return value;
	}


}
