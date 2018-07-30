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

public class Import {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook hw;
		try {
			hw = new XSSFWorkbook(new FileInputStream("province.xlsx"));
			for(int i = 1;i <= 12;i++) {
				XSSFSheet hsheet = hw.getSheetAt(i + 2);
				if(hsheet != null) {
					String date = "2017";
					if(i < 10) {
						date = date + "0" + i;
					} else {
						date = date + i;
					}
					importTotalData(hsheet,date);
				}
			}
			
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void importTotalData(XSSFSheet hsheet,String date) {
		int rows = hsheet.getPhysicalNumberOfRows();
		int flag = 0;
		int province = 0;
		String subject = "0";
		for (int i = 2; i < 37; i++) {
			if(flag == 1)
				break;
			province = i - 2;
			XSSFRow hrow = hsheet.getRow(i);
			for (int j = 1; j <= 24; j++) {
				XSSFCell incomecell = hrow.getCell(j);
				String income = getCellValue(incomecell);
				if(income.equals("")) {
					income = "0.0";
				}
				j++;
				XSSFCell tbcell = hrow.getCell(j);
				String tongbizengfu = getCellValue(tbcell);
				if(tongbizengfu.equals("")) {
					tongbizengfu = "0.0";
				}
				
				if(j < 3) {
					subject = "0";
				} else if(j < 5) {
					subject = "2";
				} else if(j < 7) {
					subject = "3";
				} else if(j < 9) {
					subject = "1";
				} else if(j < 11) {
					subject = "11";
				} else if(j < 13) {
					subject = "12";
				} else if(j < 15) {
					subject = "14";
				} else if(j < 17) {
					subject = "15";
				} else if(j < 19) {
					subject = "13";
				} else if(j < 21) {
					subject = "31";
				} else if(j < 23) {
					subject = "32";
				} else if(j < 25) {
					subject = "33";
				}
				//System.out.println(income+ "  " +tongbizengfu +"  " +subject + "  " + province);
				importdata(subject,province + "",tongbizengfu,income,date);
				
			}
		}
	}
	
	public static void importdata(String subject,String province,String tongbiIncrease,String income,String month) {
		try {
			double tongbizengfuDouble = Double.parseDouble(tongbiIncrease) * 100;
			tongbiIncrease = String.format("%.1f", tongbizengfuDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double incomeDouble = Double.parseDouble(income);
			income = String.format("%.1f", incomeDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		String sql = "insert into `provincesubjecttotal_copy`(`province`,`subject`,`income`,`tongbiIncrease`,`month`) values(?,?,?,?,?)";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1,province);
			prep.setString(2, subject);
			prep.setString(3, income);
			prep.setString(4, tongbiIncrease);
			prep.setString(5, month);
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
		String value = "";
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
			value = "";
		}
		if(value == null)
			value = "";
		return value;
	}


}
