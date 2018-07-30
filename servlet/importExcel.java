package servlet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import dbc.DataBaseConnection;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Servlet implementation class importExcel
 */
public class importExcel extends HttpServlet {
	private static final long serialVersionUID = 1L;

	/**
	 * @see HttpServlet#HttpServlet()
	 */
	public static String path = "";

	public importExcel() {
		super();
		// TODO Auto-generated constructor stub
	}

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		boolean isMultipart = ServletFileUpload.isMultipartContent(request);
		if (!isMultipart) {
			return;
		}

		// Create a factory for disk-based file items
		DiskFileItemFactory factory = new DiskFileItemFactory();

		// Set factory constraintsx
		factory.setSizeThreshold(1024); // yourMaxMemorySize

		File dir = new File("C:/tmp/");

		if (!dir.exists()) {
			dir.mkdirs();
		}
		factory.setRepository(dir); // yourTempDirectory

		// Create a new file upload handler
		ServletFileUpload upload = new ServletFileUpload(factory);
		upload.setHeaderEncoding("UTF-8");
		// setProgressListener
		// upload.setProgressListener(new FileUploadProgressListener());

		// Parse the request
		List items = null;
		try {
			items = upload.parseRequest(request);
		} catch (FileUploadException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		// Process the uploaded items
		Iterator iter = items.iterator();
		String date = "";
		while (iter.hasNext()) {
			FileItem item = (FileItem) iter.next();

			// 整个表单的所有域都会被解析，要先判断一下是普通表单域还是文件上传域
			if (item.isFormField()) {
				String name = item.getFieldName();
				String value = item.getString("utf-8");
				System.out.println(name + ":" + value);

				if (name.equals("datepicker")) {
					date = value;
				}
			} else {
				String fieldName = item.getFieldName();
				String fileName = item.getName();
				if (fileName.indexOf("\\") > 0) {
					fileName = fileName.substring(fileName.lastIndexOf("\\") + 1);
				}
				String contentType = item.getContentType();
				boolean isInMem = item.isInMemory();
				long sizeInBytes = item.getSize();
				System.out.println(fieldName + ":" + fileName);
				System.out.println("类型：" + contentType);
				System.out.println("是否在内在：" + isInMem);
				System.out.println("文件大小" + sizeInBytes);
				if (fileName.equals(""))
					continue;
				File filetmp = new File("C:/tmp/" + fileName);
				try {
					item.write(filetmp);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				if (fieldName.equals("excelfile")) {
					excelImport(date, filetmp);
					filetmp.delete();
				}
			}
		}
		response.sendRedirect("main");
	}

	public void excelImport(String date, File file) {
		XSSFWorkbook hw;
		try {
			hw = new XSSFWorkbook(new FileInputStream(file));
			XSSFSheet hsheet = hw.getSheetAt(2);
			if(hsheet != null) {
				importTotalData(hsheet,date);
			}
			XSSFSheet provhsheet = hw.getSheetAt(3);
			if(provhsheet != null) {
				importProvData(provhsheet,date);
			}
			XSSFSheet ictprovhsheet = hw.getSheetAt(4);
			if(ictprovhsheet != null) {
				importICTProvData(ictprovhsheet,date);
			}
			XSSFSheet gwprovhsheet = hw.getSheetAt(6);
			if(gwprovhsheet != null) {
				importGWProvData(gwprovhsheet,date);
			}
			XSSFSheet ywprovhsheet = hw.getSheetAt(5);
			if(ywprovhsheet != null) {
				importYWProvData(ywprovhsheet,date);
			}
			
			XSSFSheet ydyhsheet = hw.getSheetAt(21);
			if(ydyhsheet != null) {
				importYDYHData(ydyhsheet,date);
			}
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	public void importYDYHData(XSSFSheet hsheet,String date) {
		int province = 0;
		XSSFRow xfzrow = hsheet.getRow(9);
		XSSFRow czlsrow = hsheet.getRow(22);
		for(int i = 0;i < 33;i++) {
			if(i >= 1) {
				province = i + 2;
			}
			XSSFCell xfzcell = xfzrow.getCell(i * 5 + 1);
			String xfz = this.getCellValue(xfzcell);
			if(xfz.equals("")) {
				xfz = "0.0";
			}
			XSSFCell xfzHBcell = xfzrow.getCell(i * 5 + 3);
			String xfzHB = this.getCellValue(xfzHBcell);
			if(xfzHB.equals("")) {
				xfzHB = "0.0";
			}
			
			XSSFCell czlscell = czlsrow.getCell(i * 5 + 1);
			String czls = this.getCellValue(czlscell);
			if(czls.equals("")) {
				czls = "0.0";
			}
			XSSFCell czlsHBcell = czlsrow.getCell(i * 5 + 3);
			String czlsHB = this.getCellValue(czlsHBcell);
			if(czlsHB.equals("")) {
				czlsHB = "0.0";
			}
			updateYWCZ(date,province + "",xfz,xfzHB,czls,czlsHB);
		}
	}
	
	public void updateYWCZ(String date,String province,String xfz,String xfzHB,String czls,String czlsHB) {
		try {
			double xfzDouble = Double.parseDouble(xfz);
			double xfzHBDouble = Double.parseDouble(xfzHB) * 100;
			
			double czlsDouble = Double.parseDouble(czls);
			double czlsHBDouble = Double.parseDouble(czlsHB) * 100;
			xfz = String.format("%.1f", xfzDouble / 10000);
			xfzHB = String.format("%.1f", xfzHBDouble);
			
			czls = String.format("%.1f", czlsDouble / 10000);
			czlsHB = String.format("%.1f", czlsHBDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		String sql = "update provinceUserData set xfz = ?,xfzHB=?,czls=?,czlsHB=? where province = ? and month = ?";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1,xfz);
			prep.setString(2, xfzHB);
			prep.setString(3, czls);
			prep.setString(4, czlsHB);
			prep.setString(5, province);
			prep.setString(6, date);
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
	
	public void importYWProvData(XSSFSheet hsheet,String date) {
		int province = 0;
		for(int i = 4;i < 39;i++) {
			province = i - 4;
			XSSFRow hrow = hsheet.getRow(i);
			XSSFCell incomecell = hrow.getCell(8);
			String income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			XSSFCell num1cell = hrow.getCell(12);
			String num1 = this.getCellValue(num1cell);
			if(num1.equals("")) {
				num1 = "0.0";
			}
			XSSFCell num2cell = hrow.getCell(13);
			String num2 = this.getCellValue(num2cell);
			if(num2.equals("")) {
				num2 = "0.0";
			}
			
			XSSFCell huanbicell = hrow.getCell(15);
			String huanbi = this.getCellValue(huanbicell);
			if(huanbi.equals("")) {
				huanbi = "0.0";
			}
			insertYWCZ(date,province + "",num1,huanbi);
			updateARPU(income,num1,num2,date,province + "");
		}
	}
	public void insertYWCZ(String date,String province,String ywcz,String ywczHB) {
		try {
			double ywczDouble = Double.parseDouble(ywcz);
			double ywczHBDouble = Double.parseDouble(ywczHB) * 100;
			ywcz = String.format("%.1f", ywczDouble);
			ywczHB = String.format("%.1f", ywczHBDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		String sql = "insert into provinceUserData(month,province,ydcz,ydczHB) values(?,?,?,?)";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1,date);
			prep.setString(2, province);
			prep.setString(3, ywcz);
			prep.setString(4, ywczHB);
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
	public void updateARPU(String income,String num1,String num2,String date,String province) {
		String arpu = "0.0";
		double arpuDouble = 0;
		try {
			double num1Double = Double.parseDouble(num1);
			double num2Double = Double.parseDouble(num2);
			double incomeDouble = Double.parseDouble(income);
			if(num1Double + num2Double > 0) {
				arpuDouble = incomeDouble / ((num1Double + num2Double) / 2);
				arpu = String.format("%.1f", arpuDouble);
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
		String sql = "update `provincesubjecttotal` set ARPU = ? where month = ? and province = ? and subject = 0";
		String sql2 = "update nationtotal set ARPU = ? where month = ?";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1,arpu);
			prep.setString(2, date);
			prep.setString(3, province);
			prep.executeUpdate();
			if(province.equals("0")) {
				prep.close();
				prep = conn.prepareStatement(sql2);
				prep.setString(1,arpu);
				prep.setString(2, date);
				prep.executeUpdate();
			}
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
	public void importGWProvData(XSSFSheet hsheet,String date) {
		int province = 0;
		for(int i = 4;i < 39;i++) {
			province = i - 4;
			XSSFRow hrow = hsheet.getRow(i);
			XSSFCell hcell = hrow.getCell(2);
			String tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			XSSFCell incomecell = hrow.getCell(8);
			String income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			
			insertProv(31,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(11);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(17);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(32,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(20);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(23);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(33,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
		}
	}
	
	public void importICTProvData(XSSFSheet hsheet,String date) {
		int province = 0;
		for(int i = 4;i < 39;i++) {
			province = i - 4;
			XSSFRow hrow = hsheet.getRow(i);
			XSSFCell hcell = hrow.getCell(4);
			String tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			XSSFCell incomecell = hrow.getCell(7);
			String income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			
			insertProv(13,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(12);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(15);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(11,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(20);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(23);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(12,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(28);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(31);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(14,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(36);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(39);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(15,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
		}
	}
	
	public void importProvData(XSSFSheet hsheet,String date) {
		int province = 0;
		for(int i = 4;i < 39;i++) {
			province = i - 4;
			XSSFRow hrow = hsheet.getRow(i);
			XSSFCell hcell = hrow.getCell(5);
			String tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			XSSFCell incomecell = hrow.getCell(14);
			String income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			
			/////
			XSSFCell zlsrzdwcell = hrow.getCell(13);
			String zlsrzdw = this.getCellValue(zlsrzdwcell);
			if(zlsrzdw.equals("")) {
				zlsrzdw = "0.0";
			}
			
			XSSFCell zysrzbcell = hrow.getCell(12);
			String zysrzb = this.getCellValue(zysrzbcell);
			if(zysrzb.equals("")) {
				zysrzb = "0.0";
			}
			XSSFCell cxywzjkcell = hrow.getCell(25);
			String cxywzjk = this.getCellValue(cxywzjkcell);
			if(cxywzjk.equals("")) {
				cxywzjk = "0.0";
			}
			XSSFCell cxywzdwcell = hrow.getCell(26);
			String cxywzdw = this.getCellValue(cxywzdwcell);
			if(cxywzdw.equals("")) {
				cxywzdw = "0.0";
			}
			XSSFCell ywjczjkcell = hrow.getCell(40);
			String ywjczjk = this.getCellValue(ywjczjkcell);
			if(ywjczjk.equals("")) {
				ywjczjk = "0.0";
			}
			XSSFCell gwjczjkcell = hrow.getCell(55);
			String gwjczjk = this.getCellValue(gwjczjkcell);
			if(gwjczjk.equals("")) {
				gwjczjk = "0.0";
			}
			String ARPU = "0.0";
			insertProv(0,province,date,tongbizengfu,income,zysrzb,ARPU,cxywzdw,cxywzjk,ywjczjk,gwjczjk,zlsrzdw);
			if(province == 0) {
				updateNation(zysrzb,ARPU,cxywzdw,cxywzjk,ywjczjk,gwjczjk,zlsrzdw,date);
			}
			hcell = hrow.getCell(19);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(22);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(1,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(31);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(37);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(2,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
			
			hcell = hrow.getCell(46);
			tongbizengfu = this.getCellValue(hcell);
			if(tongbizengfu.equals("")) {
				tongbizengfu = "0.0";
			}
			incomecell = hrow.getCell(52);
			income = this.getCellValue(incomecell);
			if(income.equals("")) {
				income = "0.0";
			}
			insertProv(3,province,date,tongbizengfu,income,"0.0","0.0","0.0","0.0","0.0","0.0","0.0");
		}
	}
	
	public int updateNation(String zysrzb,String ARPU,String cxywzdw,String cxywzjk,String ydjczjk,String gwjczjk,String zlsrzdw,String date) {
		int result = 1;
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
		String sql = "update `nationtotal` set zysrzb = ?,ARPU = ?,cxywzdw=?,cxywzjk=?,ydjczjk=?,gwjczjk=?,zlsrzdw=? where month = ?";
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
			prep.setString(8, date);
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
		return result;
	}
	public int insertProv(int subject,int province,String date,String tongbizengfu,String income,String zysrzb,String ARPU,String cxywzdw,String cxywzjk,String ydjczjk,String gwjczjk,String zlsrzdw) {
		int result = 1;
		System.out.println(subject + "  " + province + "  " + date + "  " + tongbizengfu + "  " + income);
		try {
			double tongbizengfuDouble = Double.parseDouble(tongbizengfu) * 100;
			tongbizengfu = String.format("%.1f", tongbizengfuDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double incomeDouble = Double.parseDouble(income);
			income = String.format("%.1f", incomeDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
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
		
		String sql = "replace into `provincesubjecttotal`(`subject`,`income`,`tongbiIncrease`,`month`,`province`,zysrzb,ARPU,cxywzdw,cxywzjk,ydjczjk,gwjczjk,zlsrzdw) values(?,?,?,?,?,?,?,?,?,?,?,?)";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setInt(1,subject);
			prep.setString(2, income);
			prep.setString(3, tongbizengfu);
			prep.setString(4, date);
			prep.setString(5, province + "");
			prep.setString(6, zysrzb);
			prep.setString(7, ARPU);
			prep.setString(8, cxywzdw);
			prep.setString(9, cxywzjk);
			prep.setString(10, ydjczjk);
			prep.setString(11, gwjczjk);
			prep.setString(12, zlsrzdw);
			prep.executeUpdate();
		} catch (Exception ex) {
			ex.printStackTrace();
			result = -1;
		} finally {
			try {
				if (prep != null)
					prep.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}

			DataBaseConnection.returnConnection(conn);
		}
		return result;
	}
	
	public void importTotalData(XSSFSheet hsheet,String date) {
		int rows = hsheet.getPhysicalNumberOfRows();
		int flag = 0;
		for (int i = 0; i < rows; i++) {
			if(flag == 1)
				break;
			XSSFRow hrow = hsheet.getRow(i);
			int cols = hrow.getPhysicalNumberOfCells();
			for (int j = 0; j < cols; j++) {
				XSSFCell hcell = hrow.getCell(j);
				String cellValue = this.getCellValue(hcell);
				System.out.print(cellValue + " ");
				XSSFCell tbcell = hrow.getCell(j + 2);
				String tongbizengfu = this.getCellValue(tbcell);
				if(tongbizengfu.equals("")) {
					tongbizengfu = "0.0";
				}
				XSSFCell incomecell = hrow.getCell(j + 4);
				String income = this.getCellValue(incomecell);
				if(income.equals("")) {
					income = "0.0";
				}
				if(cellValue != null && cellValue.equals("主营业务收入")) {
					insertNationalTotal(date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("新兴ICT") >= 0) {
					insertNationalSubjectTotal(1,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("物联网") >= 0) {
					insertNationalSubjectTotal(11,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("大数据") >= 0) {
					insertNationalSubjectTotal(12,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("IT服务") >= 0) {
					insertNationalSubjectTotal(13,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("IDC") >= 0) {
					insertNationalSubjectTotal(14,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("云计算") >= 0) {
					insertNationalSubjectTotal(15,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("移网基础业务") >= 0) {
					insertNationalSubjectTotal(2,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("固网基础业务") >= 0) {
					insertNationalSubjectTotal(3,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("互联网业务") >= 0) {
					insertNationalSubjectTotal(31,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("固话") >= 0) {
					insertNationalSubjectTotal(32,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("数据网元 ") >= 0) {
					insertNationalSubjectTotal(33,date,tongbizengfu,income);
					break;
				} else if(cellValue.indexOf("分摊固网其他收入") >= 0) {
					flag = 1;
					insertNationalSubjectTotal(34,date,tongbizengfu,income);
					break;
				}
			}
		}
	}
		
	public int insertNationalTotal(String date,String tongbizengfu,String income) {
		int result = 1;
		try {
			double tongbizengfuDouble = Double.parseDouble(tongbizengfu) * 100;
			tongbizengfu = String.format("%.1f", tongbizengfuDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double incomeDouble = Double.parseDouble(income);
			income = String.format("%.1f", incomeDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
			
		String sql = "REPLACE into `nationtotal`(`income`,`tongbiIncrease`,`month`) values(?,?,?)";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1, income);
			prep.setString(2, tongbizengfu);
			prep.setString(3, date);
			prep.executeUpdate();
		} catch (Exception ex) {
			try {
				conn.rollback();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}
			result = -1;
		} finally {
			try {
				if (prep != null)
					prep.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}

			DataBaseConnection.returnConnection(conn);

		}
		return result;
	}

	public int insertNationalSubjectTotal(int subject,String date,String tongbizengfu,String income) {
		int result = 1;
		try {
			double tongbizengfuDouble = Double.parseDouble(tongbizengfu) * 100;
			tongbizengfu = String.format("%.1f", tongbizengfuDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		try {
			double incomeDouble = Double.parseDouble(income);
			income = String.format("%.1f", incomeDouble);
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		String sql = "REPLACE into `nationsubjecttotal`(`subject`,`income`,`tongbiIncrease`,`month`) values(?,?,?,?)";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		try {
			prep = conn.prepareStatement(sql);
			prep.setInt(1,subject);
			prep.setString(2, income);
			prep.setString(3, tongbizengfu);
			prep.setString(4, date);
			prep.executeUpdate();
		} catch (Exception ex) {
			ex.printStackTrace();
			result = -1;
		} finally {
			try {
				if (prep != null)
					prep.close();
			} catch (SQLException e1) {
				e1.printStackTrace();
			}

			DataBaseConnection.returnConnection(conn);

		}
		return result;
	}
	
	public String getCellValue(XSSFCell cell) {
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
							value = "";
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

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
