package servlet;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import bean.provinceBean;
import bean.queryBean;
import dbc.DataBaseConnection;
import net.sf.json.JSONObject;

/**
 * Servlet implementation class provinceData
 */
public class provinceData extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public provinceData() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		ResultSet rs = null;
		PrintWriter pw = response.getWriter();
		provinceBean bean = new provinceBean();
		String month = request.getParameter("month");
		String subject = request.getParameter("subject");
		String sql = "select income,tongbiIncrease from provincesubjecttotal where subject = ? and month = ? and province > 2 order by province asc";
		try {
			prep = conn.prepareStatement(sql);
			prep.setString(1, subject);
			prep.setString(2, month);
			rs = prep.executeQuery();
			while(rs.next()) {
				String income = rs.getString(1);
				String tongbiIncrease = rs.getString(2);
				bean.getProvinceIncomes().add(income);
				bean.getProvinceTBs().add(tongbiIncrease);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (rs != null)
					rs.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
			try {
				if (prep != null)
					prep.close();
			} catch (SQLException e) {
				e.printStackTrace();
			}
			DataBaseConnection.returnConnection(conn);
		}
		pw.write(JSONObject.fromObject(bean).toString());
		pw.close();
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
