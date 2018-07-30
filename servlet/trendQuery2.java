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

import bean.queryBean;
import dbc.DataBaseConnection;
import net.sf.json.JSONObject;

/**
 * Servlet implementation class trendQuery2
 */
public class trendQuery2 extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public trendQuery2() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		PrintWriter pw = response.getWriter();
		String province = request.getParameter("province");
		String zhibiao = request.getParameter("zhibiao");
		String sql = "select " + zhibiao + ",month from provincesubjecttotal where subject = 0 and province = " + province + " order by month asc";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		ResultSet rs = null;
		queryBean bean = new queryBean();
		try {
			prep = conn.prepareStatement(sql);
			rs = prep.executeQuery();
			String month = "";
			while(rs.next()) {
				month = rs.getString(2);
				if(month.substring(0,4).equals("2017")) {
					bean.getList2017().add(rs.getString(1));
				} else {
					bean.getList2018().add(rs.getString(1));
				}
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
