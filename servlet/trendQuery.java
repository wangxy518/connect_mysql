package servlet;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import javax.servlet.Servlet;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import bean.NationtotalBean;
import bean.queryBean;
import dbc.DataBaseConnection;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * Servlet implementation class dataQuery
 */
public class trendQuery extends HttpServlet implements Servlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public trendQuery() {
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
		String subject = request.getParameter("subject");
		String zhibiao = request.getParameter("zhibiao");
		String queryCol = "income";
		if(zhibiao.equals("2")) {
			queryCol = "tongbiIncrease";
		}
		String sql = "";
		if(province.equals("0") && subject.equals("0")) {
			sql = "select " + queryCol + ",month from nationtotal order by month asc";
		} else if(province.equals("0") && !subject.equals("0")) {
			sql = "select " + queryCol + ",month from nationsubjecttotal where subject = " + subject + " order by month asc";
		} else if(!province.equals("0")) {
			sql = "select " + queryCol + ",month from provincesubjecttotal where subject = " + subject + " and province = " + province + " order by month asc";
		}
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
