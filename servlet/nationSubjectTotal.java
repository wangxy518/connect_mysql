package servlet;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import bean.NationSubjectTotalBean;
import bean.NationSubjectTotalListBean;
import bean.NationtotalBean;
import dbc.DataBaseConnection;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * Servlet implementation class nationSubjectTotal
 */
public class nationSubjectTotal extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public nationSubjectTotal() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub

		PrintWriter pw = response.getWriter();
		String sql = "select * from nationsubjecttotal";
		Connection conn = DataBaseConnection.getConnection();
		PreparedStatement prep = null;
		ResultSet rs = null;
		NationSubjectTotalListBean listbean = new NationSubjectTotalListBean();
		try {
			prep = conn.prepareStatement(sql);
			rs = prep.executeQuery();
			while(rs.next()) {
				NationSubjectTotalBean bean = new NationSubjectTotalBean();
				bean.setId(rs.getInt(1));
				bean.setSubject(rs.getString(2));
				bean.setIncome(rs.getString(4));
				bean.setTongbiIncrease(rs.getString(3));
				bean.setMonth(rs.getString(5));
				if(bean.getSubject().equals("1")) {
					listbean.getICTlist().add(bean);
				} else if(bean.getSubject().equals("11")) {
					listbean.getWLWlist().add(bean);
				} else if(bean.getSubject().equals("12")) {
					listbean.getDSJlist().add(bean);
				} else if(bean.getSubject().equals("13")) {
					listbean.getITlist().add(bean);
				} else if(bean.getSubject().equals("14")) {
					listbean.getIDClist().add(bean);
				} else if(bean.getSubject().equals("15")) {
					listbean.getYJSlist().add(bean);
				} else if(bean.getSubject().equals("2")) {
					listbean.getYWJClist().add(bean);
				} else if(bean.getSubject().equals("3")) {
					listbean.getGWJClist().add(bean);
				} else if(bean.getSubject().equals("31")) {
					listbean.getHLWlist().add(bean);
				} else if(bean.getSubject().equals("32")) {
					listbean.getGHlist().add(bean);
				} else if(bean.getSubject().equals("33")) {
					listbean.getSJDYlist().add(bean);
				} else if(bean.getSubject().equals("34")) {
					listbean.getQTlist().add(bean);
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
		pw.write(JSONObject.fromObject(listbean).toString());
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
