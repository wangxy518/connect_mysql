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
 * Servlet implementation class userDataQuery
 */
public class userDataQuery extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public userDataQuery() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		PrintWriter pw = response.getWriter();
		String subject = request.getParameter("userSubject");
		String zhibiao = request.getParameter("userZhibiao");
		String queryCol = "";
		if(subject.equals("1")) {
			queryCol = "zysr";
		} else if(subject.equals("2")) {
			queryCol = "ydcz";
		} else if(subject.equals("3")) {
			queryCol = "dyjz";
		} else if(subject.equals("4")) {
			queryCol = "ljjz";
		} else if(subject.equals("5")) {
			queryCol = "xfz";
		} else if(subject.equals("6")) {
			queryCol = "czls";
		} else if(subject.equals("7")) {
			queryCol = "arpu";
		}
		
		if(zhibiao.equals("2")) {
			queryCol += "HB";
		}
		String sql = "";
		sql = "select " + queryCol + ",month from provinceUserData order by month asc";
	
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
