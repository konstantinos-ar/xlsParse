package gr.ztrade.parser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XParser
{
	private static final String dblink = "jdbc:sqlserver://192.168.65.14:1111";
	private static final String dbuser = "";
	private static final String dbpass = "";

	private static final String DBDRIVER = "com.microsoft.sqlserver.jdbc.SQLServerDriver";

	public static void main(String[] args) throws SQLException
	{

		read();
		/*File file = new File("openworkbook.xlsx");
		FileInputStream fIP;
		try
		{
			fIP = new FileInputStream(file);

			//Get the workbook instance for XLSX file 
			XSSFWorkbook workbook = new XSSFWorkbook(fIP);

		}
		catch (IOException e)
		{
			e.printStackTrace();
		}

		if(file.isFile() && file.exists())
		{
			System.out.println(
					"openworkbook.xlsx file open successfully.");
		}
		else
		{
			System.out.println(
					"Error to open openworkbook.xlsx file.");
		}*/

	}


	public static void read() throws SQLException
	{
		HSSFRow row;
		URL url;
		Statement st = null,st2;
		ResultSet rs = null;
		Connection con = null;
		//FileInputStream fis = null;
		InputStream urlin = null;
		String data = null;
		int tick = 0, c = 0;
		String date = null, nav = null, shares = null, assets = null;
		//SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		//SimpleDateFormat sdf2 = new SimpleDateFormat("dd-MMM-yyyy");
		try
		{
			Class.forName(DBDRIVER);
			con = DriverManager.getConnection(dblink, dbuser, dbpass);

			st = con.createStatement();
			rs = st.executeQuery("Select Sym from MarketsData.dbo.ETFYahoo where fundid=102");
		}
		catch (SQLException | ClassNotFoundException e)
		{
			e.printStackTrace();
		}

		while (rs.next()) 
		{
			try
			{
				st2 = con.createStatement();
				String ex = rs.getString("Sym");
				url = new URL("https://www.spdrs.com/site-content/xls/"+ex+"_HistoricalNav.xls");
				System.out.println("try: "+url);

				URLConnection ucon = url.openConnection();

				ucon.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows; U; Windows NT 6.0; el; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 (.NET CLR 3.5.30729)");
				ucon.setRequestProperty("Accept","text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
				ucon.setRequestProperty("Accept-Language","el-gr,el;q=0.7,en-us.;q=0.3");
				ucon.setRequestProperty("Accept-Encoding","gzip");
				ucon.setRequestProperty("Accept-Charset","ISO-8859-7,utf-8;q=0.7,*;q=0.7");
				ucon.setRequestProperty("Keep-Alive","300");
				ucon.setRequestProperty("Connection","keep-alive");
				ucon.setConnectTimeout(20000);
				ucon.setReadTimeout(20000);

				urlin = ucon.getInputStream();

				//fis = new FileInputStream(new File("C:/Users/user/Downloads/ACIM_HistoricalNav (3).xls"));

				HSSFWorkbook workbook = new HSSFWorkbook(urlin);

				HSSFSheet spreadsheet = workbook.getSheetAt(0);
				Iterator < Row > rowIterator = spreadsheet.iterator();
				while (rowIterator.hasNext()) 
				{
					row = (HSSFRow) rowIterator.next();
					Iterator < Cell > cellIterator = row.cellIterator();
					c = 0;
					while ( cellIterator.hasNext()) 
					{
						Cell cell = cellIterator.next();
						c++;
						switch (cell.getCellType()) 
						{
						case Cell.CELL_TYPE_NUMERIC:
							//System.out.print( cell.getNumericCellValue() + " \t\t " );
							data = cell.getStringCellValue();
							if (tick > 0)
							{
								if (c == 2)
									nav = data;
								if (c == 3)
									shares = data;
								if (c == 4)
									assets = data;
							}
							break;
						case Cell.CELL_TYPE_STRING:
							//System.out.print(cell.getStringCellValue() + " \t\t " );
							data = cell.getStringCellValue();
							if (data.startsWith("Performance"))
								tick = 0;

							if (tick > 0)
							{
								if (c == 1)
									date = data;
								if (c == 2)
									nav = data;
								if (c == 3)
									shares = data;
								if (c == 4)
									assets = data;
							}

							if (data.equals("Total Net Assets"))
								tick = 1;
							break;
						}
					}
					//System.out.println();
					if (tick > 0 && !data.equals("Total Net Assets"))
					{
						//date = sdf.format(sdf2.parse(date));
						try{
							st2.execute("Insert into MarketsData.dbo.ETFHist(Sym,Date,Nav,Shares,Assets) values ('"+ex+"','"+date+"',"+nav+","+shares+","+assets+")");
							System.out.println("Insert into MarketsData.dbo.ETFHist(Sym,Date,Nav,Shares,Assets) values ('"+ex+"','"+date+"',"+nav+","+shares+","+assets+")");
						}catch (Exception e){}
						//System.out.println("Date: " + date + ", Nav: " + nav + ", Shares: " + shares + ", Assets: " + assets);

					}
				}

				//fis.close();
			}
			catch (IOException | SQLException e)
			{
				e.printStackTrace();
			}
		}
	}

}
