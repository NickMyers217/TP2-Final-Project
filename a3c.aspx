<!DOCTYPE html>
<html lang="en">
<head>
	<%@Page Language="C#" Debug="true" validateRequest="false"%>
	<%@Import Namespace="System.Data" %>
	<%@Import Namespace="System.Data.SqlClient" %>
	<%@Import Namespace="System.Xml" %>

	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">

	<title>Nick Myers</title>

	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />

	<!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
	<!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
	<!--[if lt IE 9]>
		<script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
		<script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
	<![endif]-->

	<style>
		.greenBorder {
			border: 1px solid lightgreen;
		}
		td {
			border: 1px solid lightgreen;
		}
	</style>

	<!-- jQuery and JS -->
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
	<script src="js/a3.js"></script>

</head>
<body>
	<!-- NAVBAR
	================================================== -->
	<script>
		$(function() {
				$("#navbar").load("nav.htm");
		});
	</script>
	<div id="navbar"></div>
	<!-- END NAV -->

	<!-- JOURNAL VOUCHER C#
	================================================== -->
	<script language="C#" runat="server">
		int rowCount(String strConnect, String strSelect)
		{
			try 
			{
				SqlConnection objConnect = new SqlConnection(strConnect);
				objConnect.Open();
				SqlCommand objCommand = new SqlCommand(strSelect, objConnect);
				SqlDataReader objDataReader;
				objDataReader = objCommand.ExecuteReader();

				results.InnerHtml += "<p> Running: " + strSelect;

				int c  = 0;
				while(objDataReader.Read())
				{
					c++;
				}

				return c;
			} 
			catch(Exception error)
			{
				results.InnerHtml += "<p>ERROR";
				return 100;
			}
		}

		void Page_Load(Object sender, EventArgs e)
		{
			results.InnerHtml = "";
			if(IsPostBack)
			{
				String xmlString  = Request.Form["xmlout"];
				XmlDocument x = new XmlDocument();
    			x.LoadXml(xmlString);

				int nr = x.DocumentElement.ChildNodes.Count - 2;
				String srn = x.DocumentElement.ChildNodes[0].InnerText;

				results.InnerHtml += "<p>Number of journal entres: " + nr;
				results.InnerHtml += "<p>Preparting to validate srn number: " + srn;

				String connectStr = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272";
				String sql = "SELECT * FROM je WHERE sourceref = " + srn;
				int jeRows = rowCount(connectStr, sql);

				if(jeRows > 0)
				{
					results.InnerHtml += "<p>SRN was found in " + jeRows + " entries, terminating program";
				}
				else
				{
					results.InnerHtml += "<p>SRN was found in " + jeRows + " entries, proceeding with glmaster checks";

					int glValid = 0;
					for(int i = 0; i < nr; i++)
					{
						String major = x.DocumentElement.ChildNodes[i+2].ChildNodes[0].InnerText;
						String minor = x.DocumentElement.ChildNodes[i+2].ChildNodes[1].InnerText;
						String sub1  = x.DocumentElement.ChildNodes[i+2].ChildNodes[2].InnerText;
						String sub2  = x.DocumentElement.ChildNodes[i+2].ChildNodes[3].InnerText;

						sql = "SELECT * FROM glmaster WHERE major = " + major + " AND ";
						sql += " minor = " + minor + " AND ";
						sql += " sub1  = " + sub1 + " AND ";
						sql += " sub2  = " + sub2;

						results.InnerHtml += "<p>Checking journal entry " + (i + 1) + " out of " + nr;
						int glCount = rowCount(connectStr, sql);
						results.InnerHtml += "<p>Found " + glCount + " rows.";
						
						if(glCount == 1)
						{
							glValid++;
						}
					}
	
					if(glValid != nr)
					{
						results.InnerHtml += "<p>Glmaster checks failed: " + glValid + " rows found out of " + nr;
						results.InnerHtml += "<p>Program terminating";
					}
					else
					{
						results.InnerHtml += "<p>Glmaster checks passed: " + glValid + " rows found out of " + nr;
						results.InnerHtml += "<p>Proceeding with posts";

						int numa;
						SqlConnection cn = new SqlConnection(connectStr);
						cn.Open();

						results.InnerHtml += "<p>Beginning balance updates";
						for(int i = 0; i < nr; i++)
						{
							String major = x.DocumentElement.ChildNodes[i+2].ChildNodes[0].InnerText;
							String minor = x.DocumentElement.ChildNodes[i+2].ChildNodes[1].InnerText;
							String sub1  = x.DocumentElement.ChildNodes[i+2].ChildNodes[2].InnerText;
							String sub2  = x.DocumentElement.ChildNodes[i+2].ChildNodes[3].InnerText;
							String ta  = x.DocumentElement.ChildNodes[i+2].ChildNodes[4].InnerText;

							sql = "UPDATE glmaster SET balance = balance + @ta WHERE ";
							sql += "major = @major AND ";
							sql += "minor = @minor AND ";
							sql += "sub1 = @sub1 AND ";
							sql += "sub2 = @sub2";

							try
							{
								SqlCommand glUpdateCmd = new SqlCommand(sql, cn);

								glUpdateCmd.Parameters.AddWithValue("@major", major);
								glUpdateCmd.Parameters.AddWithValue("@minor", minor);
								glUpdateCmd.Parameters.AddWithValue("@sub1", sub1);
								glUpdateCmd.Parameters.AddWithValue("@sub2", sub2);
								glUpdateCmd.Parameters.AddWithValue("@ta", ta);

								numa = glUpdateCmd.ExecuteNonQuery();
								results.InnerHtml += "<p>Executed sql: " + sql + " recieved a numa of: " + numa;
							}
							catch(Exception error)
							{
								results.InnerHtml += "<p>ERROR";
							}
						}
						results.InnerHtml += "<p>Ending balance updates and beginning inserts";
						for(int i = 0; i < nr; i++)
						{
							String major = x.DocumentElement.ChildNodes[i+2].ChildNodes[0].InnerText;
							String minor = x.DocumentElement.ChildNodes[i+2].ChildNodes[1].InnerText;
							String sub1  = x.DocumentElement.ChildNodes[i+2].ChildNodes[2].InnerText;
							String sub2  = x.DocumentElement.ChildNodes[i+2].ChildNodes[3].InnerText;
							String ta = x.DocumentElement.ChildNodes[i+2].ChildNodes[4].InnerText;
							String desc = x.DocumentElement.ChildNodes[i+2].ChildNodes[5].InnerText;

							sql = "INSERT INTO je(sourceref, srseq, jemajor, jeminor, jesub1, jesub2, jedesc, jeamount)";
							sql += "VALUES (@srn, @srseq, @major, @minor, @sub1, @sub2, @desc, @ta)";

							try
							{
								SqlCommand jeInsertCmd = new SqlCommand(sql, cn);

								jeInsertCmd.Parameters.AddWithValue("@srn", srn);
								jeInsertCmd.Parameters.AddWithValue("@srseq", (i+1));
								jeInsertCmd.Parameters.AddWithValue("@major", major);
								jeInsertCmd.Parameters.AddWithValue("@minor", minor);
								jeInsertCmd.Parameters.AddWithValue("@sub1", sub1);
								jeInsertCmd.Parameters.AddWithValue("@sub2", sub2);
								jeInsertCmd.Parameters.AddWithValue("@desc", desc);
								jeInsertCmd.Parameters.AddWithValue("@ta", ta);

								numa = jeInsertCmd.ExecuteNonQuery();
								results.InnerHtml += "<p>Executed sql: " + sql + " recieved a numa of: " + numa;
							}
							catch(Exception error)
							{
								results.InnerHtml += "<p>ERROR";
							}
						}
						results.InnerHtml += "<p>Ending je inserts";
						results.InnerHtml += "<p>Finished";
					}
				}
			}
		}
	</script>
	<!-- C# END -->

	<div class="container">
		<div class="row">
			<div class="col-lg-12">

				<div id="results" runat="server"></div>
				
				<span class="text text-primary">
					<h3 class="page-header">Add Journal Voucher entries.</h3>
				</span>

				<form class="form-horizontal" id="gl" name="gl">
					<table class="table table-striped greenBorder">
						<tr><td colspan="6" align="center"><b>Journal Voucher</td></tr>
						<tr>
							<td colspan="4">
							<b>SRN:</b>
							<input id="srn" name="srn" type="text" size="4"> (4 digits)
							</td>
							<td colspan="2">
							<b>Fiscal Month:</b>
							<input id="month" name="month" type="text" size="2"> Must be 1,2,3,...12
							</td>
						</tr>
						<tr>
							<td colspan="6" align="center"><b>Journal Entries</td></tr>
						<tr>
							<td colspan="4"><b><center>General Ledger Account Number</center></td>
							<td align="center" rowspan="2"><b>Transaction<br>Amount</td>
							<td align="center" rowspan="2"><b>Description (max. 50 characters)</td>
						</tr>
						<tr>
							<td align="center"><b>Major</td>
							<td align="center"><b>Minor</td>
							<td align="center"><b>Sub<br>Acct 1</td>
							<td align="center"><b>Sub<br>Acct 2</td>
						</tr>
						<tr>
							<td><input id="gl1" type="text" size="4"></td>
							<td><input id="gl2" type="text" size="4"></td>
							<td><input id="gl3" type="text" size="4"></td>
							<td><input id="gl4" type="text" size="4"></td>
							<td align="right"><input id="gl5" type="text" size="12" align="right"></td>
							<td><input id="gl6" type="text" size="50"></td>
						</tr>
						<tr>
							<td><input id="gl7" type="text" size="4"></td>
							<td><input id="gl8" type="text" size="4"></td>
							<td><input id="gl9" type="text" size="4"></td>
							<td><input id="gl10" type="text" size="4"></td>
							<td align="right"><input id="gl11" type="text" size="12" align="right"></td>
							<td><input id="gl12" type="text" size="50"></td>
						</tr>
						<tr>
							<td><input id="gl13" type="text" size="4"></td>
							<td><input id="gl14" type="text" size="4"></td>
							<td><input id="gl15" type="text" size="4"></td>
							<td><input id="gl16" type="text" size="4"></td>
							<td align="right"><input id="gl17" type="text" size="12" align="right"></td>
							<td><input id="gl18" type="text" size="50"></td>
						</tr>
						<tr>
							<td><input id="gl19" type="text" size="4"></td>
							<td><input id="gl20" type="text" size="4"></td>
							<td><input id="gl21" type="text" size="4"></td>
							<td><input id="gl22" type="text" size="4"></td>
							<td align="right"><input id="gl23" type="text" size="12" align="right"></td>
							<td><input id="gl24" type="text" size="50"></td>
						</tr>
						<tr>
							<td><input id="gl25" type="text" size="4"></td>
							<td><input id="gl26" type="text" size="4"></td>
							<td><input id="gl27" type="text" size="4"></td>
							<td><input id="gl28" type="text" size="4"></td>
							<td align="right"><input id="gl29" type="text" size="12" align="right"></td>
							<td><input id="gl30" type="text" size="50"></td>
						</tr>
						<tr>
							<td><input id="gl31" type="text" size="4"></td>
							<td><input id="gl32" type="text" size="4"></td>
							<td><input id="gl33" type="text" size="4"></td>
							<td><input id="gl34" type="text" size="4"></td>
							<td align="right"><input id="gl35" type="text" size="12" align="right"></td>
							<td><input id="gl36" type="text" size="50"></td>
						</tr>
						<tr>
							<td colspan="4"><b>Sum of Transaction Amounts</td>
							<td align="right"><input type="text" size="12" value="0.00"></td>
							<td><b>(Must sum to 0.0)</td>
						</tr>
						<tr>
							<td colspan="6" align="center">
								<div id="errors"></div>
								<input type="button" class="btn btn-success" onclick="main();" value="Validate and Submit">
							</td>
						</tr>
					</table>
				</form>

				<!-- Hidden XML Form -->
				<form name="xmlform" method="POST" runat="server">
					<input type="hidden" id="xmlout" name="xmlout">
				</form>

			</div>		
		</div>
	</div>

</body>
</html>

  
