<!DOCTYPE html>
<%@Page Language="C#" %>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<html lang="en">
<head>
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
	</style>

	<!-- jQuery and JS -->
	<script src="https://code.jquery.com/jquery-1.10.2.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
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

	<!-- DELETE C#
	================================================== -->
	<script language="C#" runat="server">
	void Page_Load(Object sender, EventArgs e)
	{
		results.InnerHtml = "";
		if(IsPostBack)
		{
			String major = Request.Form["major"];
			String minor = Request.Form["minor"];
			String sub1 = Request.Form["sub1"];
			String sub2 = Request.Form["sub2"];

			String sql = "SELECT * FROM glmaster WHERE major = " + major;
			sql = sql + " AND minor = " + minor;
			sql = sql + " AND sub1 = " + sub1;
			sql = sql + " AND sub2 = " + sub2;

			String strConnect = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272";

			int c = 0;
			decimal bal = 1.0M;
			try
			{
				results.InnerHtml += "<p>opening a recordset with: " + sql;

				SqlConnection objConnect = new SqlConnection(strConnect);
				objConnect.Open();
				SqlCommand objCommand = new SqlCommand(sql, objConnect);
				SqlDataReader objDataReader;
				objDataReader = objCommand.ExecuteReader();

				while(objDataReader.Read())
				{
					c += 1;
					bal = Convert.ToDecimal(objDataReader["balance"]);
					results.InnerHtml += "<p>balance is " + bal;
				}
				objDataReader.Close();
				objConnect.Close();

				if(c == 0)
				{
					results.InnerHtml += "<p>ACCT NOT FOUND";
				}
				else
				{
					results.InnerHtml += "<p>" + c + " records found";

					if(bal != 0.0M)
					{
						results.InnerHtml += "<p>ACCT balance greater than 0, no deletion can be performed";
					}
					else
					{
						results.InnerHtml += "<p>Balance is zero, proceeding with deletion";

						int numa;
						SqlConnection cn = new SqlConnection(strConnect);
						cn.Open();

						sql = "DELETE FROM glmaster WHERE ";
						sql += "major = @major AND ";
						sql += "minor = @minor AND ";
		 				sql += "sub1 = @sub1 AND ";
						sql += "sub2 = @sub2";

						results.InnerHtml += "<p>Attempting delete with: " + sql;

						SqlCommand deleteGl = new SqlCommand(sql, cn);
						deleteGl.Parameters.AddWithValue("@major", major);
						deleteGl.Parameters.AddWithValue("@minor", minor);
						deleteGl.Parameters.AddWithValue("@sub1", sub1);
						deleteGl.Parameters.AddWithValue("@sub2", sub2);

						numa = deleteGl.ExecuteNonQuery();
						results.InnerHtml += "<p>" + numa + " record dleted OK.";
						cn.Close();
					}
				}
			}
			catch(Exception exc)
			{
				results.InnerHtml += "<p>ERROR";
			}
		}
	}
	</script>
	<!-- C#.NET END -->

	<div class="container">
	<div id="results" runat="server"></div>	
		<div class="row">
			<div class="col-lg-6 col-lg-offset-3">
				<div class="well bs-component">

					<form class="form-horizontal" method="POST" runat="server">
						<fieldset>

							<legend class="text text-primary">Enter a glmaster account ID to delete.</legend>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="major" id="major" placeholder="Major">
								</div>
							</div>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="minor" id="minor" placeholder="Minor">
								</div>
							</div>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="sub1" id="sub1" placeholder="Sub1">
								</div>
							</div>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="sub2" id="sub2" placeholder="Sub2">
								</div>
							</div>
						
							<div class="form-group">
								<div class="col-lg-10">
									<button type="submit" class="btn btn-success">Submit</button>
								</div>
							</div>

							<!-- Token for pass detection -->
							<input type="hidden" name="token" value="2">						

						</fieldset>
					</form>

				</div>
			</div>		
		</div>
	</div>

</body>
</html>
