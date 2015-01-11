<!DOCTYPE html>
<html lang="en">
<head>
	<% @Page Language="C#" %>
	<% @Import Namespace="System.Data" %>
	<% @Import Namespace="System.Data.SqlClient" %>

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
	
	<div class="container">
		<div id="outError" runat="server"></div>
		<div id="outResult" runat="server"></div>

	<!-- SELECT C#
	================================================== -->
	<script language="C#" runat="server">
	void Page_Load(Object sender, EventArgs e)
	{
		int c = 0;
		decimal sum = 0.0M;
		string strConnect = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272";
		string strSelect = "SELECT * FROM glmaster ORDER BY major, minor, sub1, sub2";

		string os = "<h3 class='page-header'>SQL: <small>" + strSelect + "</small></h3>";
		os += "<table class='table table-striped table-hover greenBorder'>";
		os += "<thead><tr>";
		os += "<th>Major</th>";
		os += "<th>Minor</th>";
		os += "<th>Sub1</th>";
		os += "<th>Sub2</th>";
		os += "<th>Account Description</th>";
		os += "<th>Balance</th>";
		os += "<th>Total</th>";
		os += "</tr></thead>";

		try
		{
			SqlConnection objConnect = new SqlConnection(strConnect);
			objConnect.Open();
			SqlCommand objCommand = new SqlCommand(strSelect, objConnect);
			SqlDataReader objDataReader;
			objDataReader = objCommand.ExecuteReader();

			while (objDataReader.Read())
			{
				sum=sum+Convert.ToDecimal(objDataReader["balance"]);
        c++;
				os += "<tr>";
				os += "<td>" + objDataReader["major"] + "</td>";
				os += "<td>" + objDataReader["minor"] + "</td>";
				os += "<td>" + objDataReader["sub1"] + "</td>";
				os += "<td>" + objDataReader["sub2"] + "</td>";
				os += "<td>" + objDataReader["acctdesc"] + "</td>";
				os += "<td>" + objDataReader["balance"] + "</td>";
				os += "<td>" + sum + "</td>";
				os += "</tr>";
			}

			objDataReader.Close();
			objConnect.Close();
		}
		catch (Exception error)
		{
			outError.InnerHtml = "<b>* Error</b>.<br>" + error.Message + "<br>" + error.Source;
			return;
		}

		os += "<tr><td colspan='6'>Total Balance</td><td>" + sum + "</td></tr>";
		os += "</table>";
		os += "<p>Row Count = " + c;
		DateTime t = DateTime.Now;
    os += "<P>Normal Termination " + t;

		outResult.InnerHtml = os;
	}
	</script>
	<!-- C# END -->

	</div>
</body>
</html> 

