<!DOCTYPE html>
<html lang="en">
<head>
	<%@Page Language="VB" %>
	<%@Import Namespace="System.Data" %>
	<%@Import Namespace="System.Data.SqlClient" %>

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

	<!-- SELECT VB.NET
	================================================== -->
	<script language="VB" runat="server">
	Sub Page_Load (sender as Object, e as EventArgs)
		Dim c as Integer = 0
		Dim sum as Decimal = 0.0
		Dim	strConnect as String = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272"
		Dim strSelect as String = "SELECT * FROM glmaster ORDER BY major, minor, sub1, sub2"

		Dim os as String = "<h3 class='page-header'>SQL: <small>" + strSelect + "</small></h3>"
		os += "<table class='table table-striped table-hover greenBorder'>"
		os += "<thead><tr>"
		os += "<th>Major</th>"
		os += "<th>Minor</th>"
		os += "<th>Sub1</th>"
		os += "<th>Sub2</th>"
		os += "<th>Account Description</th>"
		os += "<th>Balance</th>"
		os += "<th>Total</th>"
		os += "</tr></thead>"

		Try
			Dim  objConnect as new SqlConnection(strConnect)

			objConnect.Open()

			Dim objCommand as new SqlCommand(strSelect, objConnect)

			Dim objDataReader as SqlDataReader 

			objDataReader = objCommand.ExecuteReader()

			os += "<tbody>"
			Do while objDataReader.Read()
				c += 1
				sum += Convert.ToDecimal(objDataReader("balance"))

				os += "<tr>"
				os += "<td>" + objDataReader("major").ToString + "</td>"
				os += "<td>" + objDataReader("minor").ToString + "</td>"
				os += "<td>" + objDataReader("sub1").ToString + "</td>"
				os += "<td>" + objDataReader("sub2").ToString + "</td>"
				os += "<td>" + objDataReader("acctdesc") + "</td>"
				os += "<td>" + objDataReader("balance").ToString + "</td>"
				os += "<td>" + FormatCurrency(sum).ToString + "</td>"
				os += "</tr>"
			Loop

			objDataReader.Close()
			objConnect.Close()
		
		Catch objError as Exception 
		
			outError.InnerHtml = "<b>* Error </b>.<br />"+ objError.Message + "<br />" + objError.Source

		end Try

		os += "<tr>"
		os += "<td colspan='6'>Total Balance</td>"
		os += "<td>" + FormatCurrency(sum).ToString + "</td>"
		os += "</tr>"
		os += "</tbody></table>"
		os += "<p>Row Count = " + c.ToString
		os += "<P>Normal Termination " + Now.toString

		outResult.InnerHtml = os
	end sub	
	</script>
	<!-- VB.NET END -->

	</div>
</body>
</html> 

