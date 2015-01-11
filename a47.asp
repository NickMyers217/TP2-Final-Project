<!DOCTYPE html>
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
	
	<div class="container">

	<!-- SELECT ASP
	================================================== -->
		<%
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql_string = "SELECT * FROM je ORDER BY sourceref, srseq"
			response.write "<div class='row'>"
			response.write "<div class='col-lg-12'>"

			response.write "<h2 class='page-header'>SQL: <small>" + sql_string + "</small></h1>"
			rs.open sql_string, "DSN=gl1272;UID=gl1272;PWD=YWV85updG;"

			response.write "<table id='select' class='table table-striped table-hover greenBorder'>"

			response.write "<thead><tr>"
			for i = 0 to rs.fields.count - 1
				response.write "<th>" + cstr(rs(i).Name) + "</th>"
			next
			response.write "</tr></thead><tbody>"

			row_count = 0

			while not rs.EOF
				row_count=row_count+1

				response.write "<tr>"
				for i = 0 to rs.fields.count - 1
					 response.write "<td>" + cstr(rs(i)) + "</td>"
				next
				response.write "</tr>"

				rs.MoveNext
			wend

			response.write "</tbody></table><p>Row Count=" + cstr(row_count)
			response.write "<p>Normal Termination " + cstr(now)

			response.write "</div></div></div>"

			rs.Close
			set rs = nothing
		%>
	<!-- ASP END -->

	</div>
</body>
</html> 

