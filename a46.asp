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
	'pass1 display account id form
	sub pass1
	%>
		<div class="row">
			<div class="col-lg-6 col-lg-offset-3">
				<div class="well bs-component">

					<form class="form-horizontal" action="a46.asp" method="POST">
						<fieldset>

							<legend class="text text-primary">Enter a journal entry SRN to query.</legend>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="srn" id="srn" placeholder="SRN">
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
	<%
	end sub

	sub pass2
		Set rs = Server.CreateObject("ADODB.Recordset")

		sql = "SELECT * FROM je WHERE sourceref = " + request.form("srn")

		response.write "<div class='row'>"
		response.write "<div class='col-lg-12'>"
		response.write "<h2 class='page-header'>SQL: <small>" + sql + "</small></h1>"

		rs.open sql, "DSN=gl1272;UID=gl1272;PWD=YWV85updG;"

		response.write "<table id='select' class='table table-striped table-hover greenBorder'>"

		response.write "<thead><tr>"
		for i = 0 to rs.fields.count - 1
			response.write "<th>" + cstr(rs(i).Name) + "</th>"
		next
		response.write "</tr></thead>"

		row_count = 0
		while not rs.EOF
			row_count = row_count + 1
			response.write "<tr>"
			for i = 0 to rs.fields.count - 1
				 response.write "<td>" + cstr(rs(i)) + "</td>"
			next
			response.write "</tr>"
			rs.MoveNext
		wend

		response.write "</table><p>Row Count=" + cstr(row_count)
		response.write "<p>Normal Termination " + cstr(now)

		response.write "</div></div></div>"

		rs.Close
		set rs=nothing
	end sub

	'
	'*****MAIN
	'
	token_value=Request.form("token")
	select case token_value
	case ""
		call pass1
	case "2"
		call pass2
	end select
	%>
	<!-- ASP END -->

	</div>
</body>
</html> 

