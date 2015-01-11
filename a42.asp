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

	<!-- UPDATE ASP
	================================================== -->
	<%
	'pass1 display account id form
	sub pass1
	%>
		<div class="row">
			<div class="col-lg-6 col-lg-offset-3">
				<div class="well bs-component">

					<form class="form-horizontal" action="a42.asp" method="POST">
						<fieldset>

							<legend class="text text-primary">Enter a glmaster account ID to update.</legend>

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
									<textarea class="form-control" rows="3" name="description" id="description"></textarea>
									<span class="help-block">Enter a new description for this account.</span>
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
	
	'pass 2 find account and update description
	sub pass2
		set rs = server.createobject("ADODB.recordset")

		sql = "SELECT * FROM glmaster WHERE major = " + request.form("major")
		sql = sql + " AND minor = " + request.form("minor")
		sql = sql + " AND sub1 = " + request.form("sub1")
		sql = sql + " AND sub2 = " + request.form("sub2")

		response.write "<p>Opening recordset with: " + sql + "</p>"

		rs.open sql,"DSN=gl1272;UID=gl1272;PWD=YWV85updG;"

		response.write "<p>Recordset opened OK</p>"

		'Make sure this account exists
		c = 0
		while NOT rs.EOF
			ad = rs("acctdesc")
			c = c + 1
			rs.movenext
		wend

		rs.close
		set rs = nothing

		if c = 0 then
			response.write "<p>ACCT NOT FOUND</p>"
		else
			response.write "<p>Acct found, proceeding with update</p>"	

			set cn = server.createobject("ADODB.connection")
			cn.open "gl1272","gl1272","YWV85updG"

			sql = "UPDATE glmaster SET acctdesc = " + chr(39) + Request.form("description") + chr(39)
			sql = sql + " WHERE major = " + request.form("major") + " AND "
			sql = sql + " minor = " + request.form("minor") + " AND "
			sql = sql + " sub1 = " + request.form("sub1") + " AND "
			sql = sql + " sub2 = " + request.form("sub2")

			response.write "<p>Updating glmaster with: " + sql + "<P>"

			cn.execute sql,numa

			if numa=1 then
				response.write "<p>UPDATE OK"
			else
				response.write "<p>UPDATE FAILED"
			end if

			cn.close
			set cn = nothing
		end if

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

