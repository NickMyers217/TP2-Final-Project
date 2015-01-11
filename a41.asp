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
	
	<!-- INSERT ASP
	================================================== -->
	<%
	'
	'*** globals
	'
	dim cn
	dim rs
	dim tokenvalue

	'Pass1 produce the form
	sub pass1
	%>
		<div class="row">
			<div class="col-lg-6 col-lg-offset-3">
				<div class="well bs-component">

					<form class="form-horizontal" action="a41.asp" method="POST">
						<fieldset>

							<legend class="text text-primary">Insert a new account into glmaster.</legend>

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
									<span class="help-block">Describe this account with a short sentence.</span>
								</div>
							</div>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="balance" id="balance" placeholder="Balance">
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

	'Pass 2 insert the record
	sub pass2
		response.write "<P>Pass 2 tokenvalue = " + cstr(tokenvalue)

		set cn = Server.CreateObject("ADODB.Connection")
		cn.open "gl1272","gl1272","YWV85updG"
		response.write "<P>Connection created OK"

		set rs = Server.CreateObject("ADODB.Recordset")

		'Obtain the account id that is being inserted and validate it here
		major = request.form("major")
		minor = request.form("minor")
		sub1  = request.form("sub1")
		sub2  = request.form("sub2")
		accountID = cstr(major) + cstr(minor) + cstr(sub1) + cstr(sub2)

		sqlString = "SELECT * FROM glmaster WHERE major = " + cstr(major)
		sqlString = sqlString + " AND minor = " + cstr(minor)	
		sqlString = sqlString + " AND sub1 = " + cstr(sub1)	
		sqlString = sqlString + " AND sub2 = " + cstr(sub2)	

		response.write "<P>Opening glmaster rs with SQL--->" + cstr(sqlString) + "<------ "
		rs.open sqlString,"DSN=gl1272;UID=gl1272;PWD=YWV85updG;"
		response.write "<P>Recordset opened OK"

		'If there are more than 0 records, this account already exists
		c = 0
	 	while NOT rs.EOF
			c = c + 1
			rs.movenext
	 	wend
	 	rs.close
		set rs = nothing

		response.write "<P>Existence for account = " + cstr(accountID)
			if c = 0 then
				response.write ": Found " + cstr(c) + " records. Proceeding with Insert"

				sqlString = "INSERT INTO glmaster "
				sqlString = sqlString + "(major, minor, sub1, sub2, acctdesc, balance) VALUES ("
				sqlString = sqlString + request.form("major") + ", "
				sqlString = sqlString + request.form("minor") + ", "
				sqlString = sqlString + request.form("sub1") + ", "
				sqlString = sqlString + request.form("sub2") + ", "
				sqlString = sqlString + chr(39) + request.form("description") + chr(39) + ", "
				sqlString = sqlString + request.form("balance") + ")"

				response.write "<P>Ready for Insert with sqlString = " + cstr(sqlString)

				cn.execute sqlString,numa

				if numa = 1 then
					response.write "<P>Inserted OK numa = " + cstr(numa)
				else
					response.write "<P>Insert Failed. Number of records inserted = " + cstr(numa)
				end if

				cn.close
				set cn = nothing

				response.write "<P>Terminating normally"
			else
				if c = 1 then
					response.write ": Found " + cstr(c) + " records. No insert will be performed. Duplicate key"
				else
					response.write "<P>Found " + cstr(c) + " records. No Insert will be performed. More than one account with this ID found"
				end if
			end if
	end sub

	sub passerror
		response.write "<p>INVALID TOKEN VALUE. token=" + cstr(tokernvalue)
	end sub

	'
	'*** top of main
	'
	tokenvalue = request.form("token")
	select case tokenvalue
	case ""
		 call pass1
	case "2"
		call pass2
	case else
		 call passerror
	end select

	%>
	<!-- ASP END -->

	</div>

</body>
</html>
