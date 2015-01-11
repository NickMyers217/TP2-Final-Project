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

	<div class="container">
	
	<!-- JOURNAL VOUCHER ASP
	================================================== -->
	<%
	'Pass1 produce the journal voucher form
	sub pass1
	%>
		<div class="row">
			<div class="col-lg-12">
				
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
				<form name="xmlform" action="a3.asp" method="POST">
					<input type="hidden" id="xmlout" name="xmlout">
					<!-- Token for pass detection -->
					<input type="hidden" name="token" value="2">						
				</form>

			</div>		
		</div>
	<%
	end sub

	function rowCount(sql)
		set rs = Server.CreateObject("ADODB.Recordset")		
		response.write "<p>Opening a recordset with: " + cstr(sql)
		rs.open sql,"DSN=gl1272;UID=gl1272;PWD=YWV85updG;"
		response.write "<p>Recordset opened OK"

		c = 0
		while NOT rs.EOF
			c = c + 1
			rs.movenext
		wend
		
		rs.close
		set res = nothing

		rowCount = c
	end function

	'Pass 2 validate the existence of the gl accounts and non-existence of the srn number
	'Insert all the records
	sub pass2
		fx = request.form("xmlout")
  	set x = Server.CreateObject("Microsoft.xmldom")
		x.loadXML(fx)

		if (x.parseError <> 0) then
			response.write "<p>Parse Error Found:" + cstr(x.parseError.reason)
			response.write "<p>ASP Program is terminating."
		else
			nr = x.documentElement.childNodes.length - 2
			srn = x.documentElement.childNodes(0).text

			response.write "<p>Number of journal entries: " + cstr(nr)
			response.write "<p>Preparing to validate srn number: " + cstr(srn)

			sql = "SELECT * FROM je WHERE sourceref = " + cstr(srn)
			jeRows = rowCount(sql)
			
			if jeRows > 0 then
				response.write "<p>SRN was found in " + cstr(jeRows) + " journal entries, program terminating"
			else
				response.write "<p>SRN was found in " + cstr(jeRows) + " journal entries, proceeding with glmaster checks"

				glValid = 0
				for i = 0 to nr - 1
					major = x.documentElement.childNodes(i + 2).childNodes(0).text
					minor = x.documentElement.childNodes(i + 2).childNodes(1).text
					sub1 = x.documentElement.childNodes(i + 2).childNodes(2).text
					sub2 = x.documentElement.childNodes(i + 2).childNodes(3).text

					sql = "SELECT * FROM glmaster WHERE major = " + cstr(major) + " AND "
					sql = sql + " minor = " + cstr(minor) + " AND "
					sql = sql + " sub1 = " + cstr(sub1) + " AND "
					sql = sql + " sub2 = " + cstr(sub2)	

					response.write "<p>Checking journal entry " + cstr(i + 1) + " out of " + cstr(nr)
					glCount = rowCount(sql)

					if glCount = 1 then
						glValid = glValid + 1
					end if
				next

				if glValid <> nr then
					response.write "<p>Glmaster checks failed: " + cstr(glValid) + " rows found out of " + cstr(nr)
					response.write "<p>Program terminating."
				else
					response.write "<p>Glmaster checks passed: " + cstr(glValid) + " rows found out of " + cstr(nr)

					set cn = Server.CreateObject("ADODB.Connection")
					cn.open "gl1272","gl1272","YWV85updG"	
					cn.beginTrans
					response.write "<p>Connection opened OK, updating glmaster balances"

					updateErrors = 0
					for i = 0 to nr - 1
						major = x.documentElement.childNodes(i + 2).childNodes(0).text
						minor = x.documentElement.childNodes(i + 2).childNodes(1).text
						sub1 = x.documentElement.childNodes(i + 2).childNodes(2).text
						sub2 = x.documentElement.childNodes(i + 2).childNodes(3).text
						ta = x.documentElement.childNodes(i + 2).childNodes(4).text

						sql = "UPDATE glmaster SET balance = balance + " + cstr(ta) + " WHERE "
						sql = sql + "major = " + cstr(major) + " AND "
						sql = sql + "minor = " + cstr(minor) + " AND "
						sql = sql + "sub1 = " + cstr(sub1) + " AND "
						sql = sql + "sub2 = " + cstr(sub2)
						
						cn.execute sql,numa
						
						if numa = 1 then
							response.write "<p>glmaster record " + cstr(i + 1) + " updated with: " + sql
						else
						  updateErrors = updateErrors + 1
						end if
					next

					for i = 0 to nr - 1
						major = x.documentElement.childNodes(i + 2).childNodes(0).text
						minor = x.documentElement.childNodes(i + 2).childNodes(1).text
						sub1 = x.documentElement.childNodes(i + 2).childNodes(2).text
						sub2 = x.documentElement.childNodes(i + 2).childNodes(3).text
						ta = x.documentElement.childNodes(i + 2).childNodes(4).text
						desc = x.documentElement.childNodes(i + 2).childNodes(5).text

						sql = "INSERT INTO je(sourceref, srseq, jemajor, jeminor, jesub1, jesub2, jedesc, jeamount) VALUES ("
						sql = sql + cstr(srn) + ", "
						sql = sql + cstr(i + 1) + ", "
						sql = sql + cstr(major) + ", "
						sql = sql + cstr(minor) + ", "
						sql = sql + cstr(sub1) + ", "
						sql = sql + cstr(sub2) + ", "
						sql = sql + "'" + cstr(desc) + "', "
						sql = sql + cstr(ta) + ")"
						
						cn.execute sql,numa
						
						if numa = 1 then
							response.write "<p>New je record inserted with: " + sql
						else
						  updateErrors = updateErrors + 1
						end if
					next

					if updateErrors <> 0 then
						response.write "<p>Errors found: " + cstr(updateErrors) + " rolling back"
						cn.rollbackTrans
					else
						response.write "<p>No errors, commiting"
						cn.commitTrans
					end if

					cn.close
					set cn = nothing
					set x = nothing

				end if
			end if			
  	end if

	end sub

	sub passerror
		response.write "<p>INVALID TOKEN VALUE. token=" + cstr(tokenvalue)
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

  
