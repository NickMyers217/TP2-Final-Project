<!DOCTYPE html>
<html lang="en">
<head>
	<%@Page Language="VB" Debug="true" validateRequest="false"%>
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

	<!-- JOURNAL VOUCHER VB.NET
	================================================== -->
	<script language="vb" runat="server">
		function rowCount(strConnect as String, strSelect as String)
			Try
				Dim objConnect as new SqlConnection(strConnect)
				objConnect.open()
				Dim objCommand as new SqlCommand(strSelect, objConnect)
				Dim objDataReader as SqlDataReader
				objDataReader = objCommand.ExecuteReader()

				results.innerhtml += "<p> Running: " + strSelect

				Dim c as Integer = 0
				Do While objDataReader.Read()
					c += 1
				Loop

				rowCount = c
			Catch
				results.innerhtml += "<p>ERROR"
			end Try
		end function

		Sub Page_Load(sender As System.Object, e As System.EventArgs)
			results.innerhtml = ""
			If IsPostBack
				Dim xmlString as String = request.form("xmlout")
				Dim x = Server.CreateObject("Microsoft.xmldom")
				x.loadXML(xmlString)

				Dim nr as Integer = x.documentElement.childNodes.length - 2
				Dim srn as String = x.documentElement.childNodes(0).text

				results.innerhtml += "<p>Number of journal entres: " + cstr(nr)
				results.innerhtml += "<p>Preparting to validate srn number: " + cstr(srn)

				Dim connectStr as String = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272"
				Dim sql as String = "SELECT * FROM je WHERE sourceref = " + cstr(srn)
				Dim jeRows as String = rowCount(connectStr, sql)

				if jeRows > 0 then
					results.innerhtml += "<p>SRN was found in " + cstr(jeRows) + " entries, terminating program"
				else
					results.innerhtml += "<p>SRN was found in " + cstr(jeRows) + " entries, proceeding with glmaster checks"

					Dim glValid as Integer = 0
					for i as integer = 0 to nr - 1
						Dim major as String = x.documentElement.childNodes(i+2).childNodes(0).text
						Dim minor as String = x.documentElement.childNodes(i+2).childNodes(1).text
						Dim sub1  as String = x.documentElement.childNodes(i+2).childNodes(2).text
						Dim sub2  as String = x.documentElement.childNodes(i+2).childNodes(3).text

						sql = "SELECT * FROM glmaster WHERE major = " + cstr(major) + " AND "
						sql += " minor = " + cstr(minor) + " AND "
						sql += " sub1  = " + cstr(sub1) + " AND "
						sql += " sub2  = " + cstr(sub2)

						results.innerhtml += "<p>Checking journal entry " + cstr(i + 1) + " out of " + cstr(nr)
						Dim glCount as Integer = rowCount(connectStr, sql)
						results.innerhtml += "<p>Found " + cstr(glCount) + " rows."
						
						if glCount = 1 then
							glValid = glValid + 1
						end if
					next
	
					if glValid <> nr then
						results.innerhtml += "<p>Glmaster checks failed: " + cstr(glValid) + " rows found out of " + cstr(nr)
						results.innerhtml += "<p>Program terminating"
					else
						results.innerhtml += "<p>Glmaster checks passed: " + cstr(glValid) + " rows found out of " + cstr(nr)
						results.innerhtml += "<p>Proceeding with posts"

						Dim numa as Integer
						Dim cn = New SqlConnection(connectStr)
						cn.open()

						results.innerhtml += "<p>Beginning balance updates"
						for i as Integer = 0 to nr - 1
							Dim major as String = x.documentElement.childNodes(i+2).childNodes(0).text
							Dim minor as String = x.documentElement.childNodes(i+2).childNodes(1).text
							Dim sub1  as String = x.documentElement.childNodes(i+2).childNodes(2).text
							Dim sub2  as String = x.documentElement.childNodes(i+2).childNodes(3).text
							Dim ta as String = x.documentElement.childNodes(i+2).childNodes(4).text

							sql = "UPDATE glmaster SET balance = balance + @ta WHERE "
							sql += "major = @major AND "
							sql += "minor = @minor AND "
							sql += "sub1 = @sub1 AND "
							sql += "sub2 = @sub2"

							Try
								Dim glUpdateCmd = New SqlCommand(sql, cn)

								glUpdateCmd.Parameters.AddWithValue("@major", major)
								glUpdateCmd.Parameters.AddWithValue("@minor", minor)
								glUpdateCmd.Parameters.AddWithValue("@sub1", sub1)
								glUpdateCmd.Parameters.AddWithValue("@sub2", sub2)
								glUpdateCmd.Parameters.AddWithValue("@ta", ta)

								numa = glUpdateCmd.ExecuteNonQuery()
								results.innerhtml += "<p>Executed sql: " + sql + " recieved a numa of: " + numa.ToString
							Catch
								results.innerhtml += "<p>ERROR"
							end Try
						next
						results.innerhtml += "<p>Ending balance updates and beginning inserts"
						for i as Integer = 0 to nr - 1
							Dim major as String = x.documentElement.childNodes(i+2).childNodes(0).text
							Dim minor as String = x.documentElement.childNodes(i+2).childNodes(1).text
							Dim sub1  as String = x.documentElement.childNodes(i+2).childNodes(2).text
							Dim sub2  as String = x.documentElement.childNodes(i+2).childNodes(3).text
							Dim ta as String = x.documentElement.childNodes(i+2).childNodes(4).text
							Dim desc as String = x.documentElement.childNodes(i+2).childNodes(5).text

							sql = "INSERT INTO je(sourceref, srseq, jemajor, jeminor, jesub1, jesub2, jedesc, jeamount)"
							sql += "VALUES (@srn, @srseq, @major, @minor, @sub1, @sub2, @desc, @ta)"

							Try
								Dim jeInsertCmd = New SqlCommand(sql, cn)

								jeInsertCmd.Parameters.AddWithValue("@srn", srn)
								jeInsertCmd.Parameters.AddWithValue("@srseq", (i+1).ToString)
								jeInsertCmd.Parameters.AddWithValue("@major", major)
								jeInsertCmd.Parameters.AddWithValue("@minor", minor)
								jeInsertCmd.Parameters.AddWithValue("@sub1", sub1)
								jeInsertCmd.Parameters.AddWithValue("@sub2", sub2)
								jeInsertCmd.Parameters.AddWithValue("@desc", desc)
								jeInsertCmd.Parameters.AddWithValue("@ta", ta)

								numa = jeInsertCmd.ExecuteNonQuery()
								results.innerhtml += "<p>Executed sql: " + sql + " recieved a numa of: " + numa.ToString
							Catch
								results.innerhtml += "<p>ERROR"
							end Try
						next
						results.innerhtml += "<p>Ending je inserts"
						results.innerhtml += "<p>Finished"

					end if
				end if
			End If
		End Sub
	</script>
	<!-- VB END -->

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

  
