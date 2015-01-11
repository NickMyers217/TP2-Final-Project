<!DOCTYPE html>
<%@Page Language="VB" %>
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

	<!-- DELETE VB.NET
	================================================== -->
	<script language="vb" runat="server">
	Sub Page_Load(sender As System.Object, e As System.EventArgs)
		results.InnerHtml = ""
		if IsPostBack then
			Dim major as String = request.form("major")
			Dim minor as String = request.form("minor")
			Dim sub1 as String = request.form("sub1")
			Dim sub2 as String= request.form("sub2")

			Dim sql as String = "SELECT * FROM glmaster WHERE major = " + cstr(major)
			sql = sql + " AND minor = " + cstr(minor)
			sql = sql + " AND sub1 = " + cstr(sub1)
			sql = sql + " AND sub2 = " + cstr(sub2)

			Dim	strConnect as String = "server=AUCKLAND;uid=gl1272;pwd=YWV85updG;database=gl1272"

			Dim c as integer = 0
			Dim bal as double
			try
				results.InnerHtml += "<p>opening a recordset with: " + sql

				Dim objConnect as new SqlConnection(strConnect)
				objConnect.Open()
				Dim objCommand as new SqlCommand(sql, objConnect)
				Dim objDataReader as SqlDataReader 
				objDataReader = objCommand.ExecuteReader()

				Do while objDataReader.Read()
					c += 1
					bal = objDataReader("balance")
				Loop

				objDataReader.Close()
				objConnect.Close()
			catch objError as Exception 
				results.InnerHtml += "<p>Error counting rows."
			end try

			if c = 0 then
				results.InnerHtml += "<p>ACCT NOT FOUND"
			else
				results.InnerHtml += "<p>" + cstr(c) + " records found"

				if bal <> 0 then
					results.InnerHtml += "<p>ACCT balance greater than 0, no deletion can be performed"	
				else
					results.InnerHtml += "<p>Balance is zero, proceeding with deletion"

					Dim numa as Integer
					Dim cn = New SqlConnection(strConnect)
					cn.open()

					sql = "DELETE FROM glmaster WHERE "
					sql += "major = @major AND "
					sql += "minor = @minor AND "
	 				sql += "sub1 = @sub1 AND "
					sql += "sub2 = @sub2"

					try
						results.InnerHtml += "<p>Attempting delete with: " + sql

						Dim deleteGl = New SqlCommand(sql, cn)
						deleteGl.Parameters.AddWithValue("@major", major)
						deleteGl.Parameters.AddWithValue("@minor", minor)
						deleteGl.Parameters.AddWithValue("@sub1", sub1)
						deleteGl.Parameters.AddWithValue("@sub2", sub2)

						numa = deleteGl.ExecuteNonQuery()
						results.InnerHtml += "<p>" + numa.ToString & " record dleted OK." 
						cn.close()
					catch exc as exception
						results.InnerHtml += "<p>ERROR"				
					end try
				end if
			end if
		end if
	end sub
	</script>
	<!-- VB.NET END -->

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
