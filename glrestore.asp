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

	<!-- GLRESTORE ASP
	================================================== -->
	<%
	'*************************** functions, subs, then main
	'pass 1 display a user credentials form for glrestore
	sub pass1
	%>
		<div class="row">
			<div class="col-lg-4 col-lg-offset-4">
				<div class="well bs-component">

					<form class="form-horizontal" action="glrestore.asp" method="POST">
						<fieldset>

							<legend class="text text-primary">Sorry, but you can't drop my tables.</legend>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="text" class="form-control" name="user" id="user" placeholder="User Number">
								</div>
							</div>

							<div class="form-group">
								<div class="col-lg-10">
									<input type="password" class="form-control" name="pword" id="pword" placeholder="Password">
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
	'end of pass1

	function fixdec (amount)
	dim la, lb, dif,temp
		temp=cStr(amount)
		la=InStr(1,temp,".")
		lb=len(temp)
		dif=lb-la
		if la = 0 then
				temp=temp+".00"
		else
				if  dif=0 then
						temp=temp+"00"
			 else
						if dif=1 then
									temp=temp+"0"
						else

						end if
			end if
	 end if
	 fixdec=cStr(temp)
	end function

	sub buildglmaster (cnnew)
	dim create_string
	on error resume next
		 create_string="CREATE TABLE glmaster ("
		 create_string=create_string + "major integer NOT NULL,"
		 create_string=create_string + "minor integer NOT NULL,"
		 create_string=create_string + "sub1 integer,"
		 create_string=create_string + "sub2 integer,"
		 create_string=create_string + "acctdesc varchar(50),"
		 create_string=create_string + "balance numeric (18,2) NOT NULL "
		 create_string=create_string + "primary key (major, minor, sub1, sub2) )"
		 cnnew.execute create_string
		 if noerrors(cnnew, "Task: Create new glmaster table") then
					 Response.write "<br>4. Created new glmaster table OK"
		 else
					 Response.write "<br>4. Create new glmaster table task failed ******<br>"
		 end if
	end sub

	sub dropglmaster (cnnew)
	on error resume next
		 cnnew.execute "DROP TABLE glmaster", numa
		 if noerrors(cnnew, "Task: drop glmaster table") then
				Response.write "<br>3. Dropped old glmaster table OK"
		 else
				Response.write "<br>3. Unable to drop glmaster table. Task Failed ****<br>"
		 end if
		 buildglmaster (cnnew)
	end sub

	sub dropje (cnnew)
	on error resume next
			cnnew.execute "DROP TABLE je", numa
			if noerrors(cnnew, "Task: dropping je table") then
				 Response.write "<p>5. Dropped old je table OK"
			else
				 Response.write "<p>5. Unable to  drop je table. Task Failed *********<br>"
			end if
			buildje(cnnew)
	end sub

	sub buildje (cnnew)
	dim create_string
	on error resume next
		 create_string="CREATE TABLE je ("
		 create_string=create_string +"sourceref integer NOT NULL,"
		 create_string=create_string +"srseq integer NOT NULL,"
		 create_string=create_string +"jemajor integer NOT NULL,"
		 create_string=create_string +"jeminor integer NOT NULL,"
		 create_string=create_string +"jesub1 integer,"
		 create_string=create_string +"jesub2 integer,"
		 create_string=create_string +"jedesc char(50),"
		 create_string=create_string +"jeamount numeric (18,2) NOT NULL "
		 create_string=create_string +"primary key (sourceref, srseq) )"
		 cnnew.execute create_string
		 if noerrors(cnnew, "Task: Create new je table") then
					 Response.write "<br>6. Created new je table OK"
		 else
					 Response.write "<br>6. je table create task failed ***************<br>"
		 end if
	end sub

	Function noerrors (cn , task)
	If Err <> 0 Then
			If cn.Errors.Count = 0 Then

			Else
					 for i = 0 to cn.Errors.Count - 1
								response.write "<p>"
								response.write cn.errors(i)
					 next
			End If
			noerrors= False
	Else
			noerrors = True
	End if
	End Function

	'pass 2 initiate gl restore with provided credentials
	sub pass2
		%>
			<p><b>Restoring the General Ledger Tables:</b><p>
			1. Open SRCGL
			<br>    (the permanent copy of the general ledger database)
		<%
		
		dim cm,cnold,cnnew
		dim sumbal, rsold, Insert_String,create_string,numa,numnew,deval
		dim fdsn, fuid,fpwd
		sumbal=0
		numnew=0
		on error resume next

		'************ user's DSN, Userid and Password ****************************
		'*                                                                       *
		'*                                                                       *
		'*            CHANGE THE THREE VALUES BELOW                              *
		'*            TO THE DATABASE CREDENTIALS PROVIDED                       *
		'*            TO YOU IN CLASS                                            *
		'*                                                                       *
		'*                                                                       *
		'*************************************************************************

		fdsn=cstr(request.form("user"))
		fuid=cstr(request.form("user"))
		fpwd=cstr(request.form("pword"))

		'*********** open the original general ledger database

		set rsold = Server.CreateObject("ADODB.Recordset")
		oldsql="SELECT * FROM glmaster order by major ASC, minor ASC, sub1 ASC, sub2 ASC"
		Response.write "<br>   With sql="+oldsql

		rsold.open oldsql,"DSN=SRCGLC;UID=;PWD=;"

		Response.write "<br>   Opened SRCGL OK"

		'*********** open the user requested user database

		set cnnew = Server.CreateObject("ADODB.Connection")
		cnnew.open fdsn,fuid,fpwd

		if noerrors (cnnew, "Task: Opening database") then '******** top test for user database
							Response.write "<br>2. Opened your "
							Response.write  "<b>"+fdsn+"</b>"
							Response.write " database OK"

							call dropglmaster(cnnew) '****** drop, then create the glmaster table

							Response.write "<p><table class='table table-striped table-hover greenBorder'>"
							Response.write "<thead>"
							Response.write "<tr><th>major</th><th>minor</th><th>sub1</th><th>sub2</th><th>"
							Response.write "acctdesc</th><th>balance</th></tr>"
							Response.write "</thead><tbody>"

							while not rsold.EOF   '****** loop thru the SRCGL table rows,

										 Insert_String = "INSERT INTO glmaster (major,minor,sub1,sub2,acctdesc,balance) VALUES ("
										 Insert_String = Insert_String + cStr(rsold("major")) + ","
										 Insert_String = Insert_String + cStr(rsold("minor")) + ","
										 Insert_String = Insert_String + cStr(rsold("sub1")) + ","
										 Insert_String = Insert_String + cStr(rsold("sub2")) + ","
										 Insert_String = Insert_String + chr(39)+cStr(rsold("acctdesc")) + chr(39) + ","
										 Insert_String = Insert_String + cStr(rsold("balance")) + ")"

										 cnnew.execute Insert_String '****** add the row to the new table

										 Response.write "<tr><td>" '************ show the user
										 Response.write rsold("major")
										 Response.write "</td><td>"
										 Response.write rsold("minor")
										 Response.write "</td><td>"
										 Response.write rsold("sub1")
										 Response.write "</td><td>"
										 Response.write rsold("sub2")
										 Response.write "</td><td>"
										 Response.write rsold("acctdesc")
										 Response.write "</td><td>"

										 deval=rsold("balance") '************* get the balance
										 deval=fixdec(deval)

										 Response.write deval '************** write the balnce to the user
										 Response.write "</td></tr>"

										 sumbal=sumbal+cDbl(rsold("balance")) '************* count the rows
										 numnew=numnew+1

										 rsold.movenext '************  get next record from SRCGL

								 wend '**************************  end of add rows loop
								 rsold.close '*******************  finished with the new glmaster table, so close the SRCGL version

								Response.write "<tr><td colspan="+chr(34)+"5" +chr(34)+"align="
								Response.write chr(34)+"right"+chr(34)+">SUM OF BALANCES</td><td align="+chr(34)+"right"+chr(34)+">"
								Response.write cStr(fixdec(sumbal))
								Response.write "</td></tr><tr><td colspan="+chr(34)+"6"+chr(34)+">"
								Response.write cStr(numnew)
								Response.write " rows in the new glmaster table</td></tr></tbody></table><br>"

							 call dropje(cnnew) '***** now drop, then create the je table

		else '**************************** couldn't open the user's database -- bail out!

					 Response.write "<p>2. Open task failed on database: "
					 Response.write cStr(fdsn)
					 Response.write "<p>3. Drop current glmaster table NOT attempted"
					 Response.write "<p>4. Create glmaster table NOT attempted"
					 Response.write "<p>5. Drop current je table NOT attempted"
					 Response.write "<p>6. Create je table NOT attempted"

		end if

		'**************************** back to HTML for the finish

		cnnew.close
		set cnnew=nothing
		set rsold=nothing

		%>
			<br>
			<p>
			<b>END OF THE GENERAL LEDGER RESTORE
			<p>
		<%
	end sub
	'The end of pass 2 and gl restore


	'
	'*** top of main
	'
	tokenvalue = request.form("token")
	select case tokenvalue
	case ""
		call pass1
	case "2"
		call pass2
	end select

	%>
	<!-- ASP END -->

	</div> <!-- end of container -->

</body>
</html>



