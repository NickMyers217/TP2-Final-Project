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

	<!-- INSERT VB.NET
	================================================== -->
	<script language="vb" runat="server">
	Sub Page_Load(sender As System.Object, e As System.EventArgs)
		results.InnerHTML=""
	End Sub

	Sub AddAcct(ByVal sender As System.Object, ByVal e As System.EventArgs)
		Dim numa as Integer
		Dim oksw as Integer
		Dim adtemp as String
		Dim adtemp2 as String  
		Dim strUpdateStmt As String
		Dim errortext as String
		strUpdateStmt = "INSERT into glmaster (major,minor,sub1,sub2,acctdesc,balance) VALUES ("  & _
		"@mjrv,@mnrv,@s1v,@s2v,@adv,0.0)"

		If Page.IsValid then

			Try
				Dim myConnection = New SqlConnection("server=AUCKLAND;database=gl1272;uid=gl1272;pwd=YWV85updG;")
				Dim myCommand = New SqlCommand(strUpdateStmt, myConnection)
				adtemp=ad.Text
				results.InnerHtml=""
				error_out.innerHTML=""
				oksw=0

				If Regex.IsMatch(adtemp,"[^A-Za-z0-9\s]") Then
					error_out.innerHTML = "<b>SQL Terminiation containing data other than: alphabetic characters (A thru Z), digits (0 thru 9) and spaces<p>NO data was added the the database."
					oksw = 1
				End If

				adtemp2 = UCase(adtemp)

				If Regex.IsMatch(adtemp2, "SCRIPT") Then
					error_out.innerHTML = "<b>SQL Terminiation containing script tags encountered -- NO UPDATE PERFORMED"
					oksw = 1
				End If	

				If oksw = 0 Then	  
					myCommand.Parameters.AddWithValue( "@adv"   , adtemp )
					myCommand.Parameters.AddWithValue( "@mjrv"  , mjr.Text )
					myCommand.Parameters.AddWithValue( "@mnrv"  , mnr.Text )
					myCommand.Parameters.AddWithValue( "@s1v"   , s1.Text )
					myCommand.Parameters.AddWithValue( "@s2v"   , s2.Text )
					myConnection.Open()
					numa = myCommand.ExecuteNonQuery()
					results.InnerHtml=numa.ToString & " record inserted OK." 
					myConnection.Close()
					Call ClearForm
				End If

			Catch exc as exception

				errortext=UCase(exc.Message)
				if Instr(errortext, "DUPLICATE") > 0 then
					error_out.InnerHtml = "The Account you are trying to add <b>ALREADY EXISTS</b> in the databse and cannot be duplicated.<br>Check your data and try again."
				else
					if Instr(errortext, "DATA TYPE") > 0 then
						error_out.InnerHtml = "Invalid Data Entered. The data was NOT added to the database. Correct the data problem above and try again" 
					else
						error_out.InnerHtml = "<b>* Error while Inserting</b>.<br />" _
						+ exc.Message + "<br />" + exc.Source
					end if
				end if  

			End Try

		else
			results.InnerHTML = ""
			error_out.InnerHtml = "Errors Detected"
		end if

	End Sub

	Sub ClearForm()
		mjr.Text=""
		mnr.Text=""
		s1.Text=""
		s2.Text=""
		ad.Text=""
	End Sub	

	</script>
	<!-- VB.NET END -->
	
	<div class="row">
		<div class="col-lg-6 col-lg-offset-3">
			<div class="well bs-component">

				<form id="Form1" method="post" runat="server">
				<p>Major <asp:TextBox columns="4" ID="mjr" Runat="server" />
				<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>
				<asp:RangeValidator id="RangeValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>
				<asp:CompareValidator id="CompareValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
				<br>Minor <asp:TextBox columns="4" ID="mnr" Runat="server" />
				<asp:RequiredFieldValidator id="RequiredFieldValidator2" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>
				<asp:RangeValidator id="RangeValidator2" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>
				<asp:CompareValidator id="CompareValidator2" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
				<br>Sub 1 <asp:TextBox columns="4" ID="s1" Runat="server" />
				<asp:RequiredFieldValidator id="RequiredFieldValidator3" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>
				<asp:RangeValidator id="RangeValidator3" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>
				<asp:CompareValidator id="CompareValidator3" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
				<br>Sub 2 <asp:TextBox columns="4" ID="s2" Runat="server" />
				<asp:RequiredFieldValidator id="RequiredFieldValidator4" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>
				<asp:RangeValidator id="RangeValidator4" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>
				<asp:CompareValidator id="CompareValidator4" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
				<br>Account Description <asp:TextBox columns="50" ID="ad" Runat="server" />
				<asp:RequiredFieldValidator id="RequiredFieldValidator5" runat="server" Display="Dynamic" ControlToValidate="ad" ErrorMessage="Account description cannot be blank"></asp:RequiredFieldValidator>
				<p><asp:Button Text="Add Account" ID="btnAdd" OnClick="AddAcct" Runat="server" />
				<p><asp:Label ID="numa" ForeColor="#ff0000" Runat=Server/>
				</form>
				<p>
				<font color="#00dd00"><DIV id="results" Runat="server"></DIV></font>
				<font color="#ff0000"><DIV id="error_out" Runat="server"></DIV></font>

			</div>		
		</div>
	</div>

</body>
</html> 

