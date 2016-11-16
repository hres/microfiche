<!-- #include file="adovbs.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>ATIP - Manufacturer</title>
</head>

<body leftmargin="3" topmargin="1" rightmargin="3" bottommargin="0">

<%
Dim dcnDB                       'ADODB.connection
Dim strDBLocation               'String to hold database location
Dim rsManufacturer              'Recordset
dim SQL							'String to hold the rsAll_Products sql query
dim QryStr						'String to hold query string


strDBLocation = Server.MapPath("database\ATIP.mdb")
	
Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
dcnDB.Open

Set rsManufacturer = Server.CreateObject("ADODB.Recordset")

SQL = "SELECT Manufacturers.ID, Manufacturers.ManuCode, Manufacturers.Manufacturer1, Manufacturers.Manufacturer2, Manufacturers.Manufacturer3, Manufacturers.Mailing1, Manufacturers.Mailing2, Manufacturers.Mailing3, Manufacturers.Distributor1, Manufacturers.Distributor2, Manufacturers.Distributor3, Manufacturers.sppn, Manufacturers.gp, Manufacturers.gpm, Manufacturers.total FROM Manufacturers WHERE Manufacturers.ID >1 "

	if not request.QueryString("MFRCode") = "" then

		SQL = SQL & " AND Manufacturers.ManuCode like '%" & request.QueryString("MFRCode") & request.QueryString("Reg") & "%'"
		
	end if
%>

<br>

<form action="Manufacturer.asp" method="post">
		
	<input type="text" name="MFR" size="10">
	
	<input type="submit" name="Search" value="  Search  ">
	<a href="Manufacturer.asp">Reset And View All</a></td>	

</form>

<hr>

<%

if not request.Form("MFR") = "" then

			SQL = SQL & " AND Manufacturers.ManuCode like '%" & request.Form("MFR") & "%'"

end if



rsManufacturer.Open SQL, dcnDB, adOpenStatic, adUseClient


	If Request.QueryString("page") = "" Then
		intPage = 1	
	Else
		intPage = Request.QueryString("page")
	End If

	rsManufacturer.PageSize = 3		
	rsManufacturer.CacheSize = rsManufacturer.PageSize
	intPageCount = rsManufacturer.PageCount 
	intRecordCount = rsManufacturer.RecordCount 


	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1

	If intRecordCount > 0 Then

		rsManufacturer.AbsolutePage = intPage
		intStart = rsManufacturer.AbsolutePosition

		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsManufacturer.PageSize - 1)
		End if
		
	End If


If Not rsManufacturer.EOF Then

	
		For intRecord = 1 to rsManufacturer.PageSize

			%>												
			<br>
			<table border="0" cellspacing="0" cellpadding="0">
					
				<tr>
	                <td><%= rsManufacturer.Fields("ManuCode") %></td>					
				</tr>
				
				<Tr>
					<td width="150">Manufacturer Address</td>
					<td width="900"><%= rsManufacturer.Fields("Manufacturer1") %></td>
					<td><%= rsManufacturer.Fields("sppn") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Manufacturer2") %></td>
					<td><%= rsManufacturer.Fields("gp") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Manufacturer3") %></td>
					<td><%= rsManufacturer.Fields("gpm") %></td>
				</TR>
				
				<tr><td>&nbsp;</td></tr>
								
				<Tr>
					<td width="150">Mailing Address</td>
					<td width="900"><%= rsManufacturer.Fields("Mailing1") %></td>
					<td><%= rsManufacturer.Fields("total") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Mailing2") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Mailing3") %></td>
				</TR>
				
				<tr><td>&nbsp;</td></tr>				
				
				<Tr>
					<td width="150">Distributor Address</td>
					<td width="900"><%= rsManufacturer.Fields("Distributor1") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Distributor2") %></td>
				</TR>
	
				<Tr>
					<td></td>
					<td><%= rsManufacturer.Fields("Distributor3") %></td>
				</TR>
				
				<tr><td>&nbsp;</td></tr>
				
			</table>
			<hr>

			<%
			
				
		rsManufacturer.MoveNext
		If rsManufacturer.EOF Then Exit for
		Next
		
		%>		
		
		<br>
		<div align="center">
		<table>
			<tr>
			
				<%
					If cInt(intPage) > 1 Then
					%>		
						<%
							
						QryStr = "page=1"
						%>
						<td width="200"><a href="Manufacturer.asp?<%= QryStr %>">First page</a></td>				   										

					<%
					Else
					%>
						<td width="200"></td>				   								
					<%
					End IF


					If cInt(intPage) > 1 Then
					%>				   			
						
						<%
							
						QryStr = "page=" & intPage - 1
						
						%>												
						<td width="200"><a href="Manufacturer.asp?<%= QryStr %>">Previous Page</a></td>
					<%
					Else
					%>
						<td width="200"></td>
					<%
					End IF


					If cInt(intPage) < cInt(intPageCount) Then
					%>

						<%
							
						QryStr = "page=" & intPage + 1

						%>																   		
						<td width="200"><a href="Manufacturer.asp?<%= QryStr %>">Next Page</a></td>								
					<%
					Else
					%>
				   			<td width="200"></td>							
					<%
					End If


					If cInt(intPage) < cInt(intPageCount) Then
					%>
				   		
						<%
							
						QryStr = "page=" & intPageCount
						
						%>																		
						<td width="200"><a href="Manufacturer.asp?<%= QryStr %>">Last page</a></td>
					<%
					Else
					%>
				   			<td></td>							
					<%
					End If
					%>								
			</tr>
		</table>
		
		</div>		
		
<%		
		
Else
	%>
		No Entry Found	
	<%
End If	


	rsManufacturer.Close
	dcnDB.Close
	Set rsManufacturer = Nothing
	Set dcnDB = Nothing

%>





</body>
</html>
