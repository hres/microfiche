<!-- #include file="adovbs.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>ATIP - Ingredients</title>
</head>

<body leftmargin="3" topmargin="1" rightmargin="3" bottommargin="0">

<%
Dim dcnDB                       'ADODB.connection
Dim strDBLocation               'String to hold database location
Dim rsIngredients              'Recordset
dim SQL							'String to hold the rsAll_Products sql query
dim QryStr						'String to hold query string


strDBLocation = Server.MapPath("database\ATIP.mdb")
	
Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
dcnDB.Open

Set rsIngredients = Server.CreateObject("ADODB.Recordset")

SQL = "SELECT Medicinal_Ingredients.Field1, Medicinal_Ingredients.Number, Medicinal_Ingredients.Name, Medicinal_Ingredients.Reference FROM Medicinal_Ingredients WHERE Medicinal_Ingredients.ID > 1 "

%>

<br>

<form action="Ingredients.asp" method="post">
		
	<table>
		<tr>
			<td>Number : <input type="text" name="Number" size="10"></td>
			<td>Name : <input type="text" name="Name" size="100"></td>
			<td>Reference : <input type="text" name="Reference" size="30"></td>
			<td><input type="submit" name="Search" value="  Search  "></td>
			<td><a href="Ingredients.asp">Reset And View All</a></td>	</td>
		</tr>	
	</table>
	
</form>

<hr>

<%

if not request.Form("Number") = "" then

			SQL = SQL & " AND Medicinal_Ingredients.Number like '%" & request.Form("Number") & "%'"

end if

if not request.Form("Name") = "" then

			SQL = SQL & " AND Medicinal_Ingredients.Name like '%" & request.Form("Name") & "%'"

end if

if not request.Form("Reference") = "" then

			SQL = SQL & " AND Medicinal_Ingredients.Reference like '%" & request.Form("Reference") & "%'"

end if



rsIngredients.Open SQL, dcnDB, adOpenStatic, adUseClient


	If Request.QueryString("page") = "" Then
		intPage = 1	
	Else
		intPage = Request.QueryString("page")
	End If

	rsIngredients.PageSize = 35		
	rsIngredients.CacheSize = rsIngredients.PageSize
	intPageCount = rsIngredients.PageCount
	intRecordCount = rsIngredients.RecordCount


	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1

	If intRecordCount > 0 Then

		rsIngredients.AbsolutePage = intPage
		intStart = rsIngredients.AbsolutePosition

		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsIngredients.PageSize - 1)
		End if
		
	End If


If Not rsIngredients.EOF Then
%>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<th></th>
		<th align="left">Number</th>
		<th align="left">Name</th>
		<th align="left">Reference</th>					
	</tr>


<%	
		For intRecord = 1 to rsIngredients.PageSize

			%>												
							
				<tr>
	                <td width="100"><%= rsIngredients.Fields("Field1") %></td>					
	                <td width="150"><%= rsIngredients.Fields("Number") %></td>					
	                <td width="600"><%= rsIngredients.Fields("Name") %></td>					
	                <td width="300"><%= rsIngredients.Fields("Reference") %></td>																				
				</tr>							

			<%
			
				
		rsIngredients.MoveNext
		If rsIngredients.EOF Then Exit for
		Next
%>
		</table>		
<%
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
						<td width="200"><a href="Ingredients.asp?<%= QryStr %>">First page</a></td>				   										

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
						<td width="200"><a href="Ingredients.asp?<%= QryStr %>">Previous Page</a></td>
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
						<td width="200"><a href="Ingredients.asp?<%= QryStr %>">Next Page</a></td>								
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
						<td width="200"><a href="Ingredients.asp?<%= QryStr %>">Last page</a></td>
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


	rsIngredients.Close
	dcnDB.Close
	Set rsManufacturer = Nothing
	Set dcnDB = Nothing

%>





</body>
</html>
