<!-- #include file="adovbs.inc" -->
<html>
<head>
<title>ATIP - All Products Active</title>
</head>
<body leftmargin="3" topmargin="1" rightmargin="3" bottommargin="0">

<%
Dim dcnDB                       'ADODB.connection
Dim strDBLocation               'String to hold database location
Dim rsAll_Products                      'Recordset
dim SQL							'String to hold the rsAll_Products sql query
dim QryStr						'String to hold query string
Dim VMFR
Dim VRegion
Dim VClass	
Dim VForm
Dim VRoute
Dim VActIng
Dim VAccessNum
Dim Flist
Dim Comm
Dim VSort
Dim VSortName
Dim VProductName
Dim VDIN

SQL = "SELECT All_Products.Field1, All_Products.DiscontinuedDate, All_Products.Field3, All_Products.AccessNum, All_Products.MFRCode, All_Products.RegionCode, All_Products.ClassNum, All_Products.NotificationDate, All_Products.ProductName, All_Products.DIN, All_Products.Form, All_Products.Route, All_Products.ActiveIngGroup FROM All_Products WHERE All_Products.ID > 1 "
	
strDBLocation = Server.MapPath("database\ATIP.mdb")
	
Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBLocation
dcnDB.Open

Set rsAll_Products = Server.CreateObject("ADODB.Recordset")
%>


	<table border="0" cellspacing="0" cellpadding="0">

		<tr>
		
			<th width="120" align="right">Access #:</th>
			<th width="140" align="right">Act. Ing. Group:</th>
			<th width="120" align="right">DIN #:</th>
			<th width="500" align="right">Product Name:</th>
			<th width="250" align="right">Sort By:</th>
			<th width="80"></th>
								
		</tr>
	
		<tr>

			<form action="All_Products.asp" method="post">
		
			<td align="right">			
					
				<input type="text" name="AccessNum" size="10">
			
			</td>

			<td align="right">
			

					<%
					if not request.Form("ActIng") = "" then
					%>
					
						<input type="text" name="ActIng" value="<%= request.Form("ActIng") %>" size="10">
					
					<%
					else
					
						if not request.QueryString("ActIng") = "" then
						%>					
							<input type="text" name="ActIng" value="<%= request.QueryString("ActIng") %>" size="10">					
						<%
						else
						%>					
							<input type="text" name="ActIng" value="" size="10">					
						<%
						end if

					end if
			
					%>
			
			</td>
			
			
			<td align="right">
			

					<%
					if not request.Form("DIN") = "" then
					%>
					
						<input type="text" name="DIN" value="<%= request.Form("DIN") %>" size="10">
					
					<%
					else
					
						if not request.QueryString("DIN") = "" then
						%>					
							<input type="text" name="DIN" value="<%= request.QueryString("DIN") %>" size="10">					
						<%
						else
						%>					
							<input type="text" name="DIN" value="" size="10">					
						<%
						end if

					end if
			
					%>
			
			</td>
			
			
			<td align="right">
			

					<%
					if not request.Form("ProductName") = "" then
					%>
					
						<input type="text" name="ProductName" value="<%= request.Form("ProductName") %>" size="60">
					
					<%
					else
					
						if not request.QueryString("ProductName") = "" then
						%>					
							<input type="text" name="ProductName" value="<%= request.QueryString("ProductName") %>" size="60">					
						<%
						else
						%>					
							<input type="text" name="ProductName" value="" size="60">					
						<%
						end if

					end if
			
					%>
			
			</td>
						
			<td align="right">
				<select name="Sort">
			
					<%
					VSort = ""
					VSortName = "ProductName"
					if request.form("Sort") = "ProductName" then
						VSort = " Selected"
						VSortName = "ProductName"
					elseif request.QueryString("Sort") = "ProductName" then
						VSort = " Selected"
						VSortName = "ProductName"
					end if	
					%>
					<option value="ProductName"<%= VSort %>>Product Name</option>

					<%
					VSort = ""					
					if request.form("Sort") = "Field1" then
						VSort = " Selected"
						VSortName = "Field1"
					elseif request.QueryString("Sort") = "Field1" then
						VSort = " Selected"
						VSortName = "Field1"
					end if	
					%>
					<option value="Field1"<%= VSort %>>T1</option>
					
					<%
					VSort = ""
					if request.form("Sort") = "DiscontinuedDate" then
						VSort = " Selected"
						VSortName = "DiscontinuedDate"
					elseif request.QueryString("Sort") = "DiscontinuedDate" then
						VSort = " Selected"
						VSortName = "DiscontinuedDate"
					end if	
					%>										
					<option value="DiscontinuedDate"<%= VSort %>>Disc. Date</option>
					
					<%
					VSort = ""
					if request.form("Sort") = "Field3" then
						VSort = " Selected"
						VSortName = "Field3"
					elseif request.QueryString("Sort") = "Field3" then
						VSort = " Selected"
						VSortName = "Field3"
					end if	
					%>					
					<option value="Field3"<%= VSort %>>T3</option>
					
					<%
					VSort = ""
					if request.form("Sort") = "AccessNum" then
						VSort = " Selected"
						VSortName = "AccessNum"
					elseif request.QueryString("Sort") = "AccessNum" then
						VSort = " Selected"
						VSortName = "AccessNum"
					end if	
					%>					
					<option value="AccessNum"<%= VSort %>>Access #</option>

					<%
					VSort = ""
					if request.form("Sort") = "MFRCode" then
						VSort = " Selected"
						VSortName = "MFRCode"
					elseif request.QueryString("Sort") = "MFRCode" then
						VSort = " Selected"
						VSortName = "MFRCode"
					end if	
					%>
					<option value="MFRCode"<%= VSort %>>MFR</option>

					<%
					VSort = ""
					if request.form("Sort") = "RegionCode" then
						VSort = " Selected"
						VSortName = "RegionCode"
					elseif request.QueryString("Sort") = "RegionCode" then
						VSort = " Selected"
						VSortName = "RegionCode"
					end if	
					%>
					<option value="RegionCode"<%= VSort %>>Region</option>

					<%
					VSort = ""
					if request.form("Sort") = "ClassNum" then
						VSort = " Selected"
						VSortName = "ClassNum"
					elseif request.QueryString("Sort") = "ClassNum" then
						VSort = " Selected"
						VSortName = "ClassNum"
					end if	
					%>
					<option value="ClassNum"<%= VSort %>>Class</option>

					<%
					VSort = ""
					if request.form("Sort") = "NotificationDate" then
						VSort = " Selected"
						VSortName = "NotificationDate"
					elseif request.QueryString("Sort") = "NotificationDate" then
						VSort = " Selected"
						VSortName = "NotificationDate"
					end if	
					%>
					<option value="NotificationDate"<%= VSort %>>Not. Date</option>

					<%
					VSort = ""
					if request.form("Sort") = "DIN" then
						VSort = " Selected"
						VSortName = "DIN"
					elseif request.QueryString("Sort") = "DIN" then
						VSort = " Selected"
						VSortName = "DIN"
					end if	
					%>
					<option value="DIN"<%= VSort %>>DIN</option>

					<%
					VSort = ""
					if request.form("Sort") = "Form" then
						VSort = " Selected"
						VSortName = "Form"
					elseif request.QueryString("Sort") = "Form" then
						VSort = " Selected"
						VSortName = "Form"
					end if	
					%>
					<option value="Form"<%= VSort %>>Form</option>

					<%
					VSort = ""
					if request.form("Sort") = "Route" then
						VSort = " Selected"
						VSortName = "Route"
					elseif request.QueryString("Sort") = "Route" then
						VSort = " Selected"
						VSortName = "Route"
					end if	
					%>
					<option value="Route"<%= VSort %>>Route</option>

					<%
					VSort = ""
					if request.form("Sort") = "ActiveIngGroup" then
						VSort = " Selected"
						VSortName = "ActiveIngGroup"
					elseif request.QueryString("Sort") = "ActiveIngGroup" then
						VSort = " Selected"
						VSortName = "ActiveIngGroup"
					end if	
					%>
					<option value="ActiveIngGroup"<%= VSort %>>Act. Ing. Group</option>
			
			
			
				</select>
			</td>					

			<td width="80" align="right"><input type="submit" name="SortB" value="  Sort  "></td>
			
			
		</tr>
		
	</table>
	
	<table border="0" cellspacing="0" cellpadding="0">
	
		<tr>

			<th width="100" align="right">MFR:</th>
			<th width="100" align="right">Region:</th>
			<th width="100" align="right">Class:</th>
			<th width="115" align="right">Form:</th>
			<th width="200" align="right">Route:</th>
			<th width="500" align="right">&nbsp;</th>
			<th width="125" align="right">&nbsp;</th>
		
		</tr>
		
		
		<tr>
		
			<td align="right">
			
				<select name="MFR">

					<%
					
					if not Request.QueryString("MFRCode") = "" then
					%>
					
						<option value="<%= Request.QueryString("MFRCode") %>"><%= Request.QueryString("MFRCode") %></option>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
					
					<%
					else
					
						if not request.Form("MFR") = "" then
						%>
							<option value="<%= request.Form("MFR") %>"><%= request.Form("MFR") %></option>
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
						<%
						else

							%>										
						
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
															
							<%

						end if
					
					end if

						SQL = "SELECT DISTINCT All_Products.MFRCode FROM All_Products ORDER BY MFRCode"

						rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient
						
						if not rsAll_Products.BOF or rsAll_Products.EOF then

							do while rsAll_Products.EOF = false
							
								if not (Request.QueryString("MFRCode") = rsAll_Products.Fields("MFRCode") or rsAll_Products.Fields("MFRCode") = request.form("MFR")) then
								%>
									<option value="<%= rsAll_Products.Fields("MFRCode") %>"><%= rsAll_Products.Fields("MFRCode") %></option>								
								<%
								end if
								rsAll_Products.movenext	
								
							loop

						end if
						rsAll_Products.Close
					%>
		
		
		
				</select>
			
			</td>

			<td align="right">
			
				<select name="Region">

					<%
					if not Request.QueryString("Reg") = "" then
					%>
					
						<option value="<%= Request.QueryString("Reg") %>"><%= Request.QueryString("Reg") %></option>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
					
					<%
					else
					
						if not request.Form("Region") = "" then
						%>
							<option value="<%= request.Form("Region") %>"><%= request.Form("Region") %></option>
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
						<%
						else

							%>										
						
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
															
							<%

						end if
					
					end if

						SQL = "SELECT DISTINCT All_Products.RegionCode FROM All_Products ORDER BY RegionCode"

						rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient
						
						if not rsAll_Products.BOF or rsAll_Products.EOF then

							do while rsAll_Products.EOF = false
							
								if not (Request.QueryString("Reg") = rsAll_Products.Fields("RegionCode") or rsAll_Products.Fields("RegionCode") = request.form("Region")) then
								%>
									<option value="<%= rsAll_Products.Fields("RegionCode") %>"><%= rsAll_Products.Fields("RegionCode") %></option>								
								<%
								end if
								rsAll_Products.movenext	
								
							loop

						end if
						rsAll_Products.Close
					%>
		
		
				</select>
			
			</td>

			<td align="right">
			
				<select name="Class">

					<%
					if not Request.QueryString("Class") = "" then
					%>
					
						<option value="<%= Request.QueryString("Class") %>"><%= Request.QueryString("Class") %></option>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
					
					<%
					else
					
						if not request.Form("Class") = "" then
						%>
							<option value="<%= request.Form("Class") %>"><%= request.Form("Class") %></option>
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
						<%
						else

							%>										
						
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
															
							<%

						end if
					
					end if
					
						SQL = "SELECT DISTINCT All_Products.ClassNum FROM All_Products ORDER BY ClassNum"

						rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient
						
						if not rsAll_Products.BOF or rsAll_Products.EOF then

							do while rsAll_Products.EOF = false
							
								if not (Request.QueryString("Class") = rsAll_Products.Fields("ClassNum") or rsAll_Products.Fields("ClassNum") = request.form("Class")) then
								%>
									<option value="<%= rsAll_Products.Fields("ClassNum") %>"><%= rsAll_Products.Fields("ClassNum") %></option>								
								<%
								end if
								rsAll_Products.movenext	
								
							loop

						end if
						rsAll_Products.Close
					%>
		
		
				</select>
			
			</td>

			<td align="right">
			
				<select name="Form">

					<%
					if not Request.QueryString("Form") = "" then
					%>
					
						<option value="<%= Request.QueryString("Form") %>"><%= Request.QueryString("Form") %></option>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
					
					<%
					else
					
						if not request.Form("Form") = "" then
						%>
							<option value="<%= request.Form("Form") %>"><%= request.Form("Form") %></option>
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
						<%
						else

							%>										
						
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
															
							<%

						end if
					
					end if

						SQL = "SELECT DISTINCT All_Products.Form FROM All_Products ORDER BY Form"

						rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient
						
						if not rsAll_Products.BOF or rsAll_Products.EOF then

							do while rsAll_Products.EOF = false
							
								if not (Request.QueryString("Form") = rsAll_Products.Fields("Form") or rsAll_Products.Fields("Form") = request.form("Form")) then
								%>
									<option value="<%= rsAll_Products.Fields("Form") %>"><%= rsAll_Products.Fields("Form") %></option>								
								<%
								end if
								rsAll_Products.movenext	
								
							loop

						end if
						rsAll_Products.Close
					%>
																					
				</select>
			
			</td>

			<td align="right">
			
				<select name="Route">

					<%
					if not Request.QueryString("Route") = "" then
					%>
					
						<option value="<%= Request.QueryString("Route") %>"><%= Request.QueryString("Route") %></option>
						<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
					
					<%
					else
					
						if not request.Form("Route") = "" then
						%>
							<option value="<%= request.Form("Route") %>"><%= request.Form("Route") %></option>
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>						
						<%
						else

							%>										
						
							<option value="">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
															
							<%

						end if
					
					end if

						SQL = "SELECT DISTINCT All_Products.Route FROM All_Products ORDER BY Route"

						rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient
						
						if not rsAll_Products.BOF or rsAll_Products.EOF then

							do while rsAll_Products.EOF = false
							
								if not (Request.QueryString("Route") = rsAll_Products.Fields("Route") or rsAll_Products.Fields("Route") = request.form("Route")) then	
								%>
									<option value="<%= rsAll_Products.Fields("Route") %>"><%= rsAll_Products.Fields("Route") %></option>								
								<%
								end if
								rsAll_Products.movenext	
								
							loop

						end if
						rsAll_Products.Close
					%>

		
				</select>
			
		
			</td>
			
			<td align="right"><input type="submit" name="Search" value="  Search  "></td>
			<td align="center"><a href="All_Products.asp">Reset And <br>View All</a></td>	
				
			</form>
		
		</tr>
		
	</table>
	
	<table border="0" cellspacing="0" cellpadding="0">
			
<%

	if not Request.QueryString("MFRCode") = "" then
		VMFR = Request.QueryString("MFRCode")
	end if				
					
	if not Request.QueryString("Reg") = "" then
		VRegion = Request.QueryString("Reg")
	end if				
	
	if not Request.QueryString("Class") = "" then
		VClass = Request.QueryString("Class")
	end if				
	
	if not Request.QueryString("Form") = "" then
		VForm = Request.QueryString("Form")
	end if				
	
	if not Request.QueryString("Route") = "" then
		VRoute = Request.QueryString("Route")
	end if				
	
	if not Request.QueryString("ActIng") = "" then
		VActIng = Request.QueryString("ActIng")
	end if	

	if not Request.QueryString("DIN") = "" then
		VDIN = Request.QueryString("DIN")
	end if
	
	if not Request.QueryString("ProductName") = "" then
		VProductName = Request.QueryString("ProductName")
	end if		
	

	SQL = "SELECT All_Products.Field1, All_Products.DiscontinuedDate, All_Products.Field2, All_Products.Field3, All_Products.AccessNum, All_Products.MFRCode, All_Products.RegionCode, All_Products.ClassNum, All_Products.NotificationDate, All_Products.ProductName, All_Products.DIN, All_Products.Form, All_Products.Route, All_Products.ActiveIngGroup FROM All_Products WHERE All_Products.ID > 1 "
		

	if not VActIng = "" then
	
		SQL = SQL & " AND All_Products.ActiveIngGroup like '%" & VActIng & "%'"
		
	end if
	
	if not VClass = "" then
	
		SQL = SQL & " AND All_Products.ClassNum = '" & VClass & "'"
		
	end if
	
	if not VForm = "" then
	
		SQL = SQL & " AND All_Products.Form = '" & VForm & "'"
		
	end if
	
	if not VMFR = "" then
	
		SQL = SQL & " AND All_Products.MFRCode = '" & VMFR & "'"
		
	end if
	
	if not VRegion = "" then
	
		SQL = SQL & " AND All_Products.RegionCode = '" & VRegion & "'"
		
	end if
	
	if not VRoute = "" then
	
		SQL = SQL & " AND All_Products.Route = '" & VRoute & "'"
		
	end if
	
	if not VDIN = "" then
	
		SQL = SQL & " AND All_Products.DIN like '%" & VDIN & "%'"
		
	end if
	
	if not VProductName = "" then
	
		SQL = SQL & " AND All_Products.ProductName like '%" & VProductName & "%'"
		
	end if
	
	
	if VMFR = "" and VRegion = "" and VClass = "" and VForm = "" and VRoute = "" and VActIng = "" and VProductName = "" and VDIN = "" then
	
		if not request.Form("AccessNum") = "" then
				
			SQL = SQL & " AND All_Products.AccessNum like '%" & request.Form("AccessNum") & "%'"
			VAccessNum = request.Form("AccessNum")
		
		end if
		
		if not request.Form("ActIng") = "" then
	
			SQL = SQL & " AND All_Products.ActiveIngGroup like '%" & request.Form("ActIng") & "%'"
			VActIng =  request.Form("ActIng")
		
		end if
		
		if not request.Form("Class") = "" then
		
			SQL = SQL & " AND All_Products.ClassNum = '" & request.Form("Class") & "'"
			VClass = request.Form("Class")
			
		end if
		
		if not request.Form("Form") = "" then
		
			SQL = SQL & " AND All_Products.Form = '" & request.Form("Form") & "'"
			VForm = request.Form("Form")
			
		end if
		
		if not request.Form("MFR") = "" then
		
			SQL = SQL & " AND All_Products.MFRCode = '" & request.Form("MFR") & "'"
			VMFR = request.Form("MFR")
			
		end if
		
		if not request.Form("Region") = "" then
		
			SQL = SQL & " AND All_Products.RegionCode = '" & request.Form("Region") & "'"
			VRegion = request.Form("Region")
			
		end if
		
		if not request.Form("Route") = "" then
		
			SQL = SQL & " AND All_Products.Route = '" & request.Form("Route") & "'"
			VRoute = request.Form("Route")
			
		end if
		
		if not request.Form("DIN") = "" then
		
			SQL = SQL & " AND All_Products.DIN like '%" & request.Form("DIN") & "%'"
			VDIN = request.Form("DIN")
			
		end if
		
		if not request.Form("ProductName") = "" then
		
			SQL = SQL & " AND All_Products.ProductName like '%" & request.Form("ProductName") & "%'"
			VProductName = request.Form("ProductName")
			
		end if
		
	end if
		
	if not request.form("Sort") = "" then
		SQL = SQL & " ORDER BY All_Products." & request.form("Sort")
	else
		SQL = SQL & " ORDER BY All_Products." & VSortName
	end if
	
	if VAccessNum = "" and VMFR = "" and VRegion = "" and VClass = "" and VForm = "" and VRoute = "" and VActIng = "" and VDIN = "" and VProductName = "" then
	
		Flist = "Current Filters: None"

	else
	
		Flist = "Current Filters: "

	
		if not VAccessNum = "" then
		
			Flist = Flist & "Access # = " & VAccessNum
			Comm = ",  "
		
		end if
		
		if not VMFR = "" then
		
			Flist = Flist & Comm & "MFR = " & VMFR
			Comm = ",  "
		
		end if
		
		if not VRegion = "" then
		
			Flist = Flist & Comm & "Region = " & VRegion
			Comm = ",  "			
		
		end if
		
		if not VClass = "" then
		
			Flist = Flist & Comm & "Class = " & VClass
			Comm = ",  "			
		
		end if
		
		if not VForm = "" then
		
			Flist = Flist & Comm & "Form = " & VForm
			Comm = ",  "			
		
		end if
		
		if not VRoute = "" then
		
			Flist = Flist & Comm & "Route = " & VRoute
			Comm = ",  "
		
		end if
		
		if not VActIng = "" then
		
			Flist = Flist & Comm & "Act. Ing. Group = " & VActIng
			Comm = ",  "
			
		end if
		
		if not VDIN = "" then
		
			Flist = Flist & Comm & "DIN = " & VDIN
			Comm = ",  "
		
		end if
		
		if not VProductName = "" then
		
			Flist = Flist & Comm & "Product Name = " & VProductName
		
		end if
	
	end if	
				
	rsAll_Products.Open SQL, dcnDB, adOpenStatic, adUseClient

%>	

	<tr>
	
		<td height="25" colspan="11" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%= FList %></td>
	
	</tr>

</table>	

<hr width="1200">	

<%	
	
	If Request.QueryString("page") = "" Then
		intPage = 1	
	Else
		intPage = Request.QueryString("page")
	End If

	rsAll_Products.PageSize = 30		
	rsAll_Products.CacheSize = rsAll_Products.PageSize
	intPageCount = rsAll_Products.PageCount
	intRecordCount = rsAll_Products.RecordCount


	If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
	If CInt(intPage) <= 0 Then intPage = 1

	If intRecordCount > 0 Then

		rsAll_Products.AbsolutePage = intPage
		intStart = rsAll_Products.AbsolutePosition

		If CInt(intPage) = CInt(intPageCount) Then
			intFinish = intRecordCount
		Else
			intFinish = intStart + (rsAll_Products.PageSize - 1)
		End if
		
	End If


If Not rsAll_Products.EOF Then

		%>
		<table border="0" cellspacing="0" cellpadding="0">

			<tr>
                <th width="100" align="left" valign="top">T1</th>
                <th width="60" align="left" valign="top">T2</th>
                <th width="50" align="left" valign="top">Disc. Date</th>
                <th width="35" align="left" valign="top">T3</th>
                <th width="70" align="left" valign="top">Access #</th>
				<th width="50" align="left" valign="top">MFR</th>
                <th width="60" align="left" valign="top">Region</th>
                <th width="40" align="left" valign="top">Class</th>
                <th width="50" align="left" valign="top">Not. Date</th>
                <th width="400" align="left" valign="top">Product Name</th>
                <th width="70" align="left" valign="top">DIN</th>
                <th width="80" align="left" valign="top">Form</th>
                <th width="80" align="left" valign="top">Route</th>
                <th width="100" align="left" valign="top">Act. Ing. Group</th>
			</tr>


		<%
		
		currcol = "#0f0f0f"
		
		For intRecord = 1 to rsAll_Products.PageSize

			if currcol = "#F2F2F2" then
			
				currcol = "#FFFFFF"
			
			else
			
				currcol = "#F2F2F2"
			
			end if

			%>						
					
				<tr>
	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("Field1") %></td>
					<td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("Field2") %></td>
	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("DiscontinuedDate") %></td>
	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("Field3") %></td>
	                <td bgcolor=<%= currcol %>><a href="pdf/<%= rsAll_Products.Fields("AccessNum") %>.pdf" target="_blank"><%= rsAll_Products.Fields("AccessNum") %></a></td>

			<%

					QryStr = "MFRCode=" & rsAll_Products.Fields("MFRCode")


					if not VActIng = "" then
					
						QryStr = QryStr & "&ActIng=" & VActIng
					
					end if

					if not VClass = "" then
					
						QryStr = QryStr & "&Class=" & VClass
				
					end if

					if not VForm = "" then
					
						QryStr = QryStr & "&Form=" & VForm
				
					end if
					
					if not VRegion = "" then
				
						QryStr = QryStr & "&Reg=" & VRegion
				
					end if

					if not VRoute = "" then
				
						QryStr = QryStr & "&Route=" & VRoute
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
						QryStr = QryStr & "&Sort=" & VSortName

					%>
					
					<td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("MFRCode") %></a></td>
	
					
					
					
					<%
					QryStr = "Reg=" & rsAll_Products.Fields("RegionCode")
					
					
					if not VActIng = "" then
					
						QryStr = QryStr & "&ActIng=" & VActIng
					
					end if

					if not VClass = "" then
					
						QryStr = QryStr & "&Class=" & VClass
				
					end if

					if not VForm = "" then
					
						QryStr = QryStr & "&Form=" & VForm
				
					end if
					
					if not VMFR = "" then
				
						QryStr = QryStr & "&MFRCode=" & VMFR
				
					end if
					
					if not VRoute = "" then
				
						QryStr = QryStr & "&Route=" & VRoute
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
 					QryStr = QryStr & "&Sort=" & VSortName
					
					%>
					
					<td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("RegionCode") %></a>&nbsp;&nbsp;&nbsp;<a href="Manufacturer.asp?MFRCode=<%= rsAll_Products.Fields("MFRCode") %>&Reg=<%= rsAll_Products.Fields("RegionCode") %>" target="_blank">@</a></td>
	
					
					
					
					<%
					
					QryStr = "Class=" & rsAll_Products.Fields("ClassNum")
					
					if not VActIng = "" then
					
						QryStr = QryStr & "&ActIng=" & VActIng
					
					end if

					if not VForm = "" then
					
						QryStr = QryStr & "&Form=" & VForm
				
					end if
					
					if not VMFR = "" then
				
						QryStr = QryStr & "&MFRCode=" & VMFR
				
					end if

					if not VRegion = "" then
				
						QryStr = QryStr & "&Reg=" & VRegion
				
					end if

					if not VRoute = "" then
				
						QryStr = QryStr & "&Route=" & VRoute
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
					QryStr = QryStr & "&Sort=" & VSortName		
										
					%>

					<td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("ClassNum") %></a></td>

	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("NotificationDate") %></td>
	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("ProductName") %></td>
	                <td bgcolor=<%= currcol %>><%= rsAll_Products.Fields("DIN") %></td>

					
					
					<%
					QryStr = "Form=" & rsAll_Products.Fields("Form")
					
					
					if not VActIng = "" then
					
						QryStr = QryStr & "&ActIng=" & VActIng
					
					end if

					if not VClass = "" then
					
						QryStr = QryStr & "&Class=" & VClass
				
					end if
					
					if not VMFR = "" then
				
						QryStr = QryStr & "&MFRCode=" & VMFR
				
					end if

					if not VRegion = "" then
				
						QryStr = QryStr & "&Reg=" & VRegion
				
					end if

					if not VRoute = "" then
				
						QryStr = QryStr & "&Route=" & VRoute
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
					QryStr = QryStr & "&Sort=" & VSortName
					
					%>
					
	                <td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("Form") %></a></td>
	
					
					
					
					<%
					QryStr = "Route=" & rsAll_Products.Fields("Route")
					
					
					if not VActIng = "" then
					
						QryStr = QryStr & "&ActIng=" & VActIng
					
					end if

					if not VClass = "" then
					
						QryStr = QryStr & "&Class=" & VClass
				
					end if

					if not VForm = "" then
					
						QryStr = QryStr & "&Form=" & VForm
				
					end if
					
					if not VMFR = "" then
				
						QryStr = QryStr & "&MFRCode=" & VMFR
				
					end if

					if not VRegion = "" then
				
						QryStr = QryStr & "&Reg=" & VRegion
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
					QryStr = QryStr & "&Sort=" & VSortName
					
					%>
					
					<td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("Route") %></a></td>
	
				
				
				
				   <%
					QryStr = "ActIng=" & rsAll_Products.Fields("ActiveIngGroup")
					
					if not VClass = "" then
					
						QryStr = QryStr & "&Class=" & VClass
				
					end if

					if not VForm = "" then
					
						QryStr = QryStr & "&Form=" & VForm
				
					end if
					
					if not VMFR = "" then
				
						QryStr = QryStr & "&MFRCode=" & VMFR
				
					end if

					if not VRegion = "" then
				
						QryStr = QryStr & "&Reg=" & VRegion
				
					end if

					if not VRoute = "" then
				
						QryStr = QryStr & "&Route=" & VRoute
				
					end if
					
					if not VDIN = "" then
				
						QryStr = QryStr & "&DIN=" & VDIN
				
					end if
					
					if not VProductName = "" then
				
						QryStr = QryStr & "&ProductName=" & VProductName
				
					end if
					
					QryStr = QryStr & "&Sort=" & VSortName
					
				   %>
				
				    <td bgcolor=<%= currcol %>><a href="All_Products.asp?<%= QryStr %>"><%= rsAll_Products.Fields("ActiveIngGroup") %></a></td>
											
				</tr>
					
			<%

			
		rsAll_Products.MoveNext
		If rsAll_Products.EOF Then Exit for
		Next
		%>

		</table>
		
		<br>
		<div align="center">
		<table>
			<tr>
			
				<%
					If cInt(intPage) > 1 Then
					%>		
						<%
							
						QryStr = "page=1"
						
						if not VActIng = "" then
						
							QryStr = QryStr & "&ActIng=" & VActIng
						
						end if
						
						if not VClass = "" then
						
							QryStr = QryStr & "&Class=" & VClass
					
						end if
	
						if not VForm = "" then
						
							QryStr = QryStr & "&Form=" & VForm
					
						end if
						
						if not VMFR = "" then
					
							QryStr = QryStr & "&MFRCode=" & VMFR
					
						end if
	
						if not VRegion = "" then
					
							QryStr = QryStr & "&Reg=" & VRegion
					
						end if
	
						if not VRoute = "" then
					
							QryStr = QryStr & "&Route=" & VRoute
					
						end if
						
						if not VDIN = "" then
						
							QryStr = QryStr & "&DIN=" & VDIN
				
						end if
					
						if not VProductName = "" then
				
							QryStr = QryStr & "&ProductName=" & VProductName
				
						end if						
						
						QryStr = QryStr & "&Sort=" & VSortName

						%>
						<td width="200"><a href="All_Products.asp?<%= QryStr %>">First page</a></td>				   										

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
						
						if not VActIng = "" then
						
							QryStr = QryStr & "&ActIng=" & VActIng
						
						end if
						
						if not VClass = "" then
						
							QryStr = QryStr & "&Class=" & VClass
					
						end if
	
						if not VForm = "" then
						
							QryStr = QryStr & "&Form=" & VForm
					
						end if
						
						if not VMFR = "" then
					
							QryStr = QryStr & "&MFRCode=" & VMFR
					
						end if
	
						if not VRegion = "" then
					
							QryStr = QryStr & "&Reg=" & VRegion
					
						end if
	
						if not VRoute = "" then
					
							QryStr = QryStr & "&Route=" & VRoute
					
						end if
						
						if not VDIN = "" then
				
							QryStr = QryStr & "&DIN=" & VDIN
				
						end if
					
						if not VProductName = "" then
				
							QryStr = QryStr & "&ProductName=" & VProductName
				
						end if						
						
						QryStr = QryStr & "&Sort=" & VSortName

						%>												
						<td width="200"><a href="All_Products.asp?<%= QryStr %>">Previous Page</a></td>
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
						
						if not VActIng = "" then
						
							QryStr = QryStr & "&ActIng=" & VActIng
						
						end if
						
						if not VClass = "" then
						
							QryStr = QryStr & "&Class=" & VClass
					
						end if
	
						if not VForm = "" then
						
							QryStr = QryStr & "&Form=" & VForm
					
						end if
						
						if not VMFR = "" then
					
							QryStr = QryStr & "&MFRCode=" & VMFR
					
						end if
	
						if not VRegion = "" then
					
							QryStr = QryStr & "&Reg=" & VRegion
					
						end if
	
						if not VRoute = "" then
					
							QryStr = QryStr & "&Route=" & VRoute
					
						end if
						
						if not VDIN = "" then
				
							QryStr = QryStr & "&DIN=" & VDIN
				
						end if
					
						if not VProductName = "" then
				
							QryStr = QryStr & "&ProductName=" & VProductName
				
						end if						
						
						QryStr = QryStr & "&Sort=" & VSortName

						%>																   		
						<td width="200"><a href="All_Products.asp?<%= QryStr %>">Next Page</a></td>								
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
						
						if not VActIng = "" then
						
							QryStr = QryStr & "&ActIng=" & VActIng
						
						end if
						
						if not VClass = "" then
						
							QryStr = QryStr & "&Class=" & VClass
					
						end if
	
						if not VForm = "" then
						
							QryStr = QryStr & "&Form=" & VForm
					
						end if
						
						if not VMFR = "" then
					
							QryStr = QryStr & "&MFRCode=" & VMFR
					
						end if
	
						if not VRegion = "" then
					
							QryStr = QryStr & "&Reg=" & VRegion
					
						end if
	
						if not VRoute = "" then
					
							QryStr = QryStr & "&Route=" & VRoute
					
						end if
						
						if not VDIN = "" then
				
							QryStr = QryStr & "&DIN=" & VDIN
				
						end if
					
						if not VProductName = "" then
				
							QryStr = QryStr & "&ProductName=" & VProductName
				
						end if						
						
						QryStr = QryStr & "&Sort=" & VSortName

						%>																		
						<td width="200"><a href="All_Products.asp?<%= QryStr %>">Last page</a></td>
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


	rsAll_Products.Close
	dcnDB.Close
	Set rsAll_Products = Nothing
	Set dcnDB = Nothing

%>




</body>
</html>
