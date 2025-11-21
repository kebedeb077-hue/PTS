<%
Response.Expires = -1000 ' Makes the browser not cache this page
Response.Buffer = True   ' Buffers the content so the Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("default.asp")
End If
%>   

<%
If Request.QueryString("ToExcel") = "True" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "filename=SearchResults.xls"
	Dim u_title
	u_title = "My Excel Spreadsheet"
	Response.Write "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	Response.Write "<head>"
	Response.Write "<!--[if gte mso 9]><xml>"
	Response.Write "<x:ExcelWorkbook>"
	Response.Write "<x:ExcelWorksheets>"
	Response.Write "<x:ExcelWorksheet>"
	Response.Write "<x:Name>"& u_title &"</x:Name>"
	Response.Write "<x:WorksheetOptions>"
	Response.Write "<x:Print>"
	Response.Write "<x:ValidPrinterInfo/>"
	Response.Write "</x:Print>"
	Response.Write "</x:WorksheetOptions>"
	Response.Write "</x:ExcelWorksheet>"
	Response.Write "</x:ExcelWorksheets>"
	Response.Write "</x:ExcelWorkbook>"
	Response.Write "</xml>"
	Response.Write "<![endif]--> "
End If
%>

<% If Request.QueryString("ToExcel") = "" Then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Search</title>
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="tb, default">
</head>

    <style type="text/css" media="screen">
    .NoPrint{
        display:block;
    }
    </style>
    <style type="text/css" media="print">
    .NoPrint{
        display:none;
    }
    </style>
<% End If %>

    
 

<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#214285">
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#426262">
	<tr bgcolor="#426262">	<!--<table border="0" cellspacing="0" cellpadding="2" width="700"> -->
	<th height="10" align="center" width="120"><font face="Kortika" size=“1”><img border="0" src="Images/Flag.jpg" alt="Flag of Ethiopia" width="120" height="80"></font></th> </td>	
	<th height="10" align="center">
	<font face="Kortika" color="#FFFFFF" size=“10”> &#4840;&#4848;&#4705;&#4709;&#32;&#4770;&#4725;&#4846;&#4917;&#4843;&#32;&#4781;&#4621;&#4621;&#32;&#4632;&#4757;&#4877;&#4645;&#4725;&#32;&#4939;&#4845;&#4755;&#4757;&#4661;&#32;&#4706;&#4654;</font><b><font size="4" color="#FFFFFF" face="Verdana"><br>
	&#4840;&#4629;&#4725;&#4632;&#4725;&#32;&#4645;&#4653;&#4909;&#4725;&#32;&#4658;&#4661;&#4720;&#4637;</font><font size="5" color="#FFFFFF" face="Verdana"></font>
	<font color="#FFFFFF" face="Verdana" size="2"><br></font></b><font face="Verdana"><b><font color="#FFFFFF" size="2">South Ethiopia Regional State Bureau of Finance, Publications Distribution System</font></b></font></p></font></b></p></th>	
	<th height="16" width="120" align="center"><font face="Kortika" color="#FFFFFF" size=“1”><img border="0" src="Images/South_ETH_Flag.JPG" alt="Flag of South Ethiopia People " width="120" height="80"></font></th>
</tr>

<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#426262" width="100%" class="NoPrint">
  <tr>
    <td width="100%" bgcolor="#FFFFFF">
    <div align="center">
      <center>
      <table border="2" style="border-collapse: collapse" bordercolor="#FFFFFF" width="950" cellspacing="1">
        <tr>
          <td width="100" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none; font-weight: 700" href="Default.asp">
          <font size="2" color="#FFFFFF" face="Verdana">&#4632;&#4877;&#4706;&#4843;</font><br>
          <font color="#FFFFFF" size="1" face="Verdana">Login</font></a></td>
	  <td width="50">&nbsp;</td> 

          <td width="100" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none; font-weight: 700" href="default.asp?login=logout">
          <font size="2" color="#FFFFFF" face="Verdana">&#4632;&#4813;&#4907;</font><br>
          <font color="#FFFFFF" face="Verdana" size="1">Logout</font></a></td>
	  <td width="50">&nbsp;</td> 

          <td width="300" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none; font-weight: 700" href="office_list.asp">
          <font size="2" color="#FFFFFF" face="Verdana">&#4840;&#4706;&#4654;&#4814;&#4733; &#4829;&#4653;&#4829;&#4653;</font><br>
          <font color="#FFFFFF" size="1" face="Verdana">Office List</font></a></td>
	  <td width="50">&nbsp;</td>	 

          <td width="300" align="center" bgcolor="#426262">
  	  <p align="center">
          <a style="text-decoration: none; color="#FFFFFF" font-weight: 700" href="form_list.asp">
          <font size="2" font color="#FFFFFF" face="Verdana"><b>&#4840;&#4942;&#4653;&#4638;&#4733;&nbsp; &#4829;&#4653;&#4829;&#4653;</font></b><br>
          <font color="#FFFFFF" size="1" face="Verdana">Form List</font></a></td>
    	  <td width="50">&nbsp;</td>

	  <td width="300" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none; font-weight: 700" href="bof_store.asp">
          <font size="2" color="#FFFFFF" face="Verdana">&#4661;&#4726;&#4653;</font><br>
          <font color="#FFFFFF" size="1" face="Verdana">BoF Store</font></a></td>
	  <td width="50">&nbsp;</td>
	  
          <td width="150" bgcolor="#426262">
          <p align="center">
          <b><font size="2" color="#FFFFFF" face="Verdana">
          <a style="text-decoration: none; color="#FFFFFF" font-weight: 700" href="shipments.asp"> 
	  <font size="2" color="#FFFFFF" face="Verdana">&#4645;&#4653;&#4909;&#4725;</a></font></b>
          </a><a href="shipments.asp" style="text-decoration: none"><span style="font-weight: 700">
          <font face="Verdana" size="1" color="#FFFFFF">Distributions</font></span></a></td>
          <td width="50">&nbsp;</td>

	  <td width="150" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none" href="search.asp"><b>
          <font size="2" color="#FFFFFF" face="Verdana">&#4936;&#4621;&#4877;</font></b><br>
          <font color="#FFFFFF" face="Verdana" size="1"><b>Search</b></font></a></td>
          <td width="50">&nbsp;</td>
          
          <td width="150" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none" href="reports1.asp"><b>
          <font size="2" color="#FFFFFF" face="Verdana">&#4650;&#4950;&#4653;&#4725;</font></b><br>
          <font color="#FFFFFF" face="Verdana" size="1"><b>Reports</b></font></a></td>
          <td width="50">&nbsp;</td>

          <td width="150" align="center" bgcolor="#426262">
          <a style="text-decoration: none" href="backup.asp"><b>
          <font size="2" color="#FFFFFF" face="Verdana">&#4707;&#4779;&#4949;</font></b><br>
          <font color="#FFFFFF" face="Verdana" size="1">
          <b>Backup</b></font></a></td>
          <td width="50" align="center"> &nbsp;</td>	  
	  
          <td width="100" bgcolor="#426262">
          <p align="center">
          <a style="text-decoration: none" href="about.htm"><b>
          <font size="2" color="#FFFFFF" face="Verdana">&#4853;&#4875;&#4941;</font></b><br>        
          <span style="font-weight: 700">
	  <font color="#FFFFFF" face="Verdana" size="1">Help</span></font></a></td>
      
        </tr>
        </table>
      </center>
    </div>
    </td>
  </tr>
        </table>
      </center>
    </div>
    </td>
  </tr>
</table>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">

  <!--webbot bot="SaveResults" u-file="_private/form_results.csv" s-format="TEXT/CSV" s-label-fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" -->

<%
' Declare the variables... 
' The connection string... 
Dim CONN_STRING
CONN_STRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=Admin; Password=serbof;"


' ADO recordset
' For clarity purposes names representative of recordset are used
Dim rstSearch  					
Dim rstSearchShipment			
Dim rstSearchShipmentList
Dim rstSearchForm
Dim rstSearchFormList
Dim rstSearchOffice
Dim rstSearchOfficeList

' The SQL Query that is built
Dim strSQL     
Dim strSQLShipment
Dim strSQLShipmentList
Dim strSQLForm
Dim strSQLFormList
Dim strSQLOffice
Dim strSQLOfficeList

' The text being looked for
Dim strSearchShipment_ID	
Dim strSearchOffice_ID
Dim strSearchDate  
Dim strSearchStarting_Serial_No
Dim strSearchEnding_Serial_No
Dim strSearchForm_Name
Dim strSearchModel_No
Dim strSearchCondition
Dim strSearchComments
Dim date_not_specified


Dim strURL     ' The URL of this page so the form will work
               ' no matter what this file is named.

' Retreive the URL of this page from Server Variables
strURL = Request.ServerVariables("URL")

' Retreive the conditions being searched for. 
strSearchShipment_ID = Request.QueryString("Shipment_ID")
strSearchOffice_ID = Request.QueryString("Office_ID")
strSearchSerial_No = Request.QueryString("Serial_No")
strSearchForm = Request.QueryString("Form_Name")
strSearchModel = Request.QueryString("Model_No")
strSearchCondition = Request.QueryString("Condition")
strSearchComments = Request.QueryString("Comments")


' For date, we need to do calculations based on day, month, year
date_not_specified = false
If Request.QueryString("Month")="" AND Request.QueryString("Day")="" AND Request.QueryString("Year")="" AND Request.QueryString("Monthto")="" AND Request.QueryString("Dayto")="" AND Request.QueryString("Yearto")="" Then
	date_not_specified = true
End If

' If there is a month, day, year specified, take it, otherwise put in dummy values
' It is assumed that each of the 12 months has 30 days, using the Ethiopian Calendar
' So the value for year is multiplied by 360
' The value for month by 30
' And these values are added to the value for day
If Request.QueryString("Year")="" OR Request.QueryString("Year")="0" Then
	strSearchDate=0
Else
	strSearchDate=360 * (CInt(Request.QueryString("Year"))-1)
	If Request.QueryString("Month")<>"" AND Request.QueryString("Month")<>"0" Then
		strSearchDate=strSearchDate + (30*(CInt(Request.QueryString("Month"))-1))	
	End If
	If Request.QueryString("Day")<>"" AND Request.QueryString("Day")<>"0" Then
		strSearchDate=strSearchDate + (CInt(Request.QueryString("Day")))
	End If
End If

' Do the same as above for the second date
If Request.QueryString("Yearto")="" OR Request.QueryString("Yearto")="0" Then
	strSearchDateto=360 * 2030 ' This is the upper limit, so make sure it is larger than any possible date
Else
	strSearchDateto=360 * (CInt(Request.QueryString("Yearto"))-1)
	If Request.QueryString("Monthto")<>"" AND Request.QueryString("Monthto")<>"0" Then
		strSearchDateto=strSearchDateto + (30*(CInt(Request.QueryString("Monthto"))-1))	
	End If
	If Request.QueryString("Dayto")<>"" AND Request.QueryString("Dayto")<>"0" Then
		strSearchDateto=strSearchDateto + (CInt(Request.QueryString("Dayto")))
	End If
End If


'For the dropdown list get the Shipment Names, Form List, Office List
strSQLShipmentList = "SELECT * FROM Shipments ORDER BY Shipment_ID;"
Set rstSearchShipmentList = Server.CreateObject("ADODB.Recordset")
rstSearchShipmentList.Open strSQLShipmentList, CONN_STRING

strSQLFormList = "SELECT * FROM Forms ORDER BY Form_Name;"
Set rstSearchFormList = Server.CreateObject("ADODB.Recordset")
rstSearchFormList.Open strSQLFormList, CONN_STRING

strSQLOfficeList = "SELECT * FROM Offices ORDER BY Office_Name;"
Set rstSearchOfficeList = Server.CreateObject("ADODB.Recordset")
rstSearchOfficeList.Open strSQLOfficeList, CONN_STRING

' Names of months
Dim i
Dim Months_Array(12)
Months_Array(1)="Meskerem"
Months_Array(2)="Tikimi"
Months_Array(3)="Hidar"
Months_Array(4)="Tahesas"
Months_Array(5)="Tir"
Months_Array(6)="Yekatit"
Months_Array(7)="Megabit"
Months_Array(8)="Miazia"
Months_Array(9)="Ginbot"
Months_Array(10)="Sene"
Months_Array(11)="Hamle"
Months_Array(12)="Nehase"


' Now display the search form
%>


<% If Request.QueryString("ToExcel") = "" Then %>

  <br>

  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="100%" style="border-style: solid; border-width: 2" bordercolor="#214285">
      
  <p>
  <b>
  <font color="#006666" face="Verdana" size="4" class="NoPrint">Search Deposit Database<br>
  </font></b><br>
  <b><font face="Arial" class="NoPrint">Specify one or more criteria:</font></b></p>
      
<form action="<%= strURL %>" method="get" class="NoPrint">
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
   	                  
               <tr>
                    <td width="12%">
                <b>
				<font face="Arial">
				    Office Name:&nbsp;&nbsp; </font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
					<select name="Office_ID">
					<%
					If strSearchOffice_ID = "%" Then
						noSearchOffice_ID = 0
					%>
					<option value="%" selected>All&nbsp;</option>
					<%
					Else
						noSearchOffice_ID = CInt(strSearchOffice_ID)
					%>
					<option value="%">All&nbsp;</option>
 					<%
 					End If
					Do While Not rstSearchOfficeList.EOF
					If (rstSearchOfficeList.Fields("Office_ID").Value = noSearchOffice_ID) Then
					%>
					<option value="<%= rstSearchOfficeList.Fields("Office_ID").Value %>"  selected><%= rstSearchOfficeList.Fields("Office_Name").Value %></option>
					<%
					else
					%>
					<option value="<%= rstSearchOfficeList.Fields("Office_ID").Value %>"><%= rstSearchOfficeList.Fields("Office_Name").Value %></option>
					<%
					End If
					rstSearchOfficeList.MoveNext
					Loop
					rstSearchOfficeList.Close
					Set rstSearchOfficeList = Nothing

					%>
					
					</select>
      			</font>
                    </td>
                  </tr>
                  
                  <tr>
                    <td width="12%">
				<b>
				<font face="Arial">
				    Date:&nbsp;&nbsp; </font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
                <b>From:</b>
                <select size="1" name="Day">
                <%
				DIM x(32)
				x(Request.QueryString("Day"))="selected"
				%>
				<option value="0" selected>&nbsp;</option>
				<%For i=1 to 30%>
                <option <%=x(i)%>><%=i%></option>  
               	<%Next%> 
                </select>
                
                <select size="1" name="Month">
                <%
				DIM y(13)
				y(Request.QueryString("Month"))="selected"
				%>
				<option value="0" selected>&nbsp;</option>				
				<%For i=1 to 12%>
                <option value="<%=i%>" <%=y(i)%>><%=Months_Array(i)%></option>
                <%Next%>
                </select>

                <select size="1" name="Year">
                <%
				DIM z(40)
				If Not(Request.QueryString("Year")="") AND Not(Request.QueryString("Year")="0") Then
					z(Request.QueryString("Year")-1990)="selected"
				End if
				%>
				 <option value="0" selected>&nbsp;</option>  				
				<%For i=1991 to 2020%>
               	<option <%=z(i-1990)%>><%=i%></option>
               	<%Next%> 
                </select> 
                
                <b>To:</b><select size="1" name="Dayto">
                <%
				DIM w(32)
				w(Request.QueryString("Dayto"))="selected"
				%>
				<option value="0" selected>&nbsp;</option>
				<%For i=1 to 30%>
                <option <%=w(i)%>><%=i%></option>  
               	<%Next%>
                </select>
                
                <select size="1" name="Monthto">
                <%
					DIM m(13)
					m(Request.QueryString("Monthto"))="selected"
				%>
				<option value="0" selected>&nbsp;</option>
				<%For i=1 to 12%>
                <option value="<%=i%>" <%=m(i)%>><%=Months_Array(i)%></option>
                <%Next%>
                </select>
                
                <select size="1" name="Yearto">
                <%
					DIM r(40)
					If Not(Request.QueryString("Yearto")="") AND Not(Request.QueryString("Yearto")="0") Then
						r(Request.QueryString("Yearto")-1990)="selected"
					End if
				%>    
				<option value="0" selected>&nbsp;</option>            
				<%For i=1991 to 2020%>
               	<option <%=r(i-1990)%>><%=i%></option>
               	<%Next%> 
                </select>
             
                </font></td>
                  </tr>
                  <tr>
                    <td width="12%">
				<b>
				<font face="Arial">
				Serial_No:</font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
                    <input type="text" name="Serial_No" value="<%= strSearchSerial_No %>" size="30" /></font></td>
                  </tr>
                  <tr>
                    <td width="12%">
				<b>
				<font face="Arial">
				Form:&nbsp;&nbsp; </font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
                    
					<select name="Form">
					<%
					If strSearchForm = "%" Then
						noSearchForm = 0
					%>
					<option value="%" selected>All&nbsp;</option>
					<%
					Else
						noSearchForm = CInt(strSearchForm)
					%>
					<option value="%">All&nbsp;</option>
 					<%
 					End If

					Do While Not rstSearchFormList.EOF
					
					If (rstSearchFormList.Fields("Form_ID").Value = noSearchForm) Then
					%>
					<option value="<%= rstSearchFormList.Fields("Form_ID").Value %>" selected><%= rstSearchFormList.Fields("Form_Name").Value %></option>
					<%
					Else
					%>
					<option value="<%= rstSearchFormList.Fields("Form_ID").Value %>"><%= rstSearchFormList.Fields("Form_Name").Value %></option>
					<%
					End If
					rstSearchFormList.MoveNext
					Loop
					rstSearchFormList.Close
					Set rstSearchFormList = Nothing
					%>
					</select>
                                        
                    </td>
                  </tr>
                  <tr>
                    <td width="12%">
				<b>
				<font face="Arial">
				Model: </font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
                    <input type="text" name="Model" value="<%= strSearchModel %>" size="30" /> </font>
                    </td>
                  </tr>
                  
                   <tr>
                    <td width="12%">
				<b>
				<font face="Arial">
				Comments: </font></b></td>
                    <td width="88%">
				<font face="Arial" size="2">
                    <input type="text" name="Comments" value="<%= strSearchComments %>" size="30" /> </font>
                    </td>
                  </tr>
                </table>
                <p>
				<font face="Arial" size="2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<br />
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Search" value="Search"><font size="2">
				</font>
			    </font>
			    </p>
                <p>
				&nbsp;</p>
			</form>

  <br>
  
<% End If              ' End for ToExcel %>   
  
<%
' Since this all done in one page it is necessary to see if anyone
' has searched for something.  If they have, then we access the database to list the records found
' Otherwise just show the search form and quit

If strSearchShipment_ID <> "" OR date_not_specified = false OR strSearchSerial_No <> "" OR strSearchForm_Name <> "" OR strSearchModel_No <> "" OR strSearchComments <> "" Then

    ' Build the query based on the input.
    
	strSQL = "SELECT * FROM Deposits, Shipments, Offices, Forms " _
			& "WHERE Shipments.Office_ID LIKE '" & strSearchOffice_ID & "' " _
			& "AND (((Shipments.Year-1)*360)+((Shipments.Month-1)*30)+Shipments.Day)>= " & strSearchDate & " "_
			& "AND (((Shipments.Year-1)*360)+((Shipments.Month-1)*30)+Shipments.Day)<= " & strSearchDateto & " "_
			& "AND Starting_Serial_No LIKE '%" & Replace(strSearchStarting_Serial_No, "'", "''") & "%' " _	
			& "AND Ending_Serial_No LIKE '%" & Replace(strSearchEnding_Serial_No, "'", "''") & "%' " _			
			& "AND Form_Name LIKE '%" & strSearchForm_Name & "%' " _
			& "AND Model_No LIKE '%" & Replace(strSearchModel_No, "'", "''") & "%' " _			
			& "AND Comments LIKE '%" & Replace(strSearchComments, "'", "''") & "%' " _			
			& "AND Shipments.Office_ID = Offices.Office_ID " _
			& "AND Deposits.Office_ID = Offices.Office_ID " _
			& "AND Shipments.Form_ID = Forms.Form_ID " _		
		    	& "ORDER BY Shipment_ID;"

	Set rstSearch = Server.CreateObject("ADODB.Recordset")
	rstSearch.Open strSQL, CONN_STRING
	
	' Display a table of the data in the recordset.  Loop through the
	' recordset displaying the fields from the table and using MoveNext
	' to increment to the next record.  Stop when EOF is reached
	%>

  <font face="Verdana" size="4">
		<b><font color="#006666">Search Results</font>
		<br><br> 
	<%
	If rstSearch.EOF Then
	%>
	<font face="Arial" size="2">
	0 Forms were found.</font>
	<%
	Else
	%>
		<table border="1" cellpadding="2">
		<thead>
		<tr>
			<th height="16" class="NoPrint"><font face="Arial" size="2">Deposit_ID</font></th>
			<th height="16"><font face="Arial" size="2">Shipment_Name</font></th>
			<th height="16" class="NoPrint"><font face="Arial" size="2">Office_Name</font></th>
			<th height="16"><font face="Arial" size="2">Date</font></th>
			<th height="16" class="NoPrint"><font face="Arial" size="2" >Serial_No</font></th>
			<th height="16"><font face="Arial" size="2">Form</font></th>
			<th height="16"><font face="Arial" size="2">Model</font></th>
			<th height="16"><font face="Arial" size="2">Quantity</font></th>
			<th height="16"><font face="Arial" size="2">Price</font></th>
			<th height="16" class="NoPrint"><font face="Arial" size="2">Serial_No_of_Model_22</font></th>
			<th height="16"><font face="Arial" size="2">Condition</font></th>
			<th height="16" class="NoPrint"><font face="Arial" size="2">Comments</font></th>	
				<% If Request.QueryString("ToExcel") = "" Then %>
			<th height="16" width="55" class="NoPrint"><font face="Arial" size="2" class="NoPrint">Delete</font></th>
			<th height="16" width="55" class="NoPrint"><font face="Arial" size="2" class="NoPrint">Edit</font></th>
			<th height="16" width="55" class="NoPrint"><font size="2" class="NoPrint">Transfer</font></th>	
<% End If %>			
		</tr>
		</thead>
		<tbody>
	<%
	Do While Not rstSearch.EOF
	%>
		<tr height="16">			
			<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstSearch.Fields("Deposit_ID").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstSearch.Fields("Name").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstSearch.Fields("Office_Name").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstSearch.Fields("Day").Value %>/<%= Months_Array(rstSearch.Fields("Month").Value) %>/<%= rstSearch.Fields("Year").Value %></font></td>
			<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstSearch.Fields("Serial_No").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstSearch.Fields("Form_Name").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstSearch.Fields("Model").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstSearch.Fields("Quantity").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstSearch.Fields("Price").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstSearch.Fields("Serial_No_of_Model_22").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstSearch.Fields("Condition").Value %>&nbsp;</font></td>
			<td height="16" nowrap align="left" class="NoPrint"><font face="Arial" size="2"><%= rstSearch.Fields("Comments").Value %>&nbsp;</font></td>
<% If Request.QueryString("ToExcel") = "" Then %>
			<td height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
            <font face="Verdana" size="2">
            <a href="Deposit_for_person.asp?action=deleteconfirm&Shipment_ID=<%=rstSearch.Fields("Shipment_ID").Value%>&Shipment_Name=<%=rstSearch.Fields("Name").Value%>&Deposit_ID=<%= rstSearch.Fields("Deposit_ID").Value %>" style="text-decoration: none; font-weight: 700">
            <font color="#FFFFFF">Delete</font></a></font></td>
			<td height="16" bgcolor="#006666" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
            <a href="Deposit_for_person.asp?action=edit&Shipment_ID=<%=rstSearch.Fields("Shipment_ID").Value%>&Shipment_Name=<%=rstSearch.Fields("Name").Value%>&Deposit_ID=<%= rstSearch.Fields("Deposit_ID").Value %>" style="text-decoration: none; font-weight: 700">
            <font face="Verdana" color="#FFFFFF" size="2">Edit</font></a></td>
			<td height="16" bgcolor="#006699" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
            <a href="Deposit_for_person.asp?action=transfer&Shipment_ID=<%=rstSearch.Fields("Shipment_ID").Value%>&Shipment_Name=<%=rstSearch.Fields("Name").Value%>&Deposit_ID=<%= rstSearch.Fields("Deposit_ID").Value %>" style="text-decoration: none; font-weight: 700">
            <font face="Verdana" color="#FFFFFF" size="2">Transfer</font></a></font></td>
<% End If %>
		</tr>
		<%
		rstSearch.MoveNext
	Loop
	%>
	</table>
	<font face="Arial" size="2" class="NoPrint"><b>
	(*Deposit_ID is automatically generated by the system and cannot be edited by the user)
	</b></font>
	<br>
	
<% If Request.QueryString("ToExcel") = "" Then %>
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="500" class="NoPrint">
          <tr>
            <td width="10%" align="right"> 
			&nbsp;
			</td>
            
            <td width="45%">
            <form id="PrintForm" action="<%=strURL%>" class="NoPrint">
			<img border="0" src="images/Print%20Button.JPG" align="top">
			<input id="btnPrint" type="button" value="Print list" class="PrintButton" onclick="window.print();" style="color: #FFFFFF; background-color: #525252; font-size:10pt; font-weight:bold; font-family:Verdana"/></font>
            </form>
			</td>            
            
			<td width="45%">
			<% 
			 ' When sending these parameters as QueryString, they have to be sent as %25 
			 If strSearchShipment_ID="%" Then 
				 strSearchShipment_ID="%25"
			 End If 
			 If strSearchOffice_ID="%" Then 
				 strSearchOffice_ID="%25"
			 End If 
			 If strSearchForm="%" Then 
				 strSearchForm="%25"
			 End If 
			 If strSearchCondition="%" Then 
				 strSearchCondition="%25"
			 End If 
			 %>			
            <form id="ExcelForm" action="<%=strURL%>?ToExcel=True&Shipment_ID=<%=strSearchShipment_ID%>&Office_ID=<%=strSearchOffice_ID%>&Serial_No=<%=strSearchSerial_No%>&Form=<%=strSearchForm%>&Model=<%=strSearchModel%>&Serial_No_of_Model_22=<%=strSearchSerial_No_of_Model_22%>&Condition=<%=strSearchCondition%>&Comments=<%=strSearchComments%>&Year=<%=Request.QueryString("Year")%>&Month=<%=Request.QueryString("Month")%>&Day=<%=Request.QueryString("Day")%>&Yearto=<%=Request.QueryString("Yearto")%>&Monthto=<%=Request.QueryString("Monthto")%>&Dayto=<%=Request.QueryString("Dayto")%>" method="post">
  			<img border="0" src="images/Excel%20Button.JPG" align="top">
  			<input id="btnExcel" type="submit" value="Send list to Excel" style="color: #FFFFFF; background-color: #525252; font-size:10pt; font-weight:bold; font-family:Verdana"/></font>
            </form>
            </td>

          </tr>
        </table>
<% End If 			' End ToExcel %>

	
	<% End If 		' rstSearch.EOF %>

	<%
	' Close the recordset and connection and dispose of the objects
	rstSearch.Close
	Set rstSearch = Nothing
		
End If			' The Big If

%> 

<% If Request.QueryString("ToExcel") = "" Then %>
    </td>
    </tr>
  </table>
  <br>
<% 	End If %> <!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#214285">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%">
   <p align="center"><font color="#FFFFFF" face="Arial" size="1" ><b><i>Developed by the ICT Directorate of BoF</i></b></font></td>
  </tr>
</table>
</td></tr><!--msnavigation--></table></body>

</html>