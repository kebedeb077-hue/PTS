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
	Response.AddHeader "Content-Disposition", "filename="&Request.QueryString("report")&".xls"
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
<title>Reports</title>
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
    
<%End If%>


<body><!--msnavigation-->
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

<% If Request.QueryString("ToExcel") = "" Then %>
 <br>
  
  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="100%" style="border-style: solid; border-width: 2" bordercolor="#426262">
      
<% End If %>

<%

' The connection string... 
Dim CONN_STRING
CONN_STRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=admin; Password=serbof;"

' The name of this file so all the links and forms will still
' work if it is renamed
Dim SCRIPT_NAME
SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

' This is for the link at the bottom of the page to avoid typing it each time
Dim BACK_TO_LIST_TEXT
BACK_TO_LIST_TEXT = "<a href=""" & SCRIPT_NAME & """ style=""text-decoration: none; font-weight: 700""><font face=""Arial"" size=""2"" color=""black""><b><p>"_
	& "&#4808;&#4848; &#4747;&#4619; &#4616;&#4632;&#4632;&#4616;&#4661; &#4773;&#4826;&#4613; &#4875; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;"_
	& "<br>Click here to go back </font></p></b>"      
    

' Declare the standard variables.
Dim rstDB, rstDB1, rstDB2			' ADO objects
Dim strSQL, strSQL1, strSQL2		' To hold various SQL Strings


' 12 months from the Ethiopian Calendar are used
Dim i
Dim Months_Array(12)
Months_Array(1)="&#4632;&#4661;&#4776;&#4648;&#4637;"
Months_Array(2)="&#4901;&#4677;&#4637;&#4725;"
Months_Array(3)="&#4613;&#4851;&#4653;"
Months_Array(4)="&#4723;&#4613;&#4643;&#4661;"
Months_Array(5)="&#4901;&#4653;"
Months_Array(6)="&#4840;&#4779;&#4722;&#4725;"
Months_Array(7)="&#4632;&#4875;&#4706;&#4725;"
Months_Array(8)="&#4634;&#4843;&#4826;&#4843;"
Months_Array(9)="&#4877;&#4757;&#4710;&#4725;"
Months_Array(10)="&#4656;&#4756;"
Months_Array(11)="&#4624;&#4637;&#4620;"
Months_Array(12)="&#4752;&#4624;&#4660;"

Dim Report_Month
Dim Report_Year

Dim Report_Month_From
Dim Report_Year_From
Dim Report_Month_To
Dim Report_Year_To	
Dim strSearchDate_From
Dim strSearchDate_To

' This function is to put the print/excel buttons at the bottom of the webpage
Function Printpage
%>

		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="500" class="NoPrint">
          <tr>
            <td width="10%" align="right"> 
			&nbsp;
			</td>
            
            <td width="45%">
            <form id="PrintForm" action="<%= SCRIPT_NAME %>" class="NoPrint">
			<img border="0" src="images/Print%20Button.JPG" align="top">
			<input id="btnPrint" type="button" value="Print list" class="PrintButton" onclick="window.print();" style="color: #FFFFFF; background-color: #525252; font-size:10pt; font-weight:bold; font-family:Verdana"/></font>
            </form>
			</td>            
            
			<td width="45%">
			<%
			Select Case LCase(Trim(Request.QueryString("report")))

				Case "bof"
			%>
	            <form id="ExcelForm" action="<%= SCRIPT_NAME %>?ToExcel=True&report=<%= Request.QueryString("report") %>&Month=<%=Report_Month%>&Year=<%=Report_Year%>" method="post">
	        <%
	        	Case "bof_financial"
			%>
				<form id="ExcelForm" action="<%= SCRIPT_NAME %>?ToExcel=True&report=<%= Request.QueryString("report") %>&Month_From=<%=Report_Month_From%>&Year_From=<%=Report_Year_From%>&Month_To=<%=Report_Month_To%>&Year_To=<%=Report_Year_To%>" method="post">
			<%
				Case "office"
			%>
				<form id="ExcelForm" action="<%= SCRIPT_NAME %>?ToExcel=True&report=<%= Request.QueryString("report") %>&Office_ID=<%=Report_Office_ID%>&Month=<%=Report_Month%>&Year=<%=Report_Year%>" method="post">
			<%
			End Select
			%>			    
			<img border="0" src="images/Excel%20Button.JPG" align="top">
			<input id="btnExcel" type="submit" value="Send list to Excel" style="color: #FFFFFF; background-color: #525252; font-size:10pt; font-weight:bold; font-family:Verdana"/></font>
            </form>
            </td>

          </tr>
        </table>

<%
End Function



' Check to see which report to display
Select Case LCase(Trim(Request.QueryString("report")))

	Case "bof"
	
 		If Request.QueryString("ToExcel") = "" Then 		
			Report_Month = Request.Form("Month")
			Report_Year = Request.Form("Year")
 		Else
			Report_Month = Request.QueryString("Month")
			Report_Year = Request.QueryString("Year") 
 		End If
 		
 		If Report_Month = 1 Then
  			Previous_Month = 12
  			Previous_Months_Year = Report_Year - 1
  		Else
			Previous_Month = Report_Month - 1
			Previous_Months_Year = Report_Year
		End If
				
		strSQL = " SELECT Form_ID, Form_Name, Model_No FROM Forms ORDER BY Form_Name; "

		Set rstDB = Server.CreateObject("ADODB.Recordset")
		rstDB.Open strSQL, CONN_STRING
	%>

  		<font face="Verdana" size="4" color="#006666">
		<b>&#4840;&#4629;&#4725;&#4632;&#4725; &#4808;&#4653;&#4611;&#4810; &#4650;&#4950;&#4653;&#4725; &#4635;&#4677;&#4648;&#4706;&#4843; &#4677;&#4925;<br>
		<font size="3">Form Transactions for BoF</font>
  		</font>
		</font>
  		<font face="Verdana" size="4">
		<br>
  		</font>
  		<font face="Verdana" color="#006666">
		&#4776; </font><font size="3" face="Verdana">
		<font color="#006666">From:</font><font color="#C46200"> 01/<%=Months_Array(Report_Month)%>/<%=Report_Year%></font>
		<font color="#006666">&#4616;
		To:</font><font color="#C46200"> 30/<%=Months_Array(Report_Month)%>/<%=Report_Year%></font> </font>
		</b>

  		<font face="Verdana" size="4">
		<br><br> 
		<table border="1" cellspacing="0" cellpadding="2" width="700">
		<thead>
		<tr height="16">
			<th height="16"><font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; &#4819;&#4845;&#4752;&#4725;<br>
            Form Name</font></th>
			<th height="16"><font face="Arial" size="2">&#4632;&#4616;&#4843;<br>
            Model No</font></th>
			<th height="16"><font face="Arial" size="2">&#4779;&#4616;&#4936;&#4813; &#4808;&#4653; &#4840;&#4830;&#4648;<br>
            Quantity Prev Month<br>(30/<%=Months_Array(Previous_Month)%>/<%=Previous_Months_Year%>)</font></th>
			<th height="16"><font face="Arial" size="2">&#4704;&#4808;&#4649; &#4813;&#4661;&#4901; &#4840;&#4872;&#4707;<br>
            Quantity Received This Month</font></th>
			<th height="16"><font face="Arial" size="2">&#4896;&#4677;&#4619;&#4619; &#4853;&#4637;&#4653;<br>
            Quantity Total</font></th>									
			<th height="16"><font face="Arial" size="2">&#4704;&#4808;&#4649; &#4813;&#4661;&#4901; &#4840;&#4720;&#4640;&#4651;&#4909;<br>
            Quantity Sent This Month</font></th>
			<th height="16"><font face="Arial" size="2">&#4776;&#4808;&#4906; &#4672;&#4650; &#4704;&#4877;/&#4708;&#4725;<br>
            Quantity This Month<br>(30/<%=Months_Array(Report_Month)%>/<%=Report_Year%>)</font></th>				
		</tr>
		</thead>
		<tbody>
		    <%
   			Set rstDB1 = Server.CreateObject("ADODB.Recordset")
			Set rstDB2 = Server.CreateObject("ADODB.Recordset")

			Dim Qty_This_Month
			Dim Qty_Sent
			Dim Qty_Received
			
			' For each form run this loop
		    Do While Not rstDB.EOF
		    
		    strSQL1 = "SELECT SUM(Quantity_Received_by_BoF-Quantity_Sent_from_BoF) AS Qty_This_Month " _
 			        & "FROM Shipments " _
  			        & "WHERE Form_ID = "& rstDB.Fields("Form_ID").Value &" AND (Shipments.Year<"& Report_Year &" OR (Shipments.Year="& Report_Year &" AND Shipments.Month<="& Report_Month &")) ;"

			rstDB1.Open strSQL1, CONN_STRING
			
			' If there are no records in the shipment table then Qty_This_Month will be empty
			' so set it to zero
			If IsNumeric(rstDB1.Fields("Qty_This_Month").Value)=false Then
				Qty_This_Month = 0
			Else 
				Qty_This_Month = rstDB1.Fields("Qty_This_Month").Value
			End If
	
			strSQL2 = "SELECT SUM(Quantity_Sent_from_BoF) AS Qty_Sent, SUM(Quantity_Received_by_BoF) AS Qty_Received " _
			        & "FROM Shipments " _
	    		    & "WHERE Form_ID = "& rstDB.Fields("Form_ID").Value &" AND (Shipments.Month="& Report_Month &" And Shipments.Year="& Report_Year &") ;"

			rstDB2.Open strSQL2, CONN_STRING
			
			If IsNumeric(rstDB2.Fields("Qty_Sent").Value)=false Then
				Qty_Sent = 0
			Else 
				Qty_Sent = rstDB2.Fields("Qty_Sent").Value
			End If
			
			If IsNumeric(rstDB2.Fields("Qty_Received").Value)=false Then
				Qty_Received = 0
			Else 
				Qty_Received = rstDB2.Fields("Qty_Received").Value
			End If

  		   %>
			<tr height="16">			
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDB.Fields("Form_Name").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstDB.Fields("Model_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= (Qty_This_Month - (Qty_Received - Qty_Sent)) %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Qty_Received %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= (Qty_This_Month + Qty_Sent) %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Qty_Sent %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Qty_This_Month %>&nbsp;</font></td>
			</tr>
			<%
				rstDB1.Close
				rstDB2.Close
				rstDB.MoveNext
		    Loop
    		rstDB.Close
			Set rstDB = Nothing
			Set rstDB1 = Nothing
			Set rstDB2 = Nothing

		    %>
		</tbody>
		</table>
<% If Request.QueryString("ToExcel") = "" Then %>
        <br>
		<%
		Printpage
		Response.Write(BACK_TO_LIST_TEXT)
End If  ' End for ToExcel



	Case "bof_financial"
	
		' First find the quantity in store

 		If Request.QueryString("ToExcel") = "" Then 		
			Report_Month_From = Request.Form("Month_From")
			Report_Year_From = Request.Form("Year_From")
			Report_Month_To = Request.Form("Month_To")
			Report_Year_To = Request.Form("Year_To")
 		Else
			Report_Month_From = Request.QueryString("Month_From")
			Report_Year_From = Request.QueryString("Year_From")
			Report_Month_To = Request.QueryString("Month_To")
			Report_Year_To = Request.QueryString("Year_To")
 		End If
		
		strSearchDate_From=(360 * (Report_Year_From-1)) + (30*(Report_Month_From-1))
	    strSearchDate_To=(360 * (Report_Year_To-1)) + (30*Report_Month_To)
	    
	    If (strSearchDate_To <= strSearchDate_From) Then ' Date Compare
	    %>
			<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;</font></b><b><font face="Verdana" size="4" color="#006666"><br>
			</font>
			</font>
	  		<font face="Arial" color="#006666">
			Error</font></b><font face="Arial" size="2"><font face="Verdana" size="4" color="#006666">
			</font>
			<br>
			<br>
			<font face="Arial" size="3">
			&#4776; &#4840;&#4634;&#4616;&#4813; &#4672;&#4757; &#4616; &#4776;&#4634;&#4616;&#4813; &#4632;&#4677;&#4848;&#4637; &#4768;&#4616;&#4704;&#4725;<br>
			Date <b>from</b> should be before date <b>to</b>.
			</font>
			<br><br>
			<%Response.Write(BACK_TO_LIST_TEXT)%>
			<br><br><br><br><br><br><br><br><br><br>
		<%	    
	    Else
	
			' This report is only for the Zones
			strSQL = " SELECT Office_ID, Office_Name, Office_Type FROM Offices WHERE Office_Type In ('Zone','Special_Woreda') ORDER BY Office_Type DESC, Office_Name; "

			Set rstDB = Server.CreateObject("ADODB.Recordset")
			rstDB.Open strSQL, CONN_STRING
			
			
			If Report_Month_From = 1 Then
  				Previous_Month = 12
	  			Previous_Months_Year = Report_Year_From - 1
  			Else
				Previous_Month = Report_Month_From - 1
				Previous_Months_Year = Report_Year_From
			End If
		%>
			</font>
			<font color="#006666" face="Verdana" size="4"><b>&#4776; <font color="#C46200"> 01/<%=Months_Array(Report_Month_From)%>/<%=Report_Year_From%></font> &#4773;&#4661;&#4776; <font color="#C46200"> 30/<%=Months_Array(Report_Month_To)%>/<%=Report_Year_To%></font> &#4843;&#4616;&#4813; &#4840;&#4629;&#4725;&#4632;&#4725; &#4849;&#4708; &#4661;&#4653;&#4909;&#4725;&#4755; &#4811;&#4875; &#4632;&#4877;&#4616;&#4907;</b></font><b><font face="Arial" size="2"><font face="Verdana" size="4" color="#006666"><br>
			</font>
			</font>
	  		<font face="Verdana" color="#006666">
			Financial Report for BoF<br>
			<font size="2">
			(Only Zones)<br>
			</font>
			</font>
			</b>
			<br><br> 
			<table border="1" cellspacing="0" cellpadding="2">
			<thead>
			<tr height="16">
				<th height="16"><font face="Arial" size="2">&#4840;&#4849;&#4708; &#4661;&#4653;&#4909;&#4725; &#4720;&#4896;&#4675;&#4634;<br>
                &#4768;&#4779;&#4621; &#4632;/&#4708;&#4725; &#4661;&#4637; <br>
                Office Name</font></th>
				<th height="16"><font face="Arial" size="2">&#4779;&#4616;&#4936;&#4813; &#4840;&#4830;&#4648;<br>
                Previous Balance 
                (Birr)<br>(30/<%=Months_Array(Previous_Month)%>/<%=Previous_Months_Year%>)</font></th>									
				<th height="16"><font face="Arial" size="2">&#4840;&#4872;&#4707; &#4872;&#4706;<br>
                Deposits This Period 
                (Birr)</font></th>
				<th height="16"><font face="Arial" size="2">&#4704;&#4877;&#4635;&#4669; &#4819;&#4632;&#4721; &#4840;&#4720;&#4656;&#4651;&#4904;<br>
                Costs This Period 
                (Birr)</font></th>
				<th height="16"><font face="Arial" size="2">&#4704;&#4896;/&#4840;&#4634;&#4936;&#4616;&#4877; &#4821;&#4851;<br>
                Balance (Birr)<br>(30/<%=Months_Array(Report_Month_To)%>/<%=Report_Year_To%>)</font></th>				
			</tr>
			</thead>
			<tbody>
			    <%
   				Set rstDB1 = Server.CreateObject("ADODB.Recordset")

				Dim Previous_Amount_Deposited
				Dim Deposits_This_Period
				Dim Previous_Costs
				Dim Costs_This_Period
			
				' For each Zone run this loop
			    Do While Not rstDB.EOF

			    strSQL1 = "SELECT SUM(Amount_Deposited) AS Previous_Amount_Deposited FROM Deposits WHERE Office_ID = "& rstDB.Fields("Office_ID").Value &" "_
				        & "AND (((Year-1)*360)+((Month-1)*30)+Day)<= " & strSearchDate_From & " ;"
				rstDB1.Open strSQL1, CONN_STRING
			
				If IsNumeric(rstDB1.Fields("Previous_Amount_Deposited").Value)=false Then
					Previous_Amount_Deposited = 0
				Else 
					Previous_Amount_Deposited = rstDB1.Fields("Previous_Amount_Deposited").Value
				End If
				rstDB1.Close

				strSQL1 = "SELECT SUM(Amount_Deposited) AS Deposits_This_Period FROM Deposits WHERE Office_ID = "& rstDB.Fields("Office_ID").Value &" "_
				        & "AND (((Year-1)*360)+((Month-1)*30)+Day)> " & strSearchDate_From & " "_
	 				    & "AND (((Year-1)*360)+((Month-1)*30)+Day)<= " & strSearchDate_To & " ;"
				rstDB1.Open strSQL1, CONN_STRING
			
 				If IsNumeric(rstDB1.Fields("Deposits_This_Period").Value)=false Then
					Deposits_This_Period = 0
				Else 
					Deposits_This_Period = rstDB1.Fields("Deposits_This_Period").Value			
				End If
				rstDB1.Close
	
				strSQL1 = "SELECT SUM(Total_Price) AS Previous_Costs FROM Shipments WHERE Office_ID = "& rstDB.Fields("Office_ID").Value &" "_
			    	    & "AND (((Year-1)*360)+((Month-1)*30)+Day)<= " & strSearchDate_From & " ;"
				rstDB1.Open strSQL1, CONN_STRING
			
 				If IsNumeric(rstDB1.Fields("Previous_Costs").Value)=false Then
					Previous_Costs = 0
				Else 
					Previous_Costs = rstDB1.Fields("Previous_Costs").Value			
				End If			
				rstDB1.Close	
	
				strSQL1 = "SELECT SUM(Total_Price) AS Costs_This_Period FROM Shipments WHERE Office_ID = "& rstDB.Fields("Office_ID").Value &" "_
				        & "AND (((Year-1)*360)+((Month-1)*30)+Day)> " & strSearchDate_From & " "_
	 				    & "AND (((Year-1)*360)+((Month-1)*30)+Day)<= " & strSearchDate_To & " ;"			
				rstDB1.Open strSQL1, CONN_STRING
			
 				If IsNumeric(rstDB1.Fields("Costs_This_Period").Value)=false Then
					Costs_This_Period = 0
				Else 
					Costs_This_Period = rstDB1.Fields("Costs_This_Period").Value			
				End If			
				rstDB1.Close	

  			   %>
				<tr height="16">			
					<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDB.Fields("Office_Name").Value %>, <%= rstDB.Fields("Office_Type").Value %></font></td>
					<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= (Previous_Amount_Deposited-Previous_Costs) %>&nbsp;</font></td>
					<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Deposits_This_Period %>&nbsp;</font></td>
					<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Costs_This_Period %>&nbsp;</font></td>
					<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= (Previous_Amount_Deposited-Previous_Costs)+(Deposits_This_Period-Costs_This_Period) %>&nbsp;</font></td>
				</tr>
				<%
					rstDB.MoveNext
		    	Loop
	    		rstDB.Close
				Set rstDB1 = Nothing
				Set rstDB = Nothing

		    	%>
			</tbody>
			</table>
	<% If Request.QueryString("ToExcel") = "" Then %>
    	    <br>
			<%
			Printpage
			Response.Write(BACK_TO_LIST_TEXT)
	End If  ' End for ToExcel 
	
	End If ' End for Date Compare


	Case "office"
	
 		If Request.QueryString("ToExcel") = "" Then 		
			Report_Office_ID = Request.Form("Office_ID")
    	 	Report_Month = Request.Form("Month")
			Report_Year = Request.Form("Year")
 		Else
			Report_Office_ID = Request.QueryString("Office_ID")
	     	Report_Month = Request.QueryString("Month")
			Report_Year = Request.QueryString("Year")
 		End If		
		
		strSQL = " SELECT Form_ID, Form_Name, Model_No FROM Forms ORDER BY Form_Name;"
		
		Set rstDB = Server.CreateObject("ADODB.Recordset")
		rstDB.Open strSQL, CONN_STRING
		
		Set rstDB1 = Server.CreateObject("ADODB.Recordset")
		strSQL1 = " SELECT Office_Name FROM Offices WHERE Offices.Office_ID = "& Report_Office_ID &" "
		rstDB1.Open strSQL1, CONN_STRING

	%>
		<b><font color="#006666">
  		<font face="Verdana" size="4">
		&#4840; <font color="#C46200"><%=rstDB1.Fields("Office_Name").Value%> </font> &#4830;&#4757;/&#4621;&#4841; &#4808;&#4648;&#4851; &#4840;&#4849;&#4708; &#4629;&#4725;&#4632;&#4725; &#4645;&#4653;&#4909;&#4725; &#4840;&#4811;&#4875; &#4632;&#4877;&#4616;&#4907;</font><br>
  		<font face="Verdana">
		Shipments to </font>
		</font>
  		<font face="Verdana" size="3">
		<font color="#C46200"><%=rstDB1.Fields("Office_Name").Value%> </font>
		<br>
        </font>
		</b>
  		<font face="Verdana" color="#006666">
		<b>&#4776;
		</b> </font>
  		<font color="#006666">
  		<b>
			<font size="3" face="Verdana">
			From:</font></b></font><b><font size="3" face="Verdana"><font color="#C46200"> 01/<%=Months_Array(Report_Month)%>/<%=Report_Year%></font><font color="#006666"> </font></font>		
		</b>
  		<font color="#006666">
  		<font face="Verdana" size="4">
		<b><font size="3" face="Verdana">
		&#4616; </font>
		</b></font>
  		<b>
		<font size="3" face="Verdana">
		To:</font></b></font><b><font size="3" face="Verdana"><font color="#C46200"> 30/<%=Months_Array(Report_Month)%>/<%=Report_Year%></font> </font>		
		</b>

		<br><br> 
		<table border="1" cellspacing="0" cellpadding="2">
		<thead>
		<tr height="16">
			<th height="16" align="center"><font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; &#4819;&#4845;&#4752;&#4725;<br>
            Form Name</font></th>
			<th height="16" align="center"><font face="Arial" size="2">&#4632;&#4616;&#4843;<br>
            Model No</font></th>
			<th height="16" align="center"><font face="Arial" size="2">&#4709;&#4827;&#4725;<br>
            Quantity of Forms</font></th>
			<th height="16" align="center"><font face="Arial" size="2" color="#000000">&#4843;&#4757;&#4849; &#4811;&#4875;<br>
            Unit Price (Birr)<br>(Average for the month)</font></th>
			<th height="16" align="center"><font face="Arial" size="2" color="#000000">&#4896;&#4677;&#4619;&#4619; &#4811;&#4875;<br>
            Total Price (Birr)</font></th>
		</tr>
		</thead>
		<tbody>
		    <%
			rstDB1.Close
						
			
			Dim Overall_Sum_Total_Price
			Overall_Sum_Total_Price = 0
			
			' For each form run this loop
		    Do While Not rstDB.EOF
		    
			strSQL1 = " SELECT SUM(Quantity_Sent_from_BoF) AS No_Forms, SUM(Total_Price) AS Sum_Total_Price "_
					 &" FROM Shipments " _
					 &" WHERE Office_ID="& Report_Office_ID &" AND Month="& Report_Month &" AND Year="& Report_Year &" "_
					 &" AND Shipments.Form_ID = "& rstDB.Fields("Form_ID").Value &" ;"

		    rstDB1.Open strSQL1, CONN_STRING
		    
		    If IsNumeric(rstDB1.Fields("No_Forms").Value)=false Then
				No_Forms = 0
			Else 
				No_Forms = rstDB1.Fields("No_Forms").Value			
			End If

			If IsNumeric(rstDB1.Fields("Sum_Total_Price").Value)=false Then
				Sum_Total_Price = 0
			Else 
				Sum_Total_Price = rstDB1.Fields("Sum_Total_Price").Value			
			End If

  		   %>
			<tr height="16">			
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDB.Fields("Form_Name").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstDB.Fields("Model_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= No_Forms %>&nbsp;</font></td>
				<td height="16" nowrap align="center" >
				<font face="Arial" size="2">
				<% If No_Forms=0 OR Sum_Total_Price=0 Then %>
					-
				<% Else %>
					<%= Sum_Total_Price/No_Forms %>
				<% End If %>
				</font>
				</td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Sum_Total_Price %>&nbsp;</font></td>
			</tr>
			<%
				rstDB1.Close
				rstDB.MoveNext
				Overall_Sum_Total_Price = Overall_Sum_Total_Price + Sum_Total_Price
		    Loop
    		rstDB.Close
			Set rstDB = Nothing
			Set rstDB1 = Nothing
		    %>
		    <tr height="16">
		    	<td height="16" nowrap align="center" colspan="3">&nbsp;</td>
		    	<td height="16" nowrap align="right" ><font face="Arial" size="2"><b>Overall Total:</b></font></td>
		    	<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= Overall_Sum_Total_Price %>&nbsp;</font></td>
		    </tr>
		</tbody>
		</table>
<% If Request.QueryString("ToExcel") = "" Then %>
        <br>
		<%
		Printpage
		Response.Write(BACK_TO_LIST_TEXT)
End If  ' End for ToExcel		
	
		
	Case Else ' view
	
		' The default action... just list the menu of reports
%>
		<font color="#006666" size="4" face="Verdana"><b>&#4650;&#4950;&#4653;&#4725;<br>
		</font>
  		<font face="Verdana" color="#006666">
		Reports</font></b><br>
        <br>
  		<table border="1" cellpadding="0" cellspacing="0" width="400">
          <tr>
            <td width="100%">
	        <form action="<%= SCRIPT_NAME %>?report=bof" method="post">
            <div align="center">
              <center>
    

 <table border="1" cellpadding="0" cellspacing="0" width="550">
        <tr>
        <td width="100%">
	    <form action="<%= SCRIPT_NAME %>?report=office" method="POST">
            <div align="center">
              <center>
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="500">
            <tr>
              <td width="100%" colspan="2">
			  <b>

  		<font color="#006666" face="Arial">
				&#4808;&#4648;&#4851; &#4840;&#4849;&#4708; &#4629;&#4725;&#4632;&#4725; &#4645;&#4653;&#4909;&#4725; &#4840;&#4811;&#4875; &#4632;&#4877;&#4616;&#4907;</font><u><font face="Arial" color="#000000"><br>
              </font></u><font face="Arial" color="#000000" size="2">Monthly Report for Offices</font></b><br><br>
			  </td>
            </tr>
            <tr>
              <td width="150">
			  <b><font face="Arial" size="2">&#4840;&#4706;&#4653;&#4813; &#4661;&#4637;<br>
              Office Name:</font></b><font size="2" face="Arial">
              </font>
			  </td>
              <td width="350">
					<%
				    ' For the dropdown list get the Office Names
				    Dim strSQLOfficeList
				    
					strSQLOfficeList = "SELECT * FROM Offices ORDER BY Office_Type DESC, Office_Name;"
					Set rstOfficeList = Server.CreateObject("ADODB.Recordset")
					rstOfficeList.Open strSQLOfficeList, CONN_STRING
					%>
					<select name="Office_ID" style="font-family: Visual Geez Unicode">
 					<%
					Do While Not rstOfficeList.EOF
					%>
					<option value="<%= rstOfficeList.Fields("Office_ID").Value %>"><%= rstOfficeList.Fields("Office_Name").Value %></option>
					<%
					rstOfficeList.MoveNext
					Loop
					rstOfficeList.Close
					Set rstOfficeList = Nothing
					%>
					</select>
              </td>
            </tr>
            <tr>
              <td width="150">
			  <b><font face="Arial">&#4808;&#4653;</font><font size="2" face="Arial"><br>
              Month:</font></b><font size="2" face="Arial"> </font>
			  </td>
              <td width="350">
               <select size="1" name="Month" style="font-family: Visual Geez Unicode">
                <%For i=1 to 12%>
                <option value="<%=i%>"><%=Months_Array(i)%></option>
                <%Next%>
                </select>
              </td>
            </tr>
            <tr>
              <td width="150">
			  
  		<font face="Arial">
			  <b>&#4819;</b></font><b>&#4632;&#4725;</b></font><b><font size="2" face="Arial"><br>
              Year:</font></b><font size="2" face="Arial"> </font>
			  </td>
              <td width="350">
              <select size="1" name="Year">
              <%For i=2016 to 2050%>
              <option><%=i%></option>
              <%Next%> 
              </select>
              </td>
            </tr>
          </table>
              </center>
            </div>
          <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="submit" value="Get Report" name="Get Report" style="color: #FFFFFF; font-weight: bold; font-family: Verdana; background-color: #525252"></p>
       </form>
  	   </td>
       </tr>
   	   </table>

         <p>&nbsp;</p>

<%
End Select
%>

<% If Request.QueryString("ToExcel") = "" Then %>
  <br>
</td>
</tr>
</table>
<br>
<% End If %> <!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#426262">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
 <tr>
    <td width="100%">
   <p align="center"><font color="#FFFFFF" face="Arial" size="1" ><b><i>Developed by the ICT Directorate of BoF</i></b></font></td>
  </tr>
</table>

</td></tr><!--msnavigation--></table></body>
</html>