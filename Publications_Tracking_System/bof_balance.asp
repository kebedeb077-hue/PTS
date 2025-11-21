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
<title>BoF Balance</title>
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


<%
' Declare the variables... 
' The connection string... 
' ADO recordset
' For clarity purposes names representative of recordset are used

Dim rst  					
Dim rstShipment			
Dim rstShipmentList
Dim rstForm
Dim rstFormList
Dim rstOffice
Dim rstOfficeList

' The SQL Query that is built
Dim strSQL     
Dim strSQLShipment
Dim strSQLShipmentList
Dim strSQLForm
Dim strSQLFormList
Dim strSQLOffice
Dim strSQLOfficeList

' The text being looked for
Dim strSelectShipment_ID	
Dim strSelectOffice_ID
Dim strSelectDate  
Dim strSelectSerial_No
Dim strSelectForm
Dim strSelectModel
Dim strSelectSerial_No_of_Model_22
Dim strSelectCondition
Dim strSelectComments
Dim date_not_specified


Dim strURL     ' The URL of this page so the form will work
               ' no matter what this file is named.

' Retreive the URL of this page from Server Variables
strURL = Request.ServerVariables("URL")

' Retreive the conditions being Selected for. 
strSelectShipment_ID = Request.QueryString("Shipment_ID")
strSelectOffice_ID = Request.QueryString("Office_ID")
strSelectSerial_No = Request.QueryString("Serial_No")
strSelectForm = Request.QueryString("Form")
strSelectModel = Request.QueryString("Model")
strSelectSerial_No_of_Model_22 = Request.QueryString("Serial_No_of_Model_22")



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
	strDate=0
Else
	strDate=360 * (CInt(Request.QueryString("Year"))-1)
	If Request.QueryString("Month")<>"" AND Request.QueryString("Month")<>"0" Then
		strDate=strDate + (30*(CInt(Request.QueryString("Month"))-1))	
	End If
	If Request.QueryString("Day")<>"" AND Request.QueryString("Day")<>"0" Then
		strDate=strDate + (CInt(Request.QueryString("Day")))
	End If
End If

' Do the same as above for the second date
If Request.QueryString("Yearto")="" OR Request.QueryString("Yearto")="0" Then
	strDateto=360 * 2030 ' This is the upper limit, so make sure it is larger than any possible date
Else
	strDateto=360 * (CInt(Request.QueryString("Yearto"))-1)
	If Request.QueryString("Monthto")<>"" AND Request.QueryString("Monthto")<>"0" Then
		strDateto=strDateto + (30*(CInt(Request.QueryString("Monthto"))-1))	
	End If
	If Request.QueryString("Dayto")<>"" AND Request.QueryString("Dayto")<>"0" Then
		strDateto=strDateto + (CInt(Request.QueryString("Dayto")))
	End If
End If

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


' Now display the form
%>

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

  <br>
  
  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="100%" bordercolor="#426262" style="border-style:solid; border-width:2; ">
  		<b>
  		<font color="#006666" face="Verdana" size="4">&#4661;&#4726;&#4653; &#4813;&#4661;&#4901; &#4843;&#4617; &#4942;&#4653;&#4638;&#4733;<br>
  		</font>
  		<font color="#006666" face="Verdana">Forms Balance in BoF Store</font><font face="Arial"><br>
 		</font>

  		<BR><font face="Arial">&#4840;&#4942;&#4653;&#4638;&#4733;&#4757; &#4829;&#4813;&#4813;&#4653; &#4616;&#4635;&#4840;&#4725; &#4704;&#4942;&#4653;&#4633; &#4619;&#4845; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;<br>
  		<font size="2">Click on the form to get the details of transactions.</font></font></b><br><br>	


	  <tr>
	                    <td width="39%">
							<font face="Arial" size="3"><b><font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; 
                			&#4819;&#4845;&#4752;&#4725;<br>
			                </font></b></font>
							<b>
							<font face="Arial" size="2">
							Office Name:&nbsp;&nbsp; </font></b>
						</td>
''---------------------------------
	                  
		       
           </tr>
	 
		<%

		 ' Get the list of all the forms ordered by name
		 Set Conn1 = Server.CreateObject("ADODB.Connection")
		 Conn1.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=admin; Password=serbof;")
 
		 SQL_query = "SELECT * FROM Forms ORDER BY Form_Name"
		 Set RS = Conn1.Execute(SQL_query)

 		%>
		
  		<br>	

  		<table border="1" cellspacing="0" cellpadding="5" width="900">

    		<tr>      			
      			<td width="10" align="left"><font face="Arial" size="2"><b>&#4720;&#4651;&#32;&#4673;<br>
      			No<br>
			<td width="100" align="left"><font face="Arial" size="2"><b>&#4840;&#4632;/&#4708;&#4721;&#32;&#4645;&#4637;<br>
      			Office Name<br></font></td>
			<td width="100" align="left"><b><font face="Arial" size="2">&#4840;&#4638;&#4852;&#4621;&#32;&#4673;&#4901;&#4653;<br>
      			Model No</font></b></td>
      			<td width="300" align="left"><b><font face="Arial" size="2">&#4840;&#4645;&#4653;&#4909;&#4725;&#32;&#4672;&#4757;<br>
      			Distribution Date</font></b></td>			
      			<td width="100" align="left"><b><font face="Arial" size="2">&#4709;&#4827;&#4725;<br>
      			Quantity</font></b></td>
			<td width="100" align="left"><b><font face="Arial" size="2">&#4637;&#4653;&#4632;&#4651;<br>
      			Remark</font></b></td>
    	 	</tr>
     
		<%

		 ' List the forms in a table
 		WHILE NOT RS.EOF

		%>		
		
	
		<%
	 	SQL_query_quantity_in_store = "SELECT SUM(Quantity_Received_by_BoF - Quantity_Sent_from_BoF) AS Quantity_in_Store FROM Shipments WHERE Form_ID = "& RS("Form_ID") &";"
	 	Set RS_quantity = Conn1.Execute(SQL_query_quantity_in_store)
		%>     

     	<tr>
     		<td width="7%"><font face="Arial" size="2"><%= FormCount %> &nbsp;
	 		</font></td>
     		     
     	
	 
	 	<td width="15%" align="center"><font face="Arial" size="2"> <%=RS("Model_No")%>&nbsp;</font></td>   
	
	
     		<td width="15%"><font face="Arial" size="2">&nbsp;

		' </font></td>
	 
	 	<td width="25%" align="center"><font face="Arial" size="2">&nbsp;  
		  
			<%If Not IsNumeric(RS_quantity("Quantity_in_Store")) Then%>
     	 			0
     			<%Else%>
		 		<%=RS_quantity("Quantity_in_Store")%>
	 		<%End If%>		
			</font></td>  
		<td width="10%" align="center"><font face="Arial" size="2">&nbsp;</font></td>    
		   
     
     	</tr>
     	<%
     	RS.MoveNext
     
 		WEND
 
 		RS.Close
 		Conn1.Close
	%>
    
      </table>
  
  
<br>
</td>
</tr>
</table>    
<br>

<!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#426262">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%">
   <p align="center"><font color="#FFFFFF" face="Arial" size="1" ><b><i>Developed by the ICT Directorate of BoF</i></b></font></td>
  </tr>
</table>

</table>

</td></tr><!--msnavigation--></table></body>
</html>