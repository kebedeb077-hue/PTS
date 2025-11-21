<%
Response.Expires = -1000 ' Makes the browser not cache this page
Response.Buffer = True   ' Buffers the content so the Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("default.asp")
End If
%>  

<%
  Dim PAGE_SIZE_CONSTANT
  PAGE_SIZE_CONSTANT = 20 ' This is the number of records that will be shown on one page
%>

<%
Dim Selected_Office_ID
Selected_Office_ID = Request.QueryString("Office_ID") 
%>

<%
If Request.QueryString("ToExcel") = "True" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "filename="&Selected_Office_ID&".xls"
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
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Office Shipment History</title>
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
  <!--webbot bot="SaveResults" u-file="_private/form_results.csv" s-format="TEXT/CSV" s-label-fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" -->
  <br>
  
  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="100%" style="border-style: solid; border-width: 2" bordercolor="#426262">
<% End If %>     
 



<!-- #include file="adovbs.inc" -->

<%
' The connection string
Dim CONN_STRING
CONN_STRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=admin; Password=serbof;"

' The name of this file so all the links and forms will still
' work if it is renamed
Dim SCRIPT_NAME
SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

' Declare the standard variables
Dim cnnDBEdit, rstDBEdit, rstFormList  ' ADO objects

Dim strSQL      			' To hold various SQL Strings
Dim Selected_Shipment_ID   ' Used to keep track of the record in play
Dim Added_Shipment_ID
Dim strSQLFormList

Dim Validated_Form			' To hold the state of data entered
Dim DateEntered				' Date entered by the user in the update and new record forms

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

%>

<%
' Get Office_Name

Dim Selected_Office_Name
strSQL = "SELECT Office_Name FROM Offices WHERE Office_ID = "& Selected_Office_ID &";"
Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
rstDBEdit.Open strSQL, CONN_STRING, adOpenStatic, adLockReadOnly, adCmdText
Selected_Office_Name = rstDBEdit.Fields("Office_Name").Value
rstDBEdit.Close
Set rstDBEdit = Nothing
%>


<%

' This is the link displayed at the bottom of page
Dim BACK_TO_LIST_TEXT
BACK_TO_LIST_TEXT = "<a href=""" & SCRIPT_NAME & "?Office_ID=" &Selected_Office_ID& """ style=""text-decoration: none; font-weight: 700""><font face=""Arial"" size=""2"" color=""black""><b><p>"_
	& "&#4808;&#4848; &#4747;&#4619; &#4616;&#4632;&#4632;&#4616;&#4661; &#4773;&#4826;&#4613; &#4875; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;"_
	& "<br>Click here to go back </font></p></b>"

%>

<%

' This function is called from addsave and editsave to verify the input entered by the user
FUNCTION Error_Check

		'Form validation
		Validated_Form = true
		
		If NOT IsNumeric(Request.Form("Quantity_Sent_from_BoF")) Then
		%>
			<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
			Error</font></b>
            <br><br>
			<b><span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4720;&#4619;&#4776;&#4813; &#4709;&#4827;&#4725;</span></b><span style="font-size: 12.0pt; font-family: Arial"> 
			&#4704;&#4673;&#4901;&#4653; &#4632;&#4614;&#4757; &#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
			The <b>Quantity Sent</b> has to be a number! Please go back and correct the error.</font>
		<%
			Validated_Form = false		
	
			Else If Request.Form("Quantity_Sent_from_BoF")<1 Then
			%>
				<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
				Error</font></b>
            	<br><br>
				<b><span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4720;&#4619;&#4776;&#4813; &#4709;&#4827;&#4725;</span></b><span style="font-size: 12.0pt; font-family: Arial"> 
				&#4776;&#4828;&#4654; &#4632;&#4709;&#4616;&#4901; &#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
				The <b>Quantity Sent</b> has to be greater than 0! Please go back and correct the error.</font>				
			<%
				Validated_Form = false
				
				Else If NOT IsNumeric(Request.Form("Unit_Price")) Then
				%>
					<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
					Error</font></b>
           			<br><br>
					<b><span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4768;&#4757;&#4849; 
					&#4811;&#4875; &nbsp;</span></b><span style="font-size: 12.0pt; font-family: Arial">&#4704;&#4673;&#4901;&#4653; &#4632;&#4614;&#4757; 
					&#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
					The <b>Unit Price</b> has to be a number! Please go back and correct the error.</font>			
				<%
					Validated_Form = false
					
					Else If Request.Form("Unit_Price")<=0 Then
					%>
						<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
						Error</font></b>
       	    			<br><br>
						<b><span style="font-size: 12.0pt; font-family: Arial">
						&#4840;&#4768;&#4757;&#4849; &#4811;&#4875; </span></b><span style="font-size: 12.0pt; font-family: Arial">&nbsp;&#4776;&#4828;&#4654; 
						&#4632;&#4709;&#4616;&#4901; &#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
						The <b>Unit Price</b> has to be greater than 0! Please go back and correct the error.</font>			
					<%
						Validated_Form = false
					End If	
				End If
			End If
		End If	

END FUNCTION


' This is the main select of the script
' Choose what to do by looking at the action parameter

Select Case LCase(Trim(Request.QueryString("action")))
    
	Case "addsave"
	
		' Check for errors
		Error_Check
		
		' If data is fine, enter it in the database
		If Validated_Form Then
	
	
			' After the user enters the data in the form above and presses submit, the script comes here
			' This is where a new record is created and stored in the database
			
			' Select an empty RS
			strSQL = "SELECT * FROM Shipments;"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

			' Add our record and set it's values.  
			rstDBEdit.AddNew
		
			rstDBEdit.Fields("Office_ID").Value     		= Selected_Office_ID
			rstDBEdit.Fields("Day").Value      			= Request.Form("Day")
			rstDBEdit.Fields("Month").Value      			= Request.Form("Month")
			rstDBEdit.Fields("Year").Value      			= Request.Form("Year")
			rstDBEdit.Fields("Form_ID").Value     			= Request.Form("Form_ID")
			rstDBEdit.Fields("Receipt_No").Value     		= CStr(Replace(Request.Form("Receipt_No"), "'", "''"))
			rstDBEdit.Fields("Starting_Serial_No").Value    = CStr(Replace(Request.Form("Starting_Serial_No"), "'", "''"))
			rstDBEdit.Fields("Ending_Serial_No").Value      = CStr(Replace(Request.Form("Ending_Serial_No"), "'", "''"))
			rstDBEdit.Fields("Quantity_Sent_from_BoF").Value = Request.Form("Quantity_Sent_from_BoF")
			rstDBEdit.Fields("Unit_Price").Value     		= Request.Form("Unit_Price")
			rstDBEdit.Fields("Total_Price").Value     		= Request.Form("Unit_Price")*Request.Form("Quantity_Sent_from_BoF")
			rstDBEdit.Fields("Comments").Value      		= CStr(Replace(Request.Form("Comments"), "'", "''"))

			rstDBEdit.Update

			' Get the id of the record just added. This might cause
			' problems with some DB providers, but it works with
			' Microsoft Access and SQL Server.
			Added_Shipment_ID = rstDBEdit.Fields("Shipment_ID").Value

			rstDBEdit.Close
			Set rstDBEdit = Nothing

		%> 	
			<b>
			<span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4616;&#4768;&#4619;&#4707; &#4776;&#4720;&#4635; 
			&#4768;&#4661;&#4720;&#4851;&#4848;&#4653; &#4768;&#4850;&#4661; &#4840;&#4720;&#4632;&#4824;&#4872;&#4704; &#4645;&#4653;&#4909;&#4725;</span><font face="Verdana" size="4" color="#006666"><br>
			</font><font face="Verdana" color="#006666">Registered new 
			distribution to</font><font face="Verdana"> </font>
  	 		<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></b><font face="Arial" size="3">
  			<br><br> 
		<%
			Response.Write("<p>Shipment Id #" & Added_Shipment_ID & " added!</p>")
	 End If
		Response.Write(BACK_TO_LIST_TEXT)
	
	
	Case "deleteconfirm"
	
		'Before a record is deleted, confirmation is required
		
		Selected_Shipment_ID = Request.QueryString("Shipment_ID")
	
	%> 
		<b><span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4645;&#4653;&#4909;&#4721;&#4757; 
		&#4656;&#4653;&#4829;</span></b><font face="Arial" size="3"><b><font color="#006666" size="4" face="Verdana"><br>
		</font><font color="#006666" face="Verdana">Delete Distribution</font></b>
	    <br>
		<br>
		<font face="Arial">
		<b>&#4704;&#4773;&#4653;&#4877;&#4901; &#4616;&#4635;&#4901;&#4939;&#4725; &#4845;&#4936;&#4621;&#4875;&#4617;?</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
		</font>
		<font face="Arial">
	    <b><br>
		Are you sure you want to delete?</b></font><font face="Arial" size="3"><br>
        <br>
                
	    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="448">
	    <tr>
		<td width="60" height="16" style="border-style: solid; border-width: 0; " align="center">
	    </td>
     	<td width="60" height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center">
	    <a href="<%= SCRIPT_NAME %>?action=delete&Office_ID=<%=Selected_Office_ID%>&Shipment_ID=<%= Selected_Shipment_ID %>" style="text-decoration: none; font-weight: 700">
    	<b><font face="Verdana" size="2" color="#FFFFFF">
	    Yes</font></b></a></td>
     
	    <td width="60">
	    </td>
      
     	<td width="120" bgcolor="#006699" style="border-style: solid; border-width: 2" align="center">
	    <a href="<%= SCRIPT_NAME %>?Office_ID=<%=Selected_Office_ID%>" style="text-decoration: none; font-weight: 700">
    	<b>
	    <font color="#FFFFFF" size="2" face="Verdana">Don't 
    	Delete</font></b></a></td>
      
	    <td width="148" align="center">
    	&nbsp;</td>
		</tr>
	    <tr>
	    <td width="60" height="16" style="border-style: solid; border-width: 0; " align="center">
	    </td>
     	<td width="60" height="16" align="center">
	    </td>
     
		<td width="60">
	    &nbsp;</td>
      
      	<td width="268" align="center" colspan="2">
	    <p align="left">
    	<a href="<%= SCRIPT_NAME %>?Office_ID=<%=Selected_Office_ID%>" style="text-decoration: none; font-weight: 700">
		<font face="Arial" size="3">
	    <b>
    	<font color="#006699" size="2">(Takes back to the Distributions List)</font></b></font><font color="#006699" size="2" face="Arial">
	    </font>
    	</a>
      
	    </td>
		</tr>
	    </table>
	    <br><br>	
	
	<%
	Case "delete"
		
		' Once confirmation is received in the deleteconfirm stage the record is deleted

		' Get the id to delete
		Selected_Shipment_ID = Request.QueryString("Shipment_ID")

		strSQL = "DELETE FROM Shipments WHERE Shipment_ID=" & Selected_Shipment_ID & ";"

		Set cnnDBEdit = Server.CreateObject("ADODB.Connection")
		cnnDBEdit.Open CONN_STRING

		cnnDBEdit.Execute strSQL, adAffectAll, adExecuteNoRecords
		
		cnnDBEdit.Close
		Set cnnDBEdit = Nothing

	%> 
		<b><span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
		&#4776;&#4768;&#4619;&#4707; &#4776;&#4720;&#4635; &#4768;&#4661;&#4720;&#4851;&#4848;&#4653;&nbsp; &#4645;&#4653;&#4909;&#4725; &#4840;&#4720;&#4640;&#4648;&#4824;</span></b></font><b><font color="#006666" face="Verdana" size="4"><br>
		</font><font color="#006666" face="Verdana">Deleted Distribution from</font><font face="Verdana">
		</font>
	 	<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></b><font face="Arial" size="3">
  		<br><br> 		
	<%	
		
		Response.Write("Shipment Id #" & Selected_Shipment_ID & " deleted!")
		Response.Write(BACK_TO_LIST_TEXT)
		

		
	Case "edit"
		
		' First of a 2 part process... build a form with the
		' values from the database
		Selected_Shipment_ID = Request.QueryString("Shipment_ID")

		strSQL = "SELECT * FROM Shipments WHERE Shipment_ID=" & Selected_Shipment_ID & ";"

		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

	%> 
		<b>
		<span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4616;&#4768;&#4619;&#4707; &#4776;&#4720;&#4635; 
		&#4768;&#4661;&#4720;&#4851;&#4848;&#4653; &#4645;&#4653;&#4909;&#4721;&#4757; &#4768;&#4653;&#4637;</span></b><b><font face="Arial" size="3"><font color="#006666" face="Verdana" size="4"><br>
		</font><font color="#006666" face="Verdana">Edit Distribution for</font></font><font face="Verdana" size="3"> </font>
		<font face="Arial" size="3">
		<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></font></b><font face="Arial" size="3">
  		<br><br> 
	<%

		If Not rstDBEdit.EOF Then
			%> </b>

			<form action="<%= SCRIPT_NAME %>?action=editsave&Office_ID=<%=Selected_Office_ID%>" method="post">
				<font face="Arial" size="2">
				<input type="hidden" name="Shipment_ID" value="<%= Selected_Shipment_ID %>" />
							
				<table border="1" cellspacing="0" cellpadding="0" width="500">
                  <tr>
	                    <td width="39%">
							<b>
							<font face="Arial" size="2">
						    &#4672;&#4757;<br>
						    Date:&nbsp;&nbsp;
                		    </font>
		                    </b>
		                </td>
	                    <td width="61%">
							<font face="Arial" size="2">
							<select size="1" name="Day">
							<%					
							DIM x(32)
							x(rstDBEdit.Fields("Day").Value)="selected"
							For i=1 to 30
							%>
                			<option <%=x(i)%>><%=i%></option>  
		               		<%Next%> 
        		        	</select>
                			<select size="1" name="Month" style="font-family: Visual Geez Unicode">
		                	<%
        		        	DIM y(13)
							y(rstDBEdit.Fields("Month").Value)="selected"
		                	For i=1 to 12
        		        	%>
                		    <option value="<%=i%>" <%=y(i)%>><%=Months_Array(i)%></option>
		                	<%Next%>
        		        	</select>
                			<select size="1" name="Year">
		                	<%
							DIM z(40)
							z(rstDBEdit.Fields("Year").Value-1990)="selected"
		                	For i=2016 to 2050
        		        	%>
                			<option <%=z(i-2015)%>><%=i%></option>
		                	<%Next%> 
				            </select> 
                		    </font>
                		 </td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<font face="Arial" size="3"><b><font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; 
                			&#4819;&#4845;&#4752;&#4725;<br>
			                </font></b></font>
							<b>
							<font face="Arial" size="2">
							Form Name:&nbsp;&nbsp; </font></b>
						</td>
	                    <td width="61%">
 						    <font face="Arial" size="2">
			 				<%
						    'For the dropdown list get the Form Names
							strSQLFormList = "SELECT * FROM Forms ORDER BY Form_Name;"
							Set rstFormList = Server.CreateObject("ADODB.Recordset")
							rstFormList.Open strSQLFormList, CONN_STRING
							%>
							<select name="Form_ID" style="font-family: Visual Geez Unicode">
								
 							<%
							Do While Not rstFormList.EOF
					
							If (rstFormList.Fields("Form_ID").Value = rstDBEdit.Fields("Form_ID").Value) Then
							%>
							<option value="<%= rstFormList.Fields("Form_ID").Value %>" selected><%= rstFormList.Fields("Form_Name").Value %>, <%= rstFormList.Fields("Model_No").Value %></option>
							<%
							Else
							%>
							<option value="<%= rstFormList.Fields("Form_ID").Value %>"><%= rstFormList.Fields("Form_Name").Value %>, <%= rstFormList.Fields("Model_No").Value %></option>
							<%
							End If
							rstFormList.MoveNext
							Loop
							rstFormList.Close
							Set rstFormList = Nothing
							%>
							</select>
                  </tr>
                  <tr>
	                    <td width="39%">
							<font face="Arial" size="2">
							<b>
					        &#4808;&#4906; &#4840;&#4614;&#4752;&#4704;&#4725; &#4638;/22<br>
						    Receipt No:</b></font>
						</td>
            	        <td width="61%">
							<font face="Arial" size="2">
            		        <input type="text" name="Receipt_No" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Receipt_No").Value) %>" size="30" />
            		    </td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<b>
							<font face="Arial" size="2">
							&#4776;</font></b><font face="Arial" size="2"><b><br>
						    Starting Serial No:          
            		    </b>
            		    </td>
	                    <td width="61%">
							<font face="Arial" size="2">
    	        	        <input type="text" name="Starting_Serial_No" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Starting_Serial_No").Value) %>" size="30" /></font>
            	        </td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<b>
							<font face="Arial" size="2">
							&#4773;&#4661;&#4776;</font></b><font face="Arial" size="2"><b><br>
							Ending Serial No: </b>
						</td>
	                    <td width="61%">
							<font face="Arial" size="2">
            		        <input type="text" name="Ending_Serial_No" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Ending_Serial_No").Value) %>" size="30" />
           		        </td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<b>
							<font face="Arial" size="2">
							&#4840;&#4720;&#4619;&#4776;&#4813; &#4709;&#4827;&#4725;</font></b><font face="Arial" size="2"><b><br>
							Quantity Sent: </b>
						</td>
	                    <td width="61%">
							<font face="Arial" size="2">
                		    <input type="text" name="Quantity_Sent_from_BoF" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Quantity_Sent_from_BoF").Value) %>" size="30" />
               		    </td>
                  </tr>
                  <tr>
		                <td width="39%">
							<b>
							<font face="Arial" size="2">
						    &#4843;&#4757;&#4849; &#4811;&#4875;</font></b><font face="Arial" size="2"><b><br>
							Unit Price (Birr): </b>
						</td>
	                    <td width="61%">
							<font face="Arial" size="2">
                		    <input type="text" name="Unit_Price" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Unit_Price").Value) %>" size="30" /></font>
              		    </td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<font face="Arial" size="2">
							<b>
							&#4896;&#4677;&#4619;&#4619; &#4811;&#4875;<br>
							Total Price (Birr): </b></font>
							<b>
							<font face="Arial" size="2">
							&nbsp;&nbsp; </font></b>
						</td>
	                    <td width="61%">
							<i><span style="font-size: 10.0pt; font-family: Arial">&#4776;&#4768;&#4757;&#4849; &#4811;&#4875;&#4755; 
	        		        &#4776;&#4709;&#4827;&#4725; &#4704;&#4672;&#4901;&#4723; &#4840;&#4634;&#4656;&#4619;</span></i><font face="Arial" size="2"><i><br>
    	            		It will be automatically calculated from Unit Price and Quantity.</i></font>
                		</td>
                  </tr>
                  <tr>
	                    <td width="39%">
							<b>
							<font face="Arial" size="2">
							&#4768;&#4661;&#4720;&#4843;&#4840;&#4725;<br>
							Comments: </font></b>
						</td>
	                    <td width="61%">
							<font face="Arial" size="2">
            		        <input type="text" name="Comments" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Comments").Value) %>" size="30" /></font>
            		    </td>
                  </tr>
                  <tr>
	                    <td width="33%" colspan="2">
							<font face="Arial" size="2">
							<p align="right">
							<input type="submit" name="Update Distribution" value="Update Distribution" style="color: #FFFFFF; font-weight: bold; font-family: Verdana; background-color: #525252">
						</td>
                  </tr>
                </table>
                <p>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font size="2">&nbsp;</font></p>
			</form>
			<%
		Else
			%>
			<span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4720;&#4896;&#4840;&#4672;&#4813; &#4704;&#4829;&#4653;&#4829;&#4649; &#4813;&#4661;&#4901; 
			&#4768;&#4621;&#4720;&#4872;&#4760;&#4637;</span></font><font face="Arial"><br>
			Record not found!</font><font face="Arial" size="2">
				<%
		End If

		rstDBEdit.Close
		Set rstDBEdit = Nothing

		Response.Write(BACK_TO_LIST_TEXT)
		

	Case "editsave"

		' Check for errors
		Error_Check
		
		Selected_Shipment_ID = Request.Form("Shipment_ID")
		
		' If data is fine, enter it in the database
		If Validated_Form Then
		
			' Part 2 of 2: Here's where we save the values that the
			' user entered back to the DB.  
			' Select the record
			strSQL = "SELECT * FROM Shipments WHERE Shipment_ID=" & Selected_Shipment_ID & ";"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText
	
			rstDBEdit.Fields("Day").Value      				= Request.Form("Day")
			rstDBEdit.Fields("Month").Value      			= Request.Form("Month")
			rstDBEdit.Fields("Year").Value      			= Request.Form("Year")
			rstDBEdit.Fields("Form_ID").Value     			= Request.Form("Form_ID")
			rstDBEdit.Fields("Receipt_No").Value     		= CStr(Replace(Request.Form("Receipt_No"), "'", "''"))
			rstDBEdit.Fields("Starting_Serial_No").Value    = CStr(Replace(Request.Form("Starting_Serial_No"), "'", "''"))
			rstDBEdit.Fields("Ending_Serial_No").Value      = CStr(Replace(Request.Form("Ending_Serial_No"), "'", "''"))
			rstDBEdit.Fields("Quantity_Sent_from_BoF").Value = Request.Form("Quantity_Sent_from_BoF")
			rstDBEdit.Fields("Unit_Price").Value     		= Request.Form("Unit_Price")
			rstDBEdit.Fields("Total_Price").Value     		= Request.Form("Unit_Price")*Request.Form("Quantity_Sent_from_BoF")
			rstDBEdit.Fields("Comments").Value      		= CStr(Replace(Request.Form("Comments"), "'", "''"))
	
			rstDBEdit.Update

			rstDBEdit.Close
			Set rstDBEdit = Nothing

			%> 
   		    <b>
			<span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4616;&#4768;&#4619;&#4707; &#4776;&#4720;&#4635; 
			&#4768;&#4661;&#4720;&#4851;&#4848;&#4653; &#4840;&#4720;&#4661;&#4720;&#4779;&#4776;&#4616; &#4645;&#4653;&#4909;&#4725;</span><font color="#006666" face="Verdana" size="4"><br>
			</font></b></font><b>
			<font face="Verdana" color="#006666">
			Edited Distribution for</font></b></font><b><font face="Verdana" size="3"> </font>
			<font face="Arial" size="3">
 			<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></font></b><font face="Arial" size="2"><font face="Arial" size="3">
	  		<br><br> 	
			<%	
			Response.Write("<p>Shipment Id #" & Selected_Shipment_ID & " updated!</p>")
			Response.Write(BACK_TO_LIST_TEXT)
	 Else
	 
	 %> 
	  <a href=" <%= SCRIPT_NAME %>?action=edit&Office_ID=<%=Selected_Office_ID%>&Shipment_ID=<%= Selected_Shipment_ID %>" style="text-decoration: none; font-weight: 700">
	  <font face="Arial" size="2" color="black"><b><p>
	  &#4808;&#4848; &#4747;&#4619; &#4616;&#4632;&#4632;&#4616;&#4661; &#4773;&#4826;&#4613; &#4875; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;
	  <br>Click here to go back </font></p></b></a>
	 
	 <%
	 End If


	Case Else ' view

		' Paging of the records
		Dim iPageSize       'How big our pages are
		Dim iPageCount      'The number of pages we get back
		Dim iPageCurrent    'The page we want to show
		
		iPageSize = PAGE_SIZE_CONSTANT
	
		' The default action... just lists the records in the DB for the selected Office
		strSQL = "SELECT * FROM Shipments WHERE Office_ID = "& Selected_Office_ID &" ORDER BY (((Year-1)*360)+((Month-1)*30)+Day), Shipment_ID;"
		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.PageSize = iPageSize
		rstDBEdit.Open strSQL, CONN_STRING, adOpenStatic, adLockReadOnly, adCmdText
		
		' Get the count of the pages using the given page size
        iPageCount = rstDBEdit.PageCount

		' Retrieve page to show or default to 1
		If Request.QueryString("page") = "" Then
			iPageCurrent = iPageCount
		Else
			iPageCurrent = CInt(Request.QueryString("page"))
		End If
		
		' If the request page falls outside the acceptable range,
		' give them the closest match (1 or max)
		If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
		If iPageCurrent < 1 Then iPageCurrent = 1

		%> 
			<b><font face="Verdana" size="4" color="#006666">&#4645;&#4653;&#4909;&#4725; &#4723;&#4650;&#4781;<br>
			</font>
			<font face="Verdana" size="3"> 
			<font color="#006666">
			<font size="4">
			Distribution History</font><br>
			&#4840;&#4706;&#4653;&#4813;&#4661;&#4637; - Office Name:</font>
			<font color="#C46200"><%=Selected_Office_Name%></font></font>
			</b>
			<br><br> 
		
		<%
		If NOT iPageCount = 0 Then
			' Move to the selected page
			rstDBEdit.AbsolutePage = iPageCurrent
		%>
		 <font size="2" face="Arial" color="#000000">Page <b><%= iPageCurrent %></b> of <b><%= iPageCount %></b> 
		(</font><font face="Verdana" size="4"><font size="2" face="Arial" color="#000000">&#4704;&#4672;&#4757; 
		&#4677;&#4848;&#4637; &#4720;&#4776;&#4720;&#4621; - </font> </font> 
		<font size="2" face="Arial" color="#000000">Ordered by date)</font>
		<br><br>
		<%
		End If
		%>	
			
		<table border="1" cellspacing="0" cellpadding="0">
		<thead>
		<tr>
			<th height="16" class="NoPrint"><font face="Arial" size="2">
            Distribution ID</font></th>
			<th height="16">
			<font face="Arial" size="2">&#4672;&#4757;<br>
			Date</font></th>
			<th height="16">
            <font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; &#4819;&#4845;&#4752;&#4725;<br>
            Form Name</font></th>
            <th height="16"><b>
            <font face="Arial" size="2">&#4632;&#4610;/&#4638;&#4852;&#4621; &#4673;&#4901;&#4653;<br>
            Model No</font></th>
            <th height="16">
			<font face="Arial" size="2">&#4808;&#4906; &#4840;&#4614;&#4752;&#4704;&#4725; &#4638;/22<br>
            Receipt No</font></th>
			<th height="16">
			<font face="Arial" size="2">&#4776;<br>
			Starting Serial No</th>
			<th height="16">
			<font face="Arial" size="2">&#4773;&#4661;&#4776;<br>
            Ending Serial No</th>
			<th height="16">
			<font face="Arial" size="2">&#4840;&#4720;&#4619;&#4776;&#4813; &#4709;&#4827;&#4725;<br>
            Quantity Sent</th>
			<th height="16">
			<font face="Arial" size="2">&#4843;&#4757;&#4849; &#4811;&#4875;<br>
            Unit Price (Birr)</th>
			<th height="16">
			<font face="Arial" size="2">&#4896;&#4677;&#4619;&#4619; &#4811;&#4875;<br>
            Total Price (Birr)</th>
			<th height="16" class="NoPrint">
			<font face="Arial" size="2">&#4768;&#4661;&#4720;&#4843;&#4840;&#4725;<br>
            Comments</font></th>
<% If Request.QueryString("ToExcel") = "" Then %>
			<th height="16" width="55" class="NoPrint">
            <span style="font-family: Arial"><font size="2">&#4768;&#4901;&#4939;</font></span><font face="Arial" size="2"><br>
            Delete</font></th>
			<th height="16" width="55" class="NoPrint">
            <span style="font-family: Arial"><font size="2">&#4768;&#4653;&#4637;</font></span><font face="Arial" size="2"><br>
            Edit</font></th>
<% End If %>
		</tr>
		</thead>
		<tbody>		
		<% If iPageCount = 0 Then %>
		<tr height="16"><td colspan="12">
          <font size="2">&#4637;&#4757;&#4637; &#4840;&#4720;&#4632;&#4824;&#4872;&#4704; &#4752;&#4872;&#4653; &#4768;&#4621;&#4720;&#4872;&#4760;&#4637;&#4961;&#4961; </font>
          <font size="2" font="Arial" face="Arial">No records found.</font></td></tr>
		
		<% Else %>
		    <%
		    iRecordsShown = 0

		    Do While iRecordsShown < iPageSize And Not rstDBEdit.EOF

		    	' Get the Form list to convert ID into name when displaying
				strSQLFormList = "SELECT Form_Name, Model_No FROM Forms WHERE Form_ID = "& rstDBEdit.Fields("Form_ID").Value &"; "
				Set rstFormList = Server.CreateObject("ADODB.Recordset")
				rstFormList.Open strSQLFormList, CONN_STRING
			%>
			<tr height="16">			
				<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstDBEdit.Fields("Shipment_ID").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Day").Value %>/<%= Months_Array(rstDBEdit.Fields("Month").Value) %>/<%= rstDBEdit.Fields("Year").Value %></font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstFormList.Fields("Form_Name").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstFormList.Fields("Model_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Receipt_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Starting_Serial_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Ending_Serial_No").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Quantity_Sent_from_BoF").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Unit_Price").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Total_Price").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" class="NoPrint"><font face="Arial" size="2"><%= rstDBEdit.Fields("Comments").Value %>&nbsp;</font></td>
<% If Request.QueryString("ToExcel") = "" Then %>
				<td height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <font face="Verdana" size="2">
                <a href="<%= SCRIPT_NAME %>?action=deleteconfirm&Office_ID=<%=Selected_Office_ID%>&Shipment_ID=<%= rstDBEdit.Fields("Shipment_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font color="#FFFFFF">Delete</font></a></font></td>
				<td height="16" bgcolor="#006666" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <a href="<%= SCRIPT_NAME %>?action=edit&Office_ID=<%=Selected_Office_ID%>&Shipment_ID=<%= rstDBEdit.Fields("Shipment_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font face="Verdana" size="2" color="#FFFFFF" size="2">Edit</font></a></td>
<% End If %>
			</tr>
			<%
				rstFormList.Close
				Set rstFormList = Nothing
				
				iRecordsShown = iRecordsShown + 1
				rstDBEdit.MoveNext
		    Loop
		    %>
		    
		    <tr>
<% If Request.QueryString("ToExcel") = "" Then %>		    
		    <td colspan="13">
<% Else %> &nbsp;<td colspan="11">
<% End If %>		    
		    <%
		    ' Show "previous" and "next" page links which pass the page to view
			' and any parameters needed to rebuild the query. 
			If iPageCurrent > 1 Then
			%>
			<font face="Arial" size="2">
			<a href="<%= SCRIPT_NAME %>?Office_ID=<%=Selected_Office_ID%>&page=<%= iPageCurrent - 1 %>">[&lt;&lt; Prev]</a>
			</font>
			<%
			End If

			' You can also show page numbers:
			For I = 1 To iPageCount
				If I = iPageCurrent Then
				%>
					<font face="Arial" size="2">					
					<%= I %>
					</font>
				<%
				Else
				%>
					<font face="Arial" size="2">
					<a href="<%= SCRIPT_NAME %>?Office_ID=<%=Selected_Office_ID%>&page=<%= I %>"><%= I %></a>
					</font>
				<%
				End If
			Next 'For I

			If iPageCurrent < iPageCount Then
			%>
				<font face="Arial" size="2">
				<a href="<%= SCRIPT_NAME %>?Office_ID=<%=Selected_Office_ID%>&page=<%= iPageCurrent + 1 %>">[Next &gt;&gt;]</a>
				</font>
			<%
			End If
		    %>
		    
		    </td>
		    </tr>
		<% End If %>
		</tbody>
	</table>
	<br>
	<font face="Arial" size="2" class="NoPrint"><b>
	<font color="#000000">(*Distribution_ID is automatically generated by the system and cannot be edited by the user)</font>
	</b></font>
	<br>
		
	
		
<% If Request.QueryString("ToExcel") = "" Then %>		
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
          <tr>
            <td width="50%" valign="top" align="center">
            <br>
            <br>
      		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="250" class="NoPrint">
	        <tr>
            <td width="100%">
            <form id="ExcelForm" action="<%= SCRIPT_NAME %>?ToExcel=True&Office_ID=<%=Selected_Office_ID%>&page=<%=iPageCurrent%>" method="post">
			<img border="0" src="images/Excel%20Button.JPG" align="top">
            <font size="1">
			<input id="btnExcel" type="submit" value="Send list to Excel" style="color: #FFFFFF; background-color: #525252; font-weight:bold; font-family:Verdana; font-size:10pt"/></font></font>
            </form>
 			
            <form id="PrintForm" action="<%= SCRIPT_NAME %>" class="NoPrint">
			<img border="0" src="images/Print%20Button.JPG" align="top">
            <font size="1">
			<input id="btnPrint" type="button" value="Print list" class="PrintButton" onclick="window.print();" style="color: #FFFFFF; background-color: #525252; font-weight:bold; font-family:Verdana"/></font></font>
            </form>
			</td>            
            </tr>
            </table>
            
            </td>
            <td width="45%">
            
    <%	' Adding a new Shipment record
		' First of a 2 part process... build an empty form for the user to enter data. 
		' The fields come from the Shipments table
	%>

		<font face="Arial" size="2">
			
			<form action="<%= SCRIPT_NAME %>?action=addsave&Office_ID=<%=Selected_Office_ID%>" method="post" class="NoPrint">
				<table border="1" cellspacing="0" cellpadding="0" width="500">
                  <tr>
    	             <td width="100%" colspan="2" height="35">
  						<b>
	                    <span style="font-size: 12.0pt; font-family: Verdana; color: #006666">
        	            &#4768;&#4850;&#4661; &#4661;&#4653;&#4909;&#4725; &#4632;&#4829;&#4877;&#4709;</span></b><font face="Arial" size="3"><b><font color="#006666" face="Verdana" size="3"><br>
            	        Register a new Distribution:</font></b>          
                    </td>
                  </tr>
                  <tr>
                     <td width="39%">
						<font face="Arial" size="2">
						<b>
					    &#4672;&#4757;<br>
					    Date:&nbsp;&nbsp;          
            		   	</b>
            		</td>
                    <td width="61%">
						<font face="Arial" size="2">
        		        <select size="1" name="Day">
		                <%For i=1 to 30%>
        		        <option><%=i%></option>  
		                <%Next%> 
        		        </select>
		                <select size="1" name="Month" style="font-family: Visual Geez Unicode">
        		        <%For i=1 to 12%>
		                <option value="<%=i%>"><%=Months_Array(i)%></option>
        		        <%Next%>
		                </select>
        		        <select size="1" name="Year">
		                <%For i=2016 to 2050%>
        			    <option><%=i%></option>
		                <%Next%> 
        		        </select> </font>
                    </td>
                </tr>
                <tr>
                    <td width="39%">
						<font face="Arial" size="3"><b><font face="Arial" size="2">&#4840;&#4629;&#4725;&#4632;&#4725; 
		                &#4819;&#4845;&#4752;&#4725;<br>
		                </font></b></font>
						<b>
						<font face="Arial" size="2">
						Form Name:&nbsp;&nbsp; </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
						<%
					    ' For the dropdown list get the Form Names
						strSQLFormList = "SELECT * FROM Forms ORDER BY Form_Name;"
						Set rstFormList = Server.CreateObject("ADODB.Recordset")
						rstFormList.Open strSQLFormList, CONN_STRING
						%>
						<select name="Form_ID" style="font-family: Visual Geez Unicode">
	 					<%
						Do While Not rstFormList.EOF
						%>
						<option value="<%= rstFormList.Fields("Form_ID").Value %>"><%= rstFormList.Fields("Form_Name").Value %>, <%= rstFormList.Fields("Model_No").Value %></option>
						<%
						rstFormList.MoveNext
						Loop
						rstFormList.Close
						Set rstFormList = Nothing
						%>
						</select>
	                    </font>
                    </td>
                </tr>
                <tr>
                    <td width="39%">
		    	        <font face="Arial" size="2">
						<b>
				        &#4808;&#4906; &#4840;&#4614;&#4752;&#4704;&#4725; &#4638;/22</b></font><b><font face="Arial" size="2"><br>
						Receipt No: </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        	            <input type="text" name="Receipt_No" value="" size="30" /> </font>
                    </td>
                  </tr>
               <tr>
                    <td width="39%">
		                <font face="Arial" size="2">
						<b>
						&#4776;</b></font><b><font face="Arial" size="2"><br>
						Starting Serial No: </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        	            <input type="text" name="Starting_Serial_No" value="" size="30" /> </font>
                    </td>
                  </tr>
               <tr>
                    <td width="39%">
						<font face="Arial" size="2">
						<b>
						&#4773;&#4661;&#4776;</b></font><b><font face="Arial" size="2"><br>
						Ending Serial No: </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        	            <input type="text" name="Ending_Serial_No" value="" size="30" /> </font>
                    </td>
               </tr>
               <tr>
                    <td width="39%">
						<font face="Arial" size="2">
						<b>
						&#4840;&#4720;&#4619;&#4776;&#4813; &#4709;&#4827;&#4725;</b></font><b><font face="Arial" size="2"><br>
						Quantity Sent: </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        		        <input type="text" name="Quantity_Sent_from_BoF" value="0" size="30" /> </font>
                    </td>
               </tr>
               <tr>
                    <td width="39%">
		                <font face="Arial" size="2">
						<b>
					    &#4843;&#4757;&#4849; &#4811;&#4875;</b></font><b><font face="Arial" size="2"><br>
						Unit Price (Birr): </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        	            <input type="text" name="Unit_Price" value="0" size="30" /> </font>
                    </td>
               </tr>
               
               <tr>
                    <td width="39%">
						<font face="Arial" size="2">
						<b>
						&#4896;&#4677;&#4619;&#4619; &#4811;&#4875;</b></font><b><font face="Arial" size="2"><br>
						Total Price (Birr): </font></b>
					</td>
                    <td width="61%">
						<i><span style="font-size: 10.0pt; font-family: Arial">&#4776;&#4768;&#4757;&#4849; &#4811;&#4875;&#4755; 
        		        &#4776;&#4709;&#4827;&#4725; &#4704;&#4672;&#4901;&#4723; &#4840;&#4634;&#4656;&#4619;</span></i><font face="Arial" size="2"><i><br>
		                It will be automatically calculated from Unit Price and Quantity.</i></font>
		            </td>
               </tr>
               <tr>
                    <td width="39%">
						<font face="Arial" size="3">
						<b>
						<font face="Arial" size="2">
						&#4768;&#4661;&#4720;&#4843;&#4840;&#4725;</font></b></font><b><font face="Arial" size="2"><br>
						Comments: </font></b>
					</td>
                    <td width="61%">
						<font face="Arial" size="2">
        	            <input type="text" name="Comments" value="" size="30" /> </font>
                    </td>
               </tr>
               <tr>
	                <td width="100%" colspan="2" align="center">
						<font face="Arial" size="2">
						<input type="submit" name="Register Distribution" value="Register Distribution" style="color: #FFFFFF; font-weight: bold; font-family: Verdana; background-color: #525252">
				    </td>
               </tr>
                              
                </table>
				&nbsp;
				
			</form>

            
            </td>
            <td width="15%">
            &nbsp;
            </td>
          </tr>
		</table>
		<p>
		
<% End If %>		
	
		
		<%
			rstDBEdit.Close
			Set rstDBEdit = Nothing
			End Select
        %>

<% If Request.QueryString("ToExcel") = "" Then %>       
		 <p>&nbsp;</p>          
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