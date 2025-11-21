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
	Response.AddHeader "Content-Disposition", "filename=OfficeList.xls"
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
<title>Office List</title>
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
<%End If%>     
 
<!-- #include file="adovbs.inc" -->

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
    

' Declare the standard variables
Dim cnnDBEdit, rstDBEdit, rstOfficeList	' ADO objects

Dim strSQL, strSQLOfficeList    	' To hold various SQL Strings
Dim Selected_Office_ID   			' Used to keep track of the record
Dim Added_Office_ID

Dim Validated_Form				' To hold the state of data entered
Dim OfficeEntered				' Office entered by the user in the update and new record forms


%>

<%

' This function is called from addsave and editsave to verify the input entered by the user
FUNCTION Error_Check

		'Form validation
		Validated_Form = true
		OfficeEntered = CStr(Trim(Replace(Request.Form("Office_Name"), "'", "''")))
		If (len(OfficeEntered)<1) Then
		%>
			<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
			</font><font color="#006666" face="Verdana">Error</font></b>
            <br><br>
			<span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4706;&#4654;&#4813; &#4645;&#4637; &#4707;&#4854; &#4632;&#4614;&#4757; 
			&#4768;&#4845;&#4733;&#4621;&#4637;&#4961;&#4961;&#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4645;&#4637; &#4843;&#4661;&#4872;&#4705;&#4961;&#4961;</span><font face="Arial" size="3"><br>
			The office name cannot be empty! Please go back and enter a name.</font>
		<%
			Validated_Form = false
		
		Else
						
			' Check for duplicates
			If (Request.QueryString("action")="editsave") Then
				strSQL = "SELECT COUNT(*) AS Duplicate_Count FROM Offices WHERE Office_Name LIKE '" & Replace(OfficeEntered, "'", "''") & "' AND NOT(Office_ID = "& Selected_Office_ID &" );"
			Else
				strSQL = "SELECT COUNT(*) AS Duplicate_Count FROM Offices WHERE Office_Name LIKE '" & Replace(OfficeEntered, "'", "''") & "';"
			End If
			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText
		
			If rstDBEdit.Fields("Duplicate_Count").Value > 0 Then
			%>
				<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
				</font><font color="#006666" face="Verdana">Error</font></b>
            	<br><br>
				<span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4706;&#4654;&#4813; &#4645;&#4637; &#4672;&#4853;&#4638; 
				&#4840;&#4720;&#4632;&#4824;&#4872;&#4704; &#4661;&#4616;&#4614;&#4752; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4768;&#4850;&#4661; &#4645;&#4637; &#4843;&#4661;&#4872;&#4705;&#4961;&#4961;</span><font face="Arial" size="3"><br>
				The office <b><%=OfficeEntered%></b> already exists in the list! Please go back and enter a new name.</font>
			<%
				Validated_Form = false
			End If
			
			rstDBEdit.Close
	    	Set rstDBEdit = Nothing
    
	    End If
	    
END FUNCTION

%>

<%
' Choose what to do by looking at the action parameter
Select Case LCase(Trim(Request.QueryString("action")))
    
  
   Case "add"
   
		' First of a 2 part process... build an empty form. 
		%> 
		<b>
		<span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4768;&#4850;&#4661; &#4706;&#4654; &#4904;&#4637;&#4653;</span><font color="#006666" face="Verdana" size="4"><br>
		</font><font color="#006666" face="Verdana">Add New Office</font></b>

  		<br><br> 
			
			<form action="<%= SCRIPT_NAME %>?action=addsave" method="post">
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
                    <td width="17%">
						<font face="Arial" size="3"><b>&#4840;&#4706;&#4653;&#4813; &#4661;&#4637;</b></font><b><br>
               			<font face="Arial" size="2">Office Name:</font></b>
               		</td>
                    <td width="83%">
						<font face="Arial" size="2">
                    	<input type="text" name="Office_Name" value="Office_Name" size="30" /> </font>
                    </td>
                  </tr>
                  <tr>
                    <td width="17%">
						<font face="Arial" size="3"><b>&#4706;&#4654;&#4813;</font><font size="2" face="Arial"><br>
                		Office Type:&nbsp;</font></b>
                	</td>
                    <td width="83%">
                    	<font face="Arial" size="2">
                    	<select size="1" name="Office_Type">
                    	<option selected value="Zone">Zone</option>
                    	<option value="Regional_Burreau">Regional Burreau</option>
                    	<option value="Other">Other</option>
                		</select> 
                		</font>
                    </td>
                  </tr>
                </table>
                <p>
				<font face="Arial" size="2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<br />
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Add Office" value="Add Office" style="color: #FFFFFF; background-color: #525252; font-weight:bold; font-family:Verdana; font-size:10pt"><font size="2">
				</font>
			    </p>
			</form>
			<p>
			</p>
		<%
		Response.Write(BACK_TO_LIST_TEXT)


	Case "addsave"

		' Check for errors
		Error_Check
		
		' If data is fine, enter it in the database
		If Validated_Form Then

			strSQL = "SELECT * FROM Offices;"
   		    Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		    rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

			' Add our record and set its values.  
			rstDBEdit.AddNew
			rstDBEdit.Fields("Office_Name").Value  = OfficeEntered
			rstDBEdit.Fields("Office_Type").Value  = Request.Form("Office_Type")
			rstDBEdit.Update

			' Get the id of the record just added
			Added_Office_ID = rstDBEdit.Fields("Office_ID").Value 

		    rstDBEdit.Close
		    Set rstDBEdit = Nothing

		%> 
  			<b>
            <span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
            &#4840;&#4720;&#4904;&#4632;&#4648; &#4706;&#4654;</span><font color="#006666" size="4" face="Verdana"><br>
            </font></b></font><b>
			<font face="Verdana" color="#006666">
			Added Office</font></b><font face="Arial" size="2"><font face="Arial" size="3">
  			<br><br> 
		<%
			Response.Write("<p>Office Id #" & Added_Office_ID & " added!</p>")
			Response.Write(BACK_TO_LIST_TEXT)
	 Else
	 %>
	  		<a href=" <%= SCRIPT_NAME %>?action=add" style="text-decoration: none; font-weight: 700">
	  		<font face="Arial" size="2" color="black"><b><p>
	  		&#4808;&#4848; &#4747;&#4619; &#4616;&#4632;&#4632;&#4616;&#4661; &#4773;&#4826;&#4613; &#4875; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;
	  		<br>Click here to go back </font></p></b></a>
	 <%
	 End If
	   
	
	Case "deleteconfirm"
	
		Selected_Office_ID = Request.QueryString("Office_ID")
	
		' Check to see if there are any records in the Shipments table that refer to this Office
		' If there is a reference, then deletion will not be permitted
		
		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT COUNT(*) AS Office_Referred FROM Shipments WHERE Office_ID = " & Selected_Office_ID & ";"
		rstDBEdit.Open strSQL, CONN_STRING
		Office_Referred = rstDBEdit.Fields("Office_Referred").Value
      		rstDBEdit.Close

		strSQL = "SELECT COUNT(*) AS Office_Referred FROM Deposits WHERE Office_ID = " & Selected_Office_ID & ";"
		rstDBEdit.Open strSQL, CONN_STRING
		
		If rstDBEdit.Fields("Office_Referred").Value > 0 OR Office_Referred > 0 Then
		%>
			<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
			</font><font color="#006666" face="Verdana">Error</font></b>
            <br><br>
	 
            <font color="#000000">&#4706;&#4654;&#4813;</font><span style="font-family: Arial"><font color="#000000"> &#4704;&#4661;&#4653;&#4909;&#4725;&#4755;/ 
            &#4808;&#4845;&#4637;/ 
            &#4618;&#4661;&#4725; &#4813;&#4661;&#4901; &#4704;&#4811;&#4706;&#4752;&#4725; &#4840;&#4720;&#4896;&#4672;&#4656; &#4661;&#4616;&#4614;&#4752; &#4618;&#4656;&#4648;&#4829; &#4768;&#4845;&#4733;&#4621;&#4637;&#4961;&#4961;</font></span><font color="#000000" face="Arial"><br>
            The Office cannot be deleted because there is a reference to it in the 
            Distributions and/or Deposits list.</font>
	  
	  	<%
			Response.Write(BACK_TO_LIST_TEXT)
	    
		Else	    
	    ' Ask for confirmation of deletion
 	    %> 
 	    	<b><span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
            &#4840;&#4706;&#4654;&#4813;&#4757; &#4661;&#4637; &#4656;&#4653;&#4829;</span><font color="#006666" size="4" face="Verdana"><br>
            </font><font color="#006666" face="Verdana">Delete Office</font></b>
	    	<br>
			<br>
			<font face="Arial">
	   		<b>&#4704;&#4773;&#4653;&#4877;&#4901; &#4616;&#4635;&#4901;&#4939;&#4725; &#4845;&#4936;&#4621;&#4875;&#4617;?</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
	        </font>
		    </font>
			<font face="Arial" color="#000000">
		    <b><br>
            Are you sure you want to delete?</b></font><font face="Arial" size="3"><br>
     	    <br>
	    	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="448">
		    <tr>
		    <td width="60" height="16" style="border-style: solid; border-width: 0; " align="center">
	    	</td>
		    <td width="60" height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center">
    		<a href="<%= SCRIPT_NAME %>?action=delete&Office_ID=<%= Selected_Office_ID %>" style="text-decoration: none; font-weight: 700">
	    	<b><font face="Verdana" size="2" color="#FFFFFF">
	        Yes</font></b></a></td>
		    <td width="60">
	    	</td>
	        <td width="120" bgcolor="#006699" style="border-style: solid; border-width: 2" align="center">
    	    <a href="<%= SCRIPT_NAME %>" style="text-decoration: none; font-weight: 700">
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
        	<a href="<%= SCRIPT_NAME %>" style="text-decoration: none; font-weight: 700">
			<font face="Arial" size="3">
    	    <b>
        	<font color="#006699" size="2">(Takes back to the Office List)</font></b></font><font color="#006699" size="2" face="Arial">
	        </font>
    	    </a>
        	</td>
	        </tr>
    	    </table>
        	<br><br>	
        <%
        End If
        
        	rstDBEdit.Close
	    	Set rstDBEdit = Nothing 
        %>
	
	    <%
	    
	    
	Case "delete"
	
		' Get the id to delete
		Selected_Office_ID = Request.QueryString("Office_ID")

		strSQL = "DELETE FROM Offices WHERE Office_ID=" & Selected_Office_ID & ";"

		Set cnnDBEdit = Server.CreateObject("ADODB.Connection")
		cnnDBEdit.Open CONN_STRING

		cnnDBEdit.Execute strSQL, adAffectAll, adExecuteNoRecords
		
		cnnDBEdit.Close
		Set cnnDBEdit = Nothing
		
		%> 
		<b><span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
        &#4840;&#4720;&#4656;&#4648;&#4824;&#4813; &#4706;&#4654; &#4661;&#4637;</span><br>
		<font face="Verdana" color="#006666">
		Deleted Office</font></b><font face="Arial" size="3">
  		<br><br> 		
		<%	
		Response.Write("Office Id #" & Selected_Office_ID & " deleted!")
		Response.Write(BACK_TO_LIST_TEXT)
	
		
	Case "edit"
		' First of a 2 part process... build a form with the
		' values from the database
		Selected_Office_ID = Request.QueryString("Office_ID")

		strSQL = "SELECT * FROM Offices WHERE Office_ID=" & Selected_Office_ID & ";"

		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

		%> 
		</font><b>
        <span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
        &#4840;&#4706;&#4654;&#4813;&#4757; &#4661;&#4637; &#4768;&#4661;&#4720;&#4779;&#4781;&#4621;</span></b><font face="Arial" size="3"><b><font color="#006666" face="Verdana" size="4"><br>
        </font><font color="#006666" face="Verdana">Edit Office</font></b>
  		<br><br> 
		<%

		If Not rstDBEdit.EOF Then
			%> </b>
			<form action="<%= SCRIPT_NAME %>?action=editsave&Office_ID=<%=Selected_Office_ID%>" method="post">
				<font face="Arial" size="2">
				<input type="hidden" name="Office_ID" value="<%= Selected_Office_ID %>" />
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                   <tr>
                      <td width="17%">
						<font face="Arial" size="3"><b>&#4840;&#4706;&#4654;&#4813; &#4645;&#4637;</b></font><b><br>
  						<font face="Arial" size="2">
                    	Office Name:</font>
                		</b>
	                  </td>
                      <td width="83%">
						<font face="Arial" size="2">
                    	<input type="text" name="Office_Name" value="<%= rstDBEdit.Fields("Office_Name").Value %>" size="30" /></font>
                      </td>
                   </tr>
                   <tr>
	                  <td width="17%">
		  			     <b>
  		 				 <font face="Arial" size="3">
	                     &#4706;&#4654;&#4813;</font><font face="Arial" size="2"><br>
                         Office Type:
                         </b>
                      </td>
      	              <td width="83%">
	                    <font face="Arial" size="2">
						<select size="1" name="Office_Type">
						<option value="Zone" selected>Zone</option>
						<%if rstDBEdit.Fields("Office_Type").Value = "Regional_Bureau" then %>
						<option value="Regional_Bureau" selected>Special_Woreda</option>
						<%else%>
						<option value="Regional_Bureau">Regional_Bureau</option>
						<%end if%>
						<%if rstDBEdit.Fields("Office_Type").Value = "Other" then %>
						<option value="Other" selected>Other</option>
						<%else%>
						<option value="Other">Other</option>
						<%end if%>
    	                </select>
        	            </font>
	                  </td>
                   </tr>
                </table>
                <p>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Update Office" value="Update Office" style="color: #FFFFFF; background-color: #525252; font-weight:bold; font-family:Verdana; font-size:10pt"><font size="2">
                </font>
			    </p>
			</form>
		<%
		Else
		%> 
			</font></font>
	  		<span style="font-size: 12.0pt; font-family: Arial">
            <font color="#000000">&#4840;&#4720;&#4896;&#4840;&#4672;&#4813; &#4704;&#4829;&#4653;&#4829;&#4649; &#4813;&#4661;&#4901; 
			&#4768;&#4621;&#4720;&#4872;&#4760;&#4637;</font></span></font><font face="Arial" color="#000000"><br>
			Record not found!</font><font face="Arial" size="2">
		<%
		End If

		rstDBEdit.Close
		Set rstDBEdit = Nothing

		Response.Write(BACK_TO_LIST_TEXT)
		

	Case "editsave"
	
		' Part 2 of 2: Here is where the values that the
		' user entered are saved to the database  
		
		Selected_Office_ID = Request.Form("Office_ID")

		' Check for errors				
		Error_check
						
		If Validated_Form Then
			' Select the record
			strSQL = "SELECT * FROM Offices WHERE Office_ID=" & Selected_Office_ID & ";"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText
			rstDBEdit.Fields("Office_Name").Value   = OfficeEntered
			rstDBEdit.Fields("Office_Type").Value  = Request.Form("Office_Type")
			rstDBEdit.Update

			rstDBEdit.Close
			Set rstDBEdit = Nothing

			%> 
   	     	<b>
            <span style="font-size: 13.5pt; font-family: Verdana; color: #006666">
            &#4840;&#4720;&#4661;&#4720;&#4779;&#4776;&#4616; &#4840;&#4706;&#4654;&#4813; &#4661;&#4637;</span><font color="#006666" face="Verdana" size="4"><br>
            </font></b></font><b>
  			<font face="Verdana" color="#006666">
            Edited Office</font></b><font face="Arial" size="3"><br><br> 	
		<%	
			Response.Write("<p>Office Id #" & Selected_Office_ID & " updated!</p>")
			Response.Write(BACK_TO_LIST_TEXT)
	 Else
	 %>  
	 	   <a href=" <%= SCRIPT_NAME %>?action=edit&Office_ID=<%=Selected_Office_ID%>" style="text-decoration: none; font-weight: 700">
		   <font face="Arial" size="2" color="black"><b><p>
	 	   &#4808;&#4848; &#4747;&#4619; &#4616;&#4632;&#4632;&#4616;&#4661; &#4773;&#4826;&#4613; &#4875; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;&#4961;&#4961;
		   <br>Click here to go back </font></p></b></a>
	 <%
	 End If


	Case Else ' view
	
		' The default action... just list the records in the database except for the BoF_Store
		
		strSQL = "SELECT * FROM Offices ORDER BY Office_Type DESC, Office_Name;"
		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING, adOpenForwardOnly, adLockReadOnly, adCmdText
		
		%> 

          <b>
          <font size="4" face="Verdana" color="#006666">&#4840;&#4706;&#4654;&#4814;&#4733; &#4829;&#4653;&#4829;&#4653;</font><font color="#006666"><font face="Verdana" size="4"><br>
	      </font>

          <font face="Verdana">
	      Office List</font></font><font face="Verdana">	
    	  </font>
		  <br><br> 
		
		<table border="1" cellspacing="0" cellpadding="5">
		<thead>
		<tr>
		<!-- <th height="16"><font face="Arial" size="2">Office ID</font></th> -->
			<th height="16"><font face="Arial" size="2" color="#000000">&#4840;&#4706;&#4654;&#4813;&#4645;&#4637;<br>
            Office Name</font></th>
			<th height="16"><font face="Arial" size="2">&#4706;&#4654;&#4813;<br>
            Office Type</font></th>
<% If Request.QueryString("ToExcel") = "" Then %>			
			<th height="16" width="55" class="NoPrint"><font face="Arial" size="2">&#4768;&#4901;&#4939;<br>
            Delete</font></th>
			<th height="16" width="55" class="NoPrint"><font face="Arial" size="2">&#4768;&#4653;&#4637;<br>
            Edit</font></th>
<%End If%> 
		</tr>
		</thead>
		<tbody>
		    <%
		    Do While Not rstDBEdit.EOF
			%>
			<tr height="16">			
			<!--	<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Office_ID").Value %>&nbsp;</font></td> -->
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Office_Name").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Office_Type").Value %>&nbsp;</font></td>
<% If Request.QueryString("ToExcel") = "" Then %>
				<td height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <font face="Verdana" size="2">
                <a href="<%= SCRIPT_NAME %>?action=deleteconfirm&Office_ID=<%= rstDBEdit.Fields("Office_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font color="#FFFFFF">Delete</font></a></font></td>
				<td height="16" bgcolor="#006666" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint"><font face="Arial" size="2">
                <a href="<%= SCRIPT_NAME %>?action=edit&Office_ID=<%= rstDBEdit.Fields("Office_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font face="Verdana" color="#FFFFFF" size="2">Edit</font></a></td>
<%End If%> 
			</tr>
			<%
			rstDBEdit.MoveNext
		    Loop
		    %>
		</tbody>
		
<% If Request.QueryString("ToExcel") = "" Then %>		
		<tfoot>
		<tr class="NoPrint">
			<td colspan="1"></td>
			<td align="center" height="16" bgcolor="#808000" style="border-style: solid; border-width: 2">
            <font face="Verdana" size="2">
            <a href="<%= SCRIPT_NAME %>?action=add" style="text-decoration: none; font-weight: 700">
            <font color="#FFFFFF">Add a new Office</font></a></font></td>
			<td colspan="3">&nbsp;</td>
		</tr>
		</tfoot>
<%End If%> 

		</table>
		<font face="Arial" size="2"><font color="#000000">(*Office_ID is automatically generated by the system and cannot be edited by the user)</font> </font>
		<br>
		<br>
		
<% If Request.QueryString("ToExcel") = "" Then %>		
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
            <form id="ExcelForm" action="<%= SCRIPT_NAME %>?ToExcel=True" method="post">
			<img border="0" src="images/Excel%20Button.JPG" align="top">
			<input id="btnExcel" type="submit" value="Send list to Excel" style="color: #FFFFFF; background-color: #525252; font-size:10pt; font-weight:bold; font-family:Verdana"/></font>
            </form>
            </td>

          </tr>
        </table>
<% End If %>		
		
		

		<%
		rstDBEdit.Close
		Set rstDBEdit = Nothing
		End Select
        %>

<% If Request.QueryString("ToExcel") = "" Then %>	
      </font>
                </b>
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