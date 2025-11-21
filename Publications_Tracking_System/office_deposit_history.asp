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
<title>Office Deposit History</title>
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
Dim cnnDBEdit, rstDBEdit, rstDBBalance  ' ADO objects

Dim strSQL      			' To hold various SQL Strings
Dim strSQL_Balance
Dim Selected_Deposit_ID   ' Used to keep track of the record in play
Dim Added_Deposit_ID
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
		
		If NOT IsNumeric(Request.Form("Amount_Deposited")) Then
		%>
			<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
			</font><font color="#006666" face="Verdana">Error</font></b>
            <br><br>
			<b>
			<font face="Arial" color="#000000">
			&#4840;&#4720;&#4672;&#4635;&#4909; &#4709;&#4653; &#4709;&#4827;&#4725;</font></b></a><span style="font-size: 12.0pt; font-family: Arial"> 
			&#4704;&#4673;&#4901;&#4653; &#4632;&#4614;&#4757; &#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
			The <b>Amount Deposited</b> has to be a number! Please go back and correct the error.</font>
		<%
			Validated_Form = false		
	
		Else If Request.Form("Amount_Deposited")<1 Then
			 %>
				<b><font color="#006666" face="Verdana" size="4">&#4661;&#4629;&#4720;&#4725;<br>
				</font><font color="#006666" face="Verdana">Error</font></b>
            	<br><br>
				<b>
				<font face="Arial" color="#000000">
				&#4840;&#4720;&#4672;&#4635;&#4909; &#4709;&#4653; &#4709;&#4827;&#4725;</font></b></a><span style="font-size: 12.0pt; font-family: Arial"> 
				&#4776;&#4828;&#4654; &#4632;&#4709;&#4616;&#4901; &#4768;&#4616;&#4704;&#4725;&#4961;&#4961; &#4773;&#4707;&#4781;&#4814;&#4757; &#4808;&#4848; &#4747;&#4619; &#4612;&#4848;&#4813; &#4843;&#4661;&#4720;&#4779;&#4781;&#4617;&#4961;&#4961;</span><font face="Arial" size="3"><br>
				The <b>Amount Deposited</b> has to be greater than 0! Please go back and correct the error.</font>				
			 <%
				Validated_Form = false
				
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
			strSQL = "SELECT * FROM Deposits;"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

			' Add our record and set it's values.  
			rstDBEdit.AddNew
		
			rstDBEdit.Fields("Office_ID").Value     		= Selected_Office_ID
			rstDBEdit.Fields("Day").Value      				= Request.Form("Day")
			rstDBEdit.Fields("Month").Value      			= Request.Form("Month")
			rstDBEdit.Fields("Year").Value      			= Request.Form("Year")
			rstDBEdit.Fields("Amount_Deposited").Value  	= Request.Form("Amount_Deposited")
			rstDBEdit.Fields("Description").Value  			= Request.Form("Description")

			rstDBEdit.Update

			' Get the id of the record just added. This might cause
			' problems with some DB providers, but it works with
			' Microsoft Access and SQL Server.
			Added_Deposit_ID = rstDBEdit.Fields("Deposit_ID").Value

			rstDBEdit.Close
			Set rstDBEdit = Nothing

		%> 	<b>
			<span style="font-size: 13.5pt; font-family: Verdana; color: #006666">&#4840;&#4720;&#4632;&#4824;&#4872;&#4704; &#4768;&#4850;&#4661; 
			&#4720;&#4672;&#4635;&#4909;</span><font color="#006666" size="4" face="Verdana"><br>
			</font><font color="#006666" face="Verdana">Registered new deposit to</font>
  	 		<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></b><font face="Arial" size="3">
  			<br><br> 
		<%
			Response.Write("<p>Deposit Id #" & Added_Deposit_ID & " added!</p>")
		End If
			Response.Write(BACK_TO_LIST_TEXT)
	
	
	Case "deleteconfirm"
	
		'Before a record is deleted, confirmation is required
		
		Selected_Deposit_ID = Request.QueryString("Deposit_ID")
	
		%>
	
    	<b><font color="#006666" size="4" face="Verdana">&#4720;&#4672;&#4635;&#4905;&#4757; &#4640;&#4653;&#4829;<br>
		</font><font color="#006666" face="Verdana">Delete Deposit</font></b>
	    <br>
		<br>
		<font face="Arial">
		<b>&#4704;&#4773;&#4653;&#4877;&#4901; &#4616;&#4635;&#4901;&#4939;&#4725; &#4845;&#4936;&#4621;&#4875;&#4617;?</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<br>
		</font>
		</font>
		<font face="Arial">
	    <b>Are you sure you want to delete?</b></font><font face="Arial" size="3"><br>
        <br>
                
	    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="448">
	    <tr>
	    <td width="60" height="16" style="border-style: solid; border-width: 0; " align="center">
	    </td>
    	<td width="60" height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center">
	    <a href="<%= SCRIPT_NAME %>?action=delete&Office_ID=<%=Selected_Office_ID%>&Deposit_ID=<%= Selected_Deposit_ID %>" style="text-decoration: none; font-weight: 700">
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
        <font color="#006699" size="2">(Takes back to the Deposits List)</font></b></font><font color="#006699" size="2" face="Arial">
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
		Selected_Deposit_ID = Request.QueryString("Deposit_ID")

		strSQL = "DELETE FROM Deposits WHERE Deposit_ID=" & Selected_Deposit_ID & ";"

		Set cnnDBEdit = Server.CreateObject("ADODB.Connection")
		cnnDBEdit.Open CONN_STRING

		cnnDBEdit.Execute strSQL, adAffectAll, adExecuteNoRecords
		
		cnnDBEdit.Close
		Set cnnDBEdit = Nothing

	%>
		</font><b><font color="#006666" face="Verdana" size="4">&#4776; <font color="#C46200" face="Verdana" size="4"><%=Selected_Office_Name%></font> &#4840;&#4720;&#4640;&#4648;&#4824; &#4720;&#4672;&#4635;&#4909;<br>
	 	</font><font color="#006666" face="Verdana">Deleted Deposit from</font>
	 	<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></b><font face="Arial" size="3">
  		<br><br> 		
	<%	
		
		Response.Write("Deposit Id #" & Selected_Deposit_ID & " deleted!")
		Response.Write(BACK_TO_LIST_TEXT)
		

		
	Case "edit"
		
		' First of a 2 part process... build a form with the
		' values from the database
		Selected_Deposit_ID = Request.QueryString("Deposit_ID")

		strSQL = "SELECT * FROM Deposits WHERE Deposit_ID=" & Selected_Deposit_ID & ";"

		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

	%> 

  		<b><font color="#006666" face="Verdana" size="4">&nbsp; &#4840; <font color="#C46200" face="Verdana" size="4"><%=Selected_Office_Name%></font> &#4720;&#4672;&#4635;&#4909;&#4757; &#4768;&#4661;&#4720;&#4776;&#4781;&#4621;<br>
	    </font><font color="#006666" face="Verdana">Edit Deposit for</font>
	 	<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Arial">:</font></b>
  		<br><br> 
	<%

		If Not rstDBEdit.EOF Then
			%> </b>

			<form action="<%= SCRIPT_NAME %>?action=editsave&Office_ID=<%=Selected_Office_ID%>" method="post">
				<font face="Arial" size="2">
				<input type="hidden" name="Deposit_ID" value="<%= Selected_Deposit_ID %>" />
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="106%">
                <tr>
                    <td width="18%">
						<b>
						<font face="Arial" size="2">
					    &#4672;&#4757;<br>
					    Date:&nbsp;&nbsp;
                    	</font>
                    	</b>
                    </td>
                    
                    <td width="16%">
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
        	        	For i=2016 to 2030
            	    	%>
                		<option <%=z(i-1990)%>><%=i%></option>
                		<%Next%> 
			            </select> 
    	                </font>
    	            </td>
                    <td width="84%">&nbsp;
                    </td>
               </tr>
               <tr>
                  <td width="18%">
					  <font face="Arial" size="2">
					  <b>&#4840;&#4720;&#4672;&#4635;&#4909; &#4709;&#4653; &#4709;&#4827;&#4725;<br>
    	              Amount Deposited (Birr): </b>
				  </td>
                  <td width="16%">
					  <font face="Arial" size="2">
    	              <input type="text" name="Amount_Deposited" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Amount_Deposited").Value) %>" size="30" /></font>
    	          </td>
                  <td width="84%">&nbsp;
                  </td>
               </tr>
               <tr>
                  <td width="18%">
					  <font face="Arial" size="2">
					  <b>&#4632;&#4877;&#4616;&#4907;<br>
    	              Description: </b>
				  </td>
                  <td width="16%">
					  <font face="Arial" size="2">
            	      <input type="text" name="Description" value="<%= Server.HTMLEncode(rstDBEdit.Fields("Description").Value) %>" size="30" /></font>
            	  </td>
                  <td width="84%">&nbsp;
                  </td>
                </tr>                  
                </table>
                <p>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Update Deposit" value="Update Deposit" style="color: #FFFFFF; font-weight: bold; font-family: Verdana; background-color: #525252"><font size="2">
                </font>
			    </p>
			</form>
			<%
		Else
			%>
			<span style="font-size: 12.0pt; font-family: Arial">&#4840;&#4720;&#4896;&#4840;&#4672;&#4813; &#4704;&#4829;&#4653;&#4829;&#4649; &#4813;&#4661;&#4901; 
			&#4768;&#4621;&#4720;&#4872;&#4760;&#4637;</span><font face="Arial"><br>
			Record not found!</font>
			<%
		End If

		rstDBEdit.Close
		Set rstDBEdit = Nothing

		Response.Write(BACK_TO_LIST_TEXT)
		

	Case "editsave"

		' Check for errors
		Error_Check
		Selected_Deposit_ID = Request.Form("Deposit_ID")
		
		' If data is fine, enter it in the database
			If Validated_Form Then

			' Part 2 of 2: Here's where we save the values that the
			' user entered back to the DB.  
			
		
			' Select the record
			strSQL = "SELECT * FROM Deposits WHERE Deposit_ID=" & Selected_Deposit_ID & ";"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

			rstDBEdit.Fields("Day").Value      				= Request.Form("Day")
			rstDBEdit.Fields("Month").Value      			= Request.Form("Month")
			rstDBEdit.Fields("Year").Value      			= Request.Form("Year")
			rstDBEdit.Fields("Amount_Deposited").Value     	= Request.Form("Amount_Deposited")
			rstDBEdit.Fields("Description").Value      		= CStr(Replace(Request.Form("Description"), "'", "''"))	
	
			rstDBEdit.Update

			rstDBEdit.Close
			Set rstDBEdit = Nothing

			%> 
   	    	 <b><font color="#006666" face="Verdana" size="4">&#4616; <font color="#C46200" face="Verdana" size="4"><%=Selected_Office_Name%></font> &#4840;&#4720;&#4661;&#4720;&#4779;&#4776;&#4616; &#4720;&#4672;&#4635;&#4909;<br>
			 </font><font color="#006666" face="Verdana">Edited Deposit for</font>
	  		 <font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font><font color="#006666" face="Verdana">:</font></b><font face="Arial" size="3">
  			 <br><br> 	
			<%	
			Response.Write("<p>Deposit Id #" & Selected_Deposit_ID & " updated!</p>")
			Response.Write(BACK_TO_LIST_TEXT)
	 Else
	 
	 %>
		  <a href=" <%= SCRIPT_NAME %>?action=edit&Office_ID=<%=Selected_Office_ID%>&Deposit_ID=<%= Selected_Deposit_ID %>" style="text-decoration: none; font-weight: 700">
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
		
		' The default action... just lists the records in the DB for the selected form
		strSQL = "SELECT * FROM Deposits, Offices WHERE Deposits.Office_ID = Offices.Office_ID AND Deposits.Office_ID = "& Selected_Office_ID &" ORDER BY (((Year-1)*360)+((Month-1)*30)+Day), Deposit_ID;"
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
		
		%>  </font></font></font>
			<b>
			<font face="Verdana" size="3" color="#006666">
			<font size="4">
			&#4840;&#4720;&#4672;&#4635;&#4909; &#4723;&#4650;&#4781;<br>
			Deposit History</font><br>
			&#4840;&#4706;&#4653;&#4813;&#4661;&#4637; - Office Name:</font>
			<font color="#C46200" face="Verdana"><%=Selected_Office_Name%></font></font>
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
			<th height="16" class="NoPrint"><font face="Arial" size="2">Deposit ID</font></th>
			<th height="16">
			<font face="Arial" size="2">&#4672;&#4757;<br>
            Date</font></th>
			<th height="16">
			<font face="Arial" size="2">&#4840;&#4720;&#4672;&#4635;&#4909; &#4709;&#4653; &#4709;&#4827;&#4725;<br>
            Amount Deposited (Birr) </font></th>
			<th height="16">
    	    <font face="Arial" size="2">&#4632;&#4877;&#4616;&#4907;<br>
            Description</font></th>
<% If Request.QueryString("ToExcel") = "" Then %>
			<th height="16" width="55" class="NoPrint">
            <font face="Arial" size="2">&#4768;&#4901;&#4939;<br>
            Delete</font></th>
			<th height="16" width="55" class="NoPrint">
            <font face="Arial" size="2">&#4768;&#4653;&#4637;<br>
            Edit</font></th>
<% End If %>
		</tr>
		</thead>
		<tbody>
		<% If iPageCount = 0 Then %>
		<tr height="16"><td colspan="6">
          <span style="font-family: Arial"><font size="2">&#4637;&#4757;&#4637; &#4840;&#4720;&#4632;&#4824;&#4872;&#4704; &#4752;&#4872;&#4653; 
          &#4768;&#4621;&#4720;&#4872;&#4760;&#4637;&#4961;&#4961; </font></span>
          <font size="2" face="Arial">No records found.</font></td></tr>
		
		<% Else ' List the records %>
		    <%
		    iRecordsShown = 0

		    Do While iRecordsShown < iPageSize And Not rstDBEdit.EOF
				
			%>
			<tr height="16">			
				<td height="16" nowrap align="center" class="NoPrint"><font face="Arial" size="2"><%= rstDBEdit.Fields("Deposit_ID").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Day").Value %>/<%= Months_Array(rstDBEdit.Fields("Month").Value) %>/<%= rstDBEdit.Fields("Year").Value %></font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Amount_Deposited").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="right" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Description").Value %>&nbsp;</font></td>
<% If Request.QueryString("ToExcel") = "" Then %>
				<td height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <font face="Verdana" size="2">
                <a href="<%= SCRIPT_NAME %>?action=deleteconfirm&Office_ID=<%=Selected_Office_ID%>&Deposit_ID=<%= rstDBEdit.Fields("Deposit_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font color="#FFFFFF">Delete</font></a></font></td>
				<td height="16" bgcolor="#006666" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <a href="<%= SCRIPT_NAME %>?action=edit&Office_ID=<%=Selected_Office_ID%>&Deposit_ID=<%= rstDBEdit.Fields("Deposit_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font face="Verdana" size="2" color="#FFFFFF" size="2">Edit</font></a></td>
<% End If %>
			</tr>
			<%
				
				iRecordsShown = iRecordsShown + 1
				rstDBEdit.MoveNext
		    Loop
		    %>
		    
		    <tr>
<% If Request.QueryString("ToExcel") = "" Then %>		    
		    <td colspan="6">
<% Else %> &nbsp;<td colspan="4">
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
		    
	    <% End If ' End for List the records%>
		    <tr>
		    <td colspan="2" align="right">
		    <b><font face="Arial"><span style="font-size: 12.0pt">&#4634;&#4827;&#4757; -</span></font><span style="font-size: 12.0pt; font-family: Arial">
            </span></b>
		    <font face="Arial" size="3"><b>Balance**:</b></font>
		    </td>
		    <td colspan="1" align="center">
		    <%
		    Set rstDBBalance = Server.CreateObject("ADODB.Recordset")
		    Dim Total_Amount_Deposited
		    Dim Total_Debt
		    
		    strSQL_Balance = "SELECT SUM(Amount_Deposited) AS Total_Amount_Deposited FROM Deposits WHERE Office_ID = "& Selected_Office_ID &";"
			rstDBBalance.Open strSQL_Balance, CONN_STRING, adOpenForwardOnly, adLockReadOnly, adCmdText
			
			If IsNumeric(rstDBBalance.Fields("Total_Amount_Deposited").Value)=false Then
				Total_Amount_Deposited = 0
			Else 
				Total_Amount_Deposited = rstDBBalance.Fields("Total_Amount_Deposited").Value
			End If
			rstDBBalance.Close

		    strSQL_Balance = "SELECT SUM(Total_Price) AS Total_Debt FROM Shipments WHERE Office_ID = "& Selected_Office_ID &";"
			rstDBBalance.Open strSQL_Balance, CONN_STRING, adOpenForwardOnly, adLockReadOnly, adCmdText
			If IsNumeric(rstDBBalance.Fields("Total_Debt").Value)=false Then
				Total_Debt = 0
			Else 
				Total_Debt = rstDBBalance.Fields("Total_Debt").Value
			End If
			rstDBBalance.Close

			%>
			<font face="Arial" size="3">
			<b><%= Total_Amount_Deposited - Total_Debt %></b>
			</font>
			<%
			
			Set rstDBBalance = Nothing
			%> &nbsp;</td>
			
<% If Request.QueryString("ToExcel") = "" Then %>		    
		    <td colspan="3">
<% Else %> &nbsp;<td colspan="1">
<% End If %>			
			&nbsp;
		    </td>
		    </tr>
		
		    <tr>
<% If Request.QueryString("ToExcel") = "" Then %>		    
		    <td colspan="6" align="right">
<% Else %> &nbsp;<td colspan="4" align="right">
<% End If %>		    
		    <p align="left"><font face="Arial" size="2">(**</font><span style="font-family: Arial"><font size="2">&#4634;&#4824;&#4757;&#4963; 
            &#4776;&#4896;&#4677;&#4619;&#4619; &#4720;&#4672;&#4635;&#4909; &#4840;&#4720;&#4656;&#4651;&#4905;&#4725; &#4942;&#4653;&#4638;&#4733; &#4811;&#4875; &#4658;&#4672;&#4752;&#4661; &#4752;&#4813;&#4961;&#4961; <br> </font></span><font face="Arial" size="2">
            Balance is Total Amount Deposited minus Total Cost of forms 
            shipped.)</font></td>
		    </tr>
		</tbody>
	</table>
	<br>
	<font face="Arial" size="2" class="NoPrint"><b>
	<font color="#000000">(*Deposit_ID is automatically generated by the system and cannot be edited by the user)</font>
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
			<input id="btnExcel" type="submit" value="Send list to Excel" style="color: #FFFFFF; background-color: #525252; font-weight:bold; font-family:Verdana"/></font></font>
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
            
    <%	' Adding a new Deposit record
		' First of a 2 part process... build an empty form for the user to enter data. 
		' The fields come from the Deposits table
	%>          
      
		<font size="2">
			
			<font face="Verdana" size="3">
			
			<form action="<%= SCRIPT_NAME %>?action=addsave&Office_ID=<%=Selected_Office_ID%>" method="post" class="NoPrint">
				<table border="1" cellspacing="0" cellpadding="0" width="379">
                <tr>
                    <td width="354" colspan="2" height="35">
  						<b><span style="font-size: 13.5pt; color: #006666">&#4768;&#4850;&#4661; &#4720;&#4672;&#4635;&#4909; 
        	            &#4632;&#4829;&#4877;&#4709;</span></b><font color="#006666"><b><br>
            	        Register</b></font></font></font><font face="Arial" size="3"><font face="Arial" size="2"><font face="Arial" size="3"><b><font color="#006666" face="Verdana" size="3"> 
                	    a new deposit:</font></b>          
                    </td>
                </tr>
                <tr>
    	            <td width="155">
						<b><font face="Arial" size="2">
					    &#4672;&#4757;<br>
					    Date:&nbsp;&nbsp;
            	        </font>
                	    </b>
	                </td>
                    <td width="218">
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
        		        <%For i=2016 to 2030%>
		                <option><%=i%></option>
        		        <%Next%> 
                		</select> </font>
                    </td>
               </tr>
               <tr>
                    <td width="155">
	   				    <b>
						<font face="Arial" size="2">
						&#4840;&#4720;&#4672;&#4635;&#4909; &#4709;&#4653; &#4709;&#4827;&#4725;<br>
						Amount Deposited (Birr):
	                	</font>
		                </b>
				    </td>
                    <td width="218">
						<font face="Arial" size="2">
            	        <input type="text" name="Amount_Deposited" value="0" size="30" /> </font>
                    </td>
               </tr>
               <tr>
                    <td width="155">
					    <b><font face="Arial" size="3">
	 				    <font face="Arial" size="2">
						&#4632;&#4877;&#4616;&#4907;</font></font><font face="Arial" size="2"><br>
						Description: </font> </b>
				    </td>
                    <td width="218">
						<font face="Arial" size="2">
        	            <input type="text" name="Description" value="" size="30" /> </font>
                    </td>
               </tr>
               <tr>
	                <td width="354" colspan="2" align="right">
						<font face="Arial" size="2">
						<input type="submit" name="Register Deposit" value="Register Deposit" style="color: #FFFFFF; font-weight: bold; font-family: Verdana; background-color: #525252">
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