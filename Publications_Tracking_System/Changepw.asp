<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Change_PW</title>
<meta name="Microsoft Border" content="tb, default">
</head>

<%
Response.Expires = -5000 ' Makes the browser not cache this page
Response.Buffer = True   ' Buffers the content so the Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("default.asp")
End If
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

<% If Request.QueryString("ToExcel") = "" Then %>
  <br>
  
  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#(#006000" width="100%">
    <tr>
      <td width="100%" style="border-style: solid; border-width: 2" bordercolor="#5865F2">
<%End If%>      
 
<!-- #include file="adovbs.inc" -->

<%

' The connection string
Dim CONN_STRING
CONN_STRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=Admin; Password=serbof;"

' The name of this file so all the links and forms will still
' work if it is renamed
Dim SCRIPT_NAME
SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

' This is for the link at the bottom of the page to avoid typing it each time
Dim BACK_TO_LIST_TEXT
BACK_TO_LIST_TEXT = "<font face=""Arial"" size=""3""><b><p><br>Click <a href=""" & SCRIPT_NAME & """>" _
    & "here</a> to go back to the password list. </font></p></b>"

' Declare the standard variables
Dim cnnDBEdit, rstDBEdit	' ADO objects

Dim strSQL      					' To hold various SQL Strings
Dim Selected_User_ID   		' Used to keep track of the record
Dim Selected_UserName
Dim Added_User_ID
Dim Validated_Form					' To hold the state of data entered
Dim UserNameEntered
				' User entered by the user in the update and new record forms
Dim current


%>

<%

' This function is called from addsave and editsave to verify the input entered by the user
FUNCTION Error_Check

		'Form validation
		Validated_Form = true
		UserNameEntered = CStr(Trim(Replace(Request.Form("UserName"), "'", "''")))
		If (len(UserNameEntered)<1) Then
		%>
			<b><font color="#006666" face="Verdana" size="4">Error</font></b>
            <br><br>
			<font face="Arial" size="3"> user name cannot be empty! Please go back and enter a name.</font>
		<%
			Validated_Form = false
		Else
						
			' Check for duplicates
			Select Case LCase(Trim(Request.QueryString("action")))
			
				Case "addsave"
					strSQL = "SELECT COUNT(*) AS Duplicate_Count FROM users WHERE UserName LIKE '" & Replace(UserNameEntered, "'", "''") & "';"

				Case "editsave"
					'in case of edit we need to make sure to exclude the current user, in case the user has not changed the name
					strSQL = "SELECT COUNT(*) AS Duplicate_Count FROM users WHERE UserName LIKE '" & Replace(UserNameEntered, "'", "''") & "' AND NOT (user_ID LIKE "& Selected_User_ID &");"

			End Select
			
			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText
		
			If rstDBEdit.Fields("Duplicate_Count").Value > 0 Then
			%>
				<b><font color="#006666" face="Verdana" size="4">Error</font></b>
            	<br><br>
				<font face="Arial" size="3">The user <b><%=UsertEntered%></b> already exists in the list! Please go back and enter a new name.</font>
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
		<b><font color="#006666" face="Verdana" size="4">Add New password:</font></b>

		<font face="Arial" size="3">
  		<br><br> 
			
		<form action="<%= SCRIPT_NAME %>?action=addsave" method="post">
		<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">

               


                <tr>
                    <td width="19%">
				<b>user name:&nbsp;&nbsp; </font></b></td>
                    <td width="81%">
				<font face="Arial" size="2">
                    <input type="text" name="UserName" value="" size="30" /> </font>
                    </td>
                  </tr>


                   <tr>
                    <td width="19%">
				<b>password:&nbsp;&nbsp; </font></b></td>
                    <td width="81%">
				<font face="Arial" size="2">
                    <input type="password" name="password" value="" size="30" /> </font>
                    </td>
                  </tr>


                  </table>
                <p>
				<font face="Arial" size="2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<br />
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" name="Add password" value="Add password"><font size="2">
				</font>
			    </p>
			</form>
			<p>
			</p>
		<%
		Response.Write(BACK_TO_LIST_TEXT)
		

	Case "addsave"


                   Error_Check

                  If Validated_Form Then
                    strSQL = " SELECT * FROM users ;"
   		    Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		    rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

			' Add our record and set its values.  
			rstDBEdit.AddNew
			
                       

                        
                        rstDBEdit.Fields("UserName").Value  = UserNameEntered
                        rstDBEdit.Fields("password").Value  = CStr(Trim(Replace(Request.Form("Password"), "'", "''")))
                        rstDBEdit.Update

			' Get the id of the record just added
			' Added_User_ID = rstDBEdit.Fields("user_ID").Value
                      
		    rstDBEdit.Close
		    Set rstDBEdit = Nothing

		          %> 
   	     	<b><font color="#006666" face="Verdana" size="4">Edited Password:</font></b><font face="Arial" size="3">
  		 	<br><br> 	
				
			
			<p>password information updated!</p>
			<%
			Response.Write(BACK_TO_LIST_TEXT)
		Else
		   %>
                 
            <font face="Arial" size="3"><b><p><br>Click 
            <a href="<%= SCRIPT_NAME %>?action=edit&User_ID=<%= Selected_User_ID %>" style="text-decoration: none; font-weight: 700">
            here</a> to go back to the form </font></p></b>
			<%
		End If
	   
     Case "deleteconfirm"
	
		Selected_User_ID = Request.QueryString("User_ID")
		Selected_UserName = Request.QueryString("UserName")
	
		' Check to see if there are any records in the users table that refer to this user
		' If there is a reference, then deletion will not be permitted
		strSQL = "SELECT COUNT(*) AS UserName_Referred FROM users WHERE User_ID = " & Selected_User_ID & ";"
		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING
		
		If rstDBEdit.Fields("UserName_Referred").Value > 0 Then
		%>
			<b><font color="#006666" face="Verdana" size="4">Error</font></b>
            <br><br>
			<font face="Arial" size="3">The User cannot be deleted because there is a reference to it in the password list.</font>
		<%
			Response.Write(BACK_TO_LIST_TEXT)
			
			
	    Else
	    ' Ask for confirmation of deletion
 	    %>
    	<b><font color="#006666" size="4" face="Verdana">Delete password</font></b>
    	<br>
		<br>
		<font face="Arial" size="2">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
		<font face="Arial">
	    <b>Are you sure you want to delete?</b></font><br>
        <br>
	    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="448">
	    <tr>
	    <td width="60" height="16" style="border-style: solid; border-width: 0; " align="center">
	    </td>
	    <td width="60" height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center">
    	<a href="<%= SCRIPT_NAME %>?action=delete&User_ID=<%= Selected_User_ID %>&UserName=<%= Selected_UserName %>" style="text-decoration: none; font-weight: 700">
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
        <a href="<%= SCRIPT_NAME %>?&User_ID=<%= Selected_User_ID %>" style="text-decoration: none; font-weight: 700">
		<font face="Arial" size="3">
        <b>
        <font color="#006699" size="2">(Takes back to the password List)</font></b></font><font color="#006699" size="2" face="Arial">
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
		Selected_User_ID = Request.QueryString("User_ID")
		Selected_UserName = Request.QueryString("UserName")

		strSQL = "DELETE FROM users WHERE user_ID=" & Selected_User_ID & ";"

		Set cnnDBEdit = Server.CreateObject("ADODB.Connection")
		cnnDBEdit.Open CONN_STRING

		cnnDBEdit.Execute strSQL, adAffectAll, adExecuteNoRecords
		
		cnnDBEdit.Close
		Set cnnDBEdit = Nothing
		
		%> 
		</font>
	    <b><font color="#006666" face="Verdana" size="4">Deleted user:</font></b><font face="Arial" size="3">
  		<br><br> 		
		
		<font color="#C46200"><%= Selected_UserName %></font> deleted!
		<%	
		Response.Write(BACK_TO_LIST_TEXT)
	
		
	Case "edit"
	
		' First of a 2 part process... build a form with the
		' values from the db.
		Selected_User_ID = Request.QueryString("User_ID")

		strSQL = " SELECT * FROM users WHERE user_ID =" & Selected_User_ID & " "

		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
              	rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText

		%> 

  		<b><font color="#006666" face="Verdana" size="4">Edit Password</font><font color="#006666" size="4" face="Arial">:</font></b><font face="Arial" size="3">
  		<br><br> 
		<%

		If Not rstDBEdit.EOF Then
			%> </b>
			<form action="<%= SCRIPT_NAME %>?action=editsave&User_ID=<%=Selected_User_ID%>" method="post">
				<font face="Arial" size="2">
				<input type="hidden" name="User_ID" value="<%= Selected_User_ID %>" />
				<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                  <tr>
  					<font face="Arial" size="3">
                    <td width="19%">
				<b>user Name:&nbsp;&nbsp; </b></td>
                    </font>
                    <td width="81%">
				<font face="Arial" size="2">
                    <input type="text" name="UserName" value="<%= rstDBEdit.Fields("UserName").Value %>" size="30" /></font></td>
                   </tr>
                    

                 <tr>
                    <td width="19%">
				<b>password:&nbsp;&nbsp; </font></b></td>
                    <td width="81%">
				<font face="Arial" size="2">
                    <input type="text" name="password" value="<%= rstDBEdit.Fields("password").Value %>" size="30" /> </font>
                    </td>
                  </tr>
                               



  </table>
                <p>
				<font face="Arial" size="2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
                </font>
				<font face="Arial" size="2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
				<input type="submit" name="Update Password" value="Update Password"><font size="2">
                </font>
			    </p>
			</form>
			<%
		Else
			Response.Write "Record not found!"
		End If

		rstDBEdit.Close
		Set rstDBEdit = Nothing

		Response.Write(BACK_TO_LIST_TEXT)
		

	Case "editsave"
	
		' Part 2 of 2: Here's where the values that the
		' user entered are saved to the database  
		Selected_User_ID = Request.Form("User_ID")
              
		' Check for errors		
		Error_check
						
		If Validated_Form Then
			' Select the record
			strSQL = "SELECT * FROM users WHERE user_ID=" & Selected_User_ID & ";"

			Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
			rstDBEdit.Open strSQL, CONN_STRING, adOpenKeyset, adLockOptimistic, adCmdText
			rstDBEdit.Fields("UserName").Value = CStr(Trim(Replace(Request.Form("UserName"), "'", "''")))
                        rstDBEdit.Fields("password").Value = CStr(Trim(Replace(Request.Form("password"), "'", "''")))
                       	rstDBEdit.Update

			rstDBEdit.Close
			Set rstDBEdit = Nothing

			%> 
   	     	<b><font color="#006666" face="Verdana" size="4">Edited Password:</font></b><font face="Arial" size="3">
  		 	<br><br> 	
				
			
			<p>password information updated!</p>
			<%
			Response.Write(BACK_TO_LIST_TEXT)
		Else
		   %>
                 
            <font face="Arial" size="3"><b><p><br>Click 
            <a href="<%= SCRIPT_NAME %>?action=edit&User_ID=<%= Selected_User_ID %>" style="text-decoration: none; font-weight: 700">
            here</a> to go back to the form </font></p></b>
			<%
		End If


	Case Else ' view
	
		' The default action... just list the records in the DB
		strSQL = "SELECT * FROM Users ORDER BY UserName;"
		Set rstDBEdit = Server.CreateObject("ADODB.Recordset")
		rstDBEdit.Open strSQL, CONN_STRING, adOpenForwardOnly, adLockReadOnly, adCmdText
		
		%> 

        <font face="Verdana" size="4">
		<b><font color="#006666">Users List</font>
		<br><br> 
		
		<table border="1" cellpadding="2">
		<thead>
		<tr>
			<th height="16"><font face="Arial" size="2">User Name</font></th>
			<th height="16"><font face="Arial" size="2">Password</font></th>
                       	<% If Request.QueryString("ToExcel") = "" Then %>			
			<th height="16" width="55"><font face="Arial" size="2" class="NoPrint">Delete</font></th>
			<th height="16" width="55"><font face="Arial" size="2" class="NoPrint">Edit</font></th>
<%End If%> 			
		</tr>
		</thead>
		<tbody>
		    <%
		    Do While Not rstDBEdit.EOF
			%>
			<tr height="16">			
				<td height="16" nowrap align="left" ><font face="Arial" size="2"><%= rstDBEdit.Fields("UserName").Value %>&nbsp;</font></td>
				<td height="16" nowrap align="center" ><font face="Arial" size="2"><%= rstDBEdit.Fields("Password").Value %>&nbsp;</font></td>
                                 <% If Request.QueryString("ToExcel") = "" Then %>
				<td height="16" bgcolor="#993333" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint">
                <font face="Verdana" size="2">
                <a href="<%= SCRIPT_NAME %>?action=deleteconfirm&User_ID=<%= rstDBEdit.Fields("user_ID").Value %>&UserName=<%= rstDBEdit.Fields("UserName").Value %>" style="text-decoration: none; font-weight: 700">
                <font color="#FFFFFF">Delete</font></a></font></td>
				<td height="16" bgcolor="#006666" style="border-style: solid; border-width: 2" align="center" width="55" class="NoPrint"><font face="Arial" size="2">
                <a href="<%= SCRIPT_NAME %>?action=edit&User_ID=<%= rstDBEdit.Fields("user_ID").Value %>" style="text-decoration: none; font-weight: 700">
                <font color="#FFFFFF" size="2">Edit</font></a></td>
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
			<td colspan="1" align="center" height="16" bgcolor="#006699" style="border-style: solid; border-width: 2">
            <font face="Verdana" size="2">
            <a href="<%= SCRIPT_NAME %>?action=add" style="text-decoration: none; font-weight: 700">
            <font color="#FFFFFF">Add a new password</font></a></font></td>
			<td colspan="2"></td>
		</tr>
		</tfoot>
<%End If%> 		
		
		</table>
		
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