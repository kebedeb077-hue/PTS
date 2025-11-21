<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Offices</title>
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="tb, default">
</head>

<%
Response.Expires = -1000 ' Makes the browser not cache this page
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

  <br>
  
  <table border="0" cellpadding="15" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="100%" bordercolor="#426262" style="border-style:solid; border-width:2; ">
  <b>
          <font size="4" face="Verdana" color="#006666">&#4840;&#4706;&#4654;&#4814;&#4733; &#4829;&#4653;&#4829;&#4653;<br>
  </font>
  <font color="#006666" face="Verdana">Office List</font><font face="Arial"><br>
  </font>
  <BR></b><span style="font-size: 12.0pt; font-family: Arial; font-weight: 700">
  &#4840;&#4768;&#4757;&#4853;&#4757; &#4706;&#4654; &#4723;&#4650;&#4781; &#4616;&#4635;&#4840;&#4725; &#4704;&#4645;&#4653;&#4909;&#4725; &#4808;&#4845;&#4637; &#4704;&#4720;&#4672;&#4635;&#4909; &#4619;&#4845; &#4781;&#4618;&#4781; &#4843;&#4853;&#4653;&#4873;</span><b><font face="Arial"><br>
  <font size="2">Click on  Distributions or Deposits to view the history for 
  a certain office.</font></font></b><br><BR>
 
<%
 
 ' Get the list of all the offices ordered by name, except the BoF_Store
 Set MyConn = Server.CreateObject("ADODB.Connection")
 MyConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Forms.mdb") & "; User Id=admin; Password=serbof;"
 
 SQL_query = "SELECT * FROM Offices ORDER BY Office_Type DESC, Office_Name"
 Set RS = MyConn.Execute(SQL_query)

%>
     <table border="1" cellspacing="0" cellpadding="3" width="600">
     <tr>

      <td width="70%" align="center"><font face="Arial" size="2"><b>
      <font face="Arial" size="2">&#4840;&#4706;&#4653;&#4813; &#4661;&#4637;</font></b><br>
      <b>Office Name<br>
      </b>(&#4704; &#4677;&#4848;&#4637; &#4720;&#4776;&#4720;&#4621; - Alphabetically ordered)</font></td>
      <td width="10%" align="center"><b><font size="2" face="Arial">&#4706;&#4654;&#4813;<br>
      Office Type</font></b></td>
      <td width="10%" align="center"><b><font size="2" face="Arial">&#4645;&#4653;&#4909;&#4725;<br>
      Distributions</font></b></td>   
      </tr>
     
<%
 ' List all the offices
 WHILE NOT RS.EOF
%>
	 <tr>
     <td width="70%"><font face="Arial" size="2"><%=RS("Office_Name")%></font>&nbsp;</td>
     <td width="10%"><font face="Arial" size="2"><%=RS("Office_Type")%></font>&nbsp;</td>
     <td width="10%" align="center"><font face="Arial" size="2">
	<A  HREF="office_shipment_history.asp?Office_ID=<%=RS("Office_ID")%>">
		&#4645;&#4653;&#4909;&#4725;<br>Distributions</A></font></td>    			
	 </tr>
     <%
     
     RS.MoveNext
     
 WEND

 RS.Close
 MyConn.Close
 
 
%>
    
	 </table>  
  
<br>
</td>
</tr>
</table>    
<br>

<!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td bgcolor="#426262">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
 <tr>
    <td width="100%">
   <p align="center"><font color="#FFFFFF" face="Arial" size="1" ><b><i>Developed by the ICT Directorate of BoF</i></b></font></td>
  </tr>
</table>

</td></tr><!--msnavigation--></table></body>
</html>