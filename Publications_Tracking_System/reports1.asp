
<%If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("default.asp")
End If
%>

<%

' ****** THIS PART DEFINES THE ATTRIBUTES WHEN SENDING THE PAGE TO EXCEL ******

If Request.QueryString("ToExcel") = "True" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "filename=report.xls"
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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
	<style>
	br.page {page-break-before: always}
	</style>
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

   
<%End If%>

<% If Request.QueryString("ToExcel") = "" Then %>
 <br>
  
  <table border="0" cellpadding="15" cellspacing="0" width="100%" bgcolor="#FFFFFF" style="border-collapse: collapse" bordercolor="#426262">
    <tr>
      <td width="100%" style="border-style: solid; border-width: 2" bordercolor="#426262">
      
<% End If %>

<%

' ****** DEFINE GLOBAL VARIABLES ******

' The connection string... 
Dim CONN_STRING
CONN_STRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("Database\Payroll.mdb") & "; User Id=admin; Password=sepbof;"


' The name of this file so all the links and forms will still
' work if it is renamed
Dim SCRIPT_NAME
SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")


' This is for the link at the bottom of the page to avoid typing it each time
Dim BACK_TO_LIST_TEXT
BACK_TO_LIST_TEXT = "<a href=""" & SCRIPT_NAME & """ class=""NoPrint""><font face=""Arial"" size=""2"" color=""black""><b><p>"_
	& "Click here to go back </font></p></b>"      
    
' Declare the standard variables.
Dim rstDB 		' ADO objects
Dim strSQL		' To hold various SQL Strings

	' The default action... just list the menu of reports
%>
 <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#426262" width="100%" class="NoPrint">
	<b><font face="Verdana" size="7" color="#006666">Reports</font></b><p>
      
	<tr>
	<td width="350" bgcolor="#426262">
          <p align="center">
          <span style="font-weight: 700">
          <font face="Verdana" size="2" color="#FFFFFF">
          <a style="text-decoration: none" href="MRBoF_store.asp">
          <font color="#FFFFFF">Monthly Report for BoF Store</font></a></font></span></td>
	<td width="50" bgcolor="FFFFFF"><p>&nbsp;</p></td> 
	
	<!---
	<td width="350" bgcolor="#426262">
          <p align="center">
          <span style="font-weight: 700">
          <font face="Verdana" size="2" color="#FFFFFF">
          <a style="text-decoration: none" href="FRBoF_store.asp">
          <font color="#FFFFFF">Financial Report for BoF Store</font></a></font></span></td>
	<td width="50" bgcolor="FFFFFF"><p>&nbsp;</p></td> 
	--->
          <td width="350" bgcolor="#426262">
          <p align="center">
          <span style="font-weight: 700">
          <font face="Verdana" size="2" color="#FFFFFF">
          <a style="text-decoration: none" href="MR_Offices.asp">
          <font color="#FFFFFF">Monthly Report for Offices</font></a></font></span></td>
	<td width="50" bgcolor="FFFFFF"><p>&nbsp;</p></td> 	
	   
         <td width="350" bgcolor="#426262">
          <p align="center">
          <span style="font-weight: 700">
          <font face="Verdana" size="2" color="#FFFFFF">
          <a style="text-decoration: none" href="reports.asp">
          <font color="#FFFFFF">Reports</font></a></font></span></td>
	 <td width="50" bgcolor="FFFFFF"><p>&nbsp;</p></td> 
     	    	
	</tr>  
	 	
	</table>
		</table>
<% If Request.QueryString("ToExcel") = "" Then %>



<br>
<p> 
<p></p>
<br>
            </td></tr></table><!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation-->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%">
   <p align="center"><font color="#FFFFFF" face="Arial" size="1" ><b><i>Developed by the ICT Directorate of BoF</i></b></font></td>
  </tr>
</table>
    
</td></tr><!--msnavigation--></table></body></html><% End If %>