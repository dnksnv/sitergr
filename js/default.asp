<%@ LANGUAGE = JScript %>
<HTML>
    <HEAD>
        <TITLE>default.asp</TITLE>
 		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<meta http-equiv="Content-Language" content="ru">
    </HEAD>
<BODY topmargin="10" leftmargin="10" text="#0000FF">
   
 
        </b></font><font face="Arial"  size="2"> </h2>

<table align="center" width="80%" border="0">

<%			// Map  database to physical path
			
			filePath = Server.MapPath("site.mdb");
			// Create ADO Connection Component to connect with
			// sample database
			
			oConn = Server.CreateObject("ADODB.Connection");
			oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
			
			// Execute a SQL query and store the results within
			// recordset
			oRs = oConn.Execute("SELECT * From net_kont2  ORDER BY  dataD ");
%>
 
<thead>
   <tr>
      <td align="center"><b>Номер</b></td>
      <td align="center"><b>Имя</b></td>
      <td align="center"><b>Фамилия</b></td>
      <td align="center"><b>Адрес</b></td>
      <td align="center"><b>E-mail</b></td>      
      <td align="center"><b>Пункт выдачи</b></td>
      <td align="center"><b>Оператор</b></td>
      
   </tr>
 
<% flag=1;   i=0; %>
	
<%  while (!oRs.eof) { 

i++;  %>	
  
<tr>
  
<% if (flag==1) {%>  <td align="right" bgcolor='#99CCFF'>  <%  }
   if (flag==0) {%>  <td align="right" bgcolor='#FFFFFF'>  <%  }  %>		
    	
<% = i %>            </TD>
 
<% 		for(Index=0; Index < (oRs.fields.count); Index++) { 				
					
   if (flag==1) {%>  <td align="left" bgcolor='#99CCFF'>  <%  }
   if (flag==0) {%>  <td align="left" bgcolor='#FFFFFF'>  <%  } 
 
gofile = "update.asp?dataD="+oRs("dataD"); %>
<%= "<a href= " +gofile + "  >	" %>
    	
<% = oRs(Index)%>    </TD>
					
            <%  	}%>

</tr>
<% oRs.MoveNext();
				
if (flag==1) flag=0 ; else flag=1 ;
			
}%>

</TABLE>
<%   
	oRs.close();
	oConn.close();
%>

  <p>&nbsp;</font></p>
  <p>&nbsp;</p>

<form method="GET" action="delete.asp"> 
      <input type="submit" value="Удалить" name="B1">
      <input type="reset" value="Cбpoc" name="B2"></p>
</form>

   
		<p><a href="form.htm">Вернуться</a><span lang="en-us">
						</a></p>

   
</BODY> </HTML>