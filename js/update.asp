<%@ LANGUAGE = JScript %>
<%
	dataDl=Request.QueryString("dataD"); 	
                      %>

<HTML>
    <HEAD>
         <TITLE>update.asp</TITLE>
 		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<meta http-equiv="Content-Language" content="ru">
 </HEAD>
    <BODY topmargin="10" leftmargin="10" text="#0000FF">
   
    <h2 align="center">
   
        </b></font><font face="Arial"  size="2" >
    </h2>
<table align="center" width="80%" border="0">

<%				
			filePath = Server.MapPath("site.mdb");
	        oConn = Server.CreateObject("ADODB.Connection");
		 	oConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
     	 	SQL_Sapros="SELECT * From site_tab  WHERE dataD=" + "'"+ dataDl+ "'";
		  	oRs = oConn.Execute(SQL_Sapros);		
%>
  
 
 
</TABLE>

<p>&nbsp;</p> 	</font>
<form method="GET" action="update2.asp" >
      <p>Имя&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="firstName" size="30" value='<%=oRs("ima")%>'></p>
      <p>Фамилия&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="surName" size="30" value='<%=oRs("fam")%>'></p>
      <p>Адрес&nbsp;&nbsp;&nbsp; <input type="text" name="address" size="30" value='<%=oRs("address")%>'></p>
      <p>E-mail&nbsp;&nbsp;&nbsp; <input type="text" name="email" size="30" value='<%=oRs("email")%>'></p>
      <p>Пункт выдачи&nbsp;&nbsp;&nbsp; <input type="text" name="Punkts" size="30" value='<%=oRs("punk")%>'></p>
      <p>Оператор&nbsp;&nbsp;&nbsp; <input type="text" name="Operators" size="30" value='<%=oRs("oper")%>'></p>

      <p><input type="submit" value="Обновить" name="B1"><input type="reset" value="C6poc" name="B2"></p>
</form>
<% 	oRs.close();
	oConn.close();		
%> 
 
</BODY> </HTML>