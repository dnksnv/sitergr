
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<meta http-equiv="Content-Language" content="ru">
</head>

<%@ LANGUAGE = JScript %>
<%
	ima1=Request.QueryString("firstName");
	fam1=Request.QueryString("surNumber");      
	address1=Request.QueryString("address");  
	email1=Request.QueryString("email");  	
    punk1=Request.QueryString("Punkts");
    oper1=Request.QueryString("Operators");
            filePath = Server.MapPath("site.mdb");
	        oConnAdd = Server.CreateObject("ADODB.Connection");
		 	oConnAdd.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);	    
            SQL_Sapros="INSERT INTO site_tab VALUES('"+ima1+"','"+fam1+"','"+address1+"','"+email1+"','"+punk1+"','"+oper1+"')"; %>
   
    <%=SQL_Sapros%>
    
<%          oRsAdd = oConnAdd.Execute(SQL_Sapros);
            oConnAdd.close();	%>
	    
<p>Для завершения добавления нажмите ссылку <a href="form.htm">Далее</a>

</BODY> </HTML>