<%@ LANGUAGE = JScript %>
<%
	fam1=Request.QueryString("surName");
	        filePath = Server.MapPath("site.mdb");
	        oConnDel = Server.CreateObject("ADODB.Connection");
			oConnDel.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
			SQL_Sapros="DELETE * From site_tab WHERE fam=" + "'"+ fam1+ "'";%>
	<%=SQL_Sapros%>	
			
<%			oRsDel = oConnDel.Execute(SQL_Sapros);		   
		    oConnDel.close();
%>

<p>Для завершения удаления нажмите ссылку <a href="default.asp">Далее</a> 

</BODY> </HTML>