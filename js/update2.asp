<%@ LANGUAGE = JScript %>
<%   
	ima1=Request.QueryString("firstName");
	fam1=Request.QueryString("surName");      
	address1=Request.QueryString("address");  
	email1=Request.QueryString("email");  	
    punk1=Request.QueryString("Punkts");
    oper1=Request.QueryString("Operators");
            filePath = Server.MapPath("site.mdb");
	        oConnUp = Server.CreateObject("ADODB.Connection");
			oConnUp.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +filePath);
          	SQL_Sapros="Update site_tab SET ima= '"+ima1+"', fam= '"+fam1+"', address= '"+address1+"', email= '"+email1+"', punk= '"+punk1+"', oper= '"+oper1+"'" ;
		    oRsDel = oConnUp.Execute(SQL_Sapros);		   
		    oConnUp.close();
%>

&#65533;&#65533;&#65533; &#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533; &#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533; &#65533;&#65533;&#65533;&#65533;&#65533;&#65533;&#65533; &#65533;&#65533;&#65533;&#65533;&#65533;&#65533; <a href="default.asp">&#65533;&#65533;&#65533;&#65533;&#65533;</a> 

</BODY> </HTML>