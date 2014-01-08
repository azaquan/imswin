<%@ Language=VBScript %>

<% Response.Expires =-1%>

<%
  dim ponumb
  dim namespace
  dim Ocnn
  DIM cmd
  dim Rs
  dim Rsrec
  dim sql
  
  ponumb= Request.QueryString ("ponumb")
  
  namespace=Request.QueryString("namespace")
  
  if Request.QueryString("from")=1 and ( UCASE(NAMESPACE) <> "DEMO2" and UCASE(NAMESPACE) <> "ILFC01" and UCASE(NAMESPACE) <> "AIRINDIA") then 

	
	   Response.Redirect "enterpo.asp?ERRcode=1"
	   
  end if
  	   
  SESSION("namespace")=namespace
	
	IF UCASE(NAMESPACE) ="DEMO2" THEN 	
	
		NAMESPACE="NVLAR"
	
		SESSION("namespace")=namespace
	
	elseif UCASE(NAMESPACE) ="ILFC01" THEN 	
	
		NAMESPACE="ILFC"
	
		SESSION("namespace")=namespace
		
	elseif UCASE(NAMESPACE) ="AIRINDIA" THEN 	
	
		NAMESPACE="AIRINDIA"
	
		SESSION("namespace")=namespace
	
	END IF	
	  
%>
<!--#include file="connection.asp"-->



 <% 
 
  namespace = ucase(namespace)
  
 sql= "SELECT DISTINCT dbo.PO.po_ponumb, dbo.PO.po_date, dbo.PO.po_reqddelvdate, " 
sql=sql &  " X1.usr_username AS Buyer, X2.usr_username AS Approv,  "
sql=sql &  "    s1.sts_name AS Status, s2.sts_name AS Sdeliv,  "
sql=sql &  "    s3.sts_name AS Sship, s4.sts_name AS Sinvt,  "
sql=sql &  "    dbo.SUPPLIER.sup_name, dbo.PRIORITY.pri_desc "
sql=sql &  "FROM dbo.PO LEFT OUTER JOIN "
sql=sql &  "    dbo.STATUS s1 ON dbo.PO.po_stas = s1.sts_code AND  "
sql=sql &  "    dbo.PO.po_npecode = s1.sts_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.STATUS s2 ON dbo.PO.po_stasdelv = s2.sts_code AND  "
sql=sql &  "    dbo.PO.po_npecode = s2.sts_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.STATUS s3 ON dbo.PO.po_stasship = s3.sts_code AND  "
sql=sql &  "    dbo.PO.po_npecode = s3.sts_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.STATUS s4 ON dbo.PO.po_stasinvt = s4.sts_code AND  "
sql=sql &  "    dbo.PO.po_npecode = s4.sts_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.PRIORITY ON  "
sql=sql &  "    dbo.PO.po_priocode = dbo.PRIORITY.pri_code AND  "
sql=sql &  "    dbo.PO.po_npecode = dbo.PRIORITY.pri_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.SUPPLIER ON  "
sql=sql &  "    dbo.PO.po_suppcode = dbo.SUPPLIER.sup_code AND  "
sql=sql &  "    dbo.PO.po_npecode = dbo.SUPPLIER.sup_npecode LEFT OUTER "
sql=sql &  "     JOIN "
sql=sql &  "    dbo.XUSERPROFILE X2 ON  "
sql=sql &  "    dbo.PO.po_apprby = X2.usr_userid AND  "
sql=sql &  "    dbo.PO.po_npecode = X2.usr_npecode LEFT OUTER JOIN "
sql=sql &  "    dbo.XUSERPROFILE X1 ON  "
sql=sql &  "    dbo.PO.po_buyr = X1.usr_userid AND  "
sql=sql &  "    dbo.PO.po_npecode = X1.usr_npecode, dbo.BUYER "
sql=sql &  " WHERE (dbo.PO.po_ponumb = '" & ponumb & "') AND "
sql=sql &  "    (dbo.PO.po_npecode = '" & namespace & "') "
 
  
 ' Response.Write sql
  set ocnn= server.CreateObject("adodb.connection")
  set cmd = server.CreateObject("adodb.command") 
  Ocnn.ConnectionString = connection
  Ocnn.Open ,sa
  
  set rs = server.CreateObject("adodb.recordset")
  set rs.ActiveConnection = Ocnn
  rs.Source =sql
  rs.Open ,,3,1
  set Rsrec= server.CreateObject("adodb.recordset")
  Rsrec.Source ="select porc_recpnumb, porc_rec from porec where porc_ponumb='" & ponumb & "' and porc_npecode='" & namespace & "' order by porc_recpnumb"
  Rsrec.ActiveConnection = Ocnn
  Rsrec.Open ,,3,1
  %>
  <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<meta name="author" content="Mohammed Muzammil H">

<meta name="description" content>
<head>
	<title>View Transaction Status</title>
	

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"><!--Image Map Nav Begins, Do Not Delete-->

	

<table border="0" cellpadding="0" cellspacing="0" width="740">

	<tr>
		<td><img border="0" height="1" src="images/spacer.gif" width="1"></td>
		<td><img border="0" height="1" src="images/spacer.gif" width="30"></td>
		<td><img border="0" height="1" src="images/spacer.gif" width="1"></td>
	</tr>
	<tr>
		
		<td colspan="3"><img border="0" src="images/ohd_home.gif" style="WIDTH: 741px" useMap="#navMap" width="741"></td>
	</tr>
	
  <tr>
	
	<td bgcolor="#eee8aa" align="middle"> </td>
	<%if len(session("userid"))=0 then
	     %>
	   <td bgcolor="#eee8aa" align="right"><a href="login.asp"><img border="0" height="19" src="images/signin.gif" width="67"></a></td></td>
	<% else	%>
	<td bgcolor="#eee8aa" align="middle"><strong><font type="ariel helwetteca">Welcome <%=session("Fname")%>!</font></strong> </td>
	<td bgcolor="#eee8aa" align="right"><a href="signout.asp"><img border="0" height="19" src="images/signout.gif" width="67"></a></td></td>
	<%end if%>
	</tr>

	
	<tr>
	
	<tr></tr>
	
	</table>
    
    <table BORDER="0" width="748">
    <tbody>

    <tr>
    
    <td WIDTH="20" valign="top" bgColor="#eee8aa">
   
    <!--<table border="0" style="HEIGHT: 98px; WIDTH: 125px">    <tr>            <td height="200" style="HEIGHT: 150px" valign="top">&nbsp;<FONT color=#cc3300 size=2>             <CENTER><B>Information!</B></CENTER></FONT><BR>Listed are all the             details of the purchase order. For information&nbsp;on their             detailed line items please&nbsp;click on the View&nbsp;Line Items             button.     </td>    </tr>    <tr><td width="2" bgcolor="white"></td></tr>        </table>   -->
    </td>
    
    <td>
    <table border="0">
        <tbody>
                 <tr align="middle">
                 <td colspan="2" align="middle">
			     <img border="0" height="17" src="images/bookingheading.gif" width="401">
			     </td>
			     </tr>
    
    <b>
    <tr>
    <td align="middle">

			   
         			
			        <table align="right" border="0" width="100%" height="100%" cellspacing="0">
			        
			        
			        
			        <tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">TRANSACTION DETAILS</font></strong></center></td>
			         </tr> 
			       
			         <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
			       
			        <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			       <% if rs.RecordCount =0 then
			           Response.Write "<TR><TD colspan=4><strong>The Transaction Order does not exist. Please try again</strong></TD></TR>"
			           Response.Write "<tr><td width=2 bgcolor='#eee8aa' COLSPAN=4></td></tr>"
			           else
			        %>   
			       
			       
			        <tr bgcolor="#eee8aa"><td nowrap><font size="2" face="normal"><strong>Transaction Order :</strong></td><td colspan="3" align="left"><strong><%=rs("po_ponumb")%></strong></td></font></tr>
			        <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			       			         
			         <tr>
							<td bgcolor="white"><font size="2" face="normal"><strong>Issued Date:</strong></font></td><td bgcolor="white"><font size="2" face="normal"><%=rs("po_date")%></font></td>
							<td bgcolor="white" align="middle" nowrap><font size="2" face="normal"><strong>Requested Date Delivery:</strong>	</td><td bgcolor="white" align="middle"><font size="2" face="normal"><%=rs("po_date")%></td>
					 </tr>
			        
			         
			         			          <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			         
					<tr>
					  <td bgcolor="white" align="left"><font size="2" face="normal"><strong>Buyer:</strong></td><td align="left" bgcolor="white"><font size="2" face="normal"><%=rs("Buyer")%></td>
					  <td bgcolor="white" align="left" nowrap><font size="2" face="normal"><strong>Approved By:</strong>	</td><td bgcolor="white" align="left"><font size="2" face="normal"><%=rs("approv")%></td>
					</tr>

								          <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<tr>
					  <td bgcolor="white" align="left"><font size="2" face="normal"><strong>Supplier:</strong></td><td align="left" bgcolor="white"><font size="2" face="normal"><%=rs("sup_name")%></td>
					  <td bgcolor="white" nowrap align="left"><font size="2" face="normal"><strong>Shipped Via:	</strong></td><td bgcolor="white" align="left"><font size="2" face="normal"><%=rs("pri_desc")%></td>
					</tr>

								          

								          <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
                    
                    <tr height="25"> <td bgcolor="white">&nbsp; </td>	</tr>
                    
                    <tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white"><strong>TRANSACTION STATUS</strong></font></strong></center></td>
			         </tr>
                     
                     <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
                     
					<tr>
						<td bgcolor="white" align="left"><font size="2" face="normal"><strong>Status PO:</strong></td><td align="left"><font size="2" face="normal"><%=rs("Status")%></td>
						<td bgcolor="white" nowrap align="left"><font size="2" face="normal"><strong>Status Reception:</strong></td><td align="left"><font size="2" face="normal"><%=rs("Sdeliv")%></td>
					</tr>

								          <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<tr>
						<td bgcolor="white" nowrap align="left"><font size="2" face="normal"><strong>Status Shipping:</strong></td><td align="left"><font size="2" face="normal"><%=rs("Sship")%></td>
						<td bgcolor="white" nowrap align="left"><font size="2" face="normal"><strong>Status Inventory:</strong></td><td align="left"><font size="2" face="normal"><%=rs("Sinvt")%></td>
					</tr>

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
					<tr height="25">
					  <td bgcolor="white">&nbsp; </td>
					</tr>
					
					<tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">RECEPIENTS</font></strong></center></td>
			         </tr> 
					
					<tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
					
					<% do while not Rsrec.EOF %>
					<tr><td colspan="4" align="middle"><font size="2" face="normal"><%=rsrec("porc_rec")%></td></tr>
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					<%Rsrec.MoveNext 
					loop
					%>
								         
					
					<tr height="100">			     
						<td colspan="4" align="center">
							<table>
							<tr>
									<td colspan="1" align="right">
									<form action="messages.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form1" name="form1">
									<input id="Message" name="Messages" type="submit" value="Messages">
									</form>
									</td>

									<td colspan="2" align="right">
									<form action="poitemheader.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form2" name="form2">
									<input id="LineItem" name="LineItem" type="submit" value="Line Items">
									</form>
									</td>

									<td colspan="1" align="left">
									 <form action="NEWPO.asp" method="post" id="form2" name="form2">
									<input id="Poheader" name="Poheader" type="submit" value="New Track">
									</form>
									</td>
									
							</tr>
							</table>
						</td>
			      </tr>
					
					
					
					
					
					
					
					<% end if%>
					
</table>
</td>
     </tr></tbody>
     </table></b>
     </td>
     </tr></tbody>
     </table></td></tr></tbody></table></b></td></tr></tbody></table>
     
     <table width="750">  
      	<tr>
		<td bgcolor="#0066cc" width="740" valign="top" height="20"><img alt border="0" height="18" src="images/logo.gif" width="150"></td>
		</tr>
	
	<tr align="middle" valign="top" colspan="3">
		
		<td width="100%"><font size="1"> © 1993 - 2001 IMS Inc. All Rights Reserved.</font></td>
		</tr>
	</table>
   
</body>
</html>
