<%@ Language=VBScript %>

<% Response.Expires =-1%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<meta name="author" content="Mohammed Muzammil H">

<meta name="description" content>
<head>
	<title>View Transaction Status</title>
	

</head>
<%
  dim ponumb
  dim namespace
  dim Ocnn
  DIM cmd
  dim Rs
  dim Rsrec
  dim sql
  
  ponumb= Request.QueryString ("ponumb")
  namespace=Request.QueryString ("namespace") %>
  
  <!--#include file="connection.asp"-->
  
  <%
  sql="SELECT OB_PONUMB,ob_mesgnumb,ob_remk,ob_oper,ob_mesgdate FROM OBS WHERE OB_PONUMB='" & ponumb & "' AND OB_NPECODE='" & namespace & "'"
  
  set ocnn= server.CreateObject("adodb.connection")
  Ocnn.ConnectionString = connection
  Ocnn.Open ,sa
  
  set rs = server.CreateObject("adodb.recordset")
  set rs.ActiveConnection = Ocnn
  rs.Source =sql
  rs.Open ,,3,1
  
  %>
  
  <%
  
  sql= " select pl_manfnumb,pl_shipdate,pl_creadate,pl_remk from packinglist where pl_manfnumb in ( select pld_manfnumb from packingdetl where pld_ponum = '" & ponumb & "' and pld_npecode ='" & namespace & "') and len(isnull(pl_remk,'')) > 0"
  set RsPackingList = server.CreateObject("adodb.recordset")
  set RsPackingList.ActiveConnection = Ocnn
  RsPackingList.Source = sql
  RsPackingList.Open ,,3,1
  
  %>
  
  
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"><!--Image Map Nav Begins, Do Not Delete-->
  
	<map name="navMap">
	<area shape="RECT" coords="319,78,402,92" href="RateRequest.asp" alt="Rate Request">
	<area shape="RECT" coords="412,78,512,92" href="bookprevquote.asp" alt="Booking">
	<% if len(trim(session("userid")))=0 then %>
	<area shape="RECT" coords="525,78,605,92" href="login.asp" alt="Edit Profile">
	<% elseif len(trim(session("userid")))>0 then  %>
	<area shape="RECT" coords="525,78,605,92" href="profile.asp" alt="Edit Profile">
	<%end if%>
	
	<% if len(trim(session("userid")))=0 then %>
	<area shape="RECT" coords="615,78,700,92" href="login.asp" alt="Edit password">
    <% elseif len(trim(session("userid")))>0 then  %>
	<area shape="RECT" coords="615,78,700,92" href="changepassword.asp" alt="Edit password">
	<%end if%>
	
	</map><!--Image Map Nav Ends, Do Not Delete-->&nbsp;		

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
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">Messages</font></strong></center></td>
			         </tr> 
			         <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
			        <!--<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>-->
			        
			        <tr><td><strong><font size="2" face="normal">Transaction Order :</strong></td><td colspan="3" align="left"><strong><font size="2" face="normal"><%=ponumb%></strong></td></tr>
			        
			         <tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>       
			         <tr bgcolor="#eee8aa">
			         
							<td><strong><font size="2" face="normal">Message Number</strong></td>
							
							<td align="left"><strong><font size="2" face="normal">Operator</strong></td>
							
							<td align="left"><strong><font size="2" face="normal">Date</strong></td>
							
					 </tr>
					
					<tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
			        
			        <%if rs.EOF =true and rs.BOF =true then

			         Response.Write "<tr><td width=2 bgcolor='#eee8aa' COLSPAN=4></td></tr>"
			          Response.Write "<tr><td colspan=4><strong><font size=2>No Messages stored</strong></font></TD></TR>"
			        Response.Write "<tr><td width=2 bgcolor='#eee8aa' COLSPAN=4></td></tr>"
			        else
			        
			        %>
			        
			        
			       	<% do while not rs.EOF %>
			       	
			       	 <%if rs.AbsolutePosition  <> 1 then %>
			       <!--   <tr height="30"><td bgcolor="white" nowrap colspan="2" align="middle"></td><td bgcolor="white" colspan="2" align="left"></td></tr> -->
			       <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
			        <%end if%>
                   			        
                    
			         
					<tr>					 
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=trim(rs("ob_mesgnumb"))%></td>
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=trim(rs("ob_oper"))%></td>
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=trim(rs("ob_mesgdate"))%></td>
			         </tr>
			         
			         <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			         
					 <tr>
					  <td bgcolor="white"><strong><font size="2" face="normal">Message</strong></td></tr>
					 <!-- <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>-->
					  <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					  
					 <tr> 	
					  <td colspan="3" bgcolor="white"><font size="2" face="normal"><%=rs("ob_remk")%></td>
					 </tr>
					 
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<%rs.MoveNext 
					  loop%>
					<%end if%>  
					
					
					
			        <tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>       
			        <tr bgcolor="#eee8aa">
			         
							<td><strong><font size="2" face="normal">Manifest</strong></td>
							<td align="left"><strong><font size="2" face="normal">Shipping Date</strong></td>
							<td align="left"><strong><font size="2" face="normal">creation Date</strong></td>
		
					</tr>
					
					<tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
					
					<%if RsPackingList.RecordCount =0 then
			         Response.Write "<tr><td width=2 bgcolor='#eee8aa' COLSPAN=4></td></tr>"
			          Response.Write "<tr><td colspan=4><strong><font size=2>No Manifest Messages stored</font></strong></TD></TR>"
			        Response.Write "<tr><td width=2 bgcolor='#eee8aa' COLSPAN=4></td></tr>"
			        else
			        
			        %>
			      
			        
			       	<% do while not RsPackingList.EOF %>
			       	
			       	 <%if RsPackingList.AbsolutePosition  <> 1 then %>
			 <!--       <tr height="30"><td bgcolor="white" nowrap colspan="2" align="middle"></td><td bgcolor="white" colspan="2" align="left"></td></tr>-->
			 <tr><td width="100%" bgcolor="#eee8aa" COLSPAN="4"><img src="images/whitebar.gif" width="100%" height="1"></td></tr>
			        <%end if%>
                   			        
                     
					<tr>					 
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=RsPackingList("pl_manfnumb")%></td>
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=formatdatetime(cdate(RsPackingList("pl_shipdate")),2)%></td>
							<td bgcolor="white" align="left"><font size="2" face="normal"><%=formatdatetime(cdate(RsPackingList("pl_creadate")) ,2)%></td>
			         </tr>
			         
			         <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			         
					 <tr>
					  <td bgcolor="white"><strong><font size="2" face="normal">Message</strong></td></tr>
					  
					  <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					  
					 <tr> 	
					  <td colspan="3" bgcolor="white"><font size="2" face="normal"><%=RsPackingList("pl_remk")%></td>
					 </tr>
					 
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<%RsPackingList.MoveNext 
					  loop%>
					<%end if%>  
					
					
					
					
					
					  <tr height="100">
					  
					  <td colspan="4" align="center">
					  <table>
					   <tr>
						<td colspan="1" align="right">
						 <form action="poitemheader.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form2" name="form2">
						<input id="LineItem" name="LineItem" type="submit" value="Line Items">
						</form>
						</td>
						<td colspan="1" align="left">
						 <form action="poheader.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form2" name="form2">
						<input id="Poheader" name="Poheader" type="submit" value="Transaction Header">
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
     
     
  <!-- <br>      <div align="center" class="size">			<A class=size href="www.ims-sys.com"><font color="#cc6633">Home</font></A>			 |  			<A class=size href="services.asp"><font color="#cc6633">Company Profile</font></A> |  			<A class=size href="RateRequest.asp"><font color="#cc6633">Rate Request</font></A> |  			<A class=size href="booking.asp"><font color="#cc6633">Booking</font></A><br>  			<A class=size href="agentlist.asp"><font color="#cc6633">Agent List</font></A> |						<A class=size href="requestid.html"><font color="#cc6633">Contact Us</font></A> | 			  			 </div>      <div align="center" class="size">            <A class=size href="trademark_notice.html"><font size="2" color="#003366">1093-2001, IMS.</font></A> </div>	-->

   

</body>
</html>
