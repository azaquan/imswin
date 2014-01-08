<%@ Language=VBScript %>
<!--#include file="connection.asp"-->
<% Response.Expires =-1 %>
<%Response.Buffer =false%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<meta name="author" content="Mohammed Muzammil H">

<meta name="description" content>
<head>
	<title>View Line Items</title>
	


</head>
<%
  dim ponumb
  dim namespace
  dim Ocnn
  DIM cmd
  dim Rs
  dim Rsrec
  dim RsInvoice
  dim Rsship
  Dim Rsware
  dim sql
  
  ponumb= Request.QueryString("ponumb")
  namespace=Request.QueryString("namespace")
  
   
sql= "SELECT dbo.POITEM.poi_ponumb, dbo.POITEM.poi_liitnumb,  "
sql= sql & "    dbo.POITEM.poi_desc, dbo.POITEM.poi_unitofp,  "
sql= sql & "    dbo.POITEM.poi_primreqdqty, U1.uni_desc AS PUnit,  "
sql= sql & "    dbo.POITEM.poi_secoreqdqty, U2.uni_desc AS SUnit,  "
sql= sql & "    S1.sts_name AS Status, S2.sts_name AS SDlv,  "
sql= sql & "    S3.sts_name AS SShip, S4.sts_name AS SInvt,  "
sql= sql & "    dbo.POITEM.poi_comm, dbo.POITEM.poi_liitreqddate, "
sql= sql & "    MANUFACTURER.man_name,dbo.POITEM.poi_manupartnumb,poi_serlnumb "
sql= sql & "FROM dbo.POITEM LEFT OUTER JOIN "
sql= sql & "    dbo.MANUFACTURER ON  "
sql= sql & "    dbo.POITEM.poi_manupartnumb = dbo.MANUFACTURER.man_code "
sql= sql & "     AND  "
sql= sql & "    dbo.POITEM.poi_npecode = dbo.MANUFACTURER.man_npecode "
sql= sql & "     LEFT OUTER JOIN "
sql= sql & "    dbo.STATUS S4 ON  "
sql= sql & "    dbo.POITEM.poi_stasinvt = S4.sts_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = S4.sts_npecode LEFT OUTER JOIN "
sql= sql & "    dbo.STATUS S3 ON  "
sql= sql & "    dbo.POITEM.poi_stasship = S3.sts_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = S3.sts_npecode LEFT OUTER JOIN "
sql= sql & "    dbo.STATUS S2 ON  "
sql= sql & "    dbo.POITEM.poi_stasdlvy = S2.sts_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = S2.sts_npecode LEFT OUTER JOIN "
sql= sql & "    dbo.STATUS S1 ON  "
sql= sql & "    dbo.POITEM.poi_stasliit = S1.sts_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = S1.sts_npecode LEFT OUTER JOIN "
sql= sql & "    dbo.UNIT U1 ON  "
sql= sql & "    dbo.POITEM.poi_primuom = U1.uni_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = U1.uni_npecode LEFT OUTER JOIN "
sql= sql & "    dbo.UNIT U2 ON  "
sql= sql & "    dbo.POITEM.poi_secouom = U2.uni_code AND  "
sql= sql & "    dbo.POITEM.poi_npecode = U2.uni_npecode "
sql= sql & " WHERE (dbo.POITEM.poi_ponumb = '" & ponumb &  "') AND  "
sql= sql & "    (dbo.POITEM.poi_npecode = '" & namespace &  "')order by poitem.poi_liitnumb"
   
   

  
  set ocnn= server.CreateObject("adodb.connection")
  set cmd = server.CreateObject("adodb.command") 
  Ocnn.ConnectionString = connection
  Ocnn.Open ,sa
  
  set rs = server.CreateObject("adodb.recordset")
  set rs.ActiveConnection = Ocnn
  rs.Source =sql
  rs.Open ,,3,1
  
  if rs.RecordCount >0 then
    
    dim LIno
    LIno=rs("poi_liitnumb")
    
         'Extracting the details of the recption only that line item for which the Line Item detail is 
         'being displayed.
            
  else
    
    Response.Write "<table width='100%'>"
    Response.Write " <tr>"
	Response.Write "  <td colspan='4' bgcolor='#336699'><center><strong><font color='white'>No Line Items to Track.</font></strong></center></td> "
    Response.Write "  </tr> "
    Response.Write " </table>"
    
    Response.End 
    
   end if  
			         
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
		
		<td colspan="3"><img border="0" src="images/Ohd_home.gif" style="WIDTH: 741px" useMap="#navMap" width="741"></td>
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
   
   <!-- <table border="0" style="HEIGHT: 98px; WIDTH: 125px">       <tr>            <td height="200" style="HEIGHT: 150px" valign="top">&nbsp;<FONT color=#cc3300 size=2>             <CENTER><B>Information!</B></CENTER></FONT><BR>Listed are all the details             of the Line Items related with the Transaction             Order. For information&nbsp;on their&nbsp;messages please&nbsp;click on the View&nbsp;Messages             button.     </td>    </tr>    <tr><td width="2" bgcolor="white"></td></tr>        </table>   -->
    </td>
    
    <td>
    <table border="0">
        <tbody>
                 <tr align="middle">
                 <td colspan="2" align="middle">
			     <img border="0"  src="images/bookingheading.gif" width="401">
			     </td>
			     </tr>
    
    <b>
    <tr>
    <td align="middle">

			   
         			
			        <table align="right" border="0" width="100%" height="100%" cellspacing="0">
			        
			        
			        <tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">LINE ITEM DETAILS</font></strong></center></td>
			         </tr> 
			        <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			        <tr ><td COLSPAN=4 NOWRAP align="left"><strong>TRANSACTION ORDER :<%=rs("poi_ponumb")%></strong></td></tr>

<!-- from here starts the Loop -->
                   <%do while not rs.EOF 
                   
                   LIno=rs("poi_liitnumb")
                   %>
					
					<%if rs.AbsolutePosition  <> 1 then %>
			        <tr height="30"><td bgcolor="white" nowrap colspan="2" align="middle"></td><td bgcolor="white" colspan="2" align="left"></td></tr>
			        <%end if%>
			        
			        <tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>
			        <tr><td bgcolor="#eee8aa" nowrap colspan="2" align="middle"><strong>Line Item No.:</strong></td><td bgcolor="#eee8aa" colspan="2" align="left"><strong><%=rs("poi_liitnumb")%></strong></td></tr>
                    <tr><td width="4" bgcolor="red" COLSPAN="4"></td></tr>
                    <tr><td bgcolor="white" nowrap colspan="2" align="right"><strong>Primary</strong></td><td bgcolor="white" align="right"><strong>Secondary</strong></td></tr>
			          <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			         <tr>
							<td bgcolor="white" nowrap><strong>Quantity Requested:</strong></td>
							<td bgcolor="white" align="right"><%=rs("poi_primreqdqty")%></td>
							<td bgcolor="white" align="right"><%=rs("poi_secoreqdqty")%></td>
							<td></td>
					 </tr>
			        
			         
			         <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			         
					<tr>
					  
							<td bgcolor="white"><strong>Unit:</strong></td>
							<td bgcolor="white" align="right"><%=rs("PUnit")%></td>
							<td bgcolor="white" align="right"><%=rs("SUnit")%></td>
							<td></td>
					 
					</tr>

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<tr>
					  
					  <td bgcolor="white" colspan="1" align="left" nowrap><strong>Unit Of Purchase:</strong></td>
					  
					  <%if ucase(trim(rs("poi_unitofp")))="P" then%>
					       <td bgcolor="white" colspan="3" align="left">Primary</td>
					  <%else%>
					       <td bgcolor="white" colspan="3" align="left">Secondary</td>
					  <%end if%>
					  
					</tr>

					 <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<tr>
					  <td bgcolor="white"><strong>Date Requested:</strong></td><td bgcolor="white"><%=rs("poi_liitreqddate")%></td>
					  <td bgcolor="white" align="middle" nowrap><strong>Stock Number:</strong>	</td><td bgcolor="white" align="left"><%=rs("poi_comm")%></td>
					</tr>

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
                    
                    
                    <tr>
                    <td bgcolor="white" colspan="1" align="left" nowrap><strong>Manufacturer P/N:</strong></td>
                    <td bgcolor="white"><%=trim(rs("poi_manupartnumb") & "")%></td>
                    <td bgcolor="white" colspan="1" align="right" nowrap><strong>Serial#:</strong></td>
                    <td bgcolor="white"><%=trim(rs("poi_serlnumb") & "")%></td>
                    </tr>
                    
                    <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
                    
                    <tr>
					  <td bgcolor="white"><strong>Description:</strong></td>
					  <td bgcolor="white" align="left" wrap colspan="3"><%=rs("poi_desc")%></td>
					</tr>
					
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
					<tr height=25>
					  <td bgcolor="white">&nbsp; </td>
					</tr>
					
                    
                    <tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font  color="white">LINE ITEM STATUS</font></strong></center></td>
			         </tr>
                     
                     
                     <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
                     
					<tr>
						<td bgcolor="white"><strong>Status LI:</strong></td><td><%=rs("Status")%></td>
						<td bgcolor="white" nowrap><strong>Status Reception:</strong></td><td><%=rs("SDlv")%></td>
					</tr>

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>

					<tr>
						<td bgcolor="white" nowrap><strong>Status Shipping:</strong></td><td><%=rs("SShip")%></td>
						<td bgcolor="white" nowrap><strong>Status Inventory:</strong></td><td><%=rs("SInvt")%></td>
					</tr>

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
				
					
					<tr height=25>
					  <td bgcolor="white">&nbsp; </td>
					</tr>
					
					<tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">FREIGHT FORWARDER RECEIPT</font></strong></center></td>
			         </tr>			         
				
				<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
					<tr>
						<td bgcolor="white" nowrap align="middle"><strong>Date</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Reception #</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Qty Delivered</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Unit Price</strong></td>
					</tr>			          
					 <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
                <% 
                
                set Rsrec= server.CreateObject("adodb.recordset")  
				sql ="SELECT RECD_RECPNUMB,recd_datedlvd,recd_primqtydlvd,recd_liitnumb,recd_unitpric FROM "
				sql= sql & " rECEPTIONDETL "
				sql= sql & "WHERE RECD_RECPNUMB IN (select rec_recpnumb from reception where rec_PONUMB='" & ponumb &"' and rec_npecode='" & namespace &"') "
				sql= sql & " and recd_liitnumb='" & LIno &"' and recd_npecode='" & namespace & "'"
				Rsrec.Source = sql
				Rsrec.ActiveConnection = Ocnn
				Rsrec.Open ,,3,1
                
                do while not rsrec.EOF %>								          
					<tr>
						<td bgcolor="white" nowrap align="middle"><%=rsrec("recd_datedlvd")%></td>
						<td bgcolor="white" nowrap align="middle"><%=rsrec("RECD_RECPNUMB")%></td>
						<td bgcolor="white" nowrap align="middle"><%=rsrec("recd_primqtydlvd")%></td>
						<td bgcolor="white" nowrap align="middle"><%=rsrec("recd_unitpric")%></td>
					</tr>
					 <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					<%Rsrec.MoveNext%>
				<%loop%>	
				
				
					
					<tr height=25>
					  <td bgcolor="white">&nbsp; </td>
					</tr>
					
					<tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font size="3" color="white">INVOICE</font></strong></center></td>
			         </tr>			         
				
						<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
						
							
					<tr>
						<td bgcolor="white" nowrap align="middle"><strong>Date</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Invoice #</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Qty Invoiced</strong></td>
						<td bgcolor="white" nowrap align="middle"><strong>Unit Price</strong></td>
					</tr>			          
					 <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
			
				
				<%

				set RsInvoice = server.CreateObject("adodb.recordset")
			    sql = " SELECT dbo.INVOICEDETL.invd_liitnumb, "
				sql= sql  & " dbo.INVOICE.inv_invcdate, "
				sql= sql  & "    dbo.INVOICEDETL.invd_primreqdqty, " 
				sql= sql  & "    dbo.INVOICEDETL.invd_unitpric,  "
				sql= sql  & "    dbo.INVOICE.inv_invcnumb "
				sql= sql  & " FROM dbo.INVOICEDETL,invoice"
				sql= sql  & "  where  "
				sql= sql  & "  INVOICEDETL.invd_npecode = INVOICE.inv_npecode AND "
				sql= sql  & "     INVOICEDETL.invd_invcnumb = INVOICE.inv_invcnumb "
				sql= sql  & " and  (INVOICEDETL.invd_invcnumb = "
				sql= sql  & "        (SELECT inv_invcnumb "
				sql= sql  & "      FROM invoice "
				sql= sql  & "      WHERE inv_ponumb = '" & ponumb & "' AND inv_npecode = '" & namespace & "')) AND  "
				sql= sql  & "    (dbo.INVOICE.inv_ponumb = '" & ponumb & "') AND  "
				sql= sql  & "    (dbo.INVOICE.inv_npecode = '" & namespace & "') and "
				sql= sql  & "    invoicedetl.invd_liitnumb='" & LIno & "' "
				'Response.Write sql
				
				RsInvoice.Source = sql
				
				RsInvoice.ActiveConnection = connection
				RsInvoice.Open ,,3,1
				
				%>
				<%
				if RsInvoice.RecordCount >0 then
				
				do while not RsInvoice.EOF %>								          
					<tr>
						<td bgcolor="white" nowrap align="middle"><%=RsInvoice("inv_invcdate")%></td>
						<td bgcolor="white" nowrap align="middle"><%=RsInvoice("inv_invcnumb")%></td>
						<td bgcolor="white" nowrap align="middle"><%=RsInvoice("invd_primreqdqty")%></td>
						<td bgcolor="white" nowrap align="middle"><%=RsInvoice("invd_unitpric")%></td>
					</tr>
					 <tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
					<%RsInvoice.MoveNext%>
					
				<%loop
				
				else
				
				  %>
				  
				 <tr><td colspan="4"><strong>No invoice details.</strong></td></tr>
				  
				  <%end if %>
				
				
				  
				
					
				
				<%
				
				set RsShip = server.CreateObject("adodb.recordset")
			    sql = "select pl_shipdate,pld_manfnumb,pld_reqdqty,pld_unitpric ,pl_awbnumb,pl_dest,pl_hawbnumb, "
			    sql = sql &   " pl_fig1,pl_from1,pl_to1 ,pl_etd ,pl_eta,pl_shiprefe,pl_forwrefe " 
				sql = sql &  " from packinglist,packingdetl "
				sql = sql &  " where  (pld_liitnumb='" & LIno & "' and pld_ponum='" & ponumb & "') and "
				sql = sql &  " (pl_manfnumb =pld_manfnumb and pl_npecode='" & namespace & "' ) "
				
				rsship.Source = sql
				rsship.ActiveConnection = connection
				rsship.Open ,,3,1
				%>
				
			
				<tr height=25><td bgcolor="white">&nbsp; </td>
				
				
				<tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">SHIPPING</font></strong></center></td>
			    </tr>			         
					  <% 
				if rsship.RecordCount >0 then
				%>				
				
				<tr>
					<td bgcolor="white" nowrap align="left"><strong>Date</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_shipdate")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Manifest #</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pld_manfnumb")%></td>
				</tr>			          
				
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
			    <tr>
					<td bgcolor="white" nowrap align="left"><strong>AWB#</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_awbnumb")%></td>
			<td bgcolor="white" nowrap align="left"><strong>HAWB</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_hawbnumb")%></td>
				</tr>			          

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
			    <tr>
					<td bgcolor="white" nowrap align="left"><strong>Flight1</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_fig1")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Destination</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_dest")%></td>
				</tr>			          

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>


				<tr>
					<td bgcolor="white" nowrap align="left"><strong>From:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_from1")%></td>
					<td bgcolor="white" nowrap align="left"><strong>To:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_to1")%></td>
				</tr>			          
				
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
				<tr>
					<td bgcolor="white" nowrap align="left"><strong>ETD:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_etd")%></td>
					<td bgcolor="white" nowrap align="left"><strong>ETA:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_eta")%></td>
				</tr>			          
				
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
				<tr>
					<td bgcolor="white" nowrap align="left"><strong>Qty Shipped:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pld_reqdqty")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Unit Price:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pld_unitpric")%></td>
				</tr>			          

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
				<tr>
					<td bgcolor="white" nowrap align="left"><strong>Shipper Ref:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_shiprefe")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Forwarder Ref:</strong></td>
					<td bgcolor="white" nowrap align="left"><%=rsship("pl_forwrefe")%></td>
				</tr>			          

			  <%else %>
			  
			    <tr>
					<td bgcolor="white" nowrap align="middle"><strong>No Shipping details.</strong></td>
				</tr>			          
					
				<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
			  <%end if %>
			  
			  
			  <%	set RsWare = server.CreateObject("adodb.recordset")
			  
			  sql=" select ird_trannumb,loc_name,ir_trandate,ird_primqty,ird_unitpric,ird_newcond from invtreceiptdetl, "
			  sql= sql & " invtreceipt,location  where (ird_npecode='" & namespace & "' and ird_ponumb='" & ponumb & "' and ird_liitnumb='" & LIno & "') and "
			  sql= sql & "	(ir_trannumb=ird_trannumb and ir_npecode=ird_npecode) and "
			  sql= sql & " (loc_locacode=ird_ware and loc_compcode=ird_compcode and  loc_npecode=ird_npecode) "
			    
			    
				
				RsWare.Source = sql
				RsWare.ActiveConnection = connection
				RsWare.Open ,,3,1
				%>
				
				<tr height=25><td bgcolor="white">&nbsp; </td>
				
				
				<tr>
			          <td colspan="4" bgcolor="#336699"><center><strong><font color="white">WAREHOUSE RECEIPT</font></strong></center></td>
			         </tr>			         
								          
                <% 
				if RsWare.RecordCount >0 then
				%>				
				
				<tr>
					<td bgcolor="white" nowrap align="left"><strong>Date</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("ir_trandate")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Transaction#</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("ird_trannumb")%></td>
				</tr>			          
				
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
			    <tr>
					<td bgcolor="white" nowrap align="left"><strong>Location</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("loc_name")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Condition</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("ird_newcond")%></td>
				</tr>			          

					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
					
			    <tr>
					<td bgcolor="white" nowrap align="left"><strong>Qty Received</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("ird_primqty")%></td>
					<td bgcolor="white" nowrap align="left"><strong>Unit Price</strong></td>
					<td bgcolor="white" nowrap align="left"><%=RsWare("ird_unitpric")%></td>
				</tr>			          
		          
				
					<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
				
			  <%else %>
			  
			    <tr>
					<td bgcolor="white" nowrap align="middle"><strong>No Warehouse receipts.</strong></td>
				</tr>			          
					
				<tr><td width="2" bgcolor="#eee8aa" COLSPAN="4"></td></tr>
				
			  <%end if %>
			  
			  <% 
			     Rsrec.Close 
			     rsware.Close 
			     rsship.Close 
			     RsInvoice.Close 
			     set rsrec= nothing
			     set rsware= nothing
			     set rsship= nothing
			     set RsInvoice =  nothing
			  %>   
			  <% rs.MoveNext 
			     loop %>
			     
					<tr height="100">			     
						<td colspan=4 align=center >
							<table>
							<tr>
									<td colspan="1" align="right">
									<form action="messages.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form1" name="form1">
									<input id="Message" name="Messages" type="submit" value="Messages">
									</form>
									</td>

									<td colspan="1" align="left">
									 <form action="poheader.asp?<%="ponumb=" & ponumb & "&namespace=" & namespace %>" method="post" id="form2" name="form2">
									<input id="Poheader" name="Poheader" type="submit" value="Transaction Header">
									</form>
									</td>
									<td colspan="1" align="left">
									 <form action="ENTERPO.asp" method="post" id="form2" name="form2">
									<input id="Poheader" name="Poheader" type="submit" value="New Track">
									</form>
									</td>
									
							</TR>
							</table>
						</TD>
			      </TR>
			     
			     
			     
	
			     
</table>
</td>
     </tr>
     
     
			     
     </tbody>
     </table></b>
     </td>
     </tr>
     
      
     
     </tbody>
     </table></td></tr>
     
			    
     </tbody></table></b></td></tr>
     
     
     </tbody></table>
     
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
