<%@ Language=VBScript %>
<% Response.Expires =-1%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<meta name="keywords" content>

<meta name="description" content>
<head>
	<title>Booking a request</title>
	





<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function IMG1_onmouseup() {
document.goto.submit();  
}

//-->
</script>
</head>

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
	<!--<area shape="RECT" coords="525,78,605,92" href="profile.asp" alt="Edit Profile">	<area shape="RECT" coords="615,78,700,92" href="changepassword.asp" alt="Edit password">-->
	
	</map><!--Image Map Nav Ends, Do Not Delete-->&nbsp;		

<table border="0" cellpadding="0" cellspacing="0" width="740">

	<tr>
		<td><IMG border=0 height=1 src="images/spacer.gif" width=1></td>
		<td><IMG border=0 height=1 src="images/spacer.gif" width=30></td>
		<td><IMG border=0 height=1 src="images/spacer.gif" width=1></td>
	</tr>
	<tr>
		
		<td colspan="3"><IMG border=0 src="images/Ohd_home.gif" style="WIDTH: 741px" useMap =#navMap width=741></td>
	</tr>
	
<!--	<tr>	<td bgcolor="#eee8aa" colspan="2" align="middle"> </td>	<td bgcolor="#eee8aa" align="right"><a href="signout.asp"><img border="0" height="19" src="images/signout.gif" width="67"></a></td></td>	</tr>	-->
	<tr>
	<td bgcolor="#eee8aa" align="middle" > </td>
	<%if len(session("userid"))=0 then
	     %>
	   <td bgcolor="#eee8aa" align="right"><A href="login.asp"><IMG border=0 height=19 src="images/signin.gif" width=67></A></td></TD>
	<% else	%>
	<td bgcolor="#eee8aa" align="middle" ><strong><font type="ariel helwetteca">Welcome <%=session("Fname")%>!</font></strong> </td>
	<td bgcolor="#eee8aa" align="right"><A href="signout.asp"><IMG border=0 height=19 src="images/signout.gif" width=67></A></td></TD>
	<%end if%>
	</tr>

	
	
	<tr>
	
	<tr></tr>
	
	</table><!--d valign="top" class="size"><!--Left Nav Home Begins, Do Not Delete-->
    
    <table BORDER="0" width="748">
    <tbody>

    <tr>
    
    <td WIDTH="125" valign="top" >
    <table border="0" style="HEIGHT: 98px; WIDTH: 125px">
    <tr>
    
    
    <td height="200" style="HEIGHT: 50px" valign="top">&nbsp;
    </td>
    </tr>
    

    
    </table>
   
    </td>
    
    
    
    
    <td align="middle" Valign="top">
      
      <table width="100%" border="0" CELLSPACING="0" CELLPADDING="0">
      <tbody>
      <tr>
      <!--<TD height=20>      </TD>-->
      
      <td align="middle" colspan="2"></td>
      </tr>
      <tr style="HEIGHT: 101px; WIDTH: 616px">
      <td colspan="3" valign="top" align="middle">
      <!-- TABLE WITH THE BORDER LINES START HERE -->
      <table border="0" width="75%" height="100" cellpadding="0" cellspacing="0" bgcolor="#f7f3d7">
      
      <tr>
      <!-- horizontal-->
      <td colspan="3" width="2" bgcolor="#d1d1d1" height="2" ></td>
      </tr>
      
      <!-- vertical-->
      <tr>
      
      <td rowspan="20" align="right" height="2" width="2" bgcolor="#d1d1d1"></td>
      
     
      
      <td valign="top" align="middle">
      
      <table>
      <tr>
      <td width="5"></td>
      <td align="middle">
      <center><font color="#cc3300"><b>Information!</b><br></font> </center>
        <font size="2">
                       
             Our specially designed software has the capability of&nbsp;tracking details of any transactions 
           recorded. Use this feature to view all the details of your&nbsp;POs.        
         </font>  
        
      
      
      </td>
      <td width="5"></td>
      <td></td>
      
      </tr>
      </table>
      
      
      <td rowspan="20" align="right" height="2" width="2" bgcolor="#d1d1d1"></td>
      </tr>
      
      <!--horizontal-->
      <tr>
      <td colspan="3" width="2" height="2" bgcolor="#d1d1d1" ></td>
      </tr>
      <!--Table with the border line ends here -->
      </table>
      </td>
      </tr>
      
      <tr>
      <td height="20">
      </td>
      </tr>
             
            </tbody>
            </table>
            </td></tr></tbody></table>

<TABLE width=750 >
  <TBODY>
  
   <tr>
      <td colspan="3" align="middle">
      <IMG border=0 height=17 src="images/bookingheader.gif" width=401>
      </td>
      </tr>
      
      <tr><td width="1" bgcolor="white"></td></tr>
     
  
  
  <form action="poheader.asp" name="goto" id="goto" method="get">
			
			<tr>
			<td bgcolor="#eee8aa" align="middle" >
			  <strong> Transaction&nbsp;Number </strong>
			</td>
			
			<td align=left bgcolor="#eee8aa">
			<INPUT id=ponumb name=ponumb></td>
            
            
            <tr>
            <td bgcolor="#eee8aa" align="middle" >
			<strong>   Namespace</strong>
			</td>
			<td bgcolor="#eee8aa" align=left>
			<INPUT id=namespace name=namespace>
			</td>
           
            </tr>
            
            <tr height=50><td colspan=3 align=middle><INPUT type="submit" value="Submit" id=submit1 name=submit1></td></tr>
            </form>
  
  
  
	    <TR  >
			<TD colspan=3 bgColor=#0066cc height=20 vAlign=top width=740>
			<IMG alt="" border=0 height=18 src="images/logo.gif" width=150  >
			</TD>
		</TR>
  
  <TR align=middle vAlign=top  colspan="3">
    <TD width="100%" ><FONT size=1 ></FONT>
			</TD>
			</TR></TBODY></TABLE>




</body>
</html>
