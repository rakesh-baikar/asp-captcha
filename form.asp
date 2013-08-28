<%@ LANGUAGE="VBScript" %>
<!--#include file="library/adovbs.asp"-->
<!--#include file="library/iasutil.asp"-->
<%
' On Error Resume Next

' Classsic ASP pages created by Andre F Bruton
' E-mail: andre@bruton.co.za
' Date: 2008/01/19
contact_firstname          = SQLEncode(Request("contact_firstname"))
contact_lastname           = SQLEncode(Request("contact_lastname"))
contact_email              = SQLEncode(Request("contact_email"))
contact_tel                = SQLEncode(Request("contact_tel"))
contact_mobile             = SQLEncode(Request("contact_mobile"))
contact_fax                = SQLEncode(Request("contact_fax"))
contact_skype              = SQLEncode(Request("contact_skype"))
contact_text               = SQLEncode(Request("contact_text"))
recaptcha_challenge_field  = Request("recaptcha_challenge_field")
recaptcha_response_field   = Request("recaptcha_response_field")
recaptcha_private_key      = "6LdV2eQSAAAAAJPtRw66z3YyEMOFh98elx1VUYgp"
recaptcha_public_key       = "6LdV2eQSAAAAABdipocITMdS9G1s4ZEEiXHxK5E9"
browser                    = Request.ServerVariables("HTTP_USER_AGENT")
ip                         = Request.ServerVariables("REMOTE_HOST")
%>
<html>
<head>
<title>Recaptcha - Classic ASP</title>
<meta http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<meta name="description" content="Classic ASP Recaptcha examples">
<meta name="keywords" content="recaptcha, classic, asp, example, examples">
<meta name="rating" content="general">
<meta name="robots" content="index,follow">
<link href="css/andre.css" type=text/css rel=stylesheet>
</head>

<body>
<% 
cTemp = recaptcha_confirm(recaptcha_private_key, recaptcha_challenge_field, recaptcha_response_field)
If cTemp <> "" Then 
%>

  <p>An error occured with the recapture wording.<br>
Please try again. The error was <b><%=cTemp%></b></p>

  <form action="form2.asp" method="post">
  <table cellspacing="2" cellpadding="0" border="0" width="530">
    <tr>
      <td align="right" width="140">First Name:</td>
      <td width="390"><input name="contact_firstname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_firstname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right">Last Name:</td>
      <td><input name="contact_lastname" style=" width: 150px;" class="textSize2 black" type="text" value="<%=contact_lastname%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right">Email Address:</td>
      <td><input name="contact_email" style=" width: 200px;" class="textSize2 black" type="text" value="<%=contact_email%>" maxlength="150"></td>
    </tr>
    <tr>
      <td align="right">Phone Number:</td>
      <td><input name="contact_tel" style=" width: 100px;" class="textSize2 black" type="text" value="<%=contact_tel%>" maxlength="30"></td>
    </tr>
    <tr>
      <td align="right">Mobile Number:</td>
      <td><input name="contact_mobile" style=" width: 100px;" class="textSize2 black" type="text" value="<%=contact_mobile%>" maxlength="30"></td>
    </tr>
    <tr>
      <td align="right">Fax Number:</td>
      <td><input name="contact_fax" style=" width: 100px;" class="textSize2 black" type="text" value="<%=contact_fax%>" maxlength="30"></td>
    </tr>
    <tr>
      <td align="right">Skype ID:</td>
      <td><input name="contact_skype" style=" width: 100px;" class="textSize2 black" type="text" value="<%=contact_skype%>" maxlength="50"></td>
    </tr>
    <tr>
      <td align="right" valign="top">Human Check:</td>
      <td><%=recaptcha_challenge_writer(recaptcha_public_key)%></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td align="right" valign="top" width="140">Additional information you would like us to know (optional)</td>
      <td><textarea cols="30" name="contact_text" class="textSize2 black" rows="5"></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="submit" value=" Submit " name="submit1"></td>
    </tr>
  </table>
  </form>

<% Else %>

<% 
response.redirect("http://www.jaljyoti.com/thankyou.asp")
%>

  <p><table cellspacing="2" cellpadding="2" border="0" width="530">
    <tr>
      <td align="right" width="140">First Name:</td>
      <td width="390"><%=contact_firstname%></td>
    </tr>
    <tr>
      <td align="right">Last Name:</td>
      <td><%=contact_lastname%></td>
    </tr>
    <tr>
      <td align="right">Email Address:</td>
      <td><%=contact_email%></td>
    </tr>
    <tr>
      <td align="right">Phone Number:</td>
      <td><%=contact_tel%></td>
    </tr>
    <tr>
      <td align="right">Mobile Number:</td>
      <td><%=contact_mobile%></td>
    </tr>
    <tr>
      <td align="right">Fax Number:</td>
      <td><%=contact_fax%></td>
    </tr>
    <tr>
      <td align="right">Skype ID:</td>
      <td><%=contact_skype%></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td align="right" valign="top" width="140">Text:</td>
      <td><%=contact_text%></td>
    </tr>
  </table>



<% End If %>

<%
' The code below supplied by Mark Short 

' returns string the can be written where you would like the reCAPTCHA challenged placed on your page 
function recaptcha_challenge_writer(publickey) 
  recaptcha_challenge_writer = "<script type=""text/javascript"">" & _ 
  "var RecaptchaOptions = {" & _ 
  " theme : 'white'," & _ 
  " tabindex : 0" & _ 
  "};" & _ 
  "</script>" & _ 
  "<script type=""text/javascript"" src=""http://api.recaptcha.net/challenge?k=" & publickey & """></script>" & _ 
  "<noscript>" & _ 
  "<iframe src=""http://api.recaptcha.net/noscript?k=" & publickey & """ frameborder=""1""></iframe><br>" & _ 
  "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & _ 
  "<input type=""hidden"" name=""recaptcha_response_field"" value=""manual_challenge"">" & _ 
  "</noscript>" 
end function 

function recaptcha_confirm(privkey,rechallenge,reresponse) 
  ' Test the captcha field 
  Dim VarString 
  VarString = _ 
  "privatekey=" & privkey & _ 
  "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _ 
  "&challenge=" & rechallenge & _ 
  "&response=" & reresponse 
  Dim objXmlHttp 
  Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP") 
  objXmlHttp.open "POST", "http://api-verify.recaptcha.net/verify", False 
  objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
  objXmlHttp.send VarString 
  Dim ResponseString 
  ResponseString = split(objXmlHttp.responseText, vblf) 
  Set objXmlHttp = Nothing 
  if ResponseString(0) = "true" then 
    ' They answered correctly 
    recaptcha_confirm = "" 
  else 
    ' They answered incorrectly 
    recaptcha_confirm = ResponseString(1) 
  end if 
end function 
%>

</body>
</html>

