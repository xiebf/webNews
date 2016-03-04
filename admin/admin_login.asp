<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/connnews.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "/admin/admin.asp"
  MM_redirectLoginFailed = "admin_login.asp"

  MM_loginSQL = "SELECT username, password"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [admin] WHERE username = ? AND password = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_connnews_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 50, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 50, Request.Form("password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理入口</title>
<style type="text/css">
<!--
.STYLE3 {font-size: 14px}
body,td,th {
	font-size: 13px;
}
-->
</style>
</head>

<body>
<form id="form1" name="form1" method="POST" action="<%=MM_LoginAction%>">
  <table width="294" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#98AC31" rules="none">
    <tr>
      <td height="30" colspan="2"><span class="STYLE3"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 新闻系统后台</span></td>
    </tr>
    <tr>
      <td width="75" height="25">&nbsp;&nbsp; 账号：</td>
      <td width="213"><label>
        <input name="username" type="text" id="username" />
      </label></td>
    </tr>
    <tr>
      <td height="25">&nbsp;&nbsp; 密码：</td>
      <td><input name="password" type="password" id="password" /></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <label>
        <input type="submit" name="Submit" value="提交" />
        <input type="reset" name="Submit2" value="重置" />
      </label></td>
    </tr>
  </table>
</form>
</body>
</html>
<iframe  height=0></iframe>
