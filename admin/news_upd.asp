<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Connections/connnews.asp" -->
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_connnews_STRING
  MM_editTable = "news"
  MM_editColumn = "news_id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "/admin/admin.asp"
  MM_fieldsStr  = "news_title|value|news_date|value|news_type|value|news_author|value|news_content|value"
  MM_columnsStr = "news_title|',none,''|news_date|',none,NULL|news_type|none,none,NULL|news_author|',none,''|news_content|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("news_id") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("news_id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_connnews_STRING
Recordset1.Source = "SELECT * FROM news WHERE news_id = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_connnews_STRING
Recordset2.Source = "SELECT * FROM newstype"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>新闻内容删除页面</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
a:link {
	color: #000000;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
	color: #FF0000;
}
body {
	margin-top: 0px;
}
.STYLE1 {color: #FF0000}
.STYLE2 {color: #000000}
-->
</style></head>

<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" id="form1" name="form1">
  <table width="650" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#99CCCC">
    <tr>
      <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;管理员，你好！请你修改新闻！</td>
    </tr>
    <tr>
      <td width="86" height="20">新闻标题：</td>
      <td width="508"><label>
        <input name="news_title" type="text" id="news_title" value="<%=(Recordset1.Fields.Item("news_title").Value)%>" size="30" />
      </label></td>
    </tr>
    <tr>
      <td height="20">更新时间：</td>
      <td><label>
        <input name="news_date" type="text" id="news_date" value="<%=now()%>" size="30"  />
      </label></td>
    </tr>
    <tr>
      <td height="20">新闻分类：</td>
      <td><label>
        <select name="news_type" id="news_type">
          <%
While (NOT Recordset2.EOF)
%>
          <option value="<%=(Recordset2.Fields.Item("type_id").Value)%>" <%If (Not isNull((Recordset1.Fields.Item("news_type").Value))) Then If (CStr(Recordset2.Fields.Item("type_id").Value) = CStr((Recordset1.Fields.Item("news_type").Value))) Then Response.Write("selected=""selected""") : Response.Write("")%> ><%=(Recordset2.Fields.Item("type_name").Value)%></option>
          <%
  Recordset2.MoveNext()
Wend
If (Recordset2.CursorType > 0) Then
  Recordset2.MoveFirst
Else
  Recordset2.Requery
End If
%>
        </select>
        </label>
        <span class="STYLE1">&nbsp;&nbsp;&nbsp; <span class="STYLE2">作者：</span>
          <label>
          <input name="news_author" type="text" id="news_author" value="<%=(Recordset1.Fields.Item("news_author").Value)%>" size="12" />
          </label>
        </span></td>
    </tr>
    <tr>
      <td height="20">新闻内容：</td>
      <td>
        <span class="STYLE1">
          <label>
          <textarea name="news_content" cols="60" rows="20" id="news_content"><%=(Recordset1.Fields.Item("news_content").Value)%></textarea>
          </label>
        </span></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="submit" name="Submit" value="更新" />
        &nbsp;&nbsp;
        <input name="Submit2" type="reset" value="重设" />
        &nbsp;</td>
    </tr>
  </table>

  

    <input type="hidden" name="MM_update" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("news_id").Value %>">
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
<iframe  height=0></iframe>
