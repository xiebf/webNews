<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/connnews.asp" -->
<%
Dim Re
Dim Re_numRows

Set Re = Server.CreateObject("ADODB.Recordset")
Re.ActiveConnection = MM_connnews_STRING
Re.Source = "SELECT * FROM news ORDER BY news_id DESC"
Re.CursorType = 0
Re.CursorLocation = 2
Re.LockType = 1
Re.Open()

Re_numRows = 0
%>
<%
Dim Re1
Dim Re1_numRows

Set Re1 = Server.CreateObject("ADODB.Recordset")
Re1.ActiveConnection = MM_connnews_STRING
Re1.Source = "SELECT * FROM newstype"
Re1.CursorType = 0
Re1.CursorLocation = 2
Re1.LockType = 1
Re1.Open()

Re1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Re1_numRows = Re1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
Re_numRows = Re_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Re_total
Dim Re_first
Dim Re_last

' set the record count
Re_total = Re.RecordCount

' set the number of rows displayed on this page
If (Re_numRows < 0) Then
  Re_numRows = Re_total
Elseif (Re_numRows = 0) Then
  Re_numRows = 1
End If

' set the first and last displayed record
Re_first = 1
Re_last  = Re_first + Re_numRows - 1

' if we have the correct record count, check the other stats
If (Re_total <> -1) Then
  If (Re_first > Re_total) Then
    Re_first = Re_total
  End If
  If (Re_last > Re_total) Then
    Re_last = Re_total
  End If
  If (Re_numRows > Re_total) Then
    Re_numRows = Re_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Re
MM_rsCount   = Re_total
MM_size      = Re_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Re_first = MM_offset + 1
Re_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Re_first > MM_rsCount) Then
    Re_first = MM_rsCount
  End If
  If (Re_last > MM_rsCount) Then
    Re_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理中心</title>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	background-image: url(../images/bg.gif);
}
.style18 {color: #FFFF00}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
	color: #000000;
}
a:hover {
	text-decoration: none;
	color: #006600;
}
a:active {
	text-decoration: none;
}
#Layer1 {
	position:absolute;
	width:200px;
	height:115px;
	z-index:1;
	left: 142px;
	top: 592px;
}
.STYLE23 {font-size: 14px}
.STYLE24 {font-size: 12px}
.1 {
	font-size: 13px;
	color: #900;
}
body,td,th {
	font-size: 13px;
}
-->
</style></head>

<body>
<table width="768" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div style="height:20px"></div>
	<div style="background-color: #0066CC;height:40px;line-height:40px;"><font color="white" size="4"><strong>&nbsp;新闻系统后台管理中心</strong></font></div></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
      <tr>
        <td width="266" height="335" bgcolor="#FFFFFF"><table width="95%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="222" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    
                    <tr>
                      <td height="30">&nbsp;&nbsp; <span class="1">管理员你好！请你添加新闻信息</span></td>
                    </tr>
                    <tr>
                      <td height="25">&nbsp;&nbsp;&nbsp;<a href="/admin/news_add.asp">添加新闻</a></td>
                    </tr>
                    <tr>
                      <td height="25">&nbsp;&nbsp; <a href="/admin/type_add.asp">添加新闻分类</a></td>
                    </tr>
                  </table>
                <p>&nbsp;&nbsp; <span class="1">管理员你好！请你管理新闻分类！</span></p>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="53%" height="20">&nbsp; &nbsp;&nbsp;&nbsp; 类型 </td>
                      <td width="47%">&nbsp; 管理 </td>
                    </tr>
                  </table>
                
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <% 
While ((Repeat1__numRows <> 0) AND (NOT Re1.EOF)) 
%>
                      <tr>
                        <td height="20">&nbsp;&nbsp; <%=(Re1.Fields.Item("type_name").Value)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; [<A HREF="type_upd.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "type_id=" & Re1.Fields.Item("type_id").Value %>">修改</A>] [<A HREF="type_del.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "type_id=" & Re1.Fields.Item("type_id").Value %>">删除</A>] </td>
                      </tr>
                      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Re1.MoveNext()
Wend
%>
<tr>
                      <td><img src="/images/obj_ta_07.gif" width="200" height="1" /></td>
                    </tr>
                  </table>
                  <p>&nbsp;</p></td>
            </tr>
        </table></td>
        <td width="496" valign="top" bgcolor="#FFFFFF"><p>&nbsp;</p>
          <table width="100%" height="27" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="21"><table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td><table width="99%" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="24"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <% 
While ((Repeat2__numRows <> 0) AND (NOT Re.EOF)) 
%>
                            <tr>
                              <td width="70%" height="25">&nbsp;<a href="../newscontent.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "news_id=" & Re.Fields.Item("news_id").Value %>"><%=(Re.Fields.Item("news_title").Value)%></a></td>
                              <td width="30%"><div align="center">[<A HREF="news_upd.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "news_id=" & Re.Fields.Item("news_id").Value %>">修改</A>][<A HREF="news_del.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "news_id=" & Re.Fields.Item("news_id").Value %>">删除</A>] </div></td>
                            </tr>
                            <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  Re.MoveNext()
Wend
%>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="9"><img src="../images/obj_ta_07.gif" width="500" height="1" /></td>
                </tr>
              </table></td>
            </tr>
          </table>
          <table border="0" align="center">
            <tr>
              <td width="108"><% If MM_offset <> 0 Then %>
                  <a href="<%=MM_moveFirst%>">第一页</a>
                  <% End If ' end MM_offset <> 0 %></td>
              <td width="110"><% If MM_offset <> 0 Then %>
                  <a href="<%=MM_movePrev%>">前一页</a>
                  <% End If ' end MM_offset <> 0 %></td>
              <td width="86"><% If Not MM_atTotal Then %>
                  <a href="<%=MM_moveNext%>">下一个</a>
                  <% End If ' end Not MM_atTotal %></td>
              <td width="77"><% If Not MM_atTotal Then %>
                  <a href="<%=MM_moveLast%>">最后一页</a>
                  <% End If ' end Not MM_atTotal %></td>
              </tr>
          </table></td>
        <td width="6">&nbsp;</td>
      </tr>
      
    </table></td>
  </tr>
</table>
</body>
</html>
<%
Re.Close()
Set Re = Nothing
%>
<%
Re1.Close()
Set Re1 = Nothing
%>
