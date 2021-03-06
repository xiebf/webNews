<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/connnews.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_connnews_STRING
Recordset1.Source = "SELECT * FROM newstype"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Re1
Dim Re1_numRows
keyword= request("keyword")
Set Re1 = Server.CreateObject("ADODB.Recordset")
Re1.ActiveConnection = MM_connnews_STRING
Re1.Source = "SELECT *  FROM news  WHERE news_title like '%"&keyword&"%'  ORDER BY news_id DESC"
Re1.CursorType = 0
Re1.CursorLocation = 2
Re1.LockType = 1
Re1.Open()

Re1_numRows = 0
%>
<%

Dim Recordset1_cmd

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_connnews_STRING
Recordset1_cmd.CommandText = "SELECT * FROM newstype" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%

Dim Re1_cmd

keyword= request("keyword") 
Set Re1_cmd = Server.CreateObject ("ADODB.Command")
Re1_cmd.ActiveConnection = MM_connnews_STRING
Re1_cmd.CommandText = "SELECT * FROM news WHERE  news_title  like '%"&keyword&"%' ORDER BY news_id DESC" 
Re1_cmd.Prepared = true

Set Re1 = Re1_cmd.Execute
Re1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 10
Repeat2__index = 0
Re1_numRows = Re1_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Re1_total
Dim Re1_first
Dim Re1_last

' set the record count
Re1_total = Re1.RecordCount

' set the number of rows displayed on this page
If (Re1_numRows < 0) Then
  Re1_numRows = Re1_total
Elseif (Re1_numRows = 0) Then
  Re1_numRows = 1
End If

' set the first and last displayed record
Re1_first = 1
Re1_last  = Re1_first + Re1_numRows - 1

' if we have the correct record count, check the other stats
If (Re1_total <> -1) Then
  If (Re1_first > Re1_total) Then
    Re1_first = Re1_total
  End If
  If (Re1_last > Re1_total) Then
    Re1_last = Re1_total
  End If
  If (Re1_numRows > Re1_total) Then
    Re1_numRows = Re1_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Re1_total = -1) Then

  ' count the total records by iterating through the recordset
  Re1_total=0
  While (Not Re1.EOF)
    Re1_total = Re1_total + 1
    Re1.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Re1.CursorType > 0) Then
    Re1.MoveFirst
  Else
    Re1.Requery
  End If

  ' set the number of rows displayed on this page
  If (Re1_numRows < 0 Or Re1_numRows > Re1_total) Then
    Re1_numRows = Re1_total
  End If

  ' set the first and last displayed record
  Re1_first = 1
  Re1_last = Re1_first + Re1_numRows - 1
  
  If (Re1_first > Re1_total) Then
    Re1_first = Re1_total
  End If
  If (Re1_last > Re1_total) Then
    Re1_last = Re1_total
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

Set MM_rs    = Re1
MM_rsCount   = Re1_total
MM_size      = Re1_numRows
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
Re1_first = MM_offset + 1
Re1_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Re1_first > MM_rsCount) Then
    Re1_first = MM_rsCount
  End If
  If (Re1_last > MM_rsCount) Then
    Re1_last = MM_rsCount
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>新闻首页</title>
<style type="text/css">
body,td,th {
	font-size: 13px;
}
</style>
</head>

<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><a href="index.asp">首页</a> &gt; 公司新闻</td>
  </tr>
  <tr>
    <td width="202" background="images/images_02.gif" valign="top" ><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="60">&nbsp;</td>
      </tr>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td height="25"><a href="type.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "type_id=" & Recordset1.Fields.Item("type_id").Value %>"><%=(Recordset1.Fields.Item("type_name").Value)%></a></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
    </table></td>
    <td width="778"><img src="images/images_03.gif" width="777" height="207" /></td>
  </tr>
  <tr>
    <td width="202" background="images/images_04.gif"  valign="top"><form id="form1" name="form1" method="post" action="">
        <p>主题：
          <label for="keyword"></label>
        <input name="keyword" type="text" id="keyword" size="12" />
        <input type="submit" name="button" id="button" value="查询" />
        </p>
      </form>
    <p>&nbsp;</p></td>
    <td width="778" height="428" background="images/images_05.gif" valign="top"><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0">
      <% 
While ((Repeat2__numRows <> 0) AND (NOT Re1.EOF)) 
%>
        <tr>
          <td width="78%" height="25"><%=(Re1.Fields.Item("news_title").Value)%></td>
          <td width="22%"><%=(Re1.Fields.Item("news_date").Value)%></td>
        </tr>
        <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  Re1.MoveNext()
Wend
%>
<tr>
  <td height="25"><table border="0">
      <tr>
        <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>">第一页</a>
            <% End If ' end MM_offset <> 0 %></td>
        <td><% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>">前一页</a>
            <% End If ' end MM_offset <> 0 %></td>
        <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>">下一个</a>
            <% End If ' end Not MM_atTotal %></td>
        <td><% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>">最后一页</a>
            <% End If ' end Not MM_atTotal %></td>
      </tr>
    </table></td>
  <td>&nbsp;
记录 <%=(Re1_first)%> 到 <%=(Re1_last)%> (总共 <%=(Re1_total)%>)</td>
</tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2"><div style="background-color: #0066CC;height:25px;line-height:25px;"><font color="white">&nbsp;@版权所有</font><font color="white" style="float:right"><a>关于我们</a>&nbsp; | <a>网络广告</a> &nbsp;|<a>技术支持</a>&nbsp;</font></div></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Re1.Close()
Set Re1 = Nothing
%>
