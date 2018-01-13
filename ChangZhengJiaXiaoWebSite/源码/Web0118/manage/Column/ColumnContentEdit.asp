<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Class/UserCtrl.asp" -->
<!--#include file="../Class/LogCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

Dim Info_ID,Parent_Column,Site_Name

Info_ID=request("ID")

if Info_ID="" or not isnumeric(Info_ID) then
	response.write "参数出错"
	response.end()

end if
Info_ID=cint(Info_ID)


Dim rs_Column : Set rs_Column = db.getRecordBySQL("select *,(select Column_Name from Site_Columns where id =Parent_ID) as Parent_Column,(select Site_Name from Sys_Sites where id =Site_ID) as Site_Name,(select Type_Name from site_column_type where column_type = id) as Type_Name from Site_Columns where status<>4 and id = " & Info_ID)

if rs_Column.eof or rs_Column.bof then
	response.write "数据查询失败"
	response.End()	
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
</head>

<body>
<form name="ColumnContentEdit" method="post" action="ColumnSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="contentedit"><input type="hidden" name="Column_ID" value="<%=Info_ID%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">编辑内容</th>
		</tr>
		
		
		<tr class="tr1">
		  <td>栏目名称：</td>
		  <td><%=rs_Column("Column_Name")%></td>
	  </tr>
		<tr class="tr2">
		  <td>栏目分类：</td>
		  <td><%=rs_Column("Type_Name")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">使用子栏目内容：</td>
			<td width="70%">
			<input type="checkbox" name="UseChildContent" id="UseChildContent" value="1" <%if rs_Column("UseChildContent")=1 then
				response.Write(" checked")
			end if
			%> /></td>
	  </tr>		
		
		<tr class="tr2">
		  <td>对应模板：</td>
		  <td>
		  <%
			Set rs = db.getRecordBySQL("select * from Site_Column_Mould where status=1 and (site_id=0 or site_id=2) order by show_order desc,id desc") %>
			  <select name="Mould_Page" id="Mould_Page">
			  <option value="0" <%=object_selected(rs_column("Mould_Page"),0)%>>&nbsp;不使用模板&nbsp;</option>
				<%
			  if not(rs.eof and rs.bof) then
			  do until rs.eof
			  response.Write("<option value="""&rs("id")&""" " & object_selected(rs_column("Mould_Page"),rs("id")) & ">&nbsp;"&rs("Mould_Name")&"&nbsp;</option>")
			  rs.movenext
			  loop
			  end if
			  db.C(rs)
			  %>
			  </select>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>内容：</td>
		  <td><input type="hidden" name="Column_Content" value="<%if not isnull(rs_column("Column_Content")) then
		  	response.Write(Server.HTMLEncode(rs_column("Column_Content")))
		  end if%>"><iframe id="Column_Content" src="../editor/eWebEditor.asp?id=Column_Content&style=s_blue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
	  </tr>
  </table>
</form>
</body>
</html>
<%
	db.C(rs_Column)

	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>