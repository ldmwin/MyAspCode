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

Dim Parent_ID,Site_ID,Parent_Column,Site_Name

Parent_ID=request("Parent_ID")

if Parent_ID="" or not isnumeric(Parent_ID) then
	response.write "参数出错"
	response.end()
elseif Parent_ID=0 then
	Site_ID=request("Site_ID")
	
	if Site_ID=""  or not isnumeric(Site_ID) then
		response.write "参数出错"
		response.end()
	else
	Site_ID=cint(Site_ID)

	end if
end if
Parent_ID=cint(Parent_ID)


if Parent_ID<>0 then

	Dim rs_Column : Set rs_Column = db.getRecordBySQL("select * from Site_Columns where status<>4 and id = " & Parent_ID)
	
	if rs_Column.eof or rs_Column.bof then
		response.write "数据查询失败"
		response.End()
	else
		Site_ID = rs_Column("Site_ID")
		Parent_Column = rs_Column("Column_Name")
		
	end if
	
	db.C(rs_Column)

else
	Parent_Column = "无"
	
end if



'Dim rs_Site : Set rs_Site = db.getRecordBySQL("select *,Site=Site_Name + '(V' + Version + ')' from Sys_Sites where status<>4 and id = " & Site_ID)
Dim rs_Site : Set rs_Site = db.getRecordBySQL("select * from Sys_Sites where status<>4 and id = " & Site_ID)

if rs_Site.eof or rs_Site.bof then
	response.write "数据查询失败"
	response.End()
else
	Site_Name = rs_Site("Site_Name") & "(V " & rs_Site("Version") & ")"
end if

db.C(rs_Site)

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
<form name="ColumnAdd" method="post" action="ColumnSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add"><input type="hidden" name="Parent_id" value="<%=Parent_ID%>"><input type="hidden" name="Site_ID" value="<%=Site_ID%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建栏目</th>
		</tr>
		<tr class="tr1">
			<td width="30%">父级栏目：</td>
			<td width="70%"><%=Parent_Column%>&nbsp;</td>
		</tr>
		<tr class="tr2">
		  <td>所属站点：</td>
		  <td><%=Site_Name%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>栏目名称：</td>
		  <td><input name="Column_Name" type="text" id="Column_Name" size="50"/> 
	      *</td>
	  </tr>
		<tr class="tr2">
		  <td>栏目分类：</td>
		  <td><%
		Set rs = db.getRecordBySQL("select * from Site_Column_Type where status=1 order by show_order desc,id desc") %>
			  <select name="Column_Type" id="Column_Type">
				<%
			  if not(rs.eof and rs.bof) then
			  do until rs.eof
			  response.Write("<option value="""&rs("id")&""">&nbsp;"&rs("Type_Name")&"&nbsp;</option>")
			  rs.movenext
			  loop
			  end if
			  db.C(rs)
			  %>
			  </select>
		*&nbsp;		</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">链接地址：</td>
			<td width="70%"><input name="Link" type="text" id="Link" size="50" value="#"/> 
			  *</td>
	  </tr>		
		
		<tr class="tr2">
		  <td>打开位置：</td>
		  <td>
		  <%
			Set rs = db.getRecordBySQL("select * from Sys_Target where status=1 and (site_id=0 or site_id=" & Site_ID & ") order by show_order desc,id desc") %>
			  <select name="Target" id="Target">
				<%
			  if not(rs.eof and rs.bof) then
			  do until rs.eof
			  response.Write("<option value="""&rs("id")&""">&nbsp;"&rs("Target_Name")&"&nbsp;</option>")
			  rs.movenext
			  loop
			  end if
			  db.C(rs)
			  %>
			  </select>
			*&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>显示顺序：</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" value="0" size="10" />
*</td>
	  </tr>
  </table>
</form>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>