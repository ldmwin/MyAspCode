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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>无标题文档</title>
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
</head>

<body>
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="5">
				<input type="button" name="add" value="新建一级分类" class="button" style="margin-left:20px;"  onClick="GotoUrl('TypeAdd.asp?p_id=0');"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>			</td>
		</tr>
		<tr>
			<th colspan="5" style="text-align:center;">一级分类列表</th>
		</tr>
		<tr class="tr2">
			<td width="30%" align="center">分类名称</td>
			<td width="15%" align="center">显示顺序</td>
		    <td width="15%" align="center">状态</td>
		    <td width="20%" align="center">下级分类</td>
		    <td width="20%" align="center">操作</td>
		</tr>
		<%Dim rs_Type,sqlstr
		'sqlstr = "select *,(select count(*) from sys_floor where status=1 and sys_floor.store_id=sys_store.id) as upchildfloornum,(select count(*) from sys_floor where status=0 and sys_floor.store_id=sys_store.id) as downchildfloornum,(select org_name from Sys_Organization where sys_store.SubCompany_ID=Sys_Organization.id) as subcompany,statusshow=(case status when 0 then '下线' when 1 then '上线' end) from sys_store order by show_order asc"
		sqlstr = "select *,(select count(*) from Product_Type a where a.status=1 and a.Parent_id=Product_Type.id) as upchildfloornum,(select count(*) from Product_Type b where b.status=0 and b.parent_id=Product_Type.id) as downchildfloornum from Product_Type where parent_id=0 and status<>4 order by show_order desc,id desc"
	Set rs_Type = db.getRecordBySQL(sqlstr) 
	
	'response.Write(rs_Type.recordcount&"yes")
	
	if not (rs_Type.eof or rs_Type.bof) then
	
	do while not rs_Type.eof%>
		<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		  <td align="center"><a href="TypeView.asp?id=<%=rs_Type("id")%>"><%=rs_Type("Type_Name")%></a></td>
			<td align="center"><%=rs_Type("show_order")%></td>
	        <td align="center"><%=StatusResult(1,rs_Type("status"))%></td>
	        <td align="center">上线：<%=rs_Type("upchildfloornum")%>/下线：<%=rs_Type("downchildfloornum")%></td>
	        <td align="left"><input type="button" name="edit" class="button" value="编辑" style="margin-left:5px;" onClick="GotoUrl('TypeEdit.asp?id=<%=rs_Type("id")%>');"/>
			<input type="button" name="cancel" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('TypeSave.asp?action=cancel&id=<%=rs_Type("id")%>&p_id=<%=rs_Type("Parent_id")%>','_self','确定作废？');"/></td>
	  </tr>
	  <%
		  rs_Type.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_Type)
		  
	  %>
  </table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.CloseConn()
	
	set db=nothing
%>