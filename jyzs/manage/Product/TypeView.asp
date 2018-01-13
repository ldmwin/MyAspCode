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

dim Infos_id

Infos_id=request.QueryString("id")

if Infos_id="" or Infos_id=0 or not isnumeric(Infos_id) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Infos_id=cint(Infos_id)

Dim rs_Type : Set rs_Type = db.getRecordBySQL("select *,(select Type_Name from Product_Type  a where a.id=c.Parent_ID) as P_name from Product_Type c where status<>4 and id = " & Infos_id)

if rs_Type.eof or rs_Type.bof then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
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
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input name="edit" type="button" class="button" id="edit" style="margin-left:20px;" value="编辑" onClick="GotoUrl('TypeEdit.asp?id=<%=rs_Type("id")%>');"/>
				<input name="flooradd" type="button" class="button" id="flooradd" style="margin-left:20px;" value="添加子分类" onClick="GotoUrl('TypeAdd.asp?P_id=<%=rs_Type("id")%>');"/>
				<input type="button" name="return" class="button" value="添加同级分类" style="margin-left:20px;" onClick="GotoUrl('TypeAdd.asp?P_id=<%=rs_Type("Parent_id")%>');"/>
				<input type="button" name="cancel" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('TypeSave.asp?action=cancel&id=<%=rs_Type("id")%>&p_id=<%=rs_Type("Parent_id")%>','_self','确定作废？');"/>
				<%if rs_Type("parent_id")<>0 then%>
				<input type="button" name="return" class="button" value="返回上层" style="margin-left:20px;" onClick="GotoUrl('TypeView.asp?id=<%=rs_Type("Parent_id")%>');"/>
				<%end if%>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">分类详情</th>
		</tr>
		<tr class="tr1">
			<td width="30%">上级分类：</td>
			<td width="70%"><%if rs_Type("P_Name")="" then%>
			无
			<%else%>
			<%=rs_Type("P_Name")%>
			<%
			end if			
			%>&nbsp;
			</td>
		</tr>
		<tr class="tr2">
			<td width="30%">分类名称：</td>
			<td width="70%"><%=rs_Type("Type_Name")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">层级：</td>
			<td width="70%"><%=rs_Type("Type_Level")%>&nbsp;</td>
		</tr>
		
	  		<tr class="tr2">
		<td>显示顺序：</td>
		  <td><%=rs_Type("Show_Order")%></td>
	  </tr>
	  <tr class="tr1">
		  <td>状态：</td>
		  <td><%=StatusResult(1,rs_Type("status"))%></td>
	  </tr>
  </table>
	<br />
  <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th colspan="4" style="text-align:center;">子分类列表</th>
		</tr>
		<tr class="tr2">
			<td width="40%" align="center">分类名称</td>
			<td width="15%" align="center">显示顺序</td>
		    <td width="15%" align="center">状态</td>
		    <td width="30%" align="center">操作</td>
		</tr>
		<%Dim rs_childType
	Set rs_childType = db.getRecordBySQL("select *,(select Type_name from Product_Type a where a.Parent_ID=c.id) as P_name from Product_Type c where status<>4 and Parent_id=" & rs_Type("id") & " order by show_order desc,id desc") 
	
	if not (rs_childType.eof or rs_childType.bof) then
	
	do while not rs_childType.eof%>
		<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		  <td align="center"><a href="TypeView.asp?id=<%=rs_childType("id")%>"><%=rs_childType("Type_name")%></a></td>
			<td align="center"><%=rs_childType("show_order")%></td>
	        <td align="center"><%=StatusResult(1,rs_childType("status"))%></td>
	        <td align="left">
			<input type="button" name="edit" class="button" value="编辑" style="margin-left:5px;" onClick="GotoUrl('TypeEdit.asp?id=<%=rs_childType("id")%>');"/>
			
			<input type="button" name="cancel" class="button" value="作废" style="margin-left:5px;" onClick="GotoUrl('TypeSave.asp?action=cancel&id=<%=rs_childType("id")%>&p_id=<%=rs_Type("id")%>','_self','确定作废？');"/></td>
	  </tr>
	  <%
		  rs_childType.movenext()
			
		Loop
		
		
		end if
		
		db.C(rs_childType)
		  
	  %>
  </table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.C(rs_Type)

	db.CloseConn()
	
	set db=nothing
%>