<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>  
<%
'-----------------------------------
'文 件 名 : /Top.asp
'功    能 : 顶部导航，调用树状结构最高一级 parent_id=0 一级的
'作    者 : Mr.Lion
'建立时间 : 2011/05/12
'页面权限： system:login
'-----------------------------------
%>
<!--#include file="Class/DBCtrl.asp" -->
<!--#include file="Class/UserCtrl.asp" -->
<!--#include file="Class/LogCtrl.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"")
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

'导入一级菜单项
sub TopMenuLoad()

		
			Dim rs_tm,tempStr,i
			Set rs_tm = db.getRecordBySQL("select * from sys_menu where status=1 and parent_id=0 order by show_order desc,id desc") 
	
			if not(rs_tm.eof or rs_tm.bof) then
				do until rs_tm.EOF 
				%>
				<li><a href="top.asp" onmouseover="parent.frmleft.disp(<%=rs_tm("id")%>);" target="_self"><span><%=rs_tm("Menu_Name")%></span></a></li>
				<%
					rs_tm.MoveNext 
                loop
			
			end if

		
		db.C(rs_tm)

		
end sub

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>顶部导航</title>
<link href="css/top.css" rel="stylesheet" type="text/css" />

</head>

<body>
<div class="menu">
	<div class="system_logo"><img src="images/logo_up.gif"></div>
	<div id="tabs">
		<ul>
			<%
				Call TopMenuLoad()				
			%>
			</ul>
	</div>
	<div style="clear:both"></div>
</div>


</body>
</html>
<%
	set User = nothing
	set EventLog = nothing

	db.CloseConn()
	
	set db=nothing
%>