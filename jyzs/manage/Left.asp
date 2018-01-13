<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>  
<%
'-----------------------------------
'文 件 名 : /Left.asp
'功    能 : 左侧导航，调用树状结构除最高一级以外的级别，由top.asp传入参数，默认载入系统管理 0
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

dim divload

sub LeftTopMenuLoad()
	Dim rs_ltm,tempStr,i
	'Set rs_ltm = db.getRecordBySQL("select * from sys_menu where (needrighttoview=0 or (needrighttoview=1 and dbo.Sys_UserToMenuNum(" & User.UserID & ",id)>0)) and status=1 and parent_id=0 order by show_order desc,id desc")
	Set rs_ltm = db.getRecordBySQL("select * from sys_menu where status=1 and parent_id=0 order by show_order desc,id desc") 

		if not(rs_ltm.eof or rs_ltm.bof) then
			
			do until rs_ltm.EOF 
			%>
			<div id="leftmenu<%=rs_ltm("id")%>" style="display:<%if rs_ltm("id")= 3 then response.Write("block") else response.Write("none") end if%>" name="leftmenu">

			<%Call LeftSecMenuLoad(rs_ltm("id"))%>
			</div>
			
			<%
				rs_ltm.MoveNext 
			loop
		
		end if
	
	db.C(rs_ltm)
		
end sub

sub LeftSecMenuLoad(ParentID)	
		
		Dim rs_lsm,tempStr,i
		Set rs_lsm = db.getRecordBySQL("select m.*,(select count(*) from Sys_menu s where s.Parent_id=m.id and s.status=1 and s.menu_level>2) as childnodenum from sys_menu m where m.status=1 and m.parent_id=" & ParentID & " and m.Menu_Level=2 order by m.show_order desc,m.id desc") 

			if not(rs_lsm.eof or rs_lsm.bof) then
				do until rs_lsm.EOF 
				%>
				
				<div class="left_color"><a href="<%=rs_lsm("Link")%>" target="<%=rs_lsm("Parent_Frame")%>" <%if rs_lsm("childnodenum") > 0 then response.Write("onClick=""divcontrol('CNLTreeMenu" & rs_lsm("ID") & "')""") end if%>><%=rs_lsm("Menu_Name")%></a></div>
				<%if rs_lsm("childnodenum") > 0 then%>
				<div class="CNLTreeMenu" id="CNLTreeMenu<%=rs_lsm("ID")%>">
				<%
				divload = divload & "var MyCNLTreeMenu" & rs_lsm("ID") & "=new CNLTreeMenu(""CNLTreeMenu" & rs_lsm("ID") & """,""li"");MyCNLTreeMenu" & rs_lsm("ID") & ".InitCss(""Opened"",""Closed"",""Child"",""images/s.gif"");"
				%>
				<%response.write BuildXMLStr(rs_lsm("ID"),str,User.UserID)%>
				</div>
				<%end if%>
				<%
					str=""
					rs_lsm.MoveNext 
                loop
			
			end if
		
		db.C(rs_lsm)
		
end sub

Function BuildXMLStr(pid,str,uid) '递归类别及其子类别存入字符串
	Dim rs_menutree,tempStr,i
	Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Sys_menu s where s.Parent_id=m.id and s.status=1 and s.menu_level>2) as childnodenum from sys_menu m where m.status=1 and m.parent_id=" & pid & " and m.Menu_Level>2 order by m.show_order desc,m.id desc") 
	i  = 0
	do while not rs_menutree.eof
		if i = 0 then
			str = str & "<ul>" & vbcrlf
		end if

		if rs_menutree("childnodenum") > 0 then
			'有子目录

				str = str & "<li><a href=""" & rs_menutree("Link") & """ target=""" & rs_menutree("Parent_Frame") & """>" & rs_menutree("Menu_Name") & "</a>" & vbcrlf

		else
			'无子目录
			str = str & "<li class=""Child""><a href=""" & rs_menutree("Link") & """ target=""" & rs_menutree("Parent_Frame") & """>" & rs_menutree("Menu_Name") & "</a>" & vbcrlf
		end if
		Call BuildXMLStr(rs_menutree("ID"),str,uid) '递归调用
		rs_menutree.movenext()
		i = i + 1
		str = str & "</li>" & vbcrlf
		if rs_menutree.eof then str = str & "</ul>" & vbcrlf
	Loop
	BuildXMLStr = str
	db.C(rs_menutree)

End Function
%> 
 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>网站管理系统</title>

<link rel="stylesheet" href="css/left.css" type="text/css" />
 <script type="text/javascript" src="js/function.js" language="javascript"></script>
</head>
 
<body>
 
 
 
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" style="padding-top:10px;" id="menubar">
			
 			<%Call LeftTopMenuLoad()%>
		</td>
	</tr>
</table>
 
 
<script type="text/javascript"> 

<%response.Write(divload)%>

</script>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>