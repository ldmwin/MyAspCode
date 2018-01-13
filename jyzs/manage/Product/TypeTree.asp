<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>  
<%
'-----------------------------------
'文 件 名 : 
'功    能 : 左侧导航，调用树状结构除最高一级以外的级别，由top.asp传入参数，默认载入系统管理 0
'作    者 : Mr.Lion
'建立时间 : 2011/05/12
'页面权限： system:login
'-----------------------------------
%>
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

Function BuildXMLStr(pid,str) '递归类别及其子类别存入字符串
	Dim rs_menutree,tempStr,i
	Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Product_Type s where s.Parent_id=m.id and s.status<>4) as childnodenum from Product_Type m where m.status<>4 and m.parent_id=" & pid & " order by m.show_order desc,m.id desc") 
	i  = 0
	do while not rs_menutree.eof
		if i = 0 then
			'if pid=0 then
			'str = str & "<ul id=""tree"">" & vbcrlf
			'else
			str = str & "<ul>" & vbcrlf
			'end if
		end if

		if rs_menutree("childnodenum") > 0 then
			'有子目录

				str = str & "<li><a href=""TypeView.asp?id=" & rs_menutree("id") & """ target=""Typemagright"">" & rs_menutree("Type_Name") & "</a>" & vbcrlf

		else
			'无子目录
			str = str & "<li class=""Child""><a href=""TypeView.asp?id=" & rs_menutree("id") & """ target=""Typemagright"">" & rs_menutree("Type_Name") & "</a>" & vbcrlf
		end if
		Call BuildXMLStr(rs_menutree("ID"),str) '递归调用
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>无标题文档</title>
<link rel="stylesheet" href="../css/jquery.treeview.css" />
 
<script src="../js/jquery.js" type="text/javascript"></script>
<script src="../js/jquery.cookie.js" type="text/javascript"></script>
<script src="../js/jquery.treeview.js" type="text/javascript"></script>
 
<script type="text/javascript">
		$(function() {
			$("#tree").treeview({
				collapsed: false,
				animated: "medium",
				control:"#sidetreecontrol",
				persist: "location"
			});
		})
		
	</script>
<style>
body{
	background-color:#C9DEFA; 
	font:normal 12px ;
	SCROLLBAR-FACE-COLOR: #72A3D0; SCROLLBAR-HIGHLIGHT-COLOR: #337ABB; 
	SCROLLBAR-SHADOW-COLOR: #337ABB; SCROLLBAR-DARKSHADOW-COLOR: #337ABB; 
	SCROLLBAR-3DLIGHT-COLOR: #337ABB; SCROLLBAR-ARROW-COLOR: #FFFFFF;
	SCROLLBAR-TRACK-COLOR: #337EC0; 
}
a { color:#135294; text-decoration: none; }
a:hover { color: #ff6600; text-decoration: underline; }
</style>
</head>
 
<body>
<div style="height:25px; text-align:left; padding-left:10px" id="sidetreecontrol"><a href="#" onClick="window.location.reload();">刷新</a> | <a href="?#">展开全部</a> | <a href="?#">关闭全部</a></div>

<div>
			<ul id="tree">
			<li><span><strong><a href="TypeList.asp" target="Typemagright">分类</a></strong></span>
 			<%response.Write(BuildXMLStr(0,str))%>
			</li>
			</ul>
</div>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>