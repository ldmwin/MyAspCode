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
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript" src="../JS/My97DatePicker/WdatePicker.js"></script>
</head>

<body>
<form name="InfoAdd" method="post" action="InfoSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('InfoManage.asp');"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add">			</td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建新闻</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题：</td>
			<td width="70%"><input name="Title" type="text" id="Title" size="50" />
			*</td>
		</tr>
		
		<tr class="tr2">
		  <td>副标题：</td>
		  <td><input name="Sub_Title" type="text" id="Sub_Title" size="50" /></td>
	  </tr>
		<tr class="tr1">
			<td width="30%">新闻类型：</td>
			<td width="70%">
			<%
			Set rs = db.getRecordBySQL("select * from Information_Type where status <> 4 and Type_Level = 2 order by show_order desc,id desc") %>
		  <select name="Info_Type" id="Info_Type">
			<%
		  if not(rs.eof and rs.bof) then
		  do until rs.eof
		  response.Write("<option value=" & rs("id") & ">&nbsp;" & rs("Type_Name") & "&nbsp;</option>")
		  rs.movenext
		  loop
		  end if
		  db.C(rs)
		  %>
		  </select>*			</td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">缩略图：</td>
			<td width="70%"><input name="Pic_View" type="text" id="Pic_View" size="50"  />
		    &nbsp;<input type="button" name="Submit2" value="上传图片" onClick="window.open('../inc/upload_flash.asp?formname=InfoAdd&editname=Pic_View&uppath=../../Pictures&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=105')"></td>
		</tr>
		<tr class="tr1">
		  <td>摘要：</td>
		  <td><textarea name="Brief" cols="60" rows="6" id="Brief"></textarea></td>
	  </tr>
		<tr class="tr2">
			<td width="30%">详细介绍：Info_style</td>
			<td width="70%"><input type="hidden" name="Content"><iframe id="Content" src="../editor/eWebEditor.asp?id=Content&style=s_blue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
		</tr>
		<tr class="tr1">
		  <td>显示顺序</td>
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