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
db.dbConnStr = Config.ConnStr(0)
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

Info_ID=request("ID")

if Info_ID="" or Info_ID=0 or not isnumeric(Info_ID) then
	response.write "参数出错"
	response.end()
end if
Info_ID=cint(Info_ID)

Dim rs_Magazine : Set rs_Magazine = db.getRecordBySQL("select * from Magazines where status<>4 and id = " & Info_id)

if rs_Magazine.eof or rs_Magazine.bof then
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
<form name="PageAdd" method="post" action="PageSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="add"><input type="hidden" name="Magazine_id" value="<%=rs_Magazine("id")%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建分页</th>
		</tr>
		<tr class="tr1">
			<td width="30%">所属杂志/DM：</td>
			<td width="70%"><%=rs_Magazine("Magazine_Name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">分页标题：</td>
			<td width="70%"><input name="Title" type="text" id="Title" size="50" />*</td>
	  </tr>		
		<tr class="tr1">
		  <td>图片：</td>
		  <td><input name="Page" type="text" id="Page" size="50"  />
		    &nbsp;&nbsp;&nbsp;&nbsp;<div id="PicOperate2" style="display:none;">&nbsp;&nbsp;<div  id="ReUpload2" onClick="reupload('../inc/UploadPhotos.asp?savepath=Magazine&picshow=ViewPicShow2&picsize=1000&picheight=0&picwidth=0&picurl=Page&PicOperate=PicOperate2','PicOperate2','Page','UploadFiles2');" style="cursor:hand; display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>&nbsp;&nbsp;<div id="ViewShow1" style="cursor:hand; display:inline;" onMouseOver="PicView(0,'ViewPicShow2',this);" onMouseOut="PicView(1,'ViewPicShow2',this);">预览图片</div></div><img name="ViewPicShow2" id="ViewPicShow2" src="" width="" height="" alt="" style="display:none;"/></td>
	  </tr>
		<tr class="tr1">
		  <td>&nbsp;</td>
		  <td><iframe id="UploadFiles2" src="../inc/UploadPhotos.asp?savepath=Magazine&picshow=ViewPicShow2&picsize=1000&picheight=0&picwidth=0&picurl=Page&PicOperate=PicOperate2" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe></td>
	  </tr>
		<tr class="tr2">
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
	
	db.C(rs_Magazine)

	db.CloseConn()
	
	set db=nothing
%>