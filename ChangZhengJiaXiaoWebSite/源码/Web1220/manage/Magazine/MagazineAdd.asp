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
<form name="MagazineAdd" method="post" action="MagazineSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('MagazineManage.asp');"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/><input type="hidden" name="action" value="add"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">新建杂志/DM</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题：</td>
			<td width="70%"><input name="Magazine_Name" type="text" id="Magazine_Name" size="50" />*</td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">杂志/DM类型：</td>
			<td width="70%">
			<%
			Set rs = db.getRecordBySQL("select * from Magazine_Type where status = 1 and Type_Level = 2 order by show_order desc,id desc") %>
		  <select name="Magazine_Type" id="Magazine_Type">
			<%
		  if not(rs.eof and rs.bof) then
		  do until rs.eof
		  response.Write("<option value=" & rs("id") & ">&nbsp;" & rs("Type_Name") & "&nbsp;</option>")
		  rs.movenext
		  loop
		  end if
		  db.C(rs)
		  %>
		  </select>*</td>
		</tr>
		
		<tr class="tr1">
		  <td>期数：</td>
		  <td><input name="Magazine_Period" type="text" id="Magazine_Period" size="50" />*</td>
	  </tr>
		<tr class="tr2">
		  <td>有效期：</td>
		  <td><input name="Magazine_Validity" type="text" id="Magazine_Validity" size="50" /></td>
	  </tr>
		<tr class="tr1">
			<td width="30%">封面：</td>
			<td width="70%"><input name="Pic_View" type="text" id="Pic_View" size="50"  />
		    &nbsp;&nbsp;<div id="PicOperate1" style="display:none;">&nbsp;&nbsp;<div  id="ReUpload1" onClick="reupload('../inc/UploadPhotos.asp?savepath=Magazine&picshow=ViewPicShow1&picsize=1024&picheight=0&picwidth=0&picurl=Pic_View&PicOperate=PicOperate1','PicOperate1','Pic_View','UploadFiles1');" style="cursor:hand;display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>  <div id="ViewShow1" style="cursor:hand;display:inline;" onMouseOver="PicView(0,'ViewPicShow1',this);" onMouseOut="PicView(1,'ViewPicShow1',this);">预览图片</div></div><img name="ViewPicShow1" id="ViewPicShow1" src="" alt="" style="display:none;"/></td>
		</tr>
		<tr class="tr1">
			<td width="30%">&nbsp;</td>
			<td width="70%"><iframe id="UploadFiles1" src="../inc/UploadPhotos.asp?savepath=Magazine&picshow=ViewPicShow1&picsize=1024&picheight=0&picwidth=0&picurl=Pic_View&PicOperate=PicOperate1" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe></td>
	  </tr>
		<tr class="tr2">
		  <td>简介：</td>
		  <td><textarea name="Brief" cols="60" rows="6" id="Brief"></textarea></td>
	  </tr>
		
		<tr class="tr1">
		  <td>显示顺序：</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" value="0" size="10" />*</td>
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