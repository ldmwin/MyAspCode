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

dim Info_ID

Info_ID=request.QueryString("id")

if Info_ID="" or Info_ID=0 or not isnumeric(Info_ID) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_ID=cint(Info_ID)


sqlstr = "select *,(select type_name from Album_Type where Album_Type.id=Albums.Album_Type) as Albumtypeshow from Albums where status<>4 and id = " & Info_ID

'response.Write(sqlstr)
'response.End()


Dim rs_Album : Set rs_Album = db.getRecordBySQL(sqlstr)

if rs_Album.eof or rs_Album.bof then
	response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
	response.End()
end if

picurl1= Config.ImgUrl() & "Album/" & rs_Album("Pic_View")		
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
<form name="AlbumEdit" method="post" action="AlbumSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('AlbumView.asp?id=<%=rs_Album("id")%>');"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="edit"><input type="hidden" name="Album_ID" value="<%=rs_Album("id")%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">编辑图集</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题：</td>
			<td width="70%"><input name="Album_Name" type="text" id="Album_Name" size="50" value="<%=rs_Album("Album_Name")%>" />
			*</td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">图集类型：</td>
			<td width="70%">
			<%
			Set rs = db.getRecordBySQL("select * from Album_Type where status = 1 and Type_Level = 2 order by show_order desc,id desc") %>
		  <select name="Album_Type" id="Album_Type">
			<%
		  if not(rs.eof and rs.bof) then
		  do until rs.eof
		  response.Write("<option value=""" & rs("id") & """ " & object_selected(rs_Album("Album_Type"),rs("id")) & ">&nbsp;" & rs("Type_Name") & "&nbsp;</option>")
		  rs.movenext
		  loop
		  end if
		  db.C(rs)
		  %>
		  </select>*			</td>
		</tr>
		
		<tr class="tr1">
			<td width="30%">封面：</td>
			<td width="70%"><input name="Pic_View" type="text" id="Pic_View" size="50" value="../../Pictures/<%=rs_Album("Pic_View")%>" />
		    &nbsp;<input type="button" name="Submit2" value="上传图片" onClick="window.open('../inc/upload_flash.asp?formname=AlbumEdit&editname=Pic_View&uppath=../../Pictures&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=105')"></td>
		</tr>
		
		<tr class="tr2">
		  <td>简介：</td>
		  <td><textarea name="Brief" cols="60" rows="6" id="Brief"><%=rs_Album("Brief")%></textarea></td>
	  </tr>
		
		<tr class="tr1">
		  <td>显示顺序：</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" value="<%=rs_Album("Show_Order")%>" size="10" />
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
	
	db.C(rs_Album)

	db.CloseConn()
	
	set db=nothing
%>