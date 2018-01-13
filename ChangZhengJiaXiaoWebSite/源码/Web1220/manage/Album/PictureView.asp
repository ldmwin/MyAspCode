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

Info_ID=request("ID")

if Info_ID="" or Info_ID=0 or not isnumeric(Info_ID) then
	response.write "参数出错"
	response.end()
end if
Info_ID=cint(Info_ID)

Dim rs_Picture : Set rs_Picture = db.getRecordBySQL("select *,(select Album_Name from Albums where Albums.id = Album_Pictures.Album_ID) as Album_Name from Album_Pictures where status<>4 and id = " & Info_id)

if rs_Picture.eof or rs_Picture.bof then
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
<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<%if rs_Picture("status")=0 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('PictureSave.asp?action=singleup&id=<%=rs_Picture("id")%>');"/><input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('PictureEdit.asp?id=<%=rs_Picture("id")%>');"/>
                <%elseif rs_Picture("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('PictureSave.asp?action=singledown&id=<%=rs_Picture("id")%>');"/>
			
			<%end if%>
            
            <input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">图片预览</th>
		</tr>
		<tr class="tr1">
			<td width="30%">所属图集：</td>
			<td width="70%"><%=rs_Picture("Album_Name")%></td>
		</tr>
		<tr class="tr2">
			<td width="30%">图片标题：</td>
			<td width="70%"><%=rs_Picture("Title")%></td>
	  </tr>		
		<tr class="tr1">
		  <td>图片：</td>
		  <td><%if rs_Picture("Picture")<>"" then%>
                <img src="<%=Config.ImgUrl()%>/<%response.Write(rs_Picture("Picture"))%>" />
                <%else%>
无
<%end if%></td>
  </tr>
		<tr class="tr2">
		  <td>状态：</td>
		  <td><%=StatusResult(3,rs_Picture("status"))%></td>
  </tr>
		<tr class="tr1">
		  <td>显示顺序：</td>
		  <td><%=rs_Picture("Show_Order")%></td>
  </tr>
</table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.C(rs_Picture)

	db.CloseConn()
	
	set db=nothing
%>