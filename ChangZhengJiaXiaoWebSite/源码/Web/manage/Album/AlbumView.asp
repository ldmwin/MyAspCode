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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript"></script>
<script src="../js/function.js" type="text/javascript"></script>
<script type="text/javascript" src="../js/wbox/jquery1.4.2.js"></script> 
<script type="text/javascript" src="../js/wbox/wbox.js"></script>
<link rel="stylesheet" type="text/css" href="../js/wbox/wbox/wbox.css" />
</head>

<body>
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<%if rs_Album("status")=0 or rs_Album("status")=2 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('AlbumSave.asp?action=singleup&id=<%=rs_Album("id")%>');"/><input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('AlbumEdit.asp?id=<%=rs_Album("id")%>');"/>
                <%elseif rs_Album("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('AlbumSave.asp?action=singledown&id=<%=rs_Album("id")%>');"/>
			
			<%end if%>
            <input type="button" name="cancel" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('AlbumSave.asp?action=singlecancel&id=<%=rs_Album("id")%>','_self','确定作废？');"/>
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('AlbumManage.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">图集预览</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题：</td>
			<td width="70%"><%=rs_Album("Album_Name")%></td>
		</tr>
		
		<tr class="tr1">
			<td width="30%">图集类型：</td>
			<td width="70%"><%=rs_Album("Albumtypeshow")%></td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">封面：</td>
			<td width="70%"><%if rs_Album("Pic_View")<>"" then%>
                <img src="<%=Config.ImgUrl()%>/<%response.Write(rs_Album("Pic_View"))%>" />
                <%else%>
无
<%end if%></td>
		</tr>
		<tr class="tr1">
		  <td>状态：</td>
		  <td><%=StatusResult(1,rs_Album("status"))%></td>
	  </tr>
	  <tr class="tr2">
		  <td>显示顺序：</td>
		  <td><%=rs_Album("Show_Order")%></td>
	  </tr>
		<tr class="tr1">
		  <td>简介：</td>
		  <td><%=rs_Album("Brief")%>&nbsp;</td>
	  </tr>
</table>
<br />
<form name="frm_list" method="post" action="<%=SubmitUrl%>">
  <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th style="text-align:center;">图片列表</th>
		</tr>
		<tr class="tr2">
		  <td  align="left">
		  <%if rs_Album("status")=0 or rs_Album("status")=2 then%>
		  <input name="Pictureadd" id="Pictureadd" type="button" class="button"style="margin-left:20px;" value="新建图片" />
		   <script type="text/javascript"> 
   				$("#Pictureadd").wBox({requestType:"iframe_refresh",iframeWH:{width:850,height:550},target:"PictureAdd.asp?id=<%=rs_Album("id")%>"}); 
			</script>
		  <%end if%>
		  <input type="button" name="up" class="button" value="上线图片" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PictureSave.asp?action=up&Album_ID=<%=rs_Album("id")%>','确认上线？');"/>
			<input type="button" name="down" class="button" value="下线图片" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PictureSave.asp?action=down&Album_ID=<%=rs_Album("id")%>','确认下线？');"/>
			<input type="button" name="cancel" class="button" value="作废图片" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PictureSave.asp?action=cancel&Album_ID=<%=rs_Album("id")%>','确认作废？');"/><input type="button" name="refreshPicture" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();" style="margin-left:20px;">全选&nbsp;
		  提示：点击图片查看大图</td>
			
		</tr>
		<tr class="tr1">
			<td>
			<%Dim rs_Picture
			Set rs_Picture = db.getRecordBySQL("select * from [Album_Pictures] where Album_ID=" & rs_Album("id") & " and status<>4 order by Show_Order desc,id desc") 
	
			if not (rs_Picture.eof or rs_Picture.bof) then
			do while not rs_Picture.eof%>
			<div style="height:210px; width:130px; float:left; margin-left:5px; margin-top:5px;"><a href="<%=Config.ImgUrl()%>Album/<%=rs_Picture("Picture")%>" target="_blank"><img src="<%=Config.ImgUrl()%>/<%=rs_Picture("Picture")%>" height="150" width="120" style="border-width:thin; background-color:#000000;"/></a><span><input class="checkbox" type="checkbox" name="info_id" id="info_id" value="<%=rs_Picture("ID")%>">图片名：<strong><%=rs_Picture("Title")%></strong></span><span>状态：<%=StatusResult(3,rs_Picture("status"))%> / 排序：<%=rs_Picture("Show_Order")%></span><span>
			<a id="Pictureview<%=rs_Picture("id")%>" href="#">[浏览]</a>
				 <script type="text/javascript"> 
   					$("#Pictureview<%=rs_Picture("id")%>").wBox({requestType:"iframe",iframeWH:{width:850,height:550},target:"PictureView.asp?id=<%=rs_Picture("id")%>"}); 
				</script>
			<%if rs_Picture("status")=0 or rs_Picture("status")=2 then%>
				 <a id="Pictureedit<%=rs_Picture("id")%>" href="#">[编辑]</a>
				 <script type="text/javascript"> 
   					$("#Pictureedit<%=rs_Picture("id")%>").wBox({requestType:"iframe_refresh",iframeWH:{width:850,height:550},target:"PictureEdit.asp?id=<%=rs_Picture("id")%>"}); 
				</script>
			<%end if%></span></div>
			<%
		  rs_Picture.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_Picture)
		  
	  %>
			
			</td>
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