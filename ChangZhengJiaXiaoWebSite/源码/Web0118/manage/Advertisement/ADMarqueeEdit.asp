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

Dim rs_Marquee : Set rs_Marquee = db.getRecordBySQL("select *,(select Advertisement_name from Advertisement where Advertisement.id=Advertisement_Marquee.Advertisement_ID) as Advertisement_name from Advertisement_Marquee where status<>4 and id = " & Info_id)

if rs_Marquee.eof or rs_Marquee.bof then
	response.write "数据查询失败"
	response.End()
end if

picurl1= Config.ImgUrl() &"AD/" & rs_Marquee("Pic_Url")
picurl2= Config.ImgUrl() &"AD/" & rs_Marquee("Pic_View_Url")
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
<form name="MarqueeEdit" method="post" action="ADMarqueeSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="4">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('ADMarqueeManage.asp?AD_ID=<%=rs_Marquee("Advertisement_ID")%>');"/><input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="edit"><input type="hidden" name="Marquee_ID" value="<%=rs_Marquee("ID")%>"></td>
		</tr>
		<tr>
			<th colspan="4" style="text-align:center;">编辑广告项目</th>
		</tr>
		<tr class="tr1">
		  <td>对应页面广告：</td>
		  <td colspan="3"><%=rs_Marquee("Advertisement_Name")%>&nbsp;<input type="hidden" name="AD_ID" value="<%=rs_Marquee("Advertisement_ID")%>" /></td>
	  </tr>
		<tr class="tr2">
			<td width="20%">项目名称：</td>
			<td width="80%" colspan="3"><input name="Item_Name" type="text" id="Item_Name" size="50" value="<%=rs_Marquee("Item_Name")%>"/>
			*</td>
		</tr>
		<tr class="tr1">
		  <td width="20%">展示开始时间：</td>
		  <td width="30%"><input type="text" name="Show_Start_Time" id="Show_Start_Time" onfocus="WdatePicker({dateFmt:'yyyy-M-d H:m:s'})" class="Wdate" value="<%=rs_Marquee("Show_Start_Time")%>"/>
		    不设置为马上呈现</td>
		  <td width="20%">展示结束时间：</td>
		  <td width="30%"><input type="text" name="Show_End_Time" id="Show_End_Time" onfocus="WdatePicker({dateFmt:'yyyy-M-d H:m:s'})" class="Wdate" value="<%=rs_Marquee("Show_End_Time")%>" />
		    不设置为永久呈现</td>
	  </tr>
		
		<tr class="tr2">
			<td>链接：</td>
			<td colspan="3" class="tr2"><input name="Link" type="text" id="Link" size="50" value="<%=rs_Marquee("Link")%>"/></td>
      </tr>
		
		
		<tr class="tr1">
		  <td>打开方式：</td>
		  <td colspan="3"><label>
		    <select name="Target" id="Target">
		      <option value="_self" <%=object_selected(rs_Marquee("Target"),"_self")%>>_self</option>
		      <option value="_blank" <%=object_selected(rs_Marquee("Target"),"_blank")%>>_blank</option>
		      <option value="_parent" <%=object_selected(rs_Marquee("Target"),"_parent")%>>_parent</option>
		      <option value="_top" <%=object_selected(rs_Marquee("Target"),"_top")%>>_top</option>
	        </select>
		  </label> 
	      *</td>
	  </tr>
		
		<tr class="tr2">
		  <td>图片：</td>
		  <td colspan="3"><input name="Pic_Url" type="text" id="Pic_Url" size="50" value="<%=rs_Marquee("Pic_Url")%>"/>
		    &nbsp;&nbsp;<div id="PicOperate1" style="display: inline;">&nbsp;&nbsp;<div  id="ReUpload1" onClick="reupload('../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow1&picsize=350&picheight=0&picwidth=0&picurl=Pic_Url&PicOperate=PicOperate1','PicOperate1','Pic_Url','UploadFiles1');" style="cursor:hand;display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>  <div id="ViewShow1" style="cursor:hand;display:inline;" onMouseOver="PicView(0,'ViewPicShow1',this);" onMouseOut="PicView(1,'ViewPicShow1',this);">预览图片</div></div><img name="ViewPicShow1" id="ViewPicShow1" src="<%=picurl1%>" alt="" style="display:none;"/>&nbsp;</td>
	  </tr>
		<tr class="tr2">
		  <td>&nbsp;</td>
		  <td colspan="3"><iframe id="UploadFiles1" src="../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow1&picsize=200&picheight=0&picwidth=0&picurl=Pic_Url&PicOperate=PicOperate1&pictype=edit" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe> </td>
	  </tr>
		<tr class="tr1">
		  <td>缩略图：</td>
		  <td colspan="3"><input name="Pic_View_Url" type="text" id="Pic_View_Url" size="50" value="<%=rs_Marquee("Pic_View_Url")%>" />
		    &nbsp;&nbsp;<div id="PicOperate2" style="display: inline;">&nbsp;&nbsp;<div  id="ReUpload2" onClick="reupload('../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow2&picsize=350&picheight=0&picwidth=0&picurl=Pic_View_Url&PicOperate=PicOperate2','PicOperate2','Pic_View_Url','UploadFiles2');" style="cursor:hand;display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>  <div id="ViewShow2" style="cursor:hand;display:inline;" onMouseOver="PicView(0,'ViewPicShow2',this);" onMouseOut="PicView(1,'ViewPicShow2',this);">预览图片</div></div><img name="ViewPicShow2" id="ViewPicShow2" src="<%=picurl2%>" alt="" style="display:none;"/>&nbsp;&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>&nbsp;</td>
		  <td colspan="3"><iframe id="UploadFiles2" src="../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow2&picsize=200&picheight=0&picwidth=0&picurl=Pic_View_Url&PicOperate=PicOperate2&pictype=edit" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe> </td>
	  </tr>
		

		<tr class="tr2">
		  <td>显示顺序</td>
		  <td colspan="3"><input name="Show_Order" type="text" id="Show_Order" size="10" value="<%=rs_Marquee("Show_Order")%>"/>
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
	
	db.C(rs_Marquee)

	db.CloseConn()
	
	set db=nothing
%>