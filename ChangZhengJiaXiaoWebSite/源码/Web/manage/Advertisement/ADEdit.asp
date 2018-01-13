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

Dim rs_ad : Set rs_ad = db.getRecordBySQL("select * from Advertisement where status<>4 and id = " & Info_id)

if rs_ad.eof or rs_ad.bof then
	response.write "数据查询失败"
	response.End()
end if

picurl1= Config.ImgUrl() &"AD/" & rs_ad("Default_Pic")
picurl2= Config.ImgUrl() &"AD/" & rs_ad("Default_Pic_View")

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
<form name="ADEdit" method="post" action="ADSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="4">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('ADView.asp?id=<%=rs_ad("id")%>');"/>
				<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="edit"><input type="hidden" name="Advertisement_ID" value="<%=rs_ad("ID")%>"></td>
		</tr>
		<tr>
			<th colspan="4" style="text-align:center;">编辑页面广告</th>
		</tr>
		<tr class="tr1">
			<td width="20%">广告名称：</td>
			<td width="80%" colspan="3"><input name="Advertisement_Name" type="text" id="Advertisement_Name" size="50"  value="<%=rs_ad("Advertisement_Name")%>" />
			*</td>
		</tr>
		<tr class="tr2">
		  <td width="20%">展示开始时间：</td>
		  <td width="30%"><input type="text" name="Show_Start_Time" id="Show_Start_Time" onfocus="WdatePicker({dateFmt:'yyyy-M-d H:m:s'})" class="Wdate"  value="<%=rs_ad("Show_Start_Time")%>" />
		    不设置为马上呈现</td>
		  <td width="20%">展示结束时间：</td>
		  <td width="30%"><input type="text" name="Show_End_Time" id="Show_End_Time" onfocus="WdatePicker({dateFmt:'yyyy-M-d H:m:s'})" class="Wdate"  value="<%=rs_ad("Show_End_Time")%>" />
		    不设置为永久呈现</td>
	  </tr>
		<tr class="tr1">
		  <td>广告分类：</td>
		  <td colspan="3"><%
			Set rs = db.getRecordBySQL("select * from Advertisement_Sort where status=1 order by show_order desc,id desc") %>
		  <select name="Sort_ID" id="Sort_ID">
			<%
		  if not(rs.eof and rs.bof) then
		  do until rs.eof
		  response.Write("<option value=""" & rs("id") & """ " & object_selected(rs_ad("Sort_ID"),rs("id")) & ">&nbsp;" & rs("Sort_Name") & "&nbsp;</option>")
		  rs.movenext
		  loop
		  end if
		  db.C(rs)
		  %>
		  </select>*</td>
	  </tr>
		<tr class="tr2">
		  <td>默认标题：</td>
		  <td colspan="3"><input name="Title" type="text" id="Title" size="50"  value="<%=rs_ad("Title")%>" /></td>
	  </tr>
		<tr class="tr1">
			<td>默认链接：</td>
			<td colspan="3" class="tr2"><input name="Default_Link" type="text" id="Default_Link" size="50"  value="<%=rs_ad("Default_Link")%>" /></td>
      </tr>
		
		
		<tr class="tr2">
		  <td>默认打开方式：</td>
		  <td colspan="3"><label>
		    <select name="Default_Target" id="Default_Target">
		      <option value="_self" <%=object_selected(rs_ad("Default_Target"),"_self")%>>_self</option>
		      <option value="_blank" <%=object_selected(rs_ad("Default_Target"),"_blank")%>>_blank</option>
		      <option value="_parent" <%=object_selected(rs_ad("Default_Target"),"_parent")%>>_parent</option>
		      <option value="_top" <%=object_selected(rs_ad("Default_Target"),"_top")%>>_top</option>
	        </select>
		  </label> 
	      *</td>
	  </tr>
		<tr class="tr1">
		  <td>默认宽度(px)：</td>
		  <td><input name="Default_Width" type="text" id="Default_Width" size="15"  value="<%=rs_ad("Default_Width")%>" />
		    0为自适应</td>
	      <td>默认高度(px)：</td>
	      <td><input name="Default_Height" type="text" id="Default_Height" size="15"  value="<%=rs_ad("Default_Height")%>" />
	        0为自适应</td>
	  </tr>
		<tr class="tr2">
		  <td>自动关闭时间(秒)：</td>
		  <td><input name="Close_Time" type="text" id="Close_Time" size="15" value="<%=rs_ad("Close_Time")%>" />
		    0秒为不关闭</td>
	      <td>展示时间间隔(秒)：</td>
	      <td><input name="Show_Interval" type="text" id="Show_Interval" size="15" value="<%=rs_ad("Show_Interval")%>" />
          0秒为每次都出现</td>
	  </tr>
		<tr class="tr1">
		  <td>默认图片：</td>
		  <td colspan="3"><input name="Default_Pic" type="text" id="Default_Pic" size="50"  value="<%=rs_ad("Default_Pic")%>" />
		    &nbsp;&nbsp;<div id="PicOperate1" style="display: inline;">&nbsp;&nbsp;<div  id="ReUpload1" onClick="reupload('../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow1&picsize=350&picheight=0&picwidth=0&picurl=Default_Pic&PicOperate=PicOperate1','PicOperate1','Default_Pic','UploadFiles1');" style="cursor:hand;display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>  <div id="ViewShow1" style="cursor:hand;display:inline;" onMouseOver="PicView(0,'ViewPicShow1',this);" onMouseOut="PicView(1,'ViewPicShow1',this);">预览图片</div></div><img name="ViewPicShow1" id="ViewPicShow1" src="<%=picurl1%>" alt="" style="display:none;"/>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>&nbsp;</td>
		  <td colspan="3"><iframe id="UploadFiles1" src="../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow1&picsize=200&picheight=0&picwidth=0&picurl=Default_Pic&PicOperate=PicOperate1&pictype=edit" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe> </td>
	  </tr>
		<tr class="tr2">
		  <td>默认缩略图：</td>
		  <td colspan="3"><input name="Default_Pic_View" type="text" id="Default_Pic_View" size="50"  value="<%=rs_ad("Default_Pic_View")%>" />
		    &nbsp;&nbsp;<div id="PicOperate2" style="display: inline;">&nbsp;&nbsp;<div  id="ReUpload2" onClick="reupload('../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow2&picsize=350&picheight=0&picwidth=0&picurl=Default_Pic_View&PicOperate=PicOperate2','PicOperate2','Default_Pic_View','UploadFiles2');" style="cursor:hand;display:inline;" onMouseOver="PicView(0,null,this);" onMouseOut="PicView(1,null,this);">重新上传</div>  <div id="ViewShow2" style="cursor:hand;display:inline;" onMouseOver="PicView(0,'ViewPicShow2',this);" onMouseOut="PicView(1,'ViewPicShow2',this);">预览图片</div></div><img name="ViewPicShow2" id="ViewPicShow2" src="<%=picurl2%>" alt="" style="display:none;"/>&nbsp;&nbsp;</td>
	  </tr>
		<tr class="tr2">
		  <td>&nbsp;</td>
		  <td colspan="3"><iframe id="UploadFiles2" src="../inc/UploadPhotos.asp?savepath=AD&picshow=ViewPicShow2&picsize=200&picheight=0&picwidth=0&picurl=Default_Pic_View&PicOperate=PicOperate2&pictype=edit" frameborder=0 scrolling=no width="450" height="25" style="margin-top:2px;"></iframe>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>简介：</td>
		  <td colspan="3"><textarea name="Remark" cols="60" rows="6" id="Remark"> <%=rs_ad("Remark")%></textarea></td>
	  </tr>
		<tr class="tr2">
		  <td>调用说明：</td>
		  <td colspan="3"><textarea name="Call_Function" cols="60" rows="6" id="Call_Function"><%=rs_ad("Call_Function")%></textarea></td>
	  </tr>
		
		<tr class="tr1">
		  <td>显示顺序</td>
		  <td colspan="3"><input name="Show_Order" type="text" id="Show_Order" value="0" size="10"  value="<%=rs_ad("Show_Order")%>" />
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

	db.C(rs_ad)

	db.CloseConn()
	
	set db=nothing
%>