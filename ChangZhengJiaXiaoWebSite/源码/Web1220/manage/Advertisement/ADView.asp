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

dim Info_id

Info_id=request.QueryString("id")

if Info_id="" or Info_id=0 or not isnumeric(Info_id) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_id=cint(Info_id)


sqlstr = "select *,dbo.StatusShow(8,status) as statusshow,(select sort_name from Advertisement_Sort where Advertisement_Sort.id=Advertisement.Sort_ID) as adtypeshow,(select Manage_Url from Advertisement_Sort where Advertisement_Sort.id=Advertisement.Sort_ID) as Manage_Url,(select Rel_ShowFun from Advertisement_Sort where Advertisement_Sort.id=Advertisement.Sort_ID) as Manage_Fun from Advertisement where status<>4 and id = " & Info_id

'response.Write(sqlstr)
'response.End()


Dim rs_ad : Set rs_ad = db.getRecordBySQL(sqlstr)

if rs_ad.eof or rs_ad.bof then
	response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
	response.End()
end if

'Dim RelSetUrl:RelSetUrl=0

'		RelSetUrl = rs_ad("Manage_Url") & "?pro_id=" & rs_ad("id")
		
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
			<td align="left" colspan="4"><%if rs_ad("status")=0 or rs_ad("status")=2 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('adSave.asp?action=singleup&id=<%=rs_ad("id")%>');"/><input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('ADEdit.asp?id=<%=rs_ad("id")%>');"/>
				
                <%elseif rs_ad("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('adSave.asp?action=singledown&id=<%=rs_ad("id")%>');"/>
			
			<%end if%>
			<%if rs_ad("Manage_Url")<>"" then %><input name="adadd" type="button" class="button" id="adadd" style="margin-left:20px;" value="高级设置" onClick="GotoUrl('<%=(rs_ad("Manage_Url") & "?ad_id=" & rs_ad("id"))%>');"/>
			<%end if%>
            <input type="button" name="del" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('adSave.asp?action=singlecancel&id=<%=rs_ad("id")%>','_self','确定作废？');"/>
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('ADManage.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="4" style="text-align:center;">页面广告详情</th>
		</tr>
		<tr class="tr1">
			<td width="20%">广告名称：</td>
			<td width="80%" colspan="3"><%=rs_ad("Advertisement_Name")%></td>
		</tr>
		<tr class="tr2">
			<td width="20%">展示开始时间：</td>
			<td width="30%"><%if rs_ad("Show_Start_Time")<>"" then%>
		<%=rs_ad("Show_Start_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;</td>
	        <td width="20%">展示结束时间：</td>
	        <td width="30%"><%if rs_ad("Show_End_Time")<>"" then%>
		<%=rs_ad("Show_End_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;</td>
	  </tr>
	  		<tr class="tr1">
	  		  <td>广告分类：</td>
	  		  <td colspan="3"><%=rs_ad("adtypeshow")%>&nbsp;</td>
	  </tr>
	  		<tr class="tr2">
	  		  <td>默认标题：</td>
	  		  <td colspan="3"><%=rs_ad("Title")%>&nbsp;</td>
	  </tr>
	  		<tr class="tr1">
	  		  <td>默认链接：</td>
	  		  <td colspan="3"><%=rs_ad("Default_Link")%>&nbsp;</td>
	  </tr>
	  		<tr class="tr2">
	  		  <td>默认打开方式：</td>
	  		  <td colspan="3"><%=rs_ad("Default_Target")%>&nbsp;</td>
	  </tr>
	  		<tr class="tr1">
	  		  <td>默认宽度(px)：</td>
	  		  <td>
		<%if rs_ad("Default_Width")<>0 then%>
		<%=rs_ad("Default_Width")%>
		<%else%>
		自适应
		<%end if%>&nbsp;</td>
	          <td>默认高度(px)：</td>
	          <td><%if rs_ad("Default_Height")<>0 then%>
		<%=rs_ad("Default_Height")%>
		<%else%>
		自适应
		<%end if%>&nbsp;</td>
	  </tr>
	  		<tr class="tr2">
	  		  <td>自动关闭时间(秒)：</td>
	  		  <td><%if rs_ad("Close_Time")<>0 then%>
		<%=rs_ad("Close_Time")%>
		<%else%>
		不关闭
		<%end if%>&nbsp;</td>
	          <td>展示时间间隔(秒)：</td>
	          <td><%if rs_ad("Show_Interval")<>0 then%>
		<%=rs_ad("Show_Interval")%>
		<%else%>
		每次都出现
		<%end if%>&nbsp;</td>
	  </tr>
	  		<tr class="tr1">
	  		  <td>默认图片：</td>
	  		  <td colspan="3"><%if rs_ad("Default_Pic")<>"" then%>
                <img src="<%=Config.ImgUrl()%>AD/Default/<%response.Write(rs_ad("Default_Pic"))%>" />
                <%else%>
无
<%end if%></td>
	  </tr>
	  		<tr class="tr2">
	  		  <td>默认缩略图：</td>
	  		  <td colspan="3"><%if rs_ad("Default_Pic_View")<>"" then%>
                <img src="<%=Config.ImgUrl()%>AD/Default/<%response.Write(rs_ad("Default_Pic_View"))%>" />
                <%else%>
无
<%end if%></td>
	  </tr>
	  		
	  		
	  		<tr class="tr1">
		<td>显示顺序：</td>
		  <td colspan="3"><%=rs_ad("Show_Order")%>&nbsp;</td>
	  </tr>
	  <tr class="tr2">
		  <td>状态：</td>
		  <td colspan="3"><%=rs_ad("statusshow")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>简介：</td>
		  <td colspan="3"><%=rs_ad("Remark")%>&nbsp;</td>
	  </tr>
		<tr class="tr2">
			<td>调用说明：</td>
			<td colspan="3"><%=rs_ad("Call_Function")%>&nbsp;</td>
	  </tr>
  </table>
	<br />
	<%
'	select case rs_ad("Sort_ID")
'	case 1
'		'Call RelStoreShow()
'		execute("RelStoreShow")
'	case 2,3
'		'Call RelBrandShow()	
'		execute("RelBrandShow")
'	end select

'	response.Write(rs_ad("Manage_Fun"))
'	response.End()
	if rs_ad("Manage_Fun")<>"" then 
		execute(rs_ad("Manage_Fun"))
	end if
	%>

	<%Sub RelMarqueeShow()%>
		<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th colspan="7" style="text-align:center;">当前在线广告列表</th>
		</tr>
		<tr class="tr2">
			<td width="20%" align="center">项目名称</td>
			<td width="8%" align="center">状态</td>
			<td width="15%" align="center">对应链接</td>
			<td width="8%" align="center">打开方式</td>
			<td width="9%" align="center">图片</td>
			<td width="22%" align="center">展示开始时间/展示结束时间</td>
			<td width="18%" align="center">操作</td>
		</tr>
		<%Dim rs_Marquee
	Set rs_Marquee = db.getRecordBySQL("select *,dbo.StatusShow(8,status) as statusshow from Advertisement_Marquee where Advertisement_ID=" & rs_ad("id") & " and status=1 and dbo.AD_ShowTime(id,getdate(),2)=1 order by show_order desc,id desc") 
	
	if not (rs_Marquee.eof or rs_Marquee.bof) then
	
	do while not rs_Marquee.eof%>
		<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		  <td align="center"><%=rs_Marquee("Item_Name")%></td>
		  <td align="center"><%=rs_Marquee("statusshow")%></td>
		  <td align="center"><%=rs_Marquee("Link")%></td>
		  <td align="center"><%=rs_Marquee("Target")%></td>
		  <td align="center"><%if rs_Marquee("Pic_Url")<>"" then%>
               <a href="<%=Config.ImgUrl()%>AD/Picture/<%=rs_Marquee("Pic_Url")%>" target="_blank">点击预览</a>
                <%else%>
无
<%end if%>&nbsp;</td>
		  <td align="center"><%if rs_Marquee("Show_Start_Time")<>"" then%>
		<%=rs_Marquee("Show_Start_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;>><%if rs_Marquee("Show_End_Time")<>"" then%>
		<%=rs_Marquee("Show_End_Time")%>
		<%else%>
		结束时间无限制
		<%end if%></td>
		  <td align="center"><%if rs_Marquee("status")=0 or rs_Marquee("status")=2 then%>
						<input type="button" name="edit" class="button" value="上线" style="margin-left:5px;" onClick="GotoUrl('ADMarqueeSave.asp?id=<%=rs_Marquee("id")%>&ad_id=<%=rs_ad("id")%>&action=singleup');"/>
						<%elseif rs_Marquee("status")=1 then%>
						<input type="button" name="edit" class="button" value="下线" style="margin-left:5px;" onClick="GotoUrl('ADMarqueeSave.asp?id=<%=rs_Marquee("id")%>&ad_id=<%=rs_ad("id")%>&action=singledown');"/>
						<%end if%>&nbsp;
		  </td>
	  </tr>
	  <%
		  rs_Marquee.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_Marquee)
		  
	  %>
	  </table>
	<%End Sub%>
	
		<%Sub RelSinglePageShow()%>
		<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th colspan="7" style="text-align:center;">当前在线广告列表</th>
		</tr>
		<tr class="tr2">
			<td width="20%" align="center">项目名称</td>
			<td width="8%" align="center">状态</td>
			<td width="15%" align="center">对应链接</td>
			<td width="8%" align="center">打开方式</td>
			<td width="9%" align="center">图片</td>
			<td width="22%" align="center">展示开始时间/展示结束时间</td>
			<td width="18%" align="center">操作</td>
		</tr>
		<%Dim rs_SinglePage
	Set rs_SinglePage = db.getRecordBySQL("select *,dbo.StatusShow(1,status) as statusshow from Advertisement_SinglePage where Advertisement_ID=" & rs_ad("id") & " and status=1 and dbo.AD_ShowTime(id,getdate(),4)=1 order by show_order desc,id desc") 
	
	if not (rs_SinglePage.eof or rs_SinglePage.bof) then
	
	do while not rs_SinglePage.eof%>
		<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		  <td align="center"><%=rs_SinglePage("Item_Name")%></td>
		  <td align="center"><%=rs_SinglePage("statusshow")%></td>
		  <td align="center"><%=rs_SinglePage("Link")%></td>
		  <td align="center"><%=rs_SinglePage("Target")%></td>
		  <td align="center"><%if rs_SinglePage("Pic_Url")<>"" then%>
               <a href="<%=Config.ImgUrl()%>AD/Picture/<%=rs_SinglePage("Pic_Url")%>" target="_blank">点击预览</a>
                <%else%>
无
<%end if%>&nbsp;</td>
		  <td align="center"><%if rs_SinglePage("Show_Start_Time")<>"" then%>
		<%=rs_SinglePage("Show_Start_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;>><%if rs_SinglePage("Show_End_Time")<>"" then%>
		<%=rs_SinglePage("Show_End_Time")%>
		<%else%>
		结束时间无限制
		<%end if%></td>
		  <td align="center"><%if rs_SinglePage("status")=0 or rs_SinglePage("status")=2 then%>
						<input type="button" name="edit" class="button" value="上线" style="margin-left:5px;" onClick="GotoUrl('ADSinglePageSave.asp?id=<%=rs_SinglePage("id")%>&ad_id=<%=rs_ad("id")%>&action=singleup');"/>
						<%elseif rs_SinglePage("status")=1 then%>
						<input type="button" name="edit" class="button" value="下线" style="margin-left:5px;" onClick="GotoUrl('ADSinglePageSave.asp?id=<%=rs_SinglePage("id")%>&ad_id=<%=rs_ad("id")%>&action=singledown');"/>
						<%end if%>&nbsp;
		  </td>
	  </tr>
	  <%
		  rs_SinglePage.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_SinglePage)
		  
	  %>
	  </table>
	<%End Sub%>
	
		<%Sub RelBannerShow()%>
		<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th colspan="7" style="text-align:center;">当前在线广告列表</th>
		</tr>
		<tr class="tr2">
			<td width="20%" align="center">项目名称</td>
			<td width="8%" align="center">状态</td>
			<td width="15%" align="center">对应链接</td>
			<td width="8%" align="center">打开方式</td>
			<td width="9%" align="center">图片</td>
			<td width="22%" align="center">展示开始时间/展示结束时间</td>
			<td width="18%" align="center">操作</td>
		</tr>
		<%Dim rs_Banner
	Set rs_Banner = db.getRecordBySQL("select *,dbo.StatusShow(1,status) as statusshow from Advertisement_Banner where Advertisement_ID=" & rs_ad("id") & " and status=1 and dbo.AD_ShowTime(id,getdate(),3)=1 order by show_order desc,id desc") 
	
	if not (rs_Banner.eof or rs_Banner.bof) then
	
	do while not rs_Banner.eof%>
		<tr class="tr1" onMouseOver="this.style.backgroundColor='#C4D8ED'" onMouseOut ="this.style.backgroundColor='#F1F3F5'">
		  <td align="center"><%=rs_Banner("Item_Name")%></td>
		  <td align="center"><%=rs_Banner("statusshow")%></td>
		  <td align="center"><%=rs_Banner("Link")%></td>
		  <td align="center"><%=rs_Banner("Target")%></td>
		  <td align="center"><%if rs_Banner("Pic_Url")<>"" then%>
               <a href="<%=Config.ImgUrl()%>AD/Picture/<%=rs_Banner("Pic_Url")%>" target="_blank">点击预览</a>
                <%else%>
无
<%end if%>&nbsp;</td>
		  <td align="center"><%if rs_Banner("Show_Start_Time")<>"" then%>
		<%=rs_Banner("Show_Start_Time")%>
		<%else%>
		开始时间无限制
		<%end if%>&nbsp;>><%if rs_Banner("Show_End_Time")<>"" then%>
		<%=rs_Banner("Show_End_Time")%>
		<%else%>
		结束时间无限制
		<%end if%></td>
		  <td align="center"><%if rs_Banner("status")=0 or rs_Banner("status")=2 then%>
						<input type="button" name="edit" class="button" value="上线" style="margin-left:5px;" onClick="GotoUrl('ADBannerSave.asp?id=<%=rs_Banner("id")%>&ad_id=<%=rs_ad("id")%>&action=singleup');"/>
						<%elseif rs_Banner("status")=1 then%>
						<input type="button" name="edit" class="button" value="下线" style="margin-left:5px;" onClick="GotoUrl('ADBannerSave.asp?id=<%=rs_Banner("id")%>&ad_id=<%=rs_ad("id")%>&action=singledown');"/>
						<%end if%>&nbsp;
		  </td>
	  </tr>
	  <%
		  rs_Banner.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_Banner)
		  
	  %>
	  </table>
	<%End Sub%>
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