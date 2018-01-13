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

dim Info_ID

Info_ID=request.QueryString("id")

if Info_ID="" or Info_ID=0 or not isnumeric(Info_ID) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_ID=cint(Info_ID)


sqlstr = "select *,dbo.StatusShow(1,status) as statusshow,(select type_name from Magazine_Type where Magazine_Type.id=Magazines.Magazine_Type) as Magazinetypeshow from Magazines where status<>4 and id = " & Info_ID

'response.Write(sqlstr)
'response.End()


Dim rs_Magazine : Set rs_Magazine = db.getRecordBySQL(sqlstr)

if rs_Magazine.eof or rs_Magazine.bof then
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
				<%if rs_Magazine("status")=0 or rs_Magazine("status")=2 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('MagazineSave.asp?action=singleup&id=<%=rs_Magazine("id")%>');"/><input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('MagazineEdit.asp?id=<%=rs_Magazine("id")%>');"/>
                <%elseif rs_Magazine("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('MagazineSave.asp?action=singledown&id=<%=rs_Magazine("id")%>');"/>
			
			<%end if%>
            <input type="button" name="cancel" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('MagazineSave.asp?action=singlecancel&id=<%=rs_Magazine("id")%>','_self','确定作废？');"/>
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('MagazineManage.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">图集预览</th>
		</tr>
		<tr class="tr1">
			<td width="30%">标题：</td>
			<td width="70%"><%=rs_Magazine("Magazine_Name")%></td>
		</tr>
		
		<tr class="tr1">
			<td width="30%">图集类型：</td>
			<td width="70%"><%=rs_Magazine("Magazinetypeshow")%></td>
		</tr>
		
		<tr class="tr2">
		  <td>期数：</td>
		  <td><%=rs_Magazine("Magazine_Period")%></td>
	  </tr>
		<tr class="tr1">
		  <td>有效期：</td>
		  <td><%=rs_Magazine("Magazine_Validity")%></td>
	  </tr>
		<tr class="tr2">
			<td width="30%">封面：</td>
			<td width="70%"><%if rs_Magazine("Pic_View")<>"" then%>
                <img src="<%=Config.ImgUrl()%>Magazine/<%response.Write(rs_Magazine("Pic_View"))%>" />
                <%else%>
无
<%end if%></td>
		</tr>
		<tr class="tr1">
		  <td>状态：</td>
		  <td><%=rs_Magazine("statusshow")%></td>
	  </tr>
	  <tr class="tr2">
		  <td>显示顺序：</td>
		  <td><%=rs_Magazine("Show_Order")%></td>
	  </tr>
		<tr class="tr1">
		  <td>简介：</td>
		  <td><%=rs_Magazine("Brief")%>&nbsp;</td>
	  </tr>
</table>
<br />
<form name="frm_list" method="post" action="<%=SubmitUrl%>">
  <table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		
		<tr>
			<th style="text-align:center;">分页列表</th>
		</tr>
		<tr class="tr2">
		  <td  align="left">
		  <%if rs_Magazine("status")=0 or rs_Magazine("status")=2 then%>
		  <input name="Pageadd" id="Pageadd" type="button" class="button"style="margin-left:20px;" value="新建分页" />
		   <script type="text/javascript"> 
   				$("#Pageadd").wBox({requestType:"iframe_refresh",iframeWH:{width:850,height:550},target:"PageAdd.asp?id=<%=rs_Magazine("id")%>"}); 
			</script>
		  <%end if%>
		  <input type="button" name="up" class="button" value="上线分页" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PageSave.asp?action=up&Magazine_ID=<%=rs_Magazine("id")%>','确认上线？');"/>
			<input type="button" name="down" class="button" value="下线分页" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PageSave.asp?action=down&Magazine_ID=<%=rs_Magazine("id")%>','确认下线？');"/>
			<input type="button" name="cancel" class="button" value="作废分页" style="margin-left:20px;" onClick="GoToUrl_MorePrm('PageSave.asp?action=cancel&Magazine_ID=<%=rs_Magazine("id")%>','确认作废？');"/><input type="button" name="refreshPage" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/><input name="chk_all" type="checkbox" id="chk_all" onClick="SelectAll();" style="margin-left:20px;">全选&nbsp;
		  提示：点击分页查看大图</td>
			
		</tr>
		<tr class="tr1">
			<td>
			<%Dim rs_Page
			Set rs_Page = db.getRecordBySQL("select *,dbo.StatusShow(3,status) as statusshow from [Magazine_Pages] where Magazine_ID=" & rs_Magazine("id") & " and status<>4 order by Show_Order desc,id desc") 
	
			if not (rs_Page.eof or rs_Page.bof) then
			do while not rs_Page.eof%>
			<div style="height:210px; width:130px; float:left; margin-left:5px; margin-top:5px;"><a href="<%=Config.ImgUrl()%>Magazine/<%=rs_Page("Picture")%>" target="_blank"><img src="<%=Config.ImgUrl()%>Magazine/<%=rs_Page("Picture")%>" height="150" width="120" style="border-width:thin; background-color:#000000;"/></a><span><input class="checkbox" type="checkbox" name="info_id" id="info_id" value="<%=rs_Page("ID")%>">分页名：<strong><%=rs_Page("Title")%></strong></span><span>状态：<%=rs_Page("statusshow")%> / 排序：<%=rs_Page("Show_Order")%></span><span>
			<a id="Pageview<%=rs_Page("id")%>" href="#">[浏览]</a>
				 <script type="text/javascript"> 
   					$("#Pageview<%=rs_Page("id")%>").wBox({requestType:"iframe",iframeWH:{width:850,height:550},target:"PageView.asp?id=<%=rs_Page("id")%>"}); 
				</script>
			<%if rs_Page("status")=0 or rs_Page("status")=2 then%>
				 <a id="Pageedit<%=rs_Page("id")%>" href="#">[编辑]</a>
				 <script type="text/javascript"> 
   					$("#Pageedit<%=rs_Page("id")%>").wBox({requestType:"iframe_refresh",iframeWH:{width:850,height:550},target:"PageEdit.asp?id=<%=rs_Page("id")%>"}); 
				</script>
			<%end if%></span></div>
			<%
		  rs_Page.movenext()
			
		Loop
		
		response.Write("</ul>")
		
		end if
		
		db.C(rs_Page)
		  
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
	
	db.C(rs_Magazine)

	db.CloseConn()
	
	set db=nothing
%>