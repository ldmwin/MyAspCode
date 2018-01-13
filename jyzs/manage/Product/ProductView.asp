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

dim Info_id

Info_id=request.QueryString("id")

if Info_id="" or Info_id=0 or not isnumeric(Info_id) then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_id=cint(Info_id)


sqlstr = "select *,(select type_name from Product_Type where Product_Type.id=Products.Product_Type) as producttypeshow from Products where status<>4 and id = " & Info_id

'response.Write(sqlstr)
'response.End()


Dim rs_Product : Set rs_Product = db.getRecordBySQL(sqlstr)

if rs_Product.eof or rs_Product.bof then
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
<script language="javascript" type="text/javascript" src="../JS/My97DatePicker/WdatePicker.js"></script>
</head>

<body>
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<%if rs_Product("status")=0 or rs_Product("status")=2 then%>
				<input type="button" name="up" class="button" value="上线" style="margin-left:20px;" onClick="GotoUrl('ProductSave.asp?action=singleup&id=<%=rs_Product("id")%>');"/>
                <input type="button" name="edit" class="button" value="编辑" style="margin-left:20px;" onClick="GotoUrl('ProductEdit.asp?id=<%=rs_Product("id")%>');"/>
                <%elseif rs_Product("status")=1 then%>
			<input type="button" name="down" class="button" value="下线" style="margin-left:20px;" onClick="GotoUrl('ProductSave.asp?action=singledown&id=<%=rs_Product("id")%>');"/>
			
			<%end if%>
            <input type="button" name="del" class="button" value="作废" style="margin-left:20px;" onClick="GotoUrl('ProductSave.asp?action=singlecancel&id=<%=rs_Product("id")%>','_self','确定作废？');"/>
            <input type="button" name="return" class="button" value="返回列表" style="margin-left:20px;" onClick="GotoUrl('ProductManage.asp');"/>
<input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">产品预览</th>
		</tr>
		<tr class="tr1">
			<td width="30%">品名：</td>
			<td width="70%"><%=rs_Product("Product_Name")%></td>
		</tr>
		
		<tr class="tr2">
		  <td>品牌：</td>
		  <td><%=rs_Product("Brand")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
		  <td>规格：</td>
		  <td><%=rs_Product("Product_Unit")%></td>
	  </tr>
		<tr class="tr2">
			<td width="30%">产品类型：</td>
			<td width="70%"><%=rs_Product("producttypeshow")%></td>
		</tr>
		
		<tr class="tr1">
			<td width="30%">缩略图：</td>
			<td width="70%"><%if rs_Product("Pic_View")<>"" then%>
                <img src="<%=Config.ImgUrl()%>/<%response.Write(rs_Product("Pic_View"))%>" />
                <%else%>
无
<%end if%></td>
		</tr>
		<tr class="tr2">
		  <td>状态：</td>
		  <td><%=StatusResult(1,rs_Product("status"))%></td>
	  </tr>
	  <tr class="tr1">
		  <td>显示顺序：</td>
		  <td><%=rs_Product("Show_Order")%></td>
	  </tr>
		<tr class="tr2">
		  <td>摘要：</td>
		  <td><%=rs_Product("Brief")%>&nbsp;</td>
	  </tr>
		<tr class="tr1">
			<td width="30%">详细介绍：</td>
			<td width="70%"><%=rs_Product("Introduce")%></td>
		</tr>
		
</table>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing
	
	db.C(rs_Product)

	db.CloseConn()
	
	set db=nothing
%>