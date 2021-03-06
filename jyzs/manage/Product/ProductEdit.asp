﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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

picurl1= Config.ImgUrl() & "Info/" & rs_Product("Pic_View")		
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
<form name="ProductEdit" method="post" action="ProductSave.asp">
	<table border="0" cellspacing="1" cellpadding="5" height="1" align="center" width="100%">
		<tr class="tr2">
			<td align="left" colspan="2">
				<input type="submit" name="submit" class="button" value="保存" style="margin-left:20px;"/>
				<input type="button" name="return" class="button" value="返回" style="margin-left:20px;" onClick="GotoUrl('ProductView.asp?id=<%=rs_Product("id")%>');"/>
                <input type="button" name="refresh" value="刷新" class="button" style="margin-left:20px;"  onClick="Refresh('_self');"/>
				<input type="hidden" name="action" value="edit"><input type="hidden" name="Product_id" value="<%=rs_Product("id")%>"></td>
		</tr>
		<tr>
			<th colspan="2" style="text-align:center;">编辑产品</th>
		</tr>
		<tr class="tr1">
			<td width="30%">品名：</td>
			<td width="70%"><input name="Product_Name" type="text" id="Product_Name" size="50" value="<%=rs_Product("Product_Name")%>" />
			*</td>
		</tr>
		
		<tr class="tr2">
		  <td>品牌：</td>
		  <td><input name="Brand" type="text" id="Brand" size="30" value="<%=rs_Product("Brand")%>" /></td>
	  </tr>
		<tr class="tr1">
		  <td>规格：</td>
		  <td><input name="Product_Unit" type="text" id="Product_Unit" size="30" value="<%=rs_Product("Product_Unit")%>" /></td>
	  </tr>
		<tr class="tr1">
			<td width="30%">产品类型：</td>
			<td width="70%">
			<%
			Set rs = db.getRecordBySQL("select * from Product_Type where status <> 4 and Type_Level = 2 order by show_order desc,id desc") %>
		  <select name="Product_Type" id="Product_Type">
			<%
		  if not(rs.eof and rs.bof) then
		  do until rs.eof
		  response.Write("<option value=""" & rs("id") & """ " & object_selected(rs_Product("Product_Type"),rs("id")) & ">&nbsp;" & rs("Type_Name") & "&nbsp;</option>")
		  rs.movenext
		  loop
		  end if
		  db.C(rs)
		  %>
		  </select>*			</td>
		</tr>
		
		<tr class="tr2">
			<td width="30%">缩略图：</td>
			<td width="70%"><input name="Pic_View" type="text" id="Pic_View" size="50"  value="<%=rs_Product("Pic_View")%>"/>
		    &nbsp;<input type="button" name="Submit2" value="上传图片" onClick="window.open('../inc/upload_flash.asp?formname=ProductEdit&editname=Pic_View&uppath=../../Pictures&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=105')"></td>
		</tr>
		<tr class="tr1">
		  <td>简介：</td>
		  <td><textarea name="Brief" cols="60" rows="6" id="Brief"><%=rs_Product("Brief")%></textarea></td>
	  </tr>
		<tr class="tr2">
			<td width="30%">详细介绍：</td>
			<td width="70%"><input type="hidden" name="Introduce" value="<%=Server.HTMLEncode(rs_Product("Introduce"))%>"><iframe id="Content" src="../editor/eWebEditor.asp?id=Introduce&style=s_blue" frameborder="0" scrolling="no" width="550" height="350"></iframe></td>
		</tr>
		<tr class="tr1">
		  <td>显示顺序</td>
		  <td><input name="Show_Order" type="text" id="Show_Order" value="<%=rs_Product("Show_Order")%>" size="10" />
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
	
	db.C(rs_Product)

	db.CloseConn()
	
	set db=nothing
%>