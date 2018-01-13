<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>  
<%Response.Expires=0%>
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

Dim Product_Name,Brand,Pic_View,Brief,Introduce,Product_Type,Show_Order,Product_Unit
Dim Action:Action = request("action")

'response.Write(action)
'response.End()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Introduce-type" Introduce="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript" language="javascript"></script>
</head>

<body>
<%sub return_to_list()%>
<form action="ProductManage.asp" method=post name="frm_page">
<input type="hidden" name="me_page" value=<%=session("back_page")%>>
<input type="hidden" name="keyword" value=<%=session("back_keyword")%>>
<input type="hidden" name="RSstatus" value=<%=session("back_RSstatus")%>>
<input type="hidden" name="ProductType" value=<%=session("back_ProductType")%>>
<input type="hidden" name="SearchStartTime" value=<%=session("back_SearchStartTime")%>>
<input type="hidden" name="SearchEndTime" value=<%=session("back_SearchEndTime")%>>
</form>
<script language="javascript" type="text/javascript">
     document.frm_page.submit(); 
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()	
	
	Product_Name = trim(request.Form("Product_Name"))
	Brand = trim(request.Form("Brand"))
	Pic_View = trim(request.Form("Pic_View"))
	Brief = trim(request.Form("Brief"))
	Product_Type = trim(request.Form("Product_Type"))
	Show_Order = trim(request.Form("Show_Order"))
	Product_Unit = trim(request.Form("Product_Unit"))
	
	For i = 1 To Request.Form("Introduce").Count 
		Introduce = Introduce & Request.Form("Introduce")(i) 
	Next 

	if Product_Name = "" then
	    response.write "<script>alert('标题不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Product_Type = "" then
	    response.write "<script>alert('分类不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Show_Order = "" then
	    response.write "<script>alert('排序不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Pic_View <> "" then
	  Pic_Viewz = split(Pic_View,"/")
	  maxz = ubound(Pic_Viewz)
	  Pic_View = Pic_Viewz(maxz)
	end if 
	
	end sub

if action="add" then
'插入后的获取自动编号
	Call ReceiveData()
   
   'sql = "SET NOCOUNT ON insert into Products(Product_Name,Brand,Introduce,Pic_View,Brief,Adder_ID,Add_Time,Status,Show_Order,Product_Type) values('"& Product_Name &"','" & Brand & "','" & Introduce & "','" & Pic_View & "','" & Brief & "'," & User.UserID & ",getdate(),0," & Show_Order & "," & Product_Type & ") SELECT SCOPE_IDENTITY() SET NOCOUNT off"
'	
'	rsid = db.AddRecordBySql(sql)
	'rsid =1
	
	call db.AddRecordByRS("addnew","Products")
	
	call db.RSCmdAddPra("Product_Name",Product_Name)
	call db.RSCmdAddPra("Brand",Brand)
	call db.RSCmdAddPra("Introduce",Introduce)
	call db.RSCmdAddPra("Pic_View",Pic_View)
	call db.RSCmdAddPra("Brief",Brief)
	call db.RSCmdAddPra("Adder_ID",User.UserID)
	call db.RSCmdAddPra("Add_Time",now())
	call db.RSCmdAddPra("Status",0)
	call db.RSCmdAddPra("Show_Order",Show_Order)
	call db.RSCmdAddPra("Product_Hits",0)
	call db.RSCmdAddPra("Product_Type",Product_Type)
	call db.RSCmdAddPra("Product_Unit",Product_Unit)
	
	
	
	Info_id = db.AddRecordByRS("update","")

    if cint(Info_id) > 0 then

	   BackUrl = "ProductManage.asp"
	   Msg = "产品添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Infos_id = request.form("Product_id")
	
	if Infos_id="" or Infos_id=0 or not isnumeric(Infos_id) then
		Call CloseConn()
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_id=cint(Infos_id)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Products set Product_Name ='" & Product_Name & "',Brand='" & Brand & "',Introduce='" & Introduce & "',Pic_View='" & Pic_View & "',Brief='" &  Brief & "',Show_Order=" & Show_Order & ",Product_Type=" & Product_Type & ",Product_Unit='" & Product_Unit & "' where id=" & Infos_id
	'response.Write(SqlStr)
	'response.End()
	db.DoExecute(SqlStr)	
	
	
	   BackUrl = "ProductView.asp?id=" & Infos_id
	   Msg = "产品编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	
'---------------------------------------------
elseif action="up" or action="down" or action="del" or action="recommend" or action="unrecommend" or action="cancel" then

	i=Request.Form("info_id").Count
	if i<>0 then
		info_id=Request.Form("info_id") 
	
		select case action
		case "up"
			sqlstr="update Products set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
			'response.Write(sqlstr)
			'response.End()
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Products set status=2 where id in (" & info_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Products set status=4 where id in (" & info_id & ")"
			db.DoExecute(SqlStr)
			
		end select
	else
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
	end if
	
	
	call return_to_list()
elseif action="singleup" or action="singledown" then
	info_id=Request.QueryString("id")   

	select case action
	case "singleup"
		sqlstr="update Products set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
		db.DoExecute(SqlStr)
	case "singledown"
		sqlstr="update Products set status=2 where id in (" & info_id & ") and status=1"
		db.DoExecute(SqlStr)
		
	end select


	response.write("<script>window.location.href = 'ProductView.asp?id=" & info_id & "'; </Script>")
elseif action="singlecancel" then
		info_id=Request.QueryString("id")
		sqlstr="update Products set status=4 where id in (" & info_id & ")"
		db.DoExecute(SqlStr)
	
		response.write("<script>window.location.href = 'ProductManage.asp'; </Script>")
else  
    do_result="缺少操作参数"
	response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
	response.End()
end if
%>
</body>
</html>
<%
	set User = nothing
	set EventLog = nothing
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>