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

Dim P_id,Type_Name,Show_Order,sqlstr,New_Level
Dim Action:Action = request("action")


	sub ReceiveData()	
	
	P_id = trim(request.Form("P_id"))
	Type_Name = trim(request.Form("Type_Name"))
	Show_Order = trim(request.Form("Show_Order"))

	if P_id = "" then
	    response.write "<script>alert('父级分类不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Type_Name = "" then
	    response.write "<script>alert('分类名称不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Show_Order = "" then
	    response.write "<script>alert('显示顺序不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	end sub
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript" language="javascript"></script>
</head>

<body>
<%sub return_to_list()%>
<form action="TypeManage.asp" method=post name="frm_page">

</form>
<script language="javascript" type="text/javascript">
     document.frm_page.submit(); 
</script>
<%end sub%>

<%
if action="add" then
'插入后的获取自动编号
   call ReceiveData()
   
'   if p_id =0 then
'   sqlstr = "SET NOCOUNT ON insert into Information_Type(Type_Name,Parent_ID,Type_Level,Adder_ID,Add_Time,Status,Show_Order) values('"& Type_Name &"'," & p_id & ",1," & User.UserID & ",getdate(),1," & Show_Order & ") SELECT SCOPE_IDENTITY() SET NOCOUNT off"
'	else
'	sqlstr = "DECLARE @level int select @level=Type_Level+1 from Information_Type where id=" & p_id & " SET NOCOUNT ON insert into Information_Type(Type_Name,Parent_ID,Type_Level,Adder_ID,Add_Time,Status,Show_Order) values('"& Type_Name &"'," & p_id & ",@level," & User.UserID & ",getdate(),1," & Show_Order & ") SELECT SCOPE_IDENTITY() SET NOCOUNT off"
'	end if
'	
'	
'	rsid = db.AddRecordBySql(sqlstr)
	'rsid =1
	
	'response.Write(rsid)
	
	if p_id =0 then
	
		New_Level = 1

	else
		Set rs_Type = db.getRecordBySQL("select Type_Level from Information_Type where id=" & p_id)
		
		if rs_Type.eof or rs_Type.bof then
			response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
			response.End()
		else
			New_Level = rs_Type("Type_Level") + 1
		end if
		
		db.C(rs_Type)
		
	end if
	
		call db.AddRecordByRS("addnew","Information_Type")
		
		call db.RSCmdAddPra("Type_Name",Type_Name)
		call db.RSCmdAddPra("Parent_ID",p_id)
		call db.RSCmdAddPra("Type_Level",New_Level)
		call db.RSCmdAddPra("Adder_ID",User.UserID)
		call db.RSCmdAddPra("Add_Time",now())
		call db.RSCmdAddPra("Status",0)
		call db.RSCmdAddPra("Show_Order",Show_Order)
		
		
		Info_id = db.AddRecordByRS("update","")
	
	

    if cint(Info_id) > 0 then

	   BackUrl = "TypeView.asp?id=" & Info_id
	   Msg = "分类添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
'插入后的获取自动编号
   call ReceiveData()
   
   c_id=request.Form("id")
   
   sqlstr = "update Information_Type set Type_Name='"& Type_Name &"',Show_Order=" & Show_Order & " where id=" & c_id
	

	db.DoExecute(sqlstr)

	   BackUrl = "TypeView.asp?id=" & c_id
	   Msg = "分类编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	
'---------------------------------------------
elseif action="up" or action="down" or action="cancel" or action="recommend" or action="unrecommend" then

	Infos_id=Request.QueryString("id")
	
   	if Infos_id="" or Infos_id=0 or not isnumeric(Infos_id) then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_id=cint(Infos_id) 
	
	p_id=Request.QueryString("p_id")
	
   	if p_id="" or not isnumeric(p_id) then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	p_id=cint(p_id) 
	
		select case action
		case "up"
			sqlstr="update Information_Type set status=1 where id in (" & Infos_id & ") and status=0"
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Information_Type set status=0 where id in (" & Infos_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Information_Type set status=4 where id in (" & Infos_id & ")"
			db.DoExecute(SqlStr)
			
		end select
	
	if p_id = 0 then
		 response.write("<script>window.location.href = 'TypeList.asp'; </Script>")
	else
		response.write("<script>window.location.href = 'TypeView.asp?id=" & p_id &"'; </Script>")
	end if
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