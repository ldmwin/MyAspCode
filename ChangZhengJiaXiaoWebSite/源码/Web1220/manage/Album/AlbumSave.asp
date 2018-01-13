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

Dim Album_Name,Pic_View,Brief,Album_Type,Show_Order
Dim Action:Action = request("action")

'response.Write(action)
'response.End()
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
<form action="AlbumManage.asp" method=post name="frm_page">
<input type="hidden" name="me_page" value=<%=session("back_page")%>>
<input type="hidden" name="keyword" value=<%=session("back_keyword")%>>
<input type="hidden" name="RSstatus" value=<%=session("back_RSstatus")%>>
<input type="hidden" name="AlbumType" value=<%=session("back_AlbumType")%>>
<input type="hidden" name="SearchStartTime" value=<%=session("back_SearchStartTime")%>>
<input type="hidden" name="SearchEndTime" value=<%=session("back_SearchEndTime")%>>
</form>
<script language="javascript" type="text/javascript">
     document.frm_page.submit(); 
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()	
	
	Album_Name = trim(request.Form("Album_Name"))
	Pic_View = trim(request.Form("Pic_View"))
	Brief = trim(request.Form("Brief"))
	Album_Type = trim(request.Form("Album_Type"))
	Show_Order = trim(request.Form("Show_Order"))

	if Album_Name = "" then
	    response.write "<script>alert('标题不能为空');history.go(-1);</Script>"
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
	
	'response.Write(Pic_View)
	'response.End()
	
	end sub

if action="add" then
'插入后的获取自动编号
	Call ReceiveData()
   
   'sql = "SET NOCOUNT ON insert into Albums(Album_Name,Pic_View,Brief,Adder_ID,Add_Time,Status,Show_Order,Album_Type) values('"& Album_Name &"','" & Pic_View & "','" & Brief & "'," & User.UserID & ",getdate(),0," & Show_Order & "," & Album_Type & ") SELECT SCOPE_IDENTITY() SET NOCOUNT off"
	
	'rsid = db.AddRecordBySql(sql)
	'rsid =1
	
	call db.AddRecordByRS("addnew","Albums")
	
	call db.RSCmdAddPra("Album_Name",Album_Name)
	call db.RSCmdAddPra("Pic_View",Pic_View)
	call db.RSCmdAddPra("Brief",Brief)
	call db.RSCmdAddPra("Adder_ID",User.UserID)
	call db.RSCmdAddPra("Add_Time",now())
	call db.RSCmdAddPra("Status",0)
	call db.RSCmdAddPra("Show_Order",Show_Order)
	call db.RSCmdAddPra("Album_Type",Album_Type)
	
	
	Info_id = db.AddRecordByRS("update","")

    if cint(Info_id) > 0 then

	   BackUrl = "AlbumManage.asp"
	   Msg = "图集添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Infos_ID = request.form("Album_id")
	
	if Infos_ID="" or Infos_ID=0 or not isnumeric(Infos_ID) then
		Call CloseConn()
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_ID=cint(Infos_ID)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Albums set Album_Name ='" & Album_Name & "',Pic_View='" & Pic_View & "',Brief='" &  Brief & "',Show_Order=" & Show_Order & ",Album_Type=" & Album_Type & " where id=" & Infos_ID
	'response.Write(SqlStr)
	'response.End()
	db.DoExecute(SqlStr)	
	
	
	   BackUrl = "AlbumView.asp?id=" & Infos_ID
	   Msg = "图集编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	
'---------------------------------------------
elseif action="up" or action="down" or action="del" or action="recommend" or action="unrecommend" or action="cancel" then

	i=Request.Form("Album_id").Count
	if i<>0 then
		Album_id=Request.Form("Album_id") 
	
		select case action
		case "up"
			sqlstr="update Albums set status=1 where id in (" & Album_id & ") and (status=0 or status=2)"
			'response.Write(sqlstr)
			'response.End()
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Albums set status=2 where id in (" & Album_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Albums set status=4 where id in (" & Album_id & ")"
			db.DoExecute(SqlStr)
			
		end select
	else
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
	end if
	
	
	call return_to_list()
elseif action="singleup" or action="singledown" then
	Album_id=Request.QueryString("id")   

	select case action
	case "singleup"
		sqlstr="update Albums set status=1 where id in (" & Album_id & ") and (status=0 or status=2)"
		db.DoExecute(SqlStr)
	case "singledown"
		sqlstr="update Albums set status=2 where id in (" & Album_id & ") and status=1"
		db.DoExecute(SqlStr)
		
	end select


	response.write("<script>window.location.href = 'AlbumView.asp?id=" & Album_id & "'; </Script>")
elseif action="singlecancel" then
		Album_id=Request.QueryString("id")
		sqlstr="update Albums set status=4 where id in (" & Album_id & ")"
		db.DoExecute(SqlStr)
	
		response.write("<script>window.location.href = 'AlbumManage.asp'; </Script>")
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