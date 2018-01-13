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

Dim Title,Sub_Title,Pic_View,Brief,sContent,Info_Type,Show_Order
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
<form action="InfoManage.asp" method=post name="frm_page">
<input type="hidden" name="me_page" value=<%=session("back_page")%>>
<input type="hidden" name="keyword" value=<%=session("back_keyword")%>>
<input type="hidden" name="RSstatus" value=<%=session("back_RSstatus")%>>
<input type="hidden" name="InfoType" value=<%=session("back_InfoType")%>>
<input type="hidden" name="SearchStartTime" value=<%=session("back_SearchStartTime")%>>
<input type="hidden" name="SearchEndTime" value=<%=session("back_SearchEndTime")%>>
</form>
<script language="javascript" type="text/javascript">
     document.frm_page.submit(); 
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()	
	
	Title = trim(request.Form("Title"))
	Sub_Title = trim(request.Form("Sub_Title"))
	Pic_View = trim(request.Form("Pic_View"))
	Brief = trim(request.Form("Brief"))
	Info_Type = trim(request.Form("Info_Type"))
	Show_Order = trim(request.Form("Show_Order"))
	
	For i = 1 To Request.Form("Content").Count 
		sContent = sContent & Request.Form("Content")(i) 
	Next 

	if Title = "" then
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
	
	end sub

if action="add" then
'插入后的获取自动编号
	Call ReceiveData()
   
   'sql = "SET NOCOUNT ON insert into Informations(Title,Sub_Title,Content,Pic_View,Brief,Adder_ID,Add_Time,Status,Show_Order,Info_Type) values('"& Title &"','" & Sub_Title & "','" & sContent & "','" & Pic_View & "','" & Brief & "'," & User.UserID & ",getdate(),0," & Show_Order & "," & Info_Type & ") SELECT SCOPE_IDENTITY() SET NOCOUNT off"
'	
'	rsid = db.AddRecordBySql(sql)
	'rsid =1
	
	call db.AddRecordByRS("addnew","Informations")
	
	call db.RSCmdAddPra("Title",Title)
	call db.RSCmdAddPra("Sub_Title",Sub_Title)
	call db.RSCmdAddPra("Content",sContent)
	call db.RSCmdAddPra("Pic_View",Pic_View)
	call db.RSCmdAddPra("Brief",Brief)
	call db.RSCmdAddPra("Adder_ID",User.UserID)
	call db.RSCmdAddPra("Add_Time",now())
	call db.RSCmdAddPra("Status",0)
	call db.RSCmdAddPra("Show_Order",Show_Order)
	call db.RSCmdAddPra("Info_Hits",0)
	call db.RSCmdAddPra("Info_Type",Info_Type)
	
	
	Info_id = db.AddRecordByRS("update","")

    if cint(Info_id) > 0 then

	   BackUrl = "InfoManage.asp"
	   Msg = "新闻添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Infos_id = request.form("Info_id")
	
	if Infos_id="" or Infos_id=0 or not isnumeric(Infos_id) then
		Call CloseConn()
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_id=cint(Infos_id)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Informations set Title ='" & Title & "',Sub_Title='" & Sub_Title & "',Content='" & sContent & "',Pic_View='" & Pic_View & "',Brief='" &  Brief & "',Show_Order=" & Show_Order & ",Info_Type=" & Info_Type & " where id=" & Infos_id
	'response.Write(SqlStr)
	'response.End()
	db.DoExecute(SqlStr)	
	
	
	   BackUrl = "InfoView.asp?id=" & Infos_id
	   Msg = "新闻编辑成功"
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
			sqlstr="update Informations set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
			'response.Write(sqlstr)
			'response.End()
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Informations set status=2 where id in (" & info_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Informations set status=4 where id in (" & info_id & ")"
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
		sqlstr="update Informations set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
		db.DoExecute(SqlStr)
	case "singledown"
		sqlstr="update Informations set status=2 where id in (" & info_id & ") and status=1"
		db.DoExecute(SqlStr)
		
	end select


	response.write("<script>window.location.href = 'InfoView.asp?id=" & info_id & "'; </Script>")
elseif action="singlecancel" then
		info_id=Request.QueryString("id")
		sqlstr="update Informations set status=4 where id in (" & info_id & ")"
		db.DoExecute(SqlStr)
	
		response.write("<script>window.location.href = 'InfoManage.asp'; </Script>")
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