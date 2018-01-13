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
db.dbConnStr = Config.ConnStr(0)
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

if not IsUserInit() then

	Call EventLog.LogAdd(3,0,"system:usercheck fail" & User.UserErr)
	response.Redirect("../inc/error.asp?msg=" & User.UserErr & "。&errurl=1")

end if

Dim Item_Name,Show_End_Time,Show_Start_Time,AD_ID,Show_Order,Link,Target,Pic_Url,Pic_View_Url
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
<form action="ADBannerManage.asp" method=post name="frm_page">
<input type="hidden" name="me_page" value=<%=session("back_page")%>>
<input type="hidden" name="keyword" value=<%=session("back_keyword")%>>
<input type="hidden" name="RSstatus" value=<%=session("back_RSstatus")%>>
<input type="hidden" name="AD_ID" value=<%=session("back_AD_ID")%>>
<input type="hidden" name="SearchStartTime" value=<%=session("back_SearchStartTime")%>>
<input type="hidden" name="SearchEndTime" value=<%=session("back_SearchEndTime")%>>
</form>
<script language="javascript" type="text/javascript">
     document.frm_page.submit(); 
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()	
	
	Item_Name = trim(request.Form("Item_Name"))
	Show_Start_Time = trim(request.Form("Show_Start_Time"))
	Show_End_Time = trim(request.Form("Show_End_Time"))
	AD_ID = trim(request.Form("AD_ID"))
	Link = trim(request.Form("Link"))
	Target = trim(request.Form("Target"))
	Pic_Url = trim(request.Form("Pic_Url"))
	Pic_View_Url = trim(request.Form("Pic_View_Url"))
	Show_Order = trim(request.Form("Show_Order"))	
	

	if Item_Name="" then
	    response.write "<script>alert('广告名不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Target="" then
	    response.write "<script>alert('打开方式不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if AD_ID="" then
	    response.write "<script>alert('对应页面广告不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Show_Order="" then
	    response.write "<script>alert('显示顺序不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Show_Start_Time = "" then
		Show_Start_Time = "null"
	else
		Show_Start_Time = "'" & Show_Start_Time & "'"
	end if
	
	if Show_End_Time = "" then
		Show_End_Time = "null"
	else
		Show_End_Time = "'" & Show_End_Time & "'"
	end if
	
	end sub

if action="add" then
'插入后的获取自动编号
	Call ReceiveData()
   
   sql = "SET NOCOUNT ON insert into Advertisement_Banner(Item_Name,Show_Start_Time,Show_End_Time,Advertisement_ID,Adder_ID,Add_Time,Status,Show_Order,Link,Target,Pic_Url,Pic_View_Url) values('"& Item_Name &"'," & Show_Start_Time & "," & Show_End_Time & "," & AD_ID & "," & User.UserID & ",getdate(),0," & Show_Order & ",'" & Link & "','" & Target & "','" & Pic_Url & "','" & Pic_View_Url & "') SELECT SCOPE_IDENTITY() SET NOCOUNT off"
   
   'response.Write(sql)
	
	rsid = db.AddRecordBySql(sql)
	'rsid =1
	
	'response.Write(rsid)

    if cint(rsid) > 0 then

	   BackUrl = "ADBannerManage.asp?ad_id=" & AD_ID
	   Msg = "广告项目添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Infos_id = request.form("Banner_id")
	
	'response.Write(Infos_id)
	
	if Infos_id="" or (not isnumeric(Infos_id)) or Infos_id=0 then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_id=cint(Infos_id)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Advertisement_Banner set Item_Name ='" & Item_Name & "',Show_Start_Time=" & Show_Start_Time & ",Show_End_Time=" & Show_End_Time & ",Advertisement_ID='" & AD_ID & "',Target='" & Target & "',Link='" & Link & "',Pic_Url='" & Pic_Url & "',Pic_View_Url='" & Pic_View_Url & "',Show_Order=" & Show_Order & " where id=" & Infos_id
	'response.Write(SqlStr)
	'response.End()
	db.DoExecute(SqlStr)	
	
	
	   BackUrl = "ADBannerManage.asp?ad_id=" & AD_ID
	   Msg = "广告项目编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	
'---------------------------------------------
elseif action="up" or action="down" or action="cancel" or action="del" then

	i=Request.Form("info_id").Count
	if i<>0 then
		info_id=Request.Form("info_id") 
	
		select case action
		case "up"
			sqlstr="update Advertisement_Banner set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
			'response.Write(sqlstr)
			'response.End()
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Advertisement_Banner set status=2 where id in (" & info_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Advertisement_Banner set status=4 where id in (" & info_id & ")"
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
	ad_id=Request.QueryString("ad_id")

	select case action
	case "singleup"
		sqlstr="update Advertisement_Banner set status=1 where id in (" & info_id & ") and (status=0 or status=2)"
		db.DoExecute(SqlStr)
	case "singledown"
		sqlstr="update Advertisement_Banner set status=2 where id in (" & info_id & ") and status=1"
		db.DoExecute(SqlStr)
		
	end select


	response.write("<script>window.location.href = 'ADView.asp?id=" & ad_id & "'; </Script>")
elseif action="singlecancel" then
		info_id=Request.QueryString("id")
		ad_id=Request.QueryString("ad_id")
		
		sqlstr="update Advertisement_Banner set status=4 where id in (" & info_id & ")"
		db.DoExecute(SqlStr)
	
		response.write("<script>window.location.href = 'ADView.asp?id=" & ad_id & "'; </Script>")
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

	db.CloseConn()
	
	set db=nothing
%>