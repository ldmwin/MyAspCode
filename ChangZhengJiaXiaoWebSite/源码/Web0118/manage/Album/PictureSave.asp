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

Dim Title,Album_id,Show_Order,Picture
Dim Action:Action = request("action")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link href="../css/main.css" rel="stylesheet" type="text/css" />
<script src="../js/input.js" type="text/javascript" language="javascript"></script>
<script src="../js/function.js" type="text/javascript" language="javascript"></script>
</head>

<body>
<%sub return_to_list(Album_id)%>
<script language="javascript" type="text/javascript">
     //document.frm_page.submit(); 
	 GotoUrl('AlbumView.asp?id=<%=Album_ID%>','_self','');
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()	
	
	Title=trim(request.Form("Title"))
	Show_Order = trim(request.Form("Show_Order"))
	Picture = trim(request.Form("Picture"))
	
	if Title = "" then
	    response.write "<script>alert('图片名不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	
	if Show_Order = "" then
	    response.write "<script>alert('显示顺序不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Picture <> "" then
		Picturez = split(Picture,"/")
		maxz = ubound(Picturez)
		Picture = Picturez(maxz)
	end if 
	
	end sub

if action="add" then
'插入后的获取自动编号
	Album_id=trim(request.Form("Album_id"))
		
	if Album_id = "" then
	    response.write "<script>alert('所属图集不能为空');history.go(-1);</Script>"
		response.End()
	end if

	Call ReceiveData()
   
    'sql = "SET NOCOUNT ON insert into Album_Pictures(Title,Album_id,Adder_ID,Add_Time,Status,Show_Order,Picture) values('"& Title &"'," & Album_id & "," & User.UserID & ",getdate(),0," & Show_Order & ",'" & Picture & "') SELECT SCOPE_IDENTITY() SET NOCOUNT off"
	
	'rsid = db.AddRecordBySql(sql)
	'rsid =1
	
	call db.AddRecordByRS("addnew","Album_Pictures")
	
	call db.RSCmdAddPra("Title",Title)
	call db.RSCmdAddPra("Album_id",Album_id)
	call db.RSCmdAddPra("Adder_ID",User.UserID)
	call db.RSCmdAddPra("Add_Time",now())
	call db.RSCmdAddPra("Status",0)
	call db.RSCmdAddPra("Show_Order",Show_Order)
	call db.RSCmdAddPra("Picture",Picture)
	
	
	Info_id = db.AddRecordByRS("update","")

    if cint(Info_id) > 0 then

	   BackUrl = "PictureView.asp?id="& Info_id
	   Msg = "图片添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Infos_id = request.form("Picture_ID")
	
	if Infos_id="" or Infos_id=0 or not isnumeric(Infos_id) then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Infos_id=cint(Infos_id)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Album_Pictures set Title ='" & Title & "',Picture='" &  Picture & "',Show_Order=" & Show_Order & " where id=" & Infos_id
	'response.Write(sql)
	'response.End()
	db.DoExecute(SqlStr)
	
	   BackUrl = "PictureView.asp?id=" & Infos_id
	   Msg = "图片编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
'---------------------------------------------
elseif action="up" or action="down" or action="del" or action="recommend" or action="unrecommend" or action="cancel" then

	Album_id = request.QueryString("Album_id")
	
	if Album_id="" or Album_id=0 or not isnumeric(Album_id) then
		
		response.write("缺少操作参数")
		response.End()
		
	end if
	
	Album_id = cint(Album_id)

	i=Request.Form("info_id").Count
	if i<>0 then
		info_id=Request.Form("info_id") 
	
		select case action
		case "up"
			sqlstr="update Album_Pictures set status=1 where id in (" & info_id & ") and status=0"
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Album_Pictures set status=0 where id in (" & info_id & ") and status=1"
			db.DoExecute(SqlStr)
		case "cancel"
			sqlstr="update Album_Pictures set status=4 where id in (" & info_id & ")"
			db.DoExecute(SqlStr)
			
		end select
	else
		response.write("缺少操作参数")
		response.End()
	end if
	
	
	call return_to_list(Album_id)
elseif action="singleup" or action="singledown" then

	info_id=Request.QueryString("id")   

	select case action
	case "singleup"
		sqlstr="update Album_Pictures set status=1 where id in (" & info_id & ") and status=0"
		db.DoExecute(SqlStr)
	case "singledown"
		sqlstr="update Album_Pictures set status=0 where id in (" & info_id & ") and status=1"
		db.DoExecute(SqlStr)
		
	end select


	response.write("<script>window.location.href = 'PictureView.asp?id=" & info_id & "'; </Script>")
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