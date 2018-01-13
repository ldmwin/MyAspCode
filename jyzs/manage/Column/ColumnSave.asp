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

Dim Column_Type,Link,Show_Order,Target,Column_Name,sqlstr,New_Level
Dim Parent_ID,Site_ID
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
<%sub return_to_list(Site_ID)%>
<script language="javascript" type="text/javascript">
     //document.frm_page.submit(); 
	 GotoUrl('ColumnManage.asp?Site_ID=<%=Site_ID%>','_self','');
</script>
<%end sub%>

<%
	
	
	sub ReceiveData()

	if action = "add" then
		Parent_ID = trim(request.Form("Parent_ID"))
		
		if Parent_ID="" or not isnumeric(Parent_ID) then
			response.write "<script>alert('参数出错');history.go(-1);</Script>"
			response.end()
		elseif Parent_ID=0 then
			Site_ID=request("Site_ID")
			
			if Site_ID=""  or not isnumeric(Site_ID) then
				response.write "<script>alert('参数出错');history.go(-1);</Script>"
				response.end()
			else
			Site_ID=cint(Site_ID)
		
			end if
		end if
		Parent_ID=cint(Parent_ID)
		
		
		if Parent_ID<>0 then
		
			Dim rs_Column : Set rs_Column = db.getRecordBySQL("select * from Site_Columns where status<>4 and id = " & Parent_ID)
			
			if rs_Column.eof or rs_Column.bof then
				response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
				response.End()
			else
				Site_ID = rs_Column("Site_ID")			
			end if
			
			db.C(rs_Column)
			
		end if
		
		Dim rs_Site : Set rs_Site = db.getRecordBySQL("select * from Sys_Sites where status<>4 and id = " & Site_ID)
		
		if rs_Site.eof or rs_Site.bof then
			response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
			response.End()
		end if
		
		db.C(rs_Site)
		
		Column_Type = trim(request.Form("Column_Type"))
		
		if Column_Type = "" then
			response.write "<script>alert('栏目分类不能为空');history.go(-1);</Script>"
			response.End()
		end if
	
	end if	
	
	Column_Name = trim(request.Form("Column_Name"))
	
	Show_Order = trim(request.Form("Show_Order"))
	Link = trim(request.Form("Link"))
	Target = trim(request.Form("Target"))
	
	if Column_Name = "" then
	    response.write "<script>alert('栏目名称不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Link = "" then
	    response.write "<script>alert('链接地址不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Target = "" then
	    response.write "<script>alert('打开位置不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	if Show_Order = "" then
	    response.write "<script>alert('显示顺序不能为空');history.go(-1);</Script>"
		response.End()
	end if
	
	end sub

if action="add" then
'插入后的获取自动编号

	Call ReceiveData()
	
	if Parent_ID =0 then
	
		New_Level = 1

	else
		Set rs_Column = db.getRecordBySQL("select Column_Level from Site_Columns where id=" & Parent_ID)
		
		if rs_Column.eof or rs_Column.bof then
			response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
			response.End()
		else
			New_Level = rs_Column("Column_Level") + 1
		end if
		
		db.C(rs_Column)
		
	end if
	
		call db.AddRecordByRS("addnew","Site_Columns")
		
		call db.RSCmdAddPra("Column_Name",Column_Name)
		call db.RSCmdAddPra("Link",Link)
		call db.RSCmdAddPra("Target",Target)
		call db.RSCmdAddPra("Parent_ID",Parent_ID)
		call db.RSCmdAddPra("Column_Level",New_Level)
		call db.RSCmdAddPra("Column_Type",Column_Type)
		call db.RSCmdAddPra("Site_ID",Site_ID)
		call db.RSCmdAddPra("Adder_ID",User.UserID)
		call db.RSCmdAddPra("Add_Time",now())
		call db.RSCmdAddPra("Status",0)
		call db.RSCmdAddPra("Show_Order",Show_Order)
		call db.RSCmdAddPra("Column_Content","")
		
		
		Info_id = db.AddRecordByRS("update","")

    if cint(Info_id) > 0 then

	   BackUrl = "ColumnEdit.asp?id="& Info_id
	   Msg = "栏目添加成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
	end if
elseif action="edit" then
	Info_id = request.form("Column_ID")
	
	if Info_id="" or Info_id=0 or not isnumeric(Info_id) then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Info_id=cint(Info_id)
	
	call receivedata()
	'on error resume next
	
	SqlStr="update Site_Columns set Column_Name ='" & Column_Name & "',Link = '" & Link & "',Target='" &  Target & "',Show_Order=" & Show_Order & " where id=" & Info_id
	'response.Write(sql)
	'response.End()
	db.DoExecute(SqlStr)
	
	   BackUrl = "ColumnEdit.asp?id=" & Info_id
	   Msg = "栏目编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
'---------------------------------------------
elseif action="contentedit" then
	dim sColumn_Content,Mould_Page,UseChildContent
	Info_id = request.form("Column_ID")
	
	if Info_id="" or Info_id=0 or not isnumeric(Info_id) then
		do_result="缺少操作参数"
		response.write("<script>alert('" & do_result & "');history.go(-1);</Script>")
		response.End()
		
	end if
	
	Info_id=cint(Info_id)
	
	Mould_Page = trim(request.Form("Mould_Page"))
	
	UseChildContent = trim(request.Form("UseChildContent"))
	
	if UseChildContent <> "1" then
		UseChildContent = "0"
	end if
	
	For i = 1 To Request.Form("Column_Content").Count 
		sColumn_Content = sColumn_Content & Request.Form("Column_Content")(i) 
	Next 
	
	'on error resume next
	
	SqlStr="update Site_Columns set Mould_Page =" & Mould_Page & ",UseChildContent = " & UseChildContent & ",Column_Content='" &  sColumn_Content & "' where id=" & Info_id
	'response.Write(SqlStr)
	'response.End()
	db.DoExecute(SqlStr)
	
	   BackUrl = "ColumnContentEdit.asp?id=" & Info_id
	   Msg = "内容编辑成功"
		%>
			<!--#include file="../Inc/Massage.asp"-->
		<%
'---------------------------------------------
elseif action="up" or action="down" then
	
	Site_ID=request("Site_ID")
		
	if Site_ID=""  or not isnumeric(Site_ID) then
		response.write "<script>alert('参数出错');history.go(-1);</Script>"
		response.end()
	else
		Site_ID=cint(Site_ID)
	end if

	i=Request.Form("info_id").Count
	if i<>0 then
		info_id=Request.Form("info_id") 
	
		select case action
		case "up"
			sqlstr="update Site_Columns set status=1 where id in (" & info_id & ") and status=0"
			db.DoExecute(SqlStr)
		case "down"
			sqlstr="update Site_Columns set status=0 where id in (" & info_id & ") and status=1"
			db.DoExecute(SqlStr)
			
		end select
	else
		response.write("缺少操作参数")
		response.End()
	end if
	
	
	response.write("<script>window.location.href = 'ColumnManage.asp?Site_ID=" & Site_ID & "'; </Script>")
elseif action="singlecancel" then

	info_id=Request.QueryString("id")   

	select case action
		case "singlecancel"
			Dim rs_Column : Set rs_Column = db.getRecordBySQL("select *,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum from Site_Columns m where m.status<>4 and m.id = " & info_id)
	
			if rs_Column.eof or rs_Column.bof then
				response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
				response.End()
			else
				if rs_Column("childnodenum")>0 then
					response.write "<script>alert('当前栏目有子栏目，不能作废');history.go(-1);</Script>"
					response.End()
				else
					Site_ID = rs_Column("Site_ID")
					sqlstr="update Site_Columns set status=4 where id=" & info_id
					db.DoExecute(SqlStr)
				end if
			end if
			
			db.C(rs_Column)	
		
	end select


	response.write("<script>window.location.href = 'ColumnManage.asp?Site_ID=" & Site_ID & "'; </Script>")
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