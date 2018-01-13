<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim path : path ="../"

%>
<%
Dim RealName,Tel,Email,Content

RealName = ParaEncode(request.Form("RealName"))
Tel = ParaEncode(request.Form("Tel"))
Email = ParaEncode(request.Form("Email"))
Content = ParaEncode(request.Form("Content"))

if RealName = "" then
	response.Write("<script>alert('姓名不能为空');history.go(-1)</script>")
end if

if Tel = "" then
	response.Write("<script>alert('电话不能为空');history.go(-1)</script>")
end if

if Content = "" then
	response.Write("<script>alert('留言内容不能为空');history.go(-1)</script>")
end if
    
	'sql = "SET NOCOUNT ON insert into Messages(RealName,Tel,Email,Content,Adder_IP,Add_Time,Status) values('"& RealName &"','" & Tel & "','" & Email & "','" & Content & "','" & getip() & "','" & Now() & "',0) SELECT SCOPE_IDENTITY() SET NOCOUNT off"
	'db.AddRecordBySql(sql)
	'db.DoExecute(sql)
	
	call db.AddRecordByRS("addnew","Messages")
	
	call db.RSCmdAddPra("RealName",RealName)
	call db.RSCmdAddPra("Tel",Tel)
	call db.RSCmdAddPra("Email",Email)
	call db.RSCmdAddPra("Content",Content)
	call db.RSCmdAddPra("Adder_IP",GETIP())
	call db.RSCmdAddPra("Add_Time",now())
	call db.RSCmdAddPra("Status",0)

	infoId=db.AddRecordByRS("update","")
	
	response.Write("<script>alert('留言成功，请返回');window.location.href='index.asp'</script>")
	
	
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>