<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<%
'-----------------------------------
'文 件 名 : /logioutasp
'功    能 : 登出用户
'作    者 : Mr.Lion
'建立时间 : 2011/05/06
'-----------------------------------
%>
<!--#include file="Class/DBCtrl.asp" -->
<!--#include file="Class/UserCtrl.asp" -->
<!--#include file="Class/LogCtrl.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"")
db.OpenConn()

Dim user : Set user = New UserCtrl
Dim EventLog: Set EventLog = New LogCtrl

'Session("UserID") = 0


if IsUserInit() then
	
	Call EventLog.LogAdd(4,User.UserID,"system:logout sucess")
	
	Session("UserID") = null
	Session("UserName") = null

end if


set User = nothing
set EventLog = nothing
set Config = nothing

db.CloseConn()

set db=nothing

response.redirect("login.asp")

%>