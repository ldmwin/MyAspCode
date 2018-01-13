<%
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'★                                                                  ★
'☆                eWebEditor - eWebSoft在线编辑器                   ☆
'★                                                                  ★
'☆  版权所有: eWebSoft.com                                          ☆
'★                                                                  ★
'☆  程序制作: eWeb开发团队                                          ☆
'★            email:webmaster@webasp.net                            ★
'☆            QQ:589808                                             ☆
'★                                                                  ★
'☆  相关网址: [产品介绍]http://www.eWebSoft.com/Product/eWebEditor/ ☆
'★            [支持论坛]http://bbs.eWebSoft.com/                    ★
'☆                                                                  ☆
'★  主页地址: http://www.eWebSoft.com/   eWebSoft团队及产品         ★
'☆            http://www.webasp.net/     WEB技术及应用资源网站      ☆
'★            http://bbs.webasp.net/     WEB技术交流论坛            ★
'★                                                                  ★
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
%>

<%
Dim conn1
Dim connstr


Call OpenConn

Sub OpenConn()

    On Error Resume Next

    connstr="provider=microsoft.jet.oledb.4.0;data source="& Server.MapPath("../../data/EntBlog.mdb") &""
    Set conn1=Server.CreateObject("ADODB.Connection")
    conn1.Open connstr

    If Err Then
        Err.Clear
        Set Conn1 = Nothing
        Response.Write "数据库连接出错，请检查Conn1.asp文件中的数据库参数设置。"
        Response.End
    End If
End Sub

Sub CloseConn()
    On Error Resume Next
    If IsObject(Conn1) Then
        Conn1.Close
        Set Conn1 = Nothing
    End If
End Sub

Dim checkrs
If Session("adminid")="" or Session("adminname")="" or Session("flag")="" Then
	Response.Write("<script>top.location='../admin_login.asp';</script>")
	Response.End()
End If
Set checkrs=conn1.Execute("select * from Blog_Admin where id="& Session("adminid") &" and username='"& Session("adminname") &"' and flag="& Session("flag") &"")
If checkrs.Bof and checkrs.Eof Then
	Response.Write("<script>alert('你还没有登录,请先登录');top.location='../admin_login.asp';</script>")
	Response.End()
End If
checkrs.Close
Set checkrs=Nothing

call closeconn()
'If Session("eWebEditor_User") = "" Then
	'Response.Redirect "admin_login.asp"
	'Response.End
'End If

' 执行每天只需处理一次的事件
Call BrandNewDay()

' 初始化数据库连接
Call DBConnBegin()

' 公用变量
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))
sPosition = "位置：<a href='admin_default.asp'>后台管理</a> / "


' ********************************************
' 以下为页面公用区函数
' ********************************************
' ============================================
' 输出每页公用的顶部内容
' ============================================
Sub Header()
	Response.Write "<html><head>"
	
	' 输出 meta 标记
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"<meta name='Author' content='Andy.m'>" & _
		"<link rev=MADE href='mailto:webmaster@webasp.net'>"
	
	Response.Write "<link href='../images/admin_css.css' rel='stylesheet' type='text/css'>"
	' 输出标题
	Response.Write "<title>eWebEditor - eWebSoft在线文本编辑器 - 后台管理</title>"
	
    ' 输出每页都使用的基本样式表
	Response.Write "<link rel='stylesheet' type='text/css' href='admin/style.css'>"

	' 输出每页都使用的基本客户端脚本
	Response.Write "<script language='javaScript' SRC='admin/private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body topmargin=0 leftmargin=0>"
	Response.Write "<center>"
	
	
End Sub

' ============================================
' 输出每页公用的底部内容
' ============================================
Sub Footer()
	' 释放数据连接对象
	Call DBConnEnd()


	Response.Write "<center></body></html>"
End Sub




' ===============================================
' 初始化下拉框
'	s_FieldName	: 返回的下拉框名	
'	a_Name		: 定值名数组
'	a_Value		: 定值值数组
'	v_InitValue	: 初始值
'	s_Sql		: 从数据库中取值时,select name,value from table
'	s_AllName	: 空值的名称,如:"全部","所有","默认"
' ===============================================
Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function


%>