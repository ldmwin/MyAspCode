<%
'������������������������������������
'��                                                                  ��
'��                eWebEditor - eWebSoft���߱༭��                   ��
'��                                                                  ��
'��  ��Ȩ����: eWebSoft.com                                          ��
'��                                                                  ��
'��  ��������: eWeb�����Ŷ�                                          ��
'��            email:webmaster@webasp.net                            ��
'��            QQ:589808                                             ��
'��                                                                  ��
'��  �����ַ: [��Ʒ����]http://www.eWebSoft.com/Product/eWebEditor/ ��
'��            [֧����̳]http://bbs.eWebSoft.com/                    ��
'��                                                                  ��
'��  ��ҳ��ַ: http://www.eWebSoft.com/   eWebSoft�ŶӼ���Ʒ         ��
'��            http://www.webasp.net/     WEB������Ӧ����Դ��վ      ��
'��            http://bbs.webasp.net/     WEB����������̳            ��
'��                                                                  ��
'������������������������������������
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
        Response.Write "���ݿ����ӳ�������Conn1.asp�ļ��е����ݿ�������á�"
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
	Response.Write("<script>alert('�㻹û�е�¼,���ȵ�¼');top.location='../admin_login.asp';</script>")
	Response.End()
End If
checkrs.Close
Set checkrs=Nothing

call closeconn()
'If Session("eWebEditor_User") = "" Then
	'Response.Redirect "admin_login.asp"
	'Response.End
'End If

' ִ��ÿ��ֻ�账��һ�ε��¼�
Call BrandNewDay()

' ��ʼ�����ݿ�����
Call DBConnBegin()

' ���ñ���
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))
sPosition = "λ�ã�<a href='admin_default.asp'>��̨����</a> / "


' ********************************************
' ����Ϊҳ�湫��������
' ********************************************
' ============================================
' ���ÿҳ���õĶ�������
' ============================================
Sub Header()
	Response.Write "<html><head>"
	
	' ��� meta ���
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"<meta name='Author' content='Andy.m'>" & _
		"<link rev=MADE href='mailto:webmaster@webasp.net'>"
	
	Response.Write "<link href='../images/admin_css.css' rel='stylesheet' type='text/css'>"
	' �������
	Response.Write "<title>eWebEditor - eWebSoft�����ı��༭�� - ��̨����</title>"
	
    ' ���ÿҳ��ʹ�õĻ�����ʽ��
	Response.Write "<link rel='stylesheet' type='text/css' href='admin/style.css'>"

	' ���ÿҳ��ʹ�õĻ����ͻ��˽ű�
	Response.Write "<script language='javaScript' SRC='admin/private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body topmargin=0 leftmargin=0>"
	Response.Write "<center>"
	
	
End Sub

' ============================================
' ���ÿҳ���õĵײ�����
' ============================================
Sub Footer()
	' �ͷ��������Ӷ���
	Call DBConnEnd()


	Response.Write "<center></body></html>"
End Sub




' ===============================================
' ��ʼ��������
'	s_FieldName	: ���ص���������	
'	a_Name		: ��ֵ������
'	a_Value		: ��ֵֵ����
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
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