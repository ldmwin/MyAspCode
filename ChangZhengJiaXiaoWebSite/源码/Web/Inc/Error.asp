<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<%
'-----------------------------------
'文 件 名 : /inc/error.asp
'功    能 : 错误信息提醒 
'作    者 : Mr.Lion
'建立时间 : 2011/05/10
'-----------------------------------
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>错误信息</title>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<link href="../css/login.css" rel="stylesheet" type="text/css" />
</head>
<body>
<script language="javascript" type="text/javascript">
	if ( parent.location != document.location )
	{
		parent.location = document.location;
	}  
</script>
<%
Dim errorstr
Dim errormsg:errormsg=request.querystring("msg")
select case request.querystring("errurl")
	case "1"
		errorstr = "[<a href=../login.asp target=_top>返回</a>]"
	case "2"
		errorstr = "<br /><a href='javascript:history.go(-1);' target=_parent></a>"
	case "3"
		errorstr = "<br /><a href='javascript:history.go(-1);' target=_self></a>"
	case else
		errorstr = request.querystring("error")
end select
%>

<center>

	<div id="nifty">
		<b class="rtop"><b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b></b>
		<div style="width:403px; height:26px; line-height:26px; background:none; font-size:12px; text-align:left;">错误提示</div>
		<div style="width:403px; height:46px; background:#166CA3;"><img src="../images/error.gif" alt="" /></div>
		<div style="width:401px !important; width:403px; height:auto; background:#fff; border-left:1px solid #649EB2; border-right:1px solid #649EB2; padding-top:10px;">
            <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr>
                    <td align="center" valign="middle" style="line-height:2em;"><%=errormsg&errorstr%></td>
                </tr>
            </table>
		</div>
		<div style="width:401px !important; width:403px; height:20px; background:#F7F7E7; border:1px solid #649EB2; border-top:1px solid #ddd; margin-bottom:5px; font-size:12px; line-height:20px; "></div>
		<b class="rbottom"><b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b></b>
	</div>
</center>

</body>
</html>