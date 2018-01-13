<%
uppath=request("uppath")&"/"			'文件上传路径
filelx=request("filelx")				'文件上传类型
formName=request("formName")			'回传到上页面编辑框所在Form的Name
EditName=request("EditName")			'回传到上页面编辑框的Name
%>
<html><head><title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function mysub()
{
		esave.style.visibility="visible";
}
-->
</script>
<style type="text/css">
<!--
.style1 {color: #000000}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style2 {
	font-size: 14px;
	color: #FFFFFF;
}
-->
</style>
<link href="css.css" rel="stylesheet" type="text/css">
</head>
<body>
<form name="form1" method="post" action="upload.asp" enctype="multipart/form-data" >
<div id="esave" style="position:absolute; top:18px; left:40px; z-index:10; visibility:hidden"> 
<TABLE WIDTH=340 BORDER=0 CELLSPACING=0 CELLPADDING=0>
<TR><td width=20%></td>
<TD bgcolor=#ff0000 width="60%"> 
<TABLE WIDTH=100% height=120 BORDER=0 CELLSPACING=1 CELLPADDING=0>
<TR> 
<td bgcolor=#ffffff align=center><font color=red>正在上传文件，请稍候...</font></td>
</tr>
</table>
</td><td width=20%></td>
</tr></table></div>
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" bgcolor="#368EDE" class="style1 style2"><b>图片上传</b> 
<input type="hidden" name="filepath" value="<%=uppath%>">
<input type="hidden" name="filelx" value="<%=filelx%>">
<input type="hidden" name="EditName" value="<%=EditName%>">
<input type="hidden" name="FormName" value="<%=formName%>">
<input type="hidden" name="act" value="uploadfile">
</td>
</tr>
<tr bgcolor="#E8F1FF"> 
      <td height="80" align="center" bgcolor="#D7E9F9" class="pt9" id="upid">选择文件: 
        <input type="file" name="file1" size="40" class="tx1" value="">
<input type="submit" name="Submit" value="开始上传" class="button" onclick="javascript:mysub()">
</td>
</tr>
</table>
</form>
</body>
</html>