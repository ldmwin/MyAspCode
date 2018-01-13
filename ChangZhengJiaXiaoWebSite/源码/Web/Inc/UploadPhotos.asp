<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>
<%
dim picsize,picheight,picwidth,picuse,pictype,picshow,picurl,PicOperate
'picuse=Clng(trim(request("PhotoUrlID")))
pictype=trim(request("pictype"))
picsize=trim(request("picsize"))
picheight = trim(request("picheight"))
picwidth = trim(request("picwidth"))
picshow = trim(request("picshow"))
picurl = trim(request("picurl"))
PicOperate = trim(request("PicOperate"))
savepath = trim(request("savepath"))

if picsize = "" then
	picsize = 1024
end if

if picheight = "" then
	picheight = 0
end if

if picwidth = "" then
	picwidth = 0
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<link href="../css/main.css" rel="stylesheet" type="text/css">
<script src="../js/input.js" type="text/javascript"></script>-->
<style type="text/css">
BODY{
BACKGROUND-COLOR: #F1F3F5;
margin: 0px;
}
Body input{
margin: 0px;
padding: 0px;}
</style>
<script language="javascript" src="../js/piccheck.js" type="text/javascript"></script>
<script language=javascript>

AllowImgFileSize=<%=picsize%>; //允许上传图片文件的大小 0为无限制 单位：KB 
AllowImgWidth=<%=picwidth%>; //允许上传的图片的宽度 &#320;为无限制　单位：px(像素) 
AllowImgHeight=<%=picheight%>;
viewpic="<%=picshow%>";
</script>
</head>
<body>
<%if pictype="edit" then%>

<%else%>
<form action="UpfilePhotos.asp" method="post" name="form1" onSubmit="return check();" enctype="multipart/form-data">
  <input name="FileName" id="FileName" type="FILE" size="30" onchange= "CheckExt(this);">
  <input type="submit" name="Submit" id="Submit" value="上传"  disabled>
<!--  <input name="PhotoUrlID" type="hidden" id="PhotoUrlID" value="<%'=picuse%>">-->
  <input name="picsize" type="hidden" id="picsize" value="<%=picsize%>">
  <input name="picshow" type="hidden" id="picshow" value="<%=picshow%>">
  <input name="picurl" type="hidden" id="picurl" value="<%=picurl%>">
  <input name="PicOperate" type="hidden" id="PicOperate" value="<%=PicOperate%>">
  <input name="savepath" type="hidden" id="savepath" value="<%=savepath%>">
  </form>
<%end if%>
</body>
</html>
