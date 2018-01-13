<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.Charset="UTF-8" %>
<%
Const EnableUploadFile="Yes"        '是否开放文件上传
dim MaxFileSize
MaxFileSize=1024        '上传文件大小限制
dim SaveShowPath
'SaveShowPath="http://img.jly.com.cn/"'图片路径
SaveShowPath="http://172.16.0.2:8035/"'图片路径
dim SaveSiteTruePath
SaveSiteTruePath="E:/NJTYP/NJTYPImg/"   'E:\XTTYP\TYPImg
'SaveSiteTruePath="E:/SiteDev/JLYGroupImgSite/"
'E:\JLYGroupV2\JLYGroupImg          '实际存放上传文件的目录
'Const SaveUpFilesPath="../Upfiles/Media"        '存放上传文件的目录
'Const UpFileType="gif|jpg|bmp"
Const UpFileType="jpg|png|gif"        '允许的上传文件类型
%>
<!--#include file="UpfileClass.asp"-->
<%
const upload_type=0   '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

dim upload,oFile,formName,SavePath,orginsavepath,filename,fileExt,oFileSize,picsize,picshow,picurl,PicOperate
dim EnableUpload
dim arrUpFileType
dim ranNum
dim msg,FoundErr
dim PhotoUrlID
msg=""
FoundErr=false
EnableUpload=false

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/admin_css.css" rel="stylesheet" type="text/css">
<style type="text/css">
BODY{
BACKGROUND-COLOR: #F1F3F5;
}
</style>
<SCRIPT language="javascript" src="../js/Control.js" type="text/javascript"></SCRIPT>
</head>
<body leftmargin="2" topmargin="5" marginwidth="0" marginheight="0" >
<%
if EnableUploadFile="No" then
	response.write "系统未开放文件上传功能"
else

		select case upload_type
			case 0
				call upload_0()  '使用化境无组件上传类
			case else
				'response.write "本系统未开放插件功能"
				'response.end
		end select
	'end if
end if
%>
</body>
</html>
<%
sub upload_0()    '使用化境无组件上传类
	set upload=new upfile_class ''建立上传对象
	upload.GetData(10000*1024)   '取得上传数据,限制最大上传100M
	if upload.err > 0 then  '如果出错
		select case upload.err
			case 1
				response.write "请先选择你要上传的文件！"
			case 2
				response.write "你上传的文件总大小超出了最大限制（10M）"
		end select
		response.end
	end if
	
	orginsavepath=trim(upload.form("savepath"))
	picsize=trim(upload.form("picsize"))
	picshow=trim(upload.form("picshow"))
	picurl=trim(upload.form("picurl"))
	PicOperate=trim(upload.form("PicOperate"))
	
	
	SavePath = SaveShowPath & orginsavepath
	SaveTruePath = SaveSiteTruePath & orginsavepath

	if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
	if right(SaveTruePath,1)<>"/" then SaveTruePath=SaveTruePath & "/"	
	
	
'	response.Write(orginsavepath)
'	response.Write(picsize)
'	response.Write(picshow)
'	response.Write(picurl)
'	response.Write(PicOperate)
'	response.Write(SavePath)
'	response.Write(SaveTruePath)
	
	for each formName in upload.file '列出所有上传了的文件
	
		set ofile=upload.file(formName)  '生成一个文件对象
		
		oFileSize=ofile.filesize
		
		if oFileSize<100 then
			msg="请先选择你要上传的文件！"
			FoundErr=True
		else
			MaxFileSize = picsize	 
			if oFileSize>(MaxFileSize*1024) then
			 msg="文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
			 FoundErr=true
			end if	
		end if
		
		fileExt=lcase(ofile.FileExt)
		arrUpFileType=split(UpFileType,"|")
		
		for i=0 to ubound(arrUpFileType)
			if fileEXT=trim(arrUpFileType(i)) then
				EnableUpload=true
				exit for
			end if
		next
		
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" then
			EnableUpload=false
		end if
		
		if EnableUpload=false then
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
			FoundErr=true
		end if		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filemainname=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			fileshowname=SavePath&filemainname
			filesavename=SaveTruePath&filemainname
 
			ofile.SaveToFile filesavename '保存文件

			response.write("<span style=""forumRow"">图片上传成功！图片大小为：" & cstr(round(oFileSize/1024)) & "K</span>")
			
			strJS=strJS & "parent.document.getElementById('" & picurl & "').value='" & filemainname & "';" & vbcrlf
			strJS=strJS & "parent.document.getElementById('" & PicOperate & "').style.display='';" & vbcrlf
			strJS=strJS & "parent.document.getElementById('" & picshow & "').src='" & fileshowname & "';" & vbcrlf

		else
			strJS=strJS & "alert('" & msg & "');" & vbcrlf
		  	strJS=strJS & "history.go(-1);" & vbcrlf
		end if
		strJS=strJS & "</script>" & vbcrlf
		response.write(strJS)
		
		set file=nothing
	next
	
	set upload=nothing
end sub
%>
