// JavaScript Document
/*---------------------------------------- 
文件上传前台控制检测程序 v0.6 

远程图片检测功能 
检测上传文件类型 

　检测图片文件格式是否正确 
　检测图片文件大小 
　检测图片文件宽度 
　检测图片文件高度 
图片预览 

For 51js.com Author:333 Date:2005/08/26 
UpDate:2005/09/03 
-----------------------------------------*/ 

var ImgObj=new Image(); //建立一个图像对象 

var AllImgExt= ".jpg|.jpeg|.gif|.bmp|.png|"//全部图片格式类型 
var FileObj,ImgFileSize,ImgWidth,ImgHeight,FileExt,ErrMsg,FileMsg,HasCheked,IsImg,viewpic//全局变量 图片相关属性 

//以下为限制变量 
//var AllowExt= ".jpg|.gif|.doc|.txt|" //允许上传的文件类型 &#320;为无限制 每个扩展名后边要加一个 "| " 小写字母表示
var AllowExt= ".jpg|.png|"
//var AllowExt=0 
var AllowImgFileSize=70; //允许上传图片文件的大小 0为无限制 单位：KB 
var AllowImgWidth=450; //允许上传的图片的宽度 &#320;为无限制　单位：px(像素) 
var AllowImgHeight=450; //允许上传的图片的高度 &#320;为无限制　单位：px(像素) 

HasChecked=false; 

function CheckProperty(obj) //检测图像属性 
{ 
FileObj=obj; 
if(ErrMsg!= "") //检测是否为正确的图像文件　返回出错信息并重置 
{ 
ShowMsg(ErrMsg,false); 
return false; //返回 
} 

if(ImgObj.readyState!= "complete") //如果图像是未加载完成进行循环检测 
{ 
setTimeout( "CheckProperty(FileObj) ",500); 
return false; 
} 

ImgFileSize=Math.round(ImgObj.fileSize/1024*100)/100;//取得图片文件的大小 
ImgWidth=ImgObj.width //取得图片的宽度 
ImgHeight=ImgObj.height; //取得图片的高度 
FileMsg= "\n图片大小:"+ImgWidth+ "*"+ImgHeight+ "px"; 
FileMsg=FileMsg+ "\n图片文件大小:"+ImgFileSize+ "Kb"; 
FileMsg=FileMsg+ "\n图片文件扩展名:"+FileExt; 

//if(AllowImgWidth!=0&&AllowImgWidth <ImgWidth)
if(AllowImgWidth!=0&&AllowImgWidth !=ImgWidth)
ErrMsg=ErrMsg+ "\n图片宽度未达到限制。请上传宽度为"+AllowImgWidth+ "px的文件，当前图片宽度为"+ImgWidth+ "px"; 

if(AllowImgHeight!=0&&AllowImgHeight <ImgHeight) 
ErrMsg=ErrMsg+ "\n图片高度超过限制。请上传高度小于"+AllowImgHeight+ "px的文件，当前图片高度为"+ImgHeight+ "px"; 

if(AllowImgFileSize!=0&&AllowImgFileSize <ImgFileSize) 
ErrMsg=ErrMsg+ "\n图片文件大小超过限制。请上传小于"+AllowImgFileSize+ "KB的文件，当前文件大小为"+ImgFileSize+ "KB"; 

if(ErrMsg!= "") 
ShowMsg(ErrMsg,false); 
else 
ShowMsg(FileMsg,true); 
} 

ImgObj.onerror=function(){
	ErrMsg= "\n图片格式不正确或者图片已损坏!"; 
	ShowMsg(ErrMsg,false); 
	} 

function ShowMsg(msg,tf) //显示提示信息 tf=true 显示文件信息 tf=false 显示错误信息 msg-信息内容 
{ 
if(!tf) 
{ 
document.getElementById("Submit").disabled=true; 
FileObj.outerHTML=FileObj.outerHTML; 
//document.getElementById("MsgList").innerHTML=msg; 
alert(msg);
HasChecked=false; 
} 
else 
{ 
	document.getElementById("Submit").disabled=false; 
	if(IsImg){ 
	//document.getElementById("PreviewImg").innerHTML= " <img src= ' "+ImgObj.src+ " ' width= '60 ' height= '60 '> " 
	parent.document.getElementById(viewpic).src= ImgObj.src;
	parent.document.getElementById(viewpic).style.width = ImgObj.width;
	parent.document.getElementById(viewpic).style.height = ImgObj.height;
	parent.document.getElementById(viewpic).style.display="block";
	}
	else{ 
	alert(msg);
	HasChecked=true; 
	} 
} 
}

function CheckExt(obj) 
{ 
//obj=document.getElementById("FileName");
//alert(obj);
ErrMsg= ""; 
FileMsg= ""; 
FileObj=obj; 
IsImg=false; 
HasChecked=false; 
if(obj.value=="")return false; 
FileExt=obj.value.substr(obj.value.lastIndexOf( ".")).toLowerCase(); 
if(AllowExt!=0&&AllowExt.indexOf(FileExt+ "|")==-1) //判断文件类型是否允许上传 
{ 
ErrMsg= "\n该文件类型不允许上传。请上传 "+AllowExt+ " 类型的文件，当前文件类型为 "+FileExt; 
ShowMsg(ErrMsg,false); 
return false; 
} 

if(AllImgExt.indexOf(FileExt+ "|")!=-1) //如果图片文件，则进行图片信息处理 
{ 
IsImg=true; 
ImgObj.src=obj.value; 
CheckProperty(obj); 
return false; 
} 
else 
{ 
FileMsg= "\n文件扩展名: "+FileExt; 
ShowMsg(FileMsg,true); 
} 

} 

function check() 
{
	var strFileName=document.form1.FileName.value;
	if (strFileName=="")
	{
    	alert("请选择要上传的文件");
		document.form1.FileName.focus();
    	return false;
  	}
	else{
		parent.document.getElementById(viewpic).style.display="none";
	}
}
