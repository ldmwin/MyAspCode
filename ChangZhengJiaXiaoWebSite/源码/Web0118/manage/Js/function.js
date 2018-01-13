//var barstatus = 1;

/*调用页面，/login.asp*/
function get_Code() {
	var Dv_CodeFile = "getcode.asp";
	if(document.getElementById("imgid"))
		document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}

/*调用页面，/main.asp*/
function switchSysBar(){
     if (1 == window.barstatus){
		  window.barstatus = 0;
          switchPoint.innerHTML = '<img src="images/left.gif">';
          document.all("frmTitle").style.display="none"
     }
     else{
		  window.barstatus = 1;
          switchPoint.innerHTML = '<img src="images/right.gif">';
          document.all("frmTitle").style.display=""
     }
}

/*调用页面，/left.asp*/
function disp(n){
//隐藏所有导航的办法1，扩展性不足
//	for (var i=1;i<8;i++)
//	{
//		if (!document.getElementById("leftmenu"+i)) return;			
//		document.getElementById("leftmenu"+i).style.display="none";
//	}

//隐藏所有导航的办法2，div不支持name属性，无解
//	 var LeftMenuGroup = document.getElementsByName("leftmenu");
//	 alert(LeftMenuGroup.length);
//	for (var i=0;i<LeftMenuGroup.length;i++){
//		LeftMenuGroup[0].style.display="none";
//		alert("yes");
//	}

//隐藏所有导航的办法3，需要遍历所有div，仍然需要改善
	var divObjs =document.getElementsByTagName('div');
	for(var i=0; i<divObjs.length; i++){
	if(divObjs[i].getAttribute("name") == "leftmenu"){
	//divObjs[i].style.display !='none'?divObjs[i].style.display ='none':divObjs[i].style.display ='block';}
	divObjs[i].style.display = "none";}
}



	document.getElementById("leftmenu"+n).style.display="";
}

getElementsByName:function (name) { 　　
	var returns = document.getElementsByName(name); 　
	if(returns.length > 0) return returns; 　　
	returns = new Array(); 　
	var e = document.getElementsByTagName("td"); 　
	for(i = 0; i < e.length; i++) { 　　
		if(e[i].getAttribute("name") == name) { 　
		returns[returns.length] = e[i]; 
		} 　
	} 　　
	return returns; 　　
} 

function Ob(o){
	var o=document.getElementById(o)?document.getElementById(o):o;
	return o;
}
function Hd(o) {
	Ob(o).style.display="none";
}
function Sw(o) {
	Ob(o).style.display="";
}
function ExCls(o,a,b,n){
	var o=Ob(o);
	for(i=0;i<n;i++) {o=o.parentNode;}
	o.className=o.className==a?b:a;
}
function CNLTreeMenu(id,TagName0) {
	this.id=id;
	this.TagName0=TagName0==""?"li":TagName0;
	this.AllNodes = Ob(this.id).getElementsByTagName(TagName0);
	this.InitCss = function (ClassName0,ClassName1,ClassName2,ImgUrl) {
		this.ClassName0=ClassName0;
		this.ClassName1=ClassName1;
		this.ClassName2=ClassName2;
		this.ImgUrl=ImgUrl || "../images/s.gif";
		this.ImgBlankA ="<img src=\""+this.ImgUrl+"\" class=\"s\" onclick=\"ExCls(this,'"+ClassName0+"','"+ClassName1+"',1);\" alt=\"展开/折叠\" />";
		this.ImgBlankB ="<img src=\""+this.ImgUrl+"\" class=\"s\" />";
		for (i=0;i<this.AllNodes.length;i++ ) {
			this.AllNodes[i].className==""?this.AllNodes[i].className=ClassName1:"";
			this.AllNodes[i].innerHTML=(this.AllNodes[i].className==ClassName2?this.ImgBlankB:this.ImgBlankA)+this.AllNodes[i].innerHTML;
		}
	}
	this.SetNodes = function (n) {
		var sClsName=n==0?this.ClassName0:this.ClassName1;
		for (i=0;i<this.AllNodes.length;i++ ) {
			this.AllNodes[i].className==this.ClassName2?"":this.AllNodes[i].className=sClsName;
		}
	}
}

function divcontrol(itemid){
	if(document.getElementById(itemid).style.display=='none'){
		document.getElementById(itemid).style.display="";
	}
	else{
		document.getElementById(itemid).style.display="none";
	}
}
function divhidden(itemid){
	if(document.getElementById(itemid).style.display!='none'){
		document.getElementById(itemid).style.display="none";
	}
}

function viewPage(ipage){
	document.frm_page.me_page.value=ipage;
	document.frm_page.submit();        
}

function SelectAll()
{

	var input = document.getElementsByName("info_id");
	//var input = document.getElementById("info_id");
	var check;

	
	if(document.getElementById("chk_all").checked){
		check = true;
	}
	else{
		check = false;
	}
	
	for (var i=0;i<input.length ;i++ )
	{
		//alert(input[i].type);
		if(input[i].type=="checkbox"){
			input[i].checked = check;
			}
	}
	
}

function GotoUrl(Url,Target,msg){
	//alert(msg);
	if(msg!=''&&msg!=null&&typeof(msg)!="undefined"){
		//alert(msg);
		if(!confirm(msg)){
			return;
		}
		//alert("yes");
	}
	
//			
	if(Target == "_parent"){
		window.parent.location.href = Url; 	
		}
	else if(Target == "_top"){
		window.top.location.href = Url; 
		}
	else if(Target == "_back"){
		history.go(-1);
		}
	else{
		window.location.href = Url; 
		}
	
	}
	
//function OperateUrl(Url,Target,msg){
//	var alertmsg;
//	if(confirm(alertmsg)){
//		if(Target == "_parent"){
//			window.parent.location.href = Url; 	
//			}
//		else if(Target == "_top"){
//			window.top.location.href = Url; 
//			}
//			else if(Target == "_back"){
//			history.go(-1);
//			}
//		else{
//			window.location.href = Url; 
//			}
//		}
//	}	
	
function Refresh(Target){

	if(Target == "_parent"){
		window.parent.location.reload(); 	
		}
	else if(Target == "_top"){
		window.top.location.reload(); 
		}
	else if(Target == "_submit"){
		document.frm_search.submit();
		}
	else if(Target == "_left"){
		window.Categorymagleft.location.reload(); 
		}
	else{
		window.location.reload(); 

		}
	}
//记录状态的多参数页面切换
function GoToUrl_MorePrm(url,operation){
	//alert(operation);
	if (count_checked_items()>0){ 
		if(confirm(operation)){
			document.frm_list.action=url;
			document.frm_list.submit();
		}
	}
	else{
	  alert("请您先选择要操作的信息");
	  //return false;       
	}
}

function count_checked_items() {
	var number_checked=0;
	var box_count=document.frm_list.info_id.length;
	if ( box_count==null ) {
	if ( document.frm_list.info_id.checked==true ) {
	number_checked=1;}else {
	number_checked=0;}}
	else {
	for ( var i=0; i < (box_count); i++ ) {
	if ( document.frm_list.info_id[i].checked==true ) {
	number_checked++;}}}return number_checked;
}

function reupload(url,ope,file,fileiframe){
	document.getElementById(file).value="";
	document.getElementById(ope).style.display="none";
	document.getElementById(fileiframe).src=url;
	}
function PicView(ope,pic,picshow){
	if(ope==0){
		picshow.style.textDecoration="underline";
		if(pic!=null){
			document.getElementById(pic).style.display="block";
			}
		}
	else if(ope==1){
		picshow.style.textDecoration="none";
		if(pic!=null){
			document.getElementById(pic).style.display="none";
			}
		}
	}