function inputStyle(fEvent,oInput){
	if (!oInput.style) return;
	var put=oInput.getAttribute("type").toLowerCase();

	switch (fEvent){
		case "focus" :
			oInput.isfocus = true;
		case "mouseover" :			
			if(put=="submit" || put=="button" || put=="reset")
					oInput.className="input_on";
			else{
//				alert(oInput.className);
//				if(oInput.className != "Wdate"){
				oInput.className = "TextBoxFocus";
//				}
			}
			break;
		case "blur" :
			oInput.isfocus = false;
		case "mouseout" :
			if(put=="submit" || put=="button" || put=="reset")
				oInput.className = "input0";
		    else if(!oInput.isfocus)
//				if(oInput.className != "Wdate"){
				oInput.className = "TextBox";
//				}
			break;
		//case else :
			//if(oInput.getAttribute(fEvent+"_2"))
				//eval(oInput.getAttribute(fEvent+"_2"));
	}	
}

window.onload = function(){
	var oInput = document.getElementsByTagName("input");
	var onfocusStr = [];
	var onblurStr = [];
	//alert(oInput.length);
	try
	{
		for (var i=0; i<oInput.length; i++)
		{
			if (!oInput[i]||!oInput[i].getAttribute("type")) continue;
			var put=oInput[i].getAttribute("type").toLowerCase();
			if(put=="submit" || put=="button" || put=="reset")
			{
				oInput[i].className="input0";
			}
			if (put=="text" || put=="password" || put=="submit" || put=="button" || put=="reset")
			{
				if(oInput[i].className != "Wdate"){	//日期选择控件不变更其样式			
				
					if (document.all)
					{
						oInput[i].attachEvent("onmouseover",oInput[i].onmouseover=function(){inputStyle("mouseover",this);});
						oInput[i].attachEvent("onmouseout",oInput[i].onmouseout=function(){inputStyle("mouseout",this);});
	
					}
					else{
						oInput[i].addEventListener("onmouseover",oInput[i].onmouseover=function(){inputStyle("mouseover",this);},false);
						oInput[i].addEventListener("onmouseout",oInput[i].onmouseout=function(){inputStyle("mouseout",this);},false);				
						//ȡ
						if(oInput[i].getAttribute("onfocus")){
							oInput[i].addEventListener("onfocus",oInput[i].onblur=function(){eval(this.getAttribute("onfocus"));inputStyle("focus",this);},false);
						}else{
							oInput[i].addEventListener("onfocus",oInput[i].onfocus=function(){inputStyle("focus",this);},false);
						}
						//ʧȥ
						if(oInput[i].getAttribute("onblur")){
							oInput[i].addEventListener("onblur",oInput[i].onblur=function(){eval(this.getAttribute("onblur"));inputStyle("blur",this);},false);
						}else{
							oInput[i].addEventListener("onblur",oInput[i].onblur=function(){inputStyle("blur",this);},false);
						}
					}
				}
			}
		}
	}catch(e){}
	for(i=1;i<=8;i++)//
	{
		if(document.getElementById('con_two_'+i))
		{	
			document.getElementById('two'+i).className="hover";			
			break;
		}
	}
}