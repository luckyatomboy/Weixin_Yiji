<%

function send_template_to_all_users(fromusername,tousername)
	bp_status=load_cache("bp_status")
	if bp_status="open" then
		dim access_token
		access_token=get_token()
		dim next_openid:next_openid=""
		dim url:url="https://api.weixin.qq.com/cgi-bin/user/get?access_token="&access_token&"&next_openid="
		set http=server.createobject(xmlhttp)
			http.open "POST",url,false
			http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			http.send()
			result=http.responseText
		dim nickname:nickname=get_nickname(fromusername)
			set http=nothing
			set json=toobject(result)
			for i=0 to json.data.openid.length-1
				userid=getitem(json.data.openid,i)
				'strmsg=get_nickname(fromusername,tousername,userid)&"/"&strmsg
				strmsg=post_template(userid,tousername,nickname)
			next
	else
		send_template_to_all_users=send_text(fromusername,tousername,msg_bpclose_hint)
	end if
end function

function bp_multi_items(fromusername,tousername,picurl,url)
	bp_status=load_cache("bp_status")
	if bp_status="open" then
		bp_multi_items="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[news]]></MsgType>" &_
		"<ArticleCount>2</ArticleCount>" &_
		"<Articles><item><Title><![CDATA[牛肉期货]]></Title>" &_
		"<Description><![CDATA[]]></Description>" &_
		"<PicUrl><!CDATA["&picurl&"]]></PicUrl>" &_
		"<Url><![CDATA[" & url & "]]></Url></item>" &_
		"<item><Title><![CDATA[牛肉现货]]></Title>" &_
		"<Description><![CDATA[]]></Description>" &_
		"<PicUrl><!CDATA["&picurl&"]]></PicUrl>" &_
		"<Url><![CDATA[" & url & "]]></Url></item></Articles>" &_	
		"</xml> "
	else
		bp_multi_items=send_text(fromusername,tousername,msg_bpclose_hint)
	end if
end function

function get_resourcelist(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/material/batchget_material?access_token="&access_token
	dim str:str="" 
		str=str&"{"
		str=str&"""type"":""news"","
		str=str&"""offset"":5,"
		str=str&"""count"":1"
		str=str&"}"
	dim http, strMsg, result, itemtext
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		'set json=toobject(result)
		'for i=0 to json.item.length-1
			'itemtext=getItemProperty(json.item,0,"name")
		'next

		'get_resourcelist=send_text(FromUserName, ToUserName, itemtext)
		get_resourcelist=send_text(FromUserName, ToUserName, result)
end function

function post_template(fromusername,tousername,byuser)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/message/template/send?access_token="&access_token
	dim str:str="" 
		str=str&"{"
		str=str&"""touser"":"
		str=str&""""&fromusername&""","
		str=str&"""template_id"":"""&template_id&""","
		str=str&"""url"":"""","		
		str=str&"""data"":{"
		str=str&"""first"":{""value"":""尊敬的用户，新的报盘已生成，请点击菜单查看"",""color"":""#173177""},"
		str=str&"""keyword1"":{""value"":""一骥报盘""},"
		str=str&"""keyword2"":{""value"":""牛猪鸡期货现货报盘""}"
		'str=str&"""remark"":{""value"":""请点击菜单查看各品种报价""}"
		str=str&"}}"
	dim http, strMsg, result
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
		set http=nothing
		post_template=""
end function

function set_black_list(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/batchblacklist?access_token="&access_token
	dim str
		str="{""openid_list"":["""&fromusername&"""]}"
	dim http, strMsg, result, json
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
	set json=ToObject(result)	
		set http=nothing
		set_black_list=send_text(fromusername,tousername,json.errmsg)
end function

function get_black_list(fromusername,tousername)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/getblacklist?access_token="&access_token
	dim str: str="{""begin_openid"":""""}"
	dim http, strMsg, result, json, blacklist, blacklist_count, nickname
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
	set json=ToObject(result)	
		set http=nothing
		blacklist_count=0
		'blacklist=""
	for i=0 to json.data.openid.length-1
		userid=getitem(json.data.openid,i)				
		'if check_blacklist_tag(userid)="true" then
			blacklist_count=blacklist_count+1
			if blacklist="" then
				blacklist=userid
			else
				blacklist=blacklist&"//"&userid
			end if
		'end if
	next		
	get_black_list=send_text(fromusername,tousername,blacklist_count&":"&blacklist)
	'get_black_list=send_text(fromusername,tousername,json.data.openid.length)
end function

function set_black_list_multiple(fromusername,tousername,users)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/batchblacklist?access_token="&access_token
	dim str, user_array
		str="{""openid_list"":["""
		user_array=split(users,"//")
		for i=0 to ubound(user_array)
			if i=0 then
				str=str&user_array(i)&""""    '' &users&"""]}"
			else
				str=str&","""&user_array(i)&""""
			end if
		next
		str=str&"]}"
		
	dim http, strMsg, result, json
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		result=http.responseText
	set json=ToObject(result)	
		set http=nothing
		set_black_list=send_text(fromusername,tousername,json.errmsg)
end function

function get_user_information(userid)
	dim access_token: access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid="&userid
	dim http
	set http=server.createobject(xmlhttp)
	http.open "GET",url,false
	http.send()
	get_user_information=http.responsetext
	set http=nothing
end function

function get_nickname(userid)
	dim user_info: user_info=get_user_information(userid)
	dim json,nickname
	set json=ToObject(user_info)
	nickname=json.nickname
	'get_nickname=userid&":"&nickname
	get_nickname=nickname
end function

function check_admin_tag(userid)
	dim user_info: user_info=get_user_information(userid)
	dim json,admin_flag,tagid
	admin_flag="false"
	set json=ToObject(user_info)
	for i=0 to json.tagid_list.length-1
		tagid=getitem(json.tagid_list,i)	
		if tagid=tag_admin then
			admin_flag="true"
			exit for
		end if
	next
	check_admin_tag=admin_flag
end function

function check_empty_tag(userid)
	dim user_info: user_info=get_user_information(userid)
	dim json,empty_flag,tagid
	empty_flag="false"
	set json=ToObject(user_info)
	if json.tagid_list.length=0 then
		empty_flag="true"
	end if
	check_empty_tag=empty_flag
end function

function check_blacklist_tag(userid)
	dim user_info: user_info=get_user_information(userid)
	dim json,blacklist_flag,tagid
	blacklist_flag="false"
	set json=ToObject(user_info)
	for i=0 to json.tagid_list.length-1
		tagid=getitem(json.tagid_list,i)	
		if tagid=tag_blacklist then
			blacklist_flag="true"
			exit for
		end if
	next
	check_blacklist_tag=blacklist_flag
end function

function get_self_tag(fromusername,tousername)
	dim user_info: user_info=get_user_information(fromusername)
	dim json,tagid
	set json=ToObject(user_info)
	'for i=0 to json.tagid_list.length-1
		'tagid=getitem(json.tagid_list,i)	
		'if tagid=tag_admin then
			'check_admin_tag="true"
			'exit for
		'end if
	'next
	get_self_tag=send_text(fromusername,tousername,json.tagid_list)
	'get_self_tag=send_text(fromusername,tousername,check_blacklist_tag(fromusername))
end function

function bp(fromusername,ToUserName,title,description,picurl,url)
	bp_status=load_cache("bp_status")
	'bp_status=load_cache_sysdb("bp_status")
	if bp_status="open" then
  		bp=send_news(FromUserName,ToUserName,title,description,picurl,url)
  	else
		bp=send_text(fromusername,tousername,msg_bpclose_hint)
  	end if
end function

function closebp(fromusername,ToUserName)
  dim nickname:nickname=get_nickname(fromusername)
  set_cache "bp_status","close"
  set_cache "bpstatus_date",now()
  set_cache "bpstatus_changer",nickname
  closebp=send_text(fromusername,ToUserName,msg_bpclose_success)
end function	

function openbp(fromusername,ToUserName)
  dim nickname:nickname=get_nickname(fromusername)
  set_cache "bp_status","open"
  set_cache "bpstatus_date",now()
  set_cache "bpstatus_changer",nickname
  openbp=send_text(fromusername,ToUserName,msg_bpopen_success)
end function

function openbp_sysdb(fromusername,ToUserName)
  dim nickname:nickname=get_nickname(fromusername)
  set_cache_sysdb "bp_status","open"
  set_cache_sysdb "bpstatus_date",now()
  set_cache_sysdb "bpstatus_changer",nickname
  openbp_sysdb=send_text(fromusername,ToUserName,msg_bpopen_success)
  'openbp_sysdb=send_text(fromusername,ToUserName,msg_bpopen_success&"//"&load_cache_sysdb("bp_status")&"//"&load_cache_sysdb("bpstatus_date"))
end function

function closebp_sysdb(fromusername,ToUserName)
  dim nickname:nickname=get_nickname(fromusername)
  set_cache_sysdb "bp_status","close"
  set_cache_sysdb "bpstatus_date",now()
  set_cache_sysdb "bpstatus_changer",nickname
  closebp_sysdb=send_text(fromusername,ToUserName,msg_bpclose_success)
end function	

function get_tag_list(fromusername,ToUserName)
	dim access_token
	access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/tags/get?access_token="&access_token
	dim http, strMsg, result
	set http=server.createobject(xmlhttp)
		http.open "GET",url,false
		http.send()
		result=http.responsetext
		get_tag_list=send_text(fromusername,ToUserName,result)
end function



%>