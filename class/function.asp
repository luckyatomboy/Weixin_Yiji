<%

function get_all_users()
		dim access_token
		access_token=get_token()
		dim url:url="https://api.weixin.qq.com/cgi-bin/user/get?access_token="&access_token '&" &next_openid="
		set http=server.createobject(xmlhttp)
			http.open "POST",url,false
			http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			http.send()
			result=http.responseText
			set http=nothing
			get_all_users=result
end function

function send_template_to_all_users(fromusername,tousername)
	'dim white:white=0
	bp_status=load_cache("bp_status")
	if bp_status="open" then
		result=get_all_users()
		blacklist=get_black_list()
		set json=toobject(result)
'给所有人发模板消息'
		for i=0 to json.data.openid.length-1
			userid=getitem(json.data.openid,i)
			if instr(blacklist,userid) = 0 then '检查是否是黑名单上的人，如果是就不发'
				strmsg=post_template(userid,tousername,"")	
			end if
		next
		'send_template_to_all_users=send_text(fromusername,tousername,white)
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
	'response.write ""
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
		str=str&"""first"":{""value"":"""&load_cache("template_title")&"\n\r"",""color"":""#173177""},"
		'str=str&"""first"":{""value"":""尊敬的客户，新的报盘已生成，请点击菜单查看\n\r"",""color"":""#173177""},"
		str=str&"""keyword1"":{""value"":""上海一骥报盘""},"
		str=str&"""keyword2"":{""value"":"""&load_cache("template_brief")&"""},"
		str=str&"""remark"":{""value"":""\n\r欢迎“讨价还价”"",""color"":""#FF0000""}"
		str=str&"}}"
	dim http, strMsg, result
	set http=server.createobject(xmlhttp)
		http.open "POST",url,false
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(str)
		'result=http.responseText
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



function get_user_information(userid)
	dim access_token: access_token=get_token()
	dim url:url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid="&userid
	dim http
	set http=server.createobject(xmlhttp)
	http.open "GET",url,false
	http.send()
	get_user_information=http.responsetext
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

function set_template_title(fromusername,ToUserName,title)
	dim nickname:nickname=get_nickname(fromusername)
	set_cache "template_title",title
	'set_cache "template_title_date",now()
	'set_cache "template_title_changer",nickname
	set_template_title=send_text(fromusername,ToUserName,msg_template_title)
end function

function set_template_brief(fromusername,ToUserName,brief)
	dim nickname:nickname=get_nickname(fromusername)
	set_cache "template_brief",brief
	'set_cache "template_brief_date",now()
	'set_cache "template_brief_changer",nickname
	set_template_brief=send_text(fromusername,ToUserName,msg_template_brief)
end function

function set_black_list_multiple(openid_list,direction)
	dim access_token
	access_token=get_token()
	dim url_black:url_black="https://api.weixin.qq.com/cgi-bin/tags/members/batchblacklist?access_token="&access_token
	dim url_unblack:url_unblack="https://api.weixin.qq.com/cgi-bin/tags/members/batchunblacklist?access_token="&access_token	
	dim http, result, json
	set http=server.createobject(xmlhttp)
		if direction="B" then
			http.open "POST",url_black,false
		else
			http.open "POST",url_unblack,false
		end if
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send(openid_list)
		result=http.responseText
	set json=ToObject(result)	
		set http=nothing
		set_black_list_multiple=json.errmsg
end function

function set_black_list_by_tag(fromusername,tousername,tagid,direction)
	dim str, users
	dim strMsg, result, json, new_batch, remain
		if tagid=0 then
			users=get_all_users()
		else
			users=get_users_by_tag(tagid)
		end if
		set json=ToObject(users)
		str="{""openid_list"":["""	
		for i=0 to json.data.openid.length-1
			userid=getitem(json.data.openid,i)
			if i=0 or new_batch="true" then
				str=str&userid&""""    '' &users&"""]}"
				new_batch=""
			else
				str=str&","""&userid&""""
				if (i mod 19)=0 then
					str=str&"]}"
					result=result&"//20_"&set_black_list_multiple(str,direction)
					if i<json.data.openid.length-1 then
						str="{""openid_list"":["""	
						new_batch="true"
					else
						str=""
						new_batch=""
					end if
				end if
			end if
		next
		if str<>"" then
			str=str&"]}"
			remain=json.data.openid.length mod 20
			result=result&"//"&remain&"_"&set_black_list_multiple(str,direction)
			str=""
		end if		
		
		set_black_list_by_tag=send_text(fromusername,tousername,result)
end function

function set_tag_to_users(fromusername,ToUserName,tagid)
		blacklist=get_black_list()
		dim access_token
		access_token=get_token()
		dim url:url="https://api.weixin.qq.com/cgi-bin/tags/members/batchtagging?access_token="&access_token
		dim str: str=""
			str="{""openid_list"":["&blacklist&"],""tagid"":"&tagid&"}"   '传入的用户数不能超过50个'
		set http=server.createobject(xmlhttp)
			http.open "POST",url,false
			http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			http.send(str)
			result=http.responseText
			set_tag_to_users=send_text(fromusername,ToUserName,result)
end function

function get_users_by_tag(tagid)
		dim access_token
		access_token=get_token()
		dim url:url="https://api.weixin.qq.com/cgi-bin/user/tag/get?access_token="&access_token
		dim str: str=""
			str="{""tagid"":"&tagid&",""next_openid"":""""}"   
		set http=server.createobject(xmlhttp)
			http.open "POST",url,false
			http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			http.send(str)
			get_users_by_tag=http.responseText
end function

function get_black_list()
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
		if blacklist="" then
			blacklist=""""&userid&""""
		else
			blacklist=blacklist&","""&userid&""""
		end if
	next		
	get_black_list=blacklist
end function


%>