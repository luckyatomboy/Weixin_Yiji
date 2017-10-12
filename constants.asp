<%
'变量前缀，如一个空间下多次使用本程序的话，请每个程序配置不同的值
dim prefix:prefix="yjfood"

'微信开发配置'
dim wx_appid:wx_appid="wx65008794da566f2d"
dim wx_secret:wx_secret="f9ba1bbdc5f714f78e84bbf34be4d97f"

'发送特定代码'
dim code_send_template: code_send_template="qftz"  '发送模板消息'群发通知
dim code_close_bp: code_close_bp="gbbp"  '关闭报盘
dim code_open_bp: code_open_bp="dkbp"	'打开报盘
dim code_check_token: code_check_token="token" '检查token'
dim code_check_bpstatus: code_check_bpstatus="bpstatus" '检查报盘status'
dim code_get_news_list: code_get_news_list="newslist" '取得图文消息列表'
dim code_test_mbxx: code_test_mbxx="mbxx" '发送模板消息'
dim code_get_tag_list: code_get_tag_list="taglist" '取得所有标签列表'
dim code_get_self_tag: code_get_self_tag="selftag" '取得自己的标签列表'
dim code_set_template_title: code_set_template_title="szbt" '设置模板消息标题'
dim code_set_template_brief: code_set_template_brief="szjj" '设置模板消息简介'

'Tag ID 一定要用数字，不能用字符串！
'dim tag_beef:tag_beef=100  '牛肉'
'dim tag_pork:tag_pork=101  '猪肉'
dim tag_blacklist:tag_blacklist=116  '黑名单
dim tag_admin:tag_admin=115  '管理员

'URL'
dim url_beef_future:url_beef_future="" '牛肉期货'
dim url_beef_spot:url_beef_spot="" '牛肉现货'

dim url_pork_future:url_pork_future="" '猪肉期货'
dim url_pork_spot:url_pork_spot="" '猪肉现货'

dim url_chicken_future:url_chicken_future="" '鸡肉期货'
dim url_chicken_spot:url_chicken_spot="" '鸡肉现货'

'模板消息ID：资料'
dim template_id:template_id="aUMruJGDxP5EWsOvzMXd8-g6kQ88ZW2Ck4MwDjfMy8Q"

'菜单Event Key'
dim event_beef_future:event_beef_future="beef_future" '牛肉期货'
dim event_beef_spot:event_beef_spot="beef_spot" '牛肉现货'

dim event_pork_future:event_pork_future="pork_future" '猪肉期货'
dim event_pork_spot:event_pork_spot="pork_spot" '猪肉现货'

dim event_pork_product:event_pork_product="pork_product" '零担猪肉'
dim event_beef_product:event_beef_product="beef_product" '零担牛肉'

'图文消息URL'
'牛肉期货'	done
dim news_beef_future:news_beef_future="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000111&idx=1&sn=90c7d0a34ed969c76de078d3134a6a23&chksm=6bb295025cc51c14168c81f0a71f46c58a2e415fc028dc60a042d81444bafe2adb331d48d4cf#rd"

'牛肉现货'	done
dim news_beef_spot:news_beef_spot="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000110&idx=1&sn=3db84e3c0fd0038359ae4199298831c7&chksm=6bb295035cc51c1586cad4da5e896bc5e0f2a98ae3236537a0e6a69b747eee53c9e81f458fed#rd"

'猪肉期货'	done
dim news_pork_future:news_pork_future="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000113&idx=1&sn=aa8bda9870453b92befe4fdfb04c1237&chksm=6bb2951c5cc51c0ab9270ec5c96017595951287e56a770680bea767f740ea3741f0125749a91#rd"

'猪肉现货'	done
dim news_pork_spot:news_pork_spot="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000112&idx=1&sn=3b7a795740e788228de78d8957bbd5f4&chksm=6bb2951d5cc51c0bedc1163207903c19f3399872c77e49a33e5ed970734192f873fb20b9586a#rd"

'零担猪肉'	done
dim news_pork_product:news_pork_product="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000115&idx=1&sn=fe7ba65458029198189eadd921bbbb51&chksm=6bb2951e5cc51c08ef68acaa862eb44e13a6184e46adcfbcf43c37addca0042d3275220b7354#rd"

'零担牛肉'  done
dim news_beef_product:news_beef_product="http://mp.weixin.qq.com/s?__biz=MzI4MDc0MTEyMg==&mid=100000114&idx=1&sn=d5b71da90566daa6b5a0ddd214ea9b08&chksm=6bb2951f5cc51c099552db74201e1312ac852b3ba415789baf83ca9580b570ab3bbcac6023a3#rd"

'Message	
dim msg_bpclose_hint: msg_bpclose_hint="本次报盘已结束，敬请期待下次报盘！"
dim msg_bpclose_success: msg_bpclose_success="报盘关闭成功！"
dim msg_bpopen_success: msg_bpopen_success="报盘打开成功！"
dim msg_empty_tag: msg_empty_tag="您没有权限查看报盘，请联系客服申请开通权限！"
dim msg_new_user: msg_new_user="欢迎您关注上海一骥！您现在还没有权限查看报盘，请联系客服申请开通权限！"
dim msg_template_title: msg_template_title="模板消息标题设置成功！"
dim msg_template_brief: msg_template_brief="模板消息简介设置成功！"

'图片url'
dim pic_pork: pic_pork="http://www.yj-food.com/weixin/pork.jpg"
dim pic_beef: pic_beef="http://www.yj-food.com/weixin/cow.jpg"
dim pic_chicken: pic_chicken="http://www.yj-food.com/weixin/chicken_single.jpg"

'描述'
dim desc_pork_future: desc_pork_future="猪鸡期货报盘已更新，快点我查看吧！"
dim desc_pork_spot: desc_pork_spot="猪鸡现货报盘已更新，快点我查看吧！"
dim desc_beef_future: desc_beef_future="牛羊期货报盘已更新，快点我查看吧！"
dim desc_beef_spot: desc_beef_spot="牛羊现货报盘已更新，快点我查看吧！"
dim desc_pork_product: desc_pork_product="零担散货猪肉报盘已更新，快点我查看吧！"
dim desc_beef_product: desc_beef_product="零担散货牛肉报盘已更新，快点我查看吧！"
'dim desc_chicken: desc_chicken="鸡肉报盘已更新，快点我查看吧！"

'Title'
dim title_pork_future: title_pork_future="猪鸡期货报盘"
dim title_pork_spot: title_pork_spot="猪鸡现货报盘"
dim title_beef_future: title_beef_future="牛羊期货报盘"
dim title_beef_spot: title_beef_spot="牛羊现货报盘"
dim title_pork_product: title_pork_product="零担散货猪肉报盘"
dim title_beef_product: title_beef_product="零担散货牛肉报盘"
'dim title_chicken: title_chicken="鸡肉报盘"

%>