import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders("ethan@westwall.vip").Folders("收件箱")

outlook.SendAndReceive(False)
messages = inbox.Items
count = len(messages)

for message in messages:
	# 回复状态码 103 已回复 0 未回复
	reply_stats = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
	# 已读状态码
	read_stats = message.UnRead
	sender = message.Sender
	sendemailaddr = message.SenderEmailAddress
	subject = message.Subject
	body = message.Body
	receive_time = message.ReceivedTime


	# 保存
	# message.SaveAs("C:\\DESKTOP\\test.msg")

	# 回复全部
	# replyall_email = message.ReplyAll()
	# replyall_email.Body = "TEXT"
	# replyall_email.Send()
