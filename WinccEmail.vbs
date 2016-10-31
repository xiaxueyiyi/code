Dim wmi,xxx,u,ReportPath,BackupPath
Dim NameSpace,Email,x,y,fso,myfile,c,fs,f,uu
Set wmi=GetObject("winmgmts://./root/cimv2")
Set xxx=wmi.ExecQuery("Select * From Win32_PingStatus Where Address='www.baidu.com'")

	'***************************************
	ReportPath="C:\test"'报表路径
	BackupPath="C:\aaa\"'报表备份路径路径要存在
	'***************************************
	
For Each u In xxx
    If u.statuscode = 0 Then
    HMIRuntime.Trace "以太网连接正常!" & vbCrlf
	NameSpace = "http://schemas.microsoft.com/cdo/configuration/"'这个必须有，应该是VBS脚本链接微软网站获取某些支持应用的，删除的话vbs脚本会报错！
	Set Email = CreateObject("CDO.Message")'调用vbs邮件接口
	Email.From = "xiaxueyiyi@163.com" '发信人地址
	Email.To = "13137864@qq.com" '收信人地址（qq邮箱也可）
    'HMIRuntime.Trace "填写收信人地址" & vbCrlf
	Email.Subject = "报表文件-" & Date() '邮件主题

	Set fso=CreateObject("scripting.filesystemobject")
	Set fs=fso.GetFolder(ReportPath)
	Set f=fs.files
	If f.count<>0 Then
	For Each uu In f
	Email.AddAttachment uu.Path
	Next
	
	With Email.Configuration.Fields
	.Item(NameSpace&"sendusing") = 2
	.Item(NameSpace&"smtpserver") = "smtp.163.com" '这是163邮箱服务器地址，qq邮箱等请自行填写smtp地址
	.Item(NameSpace&"smtpserverport") = 25
	.Item(NameSpace&"smtpauthenticate") = 1
	.Item(NameSpace&"sendusername") = "xiaxueyiyi" '发信人用户名
	.Item(NameSpace&"sendpassword") = "130102" '发信人密码!
	.Update
	End With
	HMIRuntime.Trace "发送邮件！" & vbCrlf
	Email.Send
	HMIRuntime.Trace "发送完成！" & vbCrlf
	fso.MoveFile fs.path &"\*.*", BackupPath
	HMIRuntime.Trace "报表文件已转移到备份路径" & BackupPath & vbCrlf
	Set Email=Nothing
	Set fso=Nothing
	Else
	HMIRuntime.Trace "报表文件夹为空！" & vbCrlf
	End If
    Else
    HMIRuntime.Trace "以太网连接异常,停止发送邮件!" & vbCrlf
    End If
Next