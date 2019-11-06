$From = 'Jay.Wang@charteredaccountantsanz.com'
$To = 'jay.wang@charteredaccountantsanz.com'
$Subject = 'Failed to restart YSoft services'
$Body = 'YSoft services cannot be restarted. Please check the PRNT server'
$SMTPServer = 'ecp-exch-v01'
$MailMsg = New-Object System.Net.Mail.MailMessage
$MailMsg.From = $From
$MailMsg.To.Add($To)
$MailMsg.Subject = $Subject
$MailMsg.Body = $Body
$SMTP = New-Object Net.Mail.SMTPclient($SMTPServer)
$SMTP.Send($MailMsg)