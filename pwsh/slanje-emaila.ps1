# generisanje naziva datoteke

$Dir = "$env:HOME/kod/tmp/"
$Pref = 'ooo'
$Eks = '.txt'
$Dan = (Get-Date).DayOfYear

# smjesti rezultat poziva procedure u datoteku
$cmd = 'sqlcmd'
$prm = '-S .', '-Usa', '-PPass1!', '-dAdventureWorks', `
'-Q EXEC [dbo].[uspGetManagerEmployees] @BusinessEntityID = 2',`
"-o$Dir$Pref$Dan$Eks"

& $cmd $prm

# email postavke
$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
$Username = "userfrom@gmail.com"
$Password = "Pass!"

$to = "userto@gmail.com"
$cc = ""
$subject = "Naslov"
$body = "Tijelo"
# $attachment = $Dir + $Pref + $Dan + $Eks

$f = "$Pref$Dan*$Eks"
$attachments = @( gci -Path $Dir -Filter $f | % { $_.FullName } -ErrorAction Stop)

if ($attachments.Count -eq 0) {
    exit 1
}

$message = New-Object System.Net.Mail.MailMessage
$message.subject = $subject
$message.body = $body
$message.to.add($to)
# $message.cc.add($cc)
$message.from = $username

# $message.attachments.add($attachment)

$attachments | % { $message.attachments.add($_) }

$smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
$smtp.EnableSSL = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
$smtp.send($message) 

# Pomjeriti datoteku nakon slanja
$arhiva = $dir + "arhiva"
(gci -Path $Dir -Filter $f) | move -Destination $arhiva

write-host "mail poslan"

exit 0
