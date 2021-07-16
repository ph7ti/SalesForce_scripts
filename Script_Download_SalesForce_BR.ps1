#Script para acesso e Download de relatório no Sales Force
#Obs.: O mesmo pode ser alterado para acesso em outras páginas, uma vez que faz interações pseudo-humanas
##Altere seu usuário e senha
$username = "user@domain.com" 
$password = "MyPasswd"
$ie = New-Object -com InternetExplorer.Application
$ie.visible=$true
$obj = new-object -com WScript.Shell
$obj.AppActivate('Internet Explorer')
Start-Sleep -Milliseconds 1000;
##Altere o domínio do SalesForce
$ie.navigate("https://mysalesforcedomain.my.salesforce.com/")
do{Start-Sleep -Milliseconds 1000}While($ie.Busy -eq $True)
do{Start-Sleep -Milliseconds 1000}While($ie.ReadyState -ne 4)
$usernamefield = $ie.document.getElementByID("username")
$passwordfield = $ie.document.getElementByID("password")
$usernamefield.value = "$username"
$passwordfield.value = "$password"
$login = $ie.document.getElementsByName("Login") | ? {$_.value -eq "Fazer login"}
$login.click()
Start-Sleep -Milliseconds 4000;
##Altere o link para Download do relatório; O link atual efetua o download em CSV no formato UTF-8
$ie.Navigate("https://mysalesforcedomain.my.salesforce.com/00O0z000005FupyEAC?csv=1&exp=1&enc=UTF-8&isdtp=p1")
Start-Sleep -Milliseconds 3000;
$obj.SendKeys('%s')
Start-Sleep -Milliseconds 3000;
$ie.Quit()