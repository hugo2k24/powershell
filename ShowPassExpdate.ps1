$Userlogin = "panteleeva.loren"
$user= get-aduser -Identity $Userlogin -Properties msDS-UserPasswordExpiryTimeComputed
([DateTime]::FromFileTime($user.“msDS-UserPasswordExpiryTimeComputed”))