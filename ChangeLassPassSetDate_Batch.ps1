$UserCSV = Import-Csv -Path "C:\PS\PasswordBatchChange\passchange_2007.csv" -Delimiter ";" -Encoding UTF8

    foreach ($i in $UserCSV)
        {
           
            Set-ADUser  $i.SamAccountName  -Replace @{pwdLastSet='0'}
            Set-ADUser  $i.SamAccountName  -Replace @{pwdLastSet='-1'}
        }