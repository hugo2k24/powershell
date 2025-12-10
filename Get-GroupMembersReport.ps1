<#
.SYNOPSIS
    Генерирует детальный HTML-отчет о членах указанных групп Active Directory.
.DESCRIPTION
    Создает профессиональный HTML-отчет с информацией о членах групп AD.
    Включает: прямых членов, вложенные группы с раскрытием их членов,
    даты добавления (при включенном аудите), контактную информацию.
    Отчет может быть отправлен по email или сохранен локально.
.PARAMETER GroupNames
    Массив имен групп для анализа. Можно использовать шаблоны (*).
.PARAMETER GroupDNs
    Массив DistinguishedName групп для анализа.
.PARAMETER OUs
    Подразделения, из которых нужно взять все группы.
.PARAMETER ReportPath
    Путь для сохранения HTML-отчета. По умолчанию: текущая директория.
.PARAMETER SendEmail
    Отправлять отчет по электронной почте.
.PARAMETER EmailTo
    Адреса получателей отчета (через запятую).
.PARAMETER EmailFrom
    Адрес отправителя (по умолчанию: AD-отчеты@домен).
.PARAMETER SmtpServer
    SMTP-сервер для отправки (по умолчанию: реле домена).
.PARAMETER IncludeNested
    Раскрывать членов вложенных групп (рекурсивно).
.PARAMETER IncludeUserDetails
    Включать подробную информацию о пользователях (телефон, отдел, должность).
.PARAMETER IncludeInactive
    Включать неактивных пользователей (давно не входивших в систему).
.PARAMETER InactiveDays
    Порог неактивности в днях (по умолчанию: 90).
.PARAMETER DC
    Контроллер домена для запросов.
.PARAMETER ExportCSV
    Дополнительно экспортировать данные в CSV.
.EXAMPLE
    Get-GroupMembersReport.ps1 -GroupNames "Администраторы домена", "Администраторы предприятия"
    Отчет о членах критических групп безопасности.
.EXAMPLE
    Get-GroupMembersReport.ps1 -OUs "OU=Security Groups,DC=contoso,DC=local" -IncludeNested -ReportPath "C:\Audit\"
    Отчет по всем группам из указанного OU с раскрытием вложенности.
.EXAMPLE
    Get-GroupMembersReport.ps1 -GroupNames "SG_Finance_*"
.NOTES
    Author: Dmitry Plotinsky
    Version: 3.2
    Date: 2024-03-15
    Требует: Модуль ActiveDirectory, права на чтение в AD
#>

[CmdletBinding(DefaultParameterSetName = "ByName")]
param(
    [Parameter(ParameterSetName = "ByName", Mandatory=$true, Position=0)]
    [string[]]$GroupNames,
    
    [Parameter(ParameterSetName = "ByDN", Mandatory=$true)]
    [string[]]$GroupDNs,
    
    [Parameter(ParameterSetName = "ByOU", Mandatory=$true)]
    [string[]]$OUs,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = ".\AD_Group_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    
    [Parameter(Mandatory=$false)]
    [switch]$SendEmail,
    
    [Parameter(Mandatory=$false)]
    [string[]]$EmailTo,
    
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom = "ad-reports@$((Get-ADDomain).DNSRoot)",
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeNested,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeUserDetails,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeInactive,
    
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false)]
    [string]$DC,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportCSV
)

#  Инициализация и проверки
$StartTime = Get-Date
Write-Host "`n=== Генерация отчета о членах групп AD ===" -ForegroundColor Cyan
Write-Host "Время начала: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray

# Проверка и импорт модуля ActiveDirectory
if (-not (Get-Module -Name ActiveDirectory)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Verbose "Модуль ActiveDirectory загружен"
    }
    catch {
        Write-Error "Не удалось загрузить модуль ActiveDirectory. Установите RSAT Tools."
        return
    }
}

# Настройка параметров AD-запросов
$ADParams = @{}
if ($DC) { $ADParams.Server = $DC }

<#
# Проверка параметров email
if ($SendEmail -and (-not $EmailTo -or $EmailTo.Count -eq 0)) {
    Write-Warning "Параметр -SendEmail указан, но не заданы получатели (-EmailTo)."
    Write-Host "Отчет будет сохранен локально: $ReportPath" -ForegroundColor Yellow
    $SendEmail = $false
}
#>

#  Вспомогательные функции
function Get-ADGroupWithFallback {
    param([string]$Identity)
    
    try {
        return Get-ADGroup -Identity $Identity @ADParams -Properties *
    }
    catch {
        Write-Warning "Группа не найдена по имени: $Identity. Попытка поиска по шаблону..."
        $foundGroups = Get-ADGroup -Filter "Name -like '$Identity'" @ADParams -Properties *
        if ($foundGroups) {
            return $foundGroups
        }
        else {
            Write-Error "Группа '$Identity' не найдена"
            return $null
        }
    }
}

function Get-GroupMembersRecursive {
    param(
        [string]$GroupDN,
        [int]$Level = 0,
        [System.Collections.ArrayList]$Visited = (New-Object System.Collections.ArrayList)
    )
    
    # Защита от циклических ссылок
    if ($Visited -contains $GroupDN) {
        Write-Verbose "Обнаружен цикл: $GroupDN"
        return @()
    }
    $null = $Visited.Add($GroupDN)
    
    $members = @()
    
    try {
        $groupMembers = Get-ADGroupMember -Identity $GroupDN @ADParams -Recursive:$false
        
        foreach ($member in $groupMembers) {
            $memberType = $member.objectClass
            $memberDetails = $null
            
            # Получаем дополнительные свойства в зависимости от типа объекта
            switch ($memberType) {
                "user" {
                    $props = @("DisplayName", "SamAccountName", "UserPrincipalName", "Enabled", 
                              "LastLogonDate", "Created", "Department", "Title", "TelephoneNumber", 
                              "Manager", "EmailAddress")
                    if ($IncludeUserDetails) {
                        $props += "Office", "StreetAddress", "City", "PostalCode", "Company"
                    }
                    
                    $memberDetails = Get-ADUser -Identity $member.SID.Value @ADParams -Properties $props -ErrorAction SilentlyContinue
                    
                    # Проверка активности
                    $isActive = $true
                    if ($memberDetails.LastLogonDate) {
                        $daysInactive = (New-TimeSpan -Start $memberDetails.LastLogonDate -End (Get-Date)).Days
                        $isActive = $daysInactive -le $InactiveDays
                    }
                    
                    if (-not $IncludeInactive -and -not $isActive) {
                        Write-Verbose "Пропущен неактивный пользователь: $($memberDetails.SamAccountName)"
                        continue
                    }
                    
                    $members += [PSCustomObject]@{
                        ObjectClass    = "User"
                        Name           = $memberDetails.DisplayName ?? $member.Name
                        SamAccountName = $memberDetails.SamAccountName
                        UPN            = $memberDetails.UserPrincipalName
                        DN             = $member.DistinguishedName
                        SID            = $member.SID.Value
                        Enabled        = $memberDetails.Enabled
                        LastLogon      = $memberDetails.LastLogonDate
                        Department     = $memberDetails.Department
                        Title          = $memberDetails.Title
                        Email          = $memberDetails.EmailAddress
                        Phone          = $memberDetails.TelephoneNumber
                        Manager        = if ($memberDetails.Manager) { 
                            (Get-ADUser -Identity $memberDetails.Manager @ADParams -Properties DisplayName).DisplayName 
                        }
                        IsNested       = $false
                        NestingLevel   = $Level
                        SourceGroup    = $GroupDN
                        IsActive       = $isActive
                        DaysInactive   = if ($memberDetails.LastLogonDate) { 
                            (New-TimeSpan -Start $memberDetails.LastLogonDate -End (Get-Date)).Days 
                        }
                    }
                }
                
                "group" {
                    $memberDetails = Get-ADGroup -Identity $member.SID.Value @ADParams -Properties *
                    
                    $members += [PSCustomObject]@{
                        ObjectClass    = "Group"
                        Name           = $memberDetails.Name
                        SamAccountName = $memberDetails.SamAccountName
                        UPN            = ""
                        DN             = $member.DistinguishedName
                        SID            = $member.SID.Value
                        Enabled        = $true
                        LastLogon      = $null
                        Department     = ""
                        Title          = ""
                        Email          = ""
                        Phone          = ""
                        Manager        = ""
                        IsNested       = $true
                        NestingLevel   = $Level
                        SourceGroup    = $GroupDN
                        IsActive       = $true
                        DaysInactive   = $null
                    }
                    
                    # Рекурсивный обход вложенных групп
                    if ($IncludeNested) {
                        $nestedMembers = Get-GroupMembersRecursive -GroupDN $member.DistinguishedName -Level ($Level + 1) -Visited $Visited
                        $members += $nestedMembers
                    }
                }
                
                "computer" {
                    $memberDetails = Get-ADComputer -Identity $member.SID.Value @ADParams -Properties *
                    
                    $members += [PSCustomObject]@{
                        ObjectClass    = "Computer"
                        Name           = $memberDetails.Name
                        SamAccountName = $memberDetails.SamAccountName
                        UPN            = ""
                        DN             = $member.DistinguishedName
                        SID            = $member.SID.Value
                        Enabled        = $memberDetails.Enabled
                        LastLogon      = $memberDetails.LastLogonDate
                        Department     = ""
                        Title          = ""
                        Email          = ""
                        Phone          = ""
                        Manager        = ""
                        IsNested       = $false
                        NestingLevel   = $Level
                        SourceGroup    = $GroupDN
                        IsActive       = $true
                        DaysInactive   = if ($memberDetails.LastLogonDate) { 
                            (New-TimeSpan -Start $memberDetails.LastLogonDate -End (Get-Date)).Days 
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Ошибка при получении членов группы $GroupDN : $_"
    }
    
    return $members
}

function Get-CreationDateFromEventLog {
    param([string]$ObjectSID, [string]$ObjectName)
    
    # Попытка получить дату добавления в группу из журналов событий
    # Работает только если включен аудит изменений групп AD
    try {
        $event = Get-WinEvent -LogName "Security" -FilterXPath @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[(EventID=4728 or EventID=4729 or EventID=4732 or EventID=4756)]] 
      and *[EventData[Data[@Name='MemberSid']='$ObjectSID']]
    </Select>
  </Query>
</QueryList>
"@ -MaxEvents 1 -ErrorAction SilentlyContinue
        
        if ($event) {
            return $event.TimeCreated
        }
    }
    catch {
        Write-Verbose "Не удалось получить дату добавления из журналов событий для $ObjectName"
    }
    
    return $null
}

function New-HTMLReport {
    param(
        [array]$Groups,
        [hashtable]$GroupMembers,
        [string]$FilePath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Аудит групп Active Directory</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 40px; 
            background-color: #f5f5f5;
            color: #333;
        }
        .header { 
            background: linear-gradient(135deg, #2c3e50, #4a6491); 
            color: white; 
            padding: 25px; 
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .report-title { 
            font-size: 28px; 
            font-weight: 300; 
            margin-bottom: 10px;
        }
        .report-subtitle { 
            font-size: 14px; 
            opacity: 0.9; 
            margin-bottom: 5px;
        }
        .summary-box {
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .group-card {
            background: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 25px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-left: 4px solid #3498db;
        }
        .group-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 1px solid #eee;
        }
        .group-name {
            font-size: 20px;
            font-weight: 600;
            color: #2c3e50;
        }
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 10px;
        }
        .badge-primary { background: #3498db; color: white; }
        .badge-success { background: #27ae60; color: white; }
        .badge-warning { background: #f39c12; color: white; }
        .badge-danger { background: #e74c3c; color: white; }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        th {
            background-color: #f8f9fa;
            padding: 12px 15px;
            text-align: left;
            font-weight: 600;
            color: #2c3e50;
            border-bottom: 2px solid #dee2e6;
        }
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #dee2e6;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .user-row { background-color: #f8fff8; }
        .group-row { background-color: #f8f8ff; }
        .computer-row { background-color: #fff8f8; }
        .inactive { opacity: 0.6; }
        .level-indent { 
            padding-left: 20px;
            position: relative;
        }
        .level-indent:before {
            content: "↳ ";
            position: absolute;
            left: 5px;
            color: #7f8c8d;
        }
        .type-icon {
            display: inline-block;
            width: 20px;
            text-align: center;
            margin-right: 8px;
        }
        .legend {
            display: flex;
            gap: 20px;
            margin-top: 20px;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .legend-color {
            width: 20px;
            height: 20px;
            border-radius: 4px;
        }
        .footer {
            margin-top: 40px;
            text-align: center;
            color: #7f8c8d;
            font-size: 12px;
            padding-top: 20px;
            border-top: 1px solid #eee;
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="report-title">🎯 Аудит групп Active Directory</div>
        <div class="report-subtitle">Домен: $(Get-ADDomain @ADParams | Select-Object -ExpandProperty DNSRoot)</div>
        <div class="report-subtitle">Дата генерации: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</div>
        <div class="report-subtitle">Контроллер домена: $(if($DC){$DC}else{"Auto"})</div>
    </div>
    
    <div class="summary-box">
        <h3 style="margin-top: 0;">📊 Сводная информация</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
            <div>
                <div style="font-size: 24px; font-weight: bold; color: #2c3e50;">$($Groups.Count)</div>
                <div style="font-size: 14px; color: #7f8c8d;">Всего групп</div>
            </div>
            <div>
                <div style="font-size: 24px; font-weight: bold; color: #27ae60;">$(($GroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User'}).Count)</div>
                <div style="font-size: 14px; color: #7f8c8d;">Всего пользователей</div>
            </div>
            <div>
                <div style="font-size: 24px; font-weight: bold; color: #3498db;">$(($GroupMembers.Values | Where-Object {$_.ObjectClass -eq 'Group' -and $_.IsNested}).Count)</div>
                <div style="font-size: 14px; color: #7f8c8d;">Вложенных групп</div>
            </div>
            <div>
                <div style="font-size: 24px; font-weight: bold; color: #e74c3c;">$(($GroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User' -and !$_.IsActive}).Count)</div>
                <div style="font-size: 14px; color: #7f8c8d;">Неактивных пользователей</div>
            </div>
        </div>
    </div>
    
    <div class="legend">
        <div class="legend-item">
            <div class="legend-color" style="background-color: #f8fff8;"></div>
            <span>Пользователи</span>
        </div>
        <div class="legend-item">
            <div class="legend-color" style="background-color: #f8f8ff;"></div>
            <span>Группы</span>
        </div>
        <div class="legend-item">
            <div class="legend-color" style="background-color: #fff8f8;"></div>
            <span>Компьютеры</span>
        </div>
        <div class="legend-item">
            <div style="opacity: 0.6;">▪</div>
            <span>Неактивные объекты</span>
        </div>
    </div>
"@

    # Добавляем информацию по каждой группе
    foreach ($group in $Groups) {
        $members = $GroupMembers[$group.DistinguishedName]
        $directUsers = $members | Where-Object {$_.NestingLevel -eq 0 -and $_.ObjectClass -eq 'User'} | Measure-Object | Select-Object -ExpandProperty Count
        $directGroups = $members | Where-Object {$_.NestingLevel -eq 0 -and $_.ObjectClass -eq 'Group'} | Measure-Object | Select-Object -ExpandProperty Count
        $totalMembers = $members | Where-Object {$_.ObjectClass -eq 'User'} | Select-Object -Unique -Property SamAccountName | Measure-Object | Select-Object -ExpandProperty Count
        
        $html += @"
    <div class="group-card">
        <div class="group-header">
            <div class="group-name">$($group.Name)</div>
            <div>
                <span class="badge badge-primary">Участники: $directUsers</span>
                <span class="badge badge-success">Группы: $directGroups</span>
                <span class="badge badge-warning">Всего уникальных: $totalMembers</span>
                <span class="badge badge-$(if($group.GroupScope -eq 'DomainLocal'){'warning'}elseif($group.GroupScope -eq 'Global'){'primary'}else{'success'})">$($group.GroupScope)</span>
            </div>
        </div>
        
        <div style="margin-bottom: 15px;">
            <strong>Описание:</strong> $(if($group.Description){$group.Description}else{"—"})<br>
            <strong>Email группы:</strong> $(if($group.mail){$group.mail}else{"—"})<br>
            <strong>Категория:</strong> $($group.GroupCategory) | <strong>Область:</strong> $($group.GroupScope)<br>
            <strong>SID:</strong> <code>$($group.SID.Value)</code><br>
            <strong>DistinguishedName:</strong> <code>$($group.DistinguishedName)</code>
        </div>
"@

        if ($members) {
            $html += @"
        <table>
            <thead>
                <tr>
                    <th>Тип</th>
                    <th>Имя</th>
                    <th>Логин</th>
                    <th>Статус</th>
                    <th>Отдел</th>
                    <th>Последний вход</th>
                    <th>Уровень</th>
                </tr>
            </thead>
            <tbody>
"@

            foreach ($member in $members | Sort-Object NestingLevel, ObjectClass, Name) {
                $rowClass = switch ($member.ObjectClass) {
                    "User" { "user-row" }
                    "Group" { "group-row" }
                    "Computer" { "computer-row" }
                }
                
                if (-not $member.IsActive) { $rowClass += " inactive" }
                
                $typeIcon = switch ($member.ObjectClass) {
                    "User" { "👤" }
                    "Group" { "👥" }
                    "Computer" { "💻" }
                }
                
                $statusBadge = if ($member.ObjectClass -eq "User") {
                    if ($member.Enabled) {
                        if ($member.IsActive) {
                            "<span class='badge badge-success'>Активен</span>"
                        } else {
                            "<span class='badge badge-danger'>Неактивен ($($member.DaysInactive) дн.)</span>"
                        }
                    } else {
                        "<span class='badge badge-danger'>Отключен</span>"
                    }
                } else {
                    "<span class='badge badge-primary'>$($member.ObjectClass)</span>"
                }
                
                $nestingDisplay = if ($member.NestingLevel -gt 0) {
                    "class='level-indent' style='padding-left: $(20 + ($member.NestingLevel * 20))px;'"
                } else { "" }
                
                $html += @"
                <tr class="$rowClass">
                    <td><span class="type-icon">$typeIcon</span></td>
                    <td $nestingDisplay>$($member.Name)</td>
                    <td>$($member.SamAccountName)</td>
                    <td>$statusBadge</td>
                    <td>$($member.Department)</td>
                    <td>$(if($member.LastLogon) {$member.LastLogon.ToString("dd.MM.yyyy HH:mm")} else {"—"})</td>
                    <td>$($member.NestingLevel)</td>
                </tr>
"@
            }
            
            $html += @"
            </tbody>
        </table>
"@
        } else {
            $html += @"
        <div style="text-align: center; padding: 30px; color: #7f8c8d;">
            <em>Группа не содержит членов</em>
        </div>
"@
        }
        
        $html += @"
    </div>
"@
    }

    $html += @"
    <div class="footer">
        <p>Отчет сгенерирован автоматически. Время выполнения: $((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds.ToString('0.00')) сек.</p>
        <p>Система аудита Active Directory | $(Get-ADDomain @ADParams | Select-Object -ExpandProperty Forest)</p>
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Host "HTML отчет сохранен: $FilePath" -ForegroundColor Green
}
#endregion

# Основная логика
try {
    $GroupsToProcess = @()
    
    # Определение групп для обработки в зависимости от параметра
    switch ($PSCmdlet.ParameterSetName) {
        "ByName" {
            Write-Host "Поиск групп по именам: $($GroupNames -join ', ')" -ForegroundColor Yellow
            foreach ($groupName in $GroupNames) {
                $foundGroups = Get-ADGroupWithFallback -Identity $groupName
                if ($foundGroups) {
                    $GroupsToProcess += $foundGroups
                }
            }
        }
        
        "ByDN" {
            Write-Host "Обработка групп по DistinguishedName" -ForegroundColor Yellow
            foreach ($dn in $GroupDNs) {
                try {
                    $group = Get-ADGroup -Identity $dn @ADParams -Properties *
                    $GroupsToProcess += $group
                }
                catch {
                    Write-Warning "Не удалось найти группу по DN: $dn"
                }
            }
        }
        
        "ByOU" {
            Write-Host "Поиск всех групп в OU: $($OUs -join ', ')" -ForegroundColor Yellow
            foreach ($ou in $OUs) {
                try {
                    $groupsInOU = Get-ADGroup -SearchBase $ou -Filter * @ADParams -Properties *
                    $GroupsToProcess += $groupsInOU
                }
                catch {
                    Write-Warning "Ошибка при поиске групп в OU $ou : $_"
                }
            }
        }
    }
    
    if ($GroupsToProcess.Count -eq 0) {
        Write-Error "Не найдено ни одной группы для обработки."
        return
    }
    
    Write-Host "Найдено групп для обработки: $($GroupsToProcess.Count)" -ForegroundColor Green
    
    # Сбор информации о членах групп
    $AllGroupMembers = @{}
    $groupCounter = 0
    
    foreach ($group in $GroupsToProcess) {
        $groupCounter++
        Write-Progress -Activity "Анализ групп" -Status "Обработка: $($group.Name)" `
            -PercentComplete (($groupCounter / $GroupsToProcess.Count) * 100)
        
        Write-Verbose "Обработка группы: $($group.Name) ($($group.DistinguishedName))"
        
        $members = Get-GroupMembersRecursive -GroupDN $group.DistinguishedName
        $AllGroupMembers[$group.DistinguishedName] = $members
        
        Write-Host "  $($group.Name): $($members.Count) объектов" -ForegroundColor Gray
    }
    
    Write-Progress -Activity "Анализ групп" -Completed
    
    # Генерация HTML отчета
    $reportFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    New-HTMLReport -Groups $GroupsToProcess -GroupMembers $AllGroupMembers -FilePath $reportFullPath
    
    # Экспорт в CSV (если требуется)
    if ($ExportCSV) {
        $csvPath = [System.IO.Path]::ChangeExtension($reportFullPath, "csv")
        $AllGroupMembers.Values | Select-Object * | Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation
        Write-Host "CSV данные сохранены: $csvPath" -ForegroundColor Green
    }
  <#  
    # Отправка email (если требуется)
    if ($SendEmail) {
        if (-not $SmtpServer) {
            $SmtpServer = (Get-ADDomainController @ADParams | Select-Object -First 1).HostName
        }
        
        $emailSubject = "Аудит групп AD: $($GroupsToProcess.Count) групп - $(Get-Date -Format 'dd.MM.yyyy')"
        $emailBody = @"
Уважаемые коллеги,

Во вложении представлен отчет по членам групп Active Directory.

<b>Сводная информация:</b>
- Всего групп: $($GroupsToProcess.Count)
- Всего пользователей: $(($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User'} | Select-Object -Unique -Property SamAccountName).Count)
- Вложенных групп: $(($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'Group' -and $_.IsNested}).Count)
- Неактивных пользователей: $(($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User' -and !$_.IsActive}).Count)

Отчет сгенерирован автоматически.
Время выполнения: $((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds.ToString('0.00')) сек.

--
Система аудита Active Directory
$(Get-ADDomain @ADParams | Select-Object -ExpandProperty DNSRoot)
"@
        
        try {
            Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $emailSubject `
                -Body $emailBody -BodyAsHtml -Attachments $reportFullPath `
                -SmtpServer $SmtpServer -Encoding UTF8
            
            Write-Host "Отчет отправлен по email: $($EmailTo -join ', ')" -ForegroundColor Green
        }
        catch {
            Write-Warning "Не удалось отправить email: $_"
        }
    }
#>

    # Вывод краткой статистики в консоль
    Write-Host "`n=== СТАТИСТИКА ===" -ForegroundColor Cyan
    $totalUsers = ($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User'} | Select-Object -Unique -Property SamAccountName).Count
    $totalNestedGroups = ($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'Group' -and $_.IsNested}).Count
    $inactiveUsers = ($AllGroupMembers.Values | Where-Object {$_.ObjectClass -eq 'User' -and !$_.IsActive}).Count
    
    Write-Host "Обработано групп: $($GroupsToProcess.Count)" -ForegroundColor White
    Write-Host "Найдено пользователей: $totalUsers" -ForegroundColor White
    Write-Host "Вложенных групп: $totalNestedGroups" -ForegroundColor White
    Write-Host "Неактивных пользователей: $inactiveUsers" -ForegroundColor $(if($inactiveUsers -gt 0){'Red'}else{'White'})
    Write-Host "Файл отчета: $reportFullPath" -ForegroundColor Green
    Write-Host "Время выполнения: $((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds.ToString('0.00')) сек." -ForegroundColor Gray
    
}
catch {
    Write-Error "Критическая ошибка: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    Write-Progress -Activity "Анализ групп" -Completed
}
