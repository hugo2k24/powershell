<#
.SYNOPSIS
    Комплексная проверка работоспособности инфраструктуры Active Directory.
.DESCRIPTION
    Выполняет всестороннюю диагностику контроллеров домена и служб AD:
    репликация, DNS, DFS, системные ресурсы, журналы событий, службы.
    Формирует HTML-отчет с цветовой индикацией проблем.
.PARAMETER Domain
    Домен для проверки (по умолчанию - текущий).
.PARAMETER DCHostnames
    Список конкретных контроллеров домена для проверки.
.PARAMETER CheckDC
    Проверять контроллеры домена (включено по умолчанию).
.PARAMETER CheckDNS
    Проверять DNS серверы.
.PARAMETER CheckDFS
    Проверять репликацию DFS.
.PARAMETER CheckSysvol
    Проверять репликацию SYSVOL.
.PARAMETER CheckServices
    Проверять состояние критических служб.
.PARAMETER CheckEventLogs
    Проверять критические события в журналах.
.PARAMETER CheckDiskSpace
    Проверять свободное место на дисках.
.PARAMETER DiskSpaceThreshold
    Порог свободного места в % (по умолчанию: 15).
.PARAMETER CPUThreshold
    Порог загрузки CPU в % (по умолчанию: 80).
.PARAMETER RAMThreshold
    Порог использования RAM в % (по умолчанию: 85).
.PARAMETER DaysBack
    Количество дней для проверки журналов событий (по умолчанию: 1).
.PARAMETER SendEmail
    Отправлять отчет по электронной почте.
.PARAMETER EmailTo
    Адреса получателей (через запятую).
.PARAMETER EmailFrom
    Адрес отправителя.
.PARAMETER SmtpServer
    SMTP-сервер.
.PARAMETER ReportPath
    Путь для сохранения HTML-отчета.
.PARAMETER Detailed
    Подробный режим с дополнительной диагностикой.
.PARAMETER Credential
    Учетные данные для подключения (если нужны).
.EXAMPLE
    Get-ADHealthReport.ps1
    Базовая проверка всех контроллеров домена.
.EXAMPLE
    Get-ADHealthReport.ps1 -Domain "domain.local" -CheckDNS -CheckDFS -Detailed -SendEmail -EmailTo "admin@contoso.local"
    Полная проверка с отправкой email.
.EXAMPLE
    Get-ADHealthReport.ps1 -DCHostnames @("DC01", "DC02") -CheckEventLogs -DaysBack 3
    Проверка конкретных DC за последние 3 дня.
.NOTES
    Author: Дмитрий Плотинский
    Version: 12.7.3
    Date: 2025-10-15
    Требует: PowerShell 5.1+, модули ActiveDirectory, Dfsr, админские права
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$Domain = $env:USERDNSDOMAIN,
    
    [Parameter(Mandatory=$false)]
    [string[]]$DCHostnames,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckDC = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckDNS,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckDFS,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckSysvol = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckServices = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckEventLogs,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckDiskSpace = $true,
    
    [Parameter(Mandatory=$false)]
    [int]$DiskSpaceThreshold = 15,
    
    [Parameter(Mandatory=$false)]
    [int]$CPUThreshold = 80,
    
    [Parameter(Mandatory=$false)]
    [int]$RAMThreshold = 85,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 1,
    
    [Parameter(Mandatory=$false)]
    [switch]$SendEmail,
    
    [Parameter(Mandatory=$false)]
    [string[]]$EmailTo,
    
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom = "ad-health@$Domain",
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = ".\AD_Health_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    
    [Parameter(Mandatory=$false)]
    [switch]$Detailed,
    
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)

# Инициализация и конфигурация
$StartTime = Get-Date
$ErrorActionPreference = 'Stop'
$WarningPreference = 'Continue'

# Цветовая схема для статусов
$StatusColors = @{
    'Healthy'   = '#2ecc71'  # Зеленый
    'Warning'   = '#f39c12'  # Оранжевый
    'Critical'  = '#e74c3c'  # Красный
    'Unknown'   = '#95a5a6'  # Серый
}

# Критические службы AD
$CriticalServices = @(
    'NTDS',
    'DNS',
    'DFS Replication',
    'Netlogon',
    'IsmServ',
    'KDC',
    'W32Time',
    'ADWS',
    'Dfs',
    'SamSs'
)

# Критические события AD (ID, LogName, Description)
$CriticalEvents = @(
    @{ID=1980; Log='Directory Service'; Desc='В SYSVOL обнаружены расхождения'}
    @{ID=1988; Log='Directory Service'; Desc='Ошибка репликации DFS-R'}
    @{ID=1126; Log='Directory Service'; Desc='Ошибка репликации между DC'}
    @{ID=1925; Log='DNS Server'; Desc='DNS сервер не загрузил зоны AD'}
    @{ID=4013; Log='DNS Server'; Desc='DNS сервер не может загрузить зоны'}
    @{ID=13508; Log='FRS'; Desc='Ошибка репликации FRS'}
    @{ID=13511; Log='FRS'; Desc='FRS не может работать с SYSVOL'}
    @{ID=7023; Log='System'; Desc='Критическая служба завершилась с ошибкой'}
    @{ID=6008; Log='System'; Desc='Неожиданное выключение системы'}
    @{ID=10010; Log='System'; Desc='Сервер не отвечает на запросы DCOM'}
    @{ID=1058; Log='System'; Desc='Ошибка службы аудита'}
    @{ID=1030; Log='DFS Replication'; Desc='Ошибка репликации DFS'}
    @{ID=6702; Log='DFS Replication'; Desc='Повторяющиеся ошибки DFS'}
)

# Файлы репликации SYSVOL для проверки
$SysvolFolders = @(
    '\\SYSVOL',
    '\\NETLOGON',
    '\Windows\SYSVOL\domain',
    '\Windows\SYSVOL\staging',
    '\Windows\SYSVOL\sysvol'
)

# Проверка и загрузка модулей
$RequiredModules = @('ActiveDirectory')
if ($CheckDFS) { $RequiredModules += 'DFSR' }

foreach ($module in $RequiredModules) {
    if (-not (Get-Module -Name $module -ErrorAction SilentlyContinue)) {
        try {
            Import-Module $module -ErrorAction Stop
            Write-Verbose "Модуль $module загружен"
        }
        catch {
            Write-Warning "Не удалось загрузить модуль $module. Некоторые проверки будут пропущены."
        }
    }
}

# Настройка параметров AD
$ADParams = @{}
if ($Credential) { $ADParams.Credential = $Credential }

# Глобальная переменная для сбора результатов
$HealthReport = @{
    DomainInfo = $null
    DomainControllers = @()
    ReplicationStatus = @()
    DNSChecks = @()
    DFSChecks = @()
    SysvolChecks = @()
    ServiceStatus = @()
    EventLogChecks = @()
    SystemResources = @()
    OverallStatus = 'Healthy'
    Summary = @{
        TotalDCs = 0
        HealthyDCs = 0
        WarningDCs = 0
        CriticalDCs = 0
        TotalIssues = 0
        CriticalIssues = 0
        WarningIssues = 0
    }
}


# Вспомогательные функции
function Get-StatusColor {
    param([string]$Status)
    return $StatusColors[$Status]
}

function Get-PerformanceCounter {
    param(
        [string]$ComputerName,
        [string]$Counter,
        [string]$Instance = ''
    )
    
    try {
        $counterPath = "\\$ComputerName\$Counter"
        if ($Instance) { $counterPath += "($Instance)" }
        
        $counterData = Get-Counter -Counter $counterPath -ErrorAction Stop
        return [Math]::Round($counterData.CounterSamples.CookedValue, 2)
    }
    catch {
        Write-Verbose "Не удалось получить счетчик $Counter с $ComputerName : $_"
        return $null
    }
}

function Test-Port {
    param(
        [string]$ComputerName,
        [int]$Port,
        [int]$Timeout = 1000
    )
    
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($ComputerName, $Port, $null, $null)
        $waitResult = $asyncResult.AsyncWaitHandle.WaitOne($Timeout, $false)
        
        if ($waitResult) {
            $tcpClient.EndConnect($asyncResult) | Out-Null
            $tcpClient.Close()
            return $true
        }
        else {
            $tcpClient.Close()
            return $false
        }
    }
    catch {
        return $false
    }
}

function Get-EventLogSummary {
    param(
        [string]$ComputerName,
        [int]$HoursBack = 24
    )
    
    $eventsSummary = @()
    $startTime = (Get-Date).AddHours(-$HoursBack)
    
    foreach ($eventConfig in $CriticalEvents) {
        try {
            $events = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                LogName = $eventConfig.Log
                ID = $eventConfig.ID
                StartTime = $startTime
            } -MaxEvents 5 -ErrorAction SilentlyContinue
            
            if ($events) {
                foreach ($event in $events) {
                    $eventsSummary += [PSCustomObject]@{
                        ComputerName = $ComputerName
                        TimeCreated = $event.TimeCreated
                        ID = $event.Id
                        LogName = $event.LogName
                        Level = $event.LevelDisplayName
                        Message = $event.Message.Substring(0, [Math]::Min(200, $event.Message.Length))
                        ProviderName = $event.ProviderName
                        Description = $eventConfig.Desc
                    }
                }
            }
        }
        catch {
            Write-Verbose "Ошибка чтения событий с $ComputerName : $_"
        }
    }
    
    return $eventsSummary
}

function Test-DNSService {
    param([string]$ComputerName)
    
    $dnsTests = @()
    
    # Проверка порта DNS (53)
    $dnsPort = Test-Port -ComputerName $ComputerName -Port 53
    $dnsTests += [PSCustomObject]@{
        Test = 'DNS Port 53'
        Status = if ($dnsPort) { 'Healthy' } else { 'Critical' }
        Details = if ($dnsPort) { 'Порт открыт' } else { 'Порт закрыт' }
    }
    
    # Проверка резолвинга
    try {
        $resolveTest = Resolve-DnsName -Name $ComputerName -Server $ComputerName -ErrorAction Stop
        $dnsTests += [PSCustomObject]@{
            Test = 'DNS Resolution'
            Status = 'Healthy'
            Details = "Успешно: $($resolveTest.IPAddress -join ', ')"
        }
    }
    catch {
        $dnsTests += [PSCustomObject]@{
            Test = 'DNS Resolution'
            Status = 'Critical'
            Details = "Ошибка: $_"
        }
    }
    
    # Проверка рекурсии
    try {
        $recursionTest = Resolve-DnsName -Name 'microsoft.com' -Server $ComputerName -ErrorAction Stop
        $dnsTests += [PSCustomObject]@{
            Test = 'DNS Recursion'
            Status = 'Healthy'
            Details = 'Рекурсия работает'
        }
    }
    catch {
        $dnsTests += [PSCustomObject]@{
            Test = 'DNS Recursion'
            Status = 'Warning'
            Details = 'Проблемы с рекурсией'
        }
    }
    
    return $dnsTests
}

function Test-DFSReplication {
    param([string]$ComputerName)
    
    $dfsTests = @()
    
    try {
        # Проверка службы DFS
        $dfsService = Get-Service -ComputerName $ComputerName -Name 'DFS Replication' -ErrorAction Stop
        $dfsTests += [PSCustomObject]@{
            Test = 'DFS Service'
            Status = if ($dfsService.Status -eq 'Running') { 'Healthy' } else { 'Critical' }
            Details = "Состояние: $($dfsService.Status)"
        }
        
        if ($dfsService.Status -eq 'Running') {
            # Проверка состояния репликации
            $dfsState = Get-DfsrState -ComputerName $ComputerName -ErrorAction Stop | 
                Where-Object { $_.State -ne 'Normal' } | 
                Select-Object -First 5
            
            if ($dfsState) {
                $dfsTests += [PSCustomObject]@{
                    Test = 'DFS Replication State'
                    Status = 'Warning'
                    Details = "Проблемы в $($dfsState.Count) репликациях"
                }
            }
            else {
                $dfsTests += [PSCustomObject]@{
                    Test = 'DFS Replication State'
                    Status = 'Healthy'
                    Details = 'Все репликации в норме'
                }
            }
        }
    }
    catch {
        $dfsTests += [PSCustomObject]@{
            Test = 'DFS Service'
            Status = 'Critical'
            Details = "Ошибка: $_"
        }
    }
    
    return $dfsTests
}

function Test-SysvolReplication {
    param([string]$ComputerName)
    
    $sysvolTests = @()
    
    # Проверка доступности SYSVOL
    $sysvolPath = "\\$ComputerName\SYSVOL"
    $netlogonPath = "\\$ComputerName\NETLOGON"
    
    try {
        $sysvolAccess = Test-Path $sysvolPath -ErrorAction Stop
        $sysvolTests += [PSCustomObject]@{
            Test = 'SYSVOL Share Access'
            Status = if ($sysvolAccess) { 'Healthy' } else { 'Critical' }
            Details = if ($sysvolAccess) { 'Доступен' } else { 'Недоступен' }
        }
        
        $netlogonAccess = Test-Path $netlogonPath -ErrorAction Stop
        $sysvolTests += [PSCustomObject]@{
            Test = 'NETLOGON Share Access'
            Status = if ($netlogonAccess) { 'Healthy' } else { 'Critical' }
            Details = if ($netlogonAccess) { 'Доступен' } else { 'Недоступен' }
        }
        
        if ($Detailed) {
            # Проверка целостности SYSVOL
            try {
                $usnJournal = Get-WmiObject -ComputerName $ComputerName -Class Win32_USNJournal -ErrorAction Stop
                $sysvolTests += [PSCustomObject]@{
                    Test = 'USN Journal'
                    Status = if ($usnJournal) { 'Healthy' } else { 'Warning' }
                    Details = if ($usnJournal) { 'Активен' } else { 'Неактивен' }
                }
            }
            catch {
                $sysvolTests += [PSCustomObject]@{
                    Test = 'USN Journal'
                    Status = 'Warning'
                    Details = 'Не удалось проверить'
                }
            }
        }
    }
    catch {
        $sysvolTests += [PSCustomObject]@{
            Test = 'SYSVOL Check'
            Status = 'Critical'
            Details = "Ошибка доступа: $_"
        }
    }
    
    return $sysvolTests
}

function Get-SystemResources {
    param([string]$ComputerName)
    
    $resources = @()
    
    # Процессор
    try {
        $cpuUsage = Get-PerformanceCounter -ComputerName $ComputerName -Counter "processor(_Total)\% processor time"
        $cpuStatus = if ($cpuUsage -ge $CPUThreshold) { 'Warning' } elseif ($cpuUsage -eq $null) { 'Unknown' } else { 'Healthy' }
        
        $resources += [PSCustomObject]@{
            Resource = 'CPU Usage'
            Value = if ($cpuUsage) { "$cpuUsage%" } else { 'N/A' }
            Status = $cpuStatus
            Threshold = "$CPUThreshold%"
        }
    }
    catch {
        Write-Verbose "Ошибка получения CPU с $ComputerName : $_"
    }
    
    # Память
    try {
        $memory = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction Stop
        $totalMemory = [Math]::Round($memory.TotalVisibleMemorySize / 1MB, 2)
        $freeMemory = [Math]::Round($memory.FreePhysicalMemory / 1MB, 2)
        $usedMemory = $totalMemory - $freeMemory
        $memoryPercent = [Math]::Round(($usedMemory / $totalMemory) * 100, 2)
        
        $memoryStatus = if ($memoryPercent -ge $RAMThreshold) { 'Warning' } else { 'Healthy' }
        
        $resources += [PSCustomObject]@{
            Resource = 'Memory Usage'
            Value = "$memoryPercent% ($usedMemory/$totalMemory GB)"
            Status = $memoryStatus
            Threshold = "$RAMThreshold%"
        }
    }
    catch {
        Write-Verbose "Ошибка получения памяти с $ComputerName : $_"
    }
    
    # Дисковое пространство
    if ($CheckDiskSpace) {
        try {
            $disks = Get-WmiObject -ComputerName $ComputerName -Class Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
            
            foreach ($disk in $disks) {
                $freeGB = [Math]::Round($disk.FreeSpace / 1GB, 2)
                $totalGB = [Math]::Round($disk.Size / 1GB, 2)
                $usedPercent = [Math]::Round(100 - ($freeGB / $totalGB * 100), 2)
                $freePercent = 100 - $usedPercent
                
                $diskStatus = if ($freePercent -le $DiskSpaceThreshold) { 'Warning' } else { 'Healthy' }
                
                $resources += [PSCustomObject]@{
                    Resource = "Disk $($disk.DeviceID)"
                    Value = "$usedPercent% used ($freeGB/$totalGB GB free)"
                    Status = $diskStatus
                    Threshold = "$DiskSpaceThreshold% free"
                }
            }
        }
        catch {
            Write-Verbose "Ошибка получения информации о дисках с $ComputerName : $_"
        }
    }
    
    return $resources
}

function New-HealthHTMLReport {
    param(
        [hashtable]$HealthData,
        [string]$FilePath
    )
    
    $domain = $HealthData.DomainInfo.DNSRoot
    $reportTime = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
    $executionTime = (New-TimeSpan -Start $StartTime -End (Get-Date)).TotalMinutes.ToString('0.00')
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>AD Health Report - $domain</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            background-attachment: fixed;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50, #34495e);
            color: white;
            padding: 30px;
            text-align: center;
            position: relative;
        }
        .header:before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, #2ecc71, #f39c12, #e74c3c);
        }
        .header h1 {
            margin: 0;
            font-size: 36px;
            font-weight: 300;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
        }
        .header h1 i {
            font-size: 40px;
        }
        .header-info {
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-top: 15px;
            font-size: 14px;
            opacity: 0.9;
        }
        .status-summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        .status-card {
            background: white;
            border-radius: 10px;
            padding: 25px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            transition: transform 0.3s ease;
        }
        .status-card:hover {
            transform: translateY(-5px);
        }
        .status-card.healthy { border-top: 5px solid #2ecc71; }
        .status-card.warning { border-top: 5px solid #f39c12; }
        .status-card.critical { border-top: 5px solid #e74c3c; }
        .status-card.unknown { border-top: 5px solid #95a5a6; }
        .status-icon {
            font-size: 48px;
            margin-bottom: 15px;
        }
        .status-count {
            font-size: 42px;
            font-weight: bold;
            margin: 10px 0;
        }
        .status-label {
            font-size: 16px;
            color: #7f8c8d;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .dc-section {
            padding: 30px;
        }
        .section-title {
            font-size: 24px;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ecf0f1;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .dc-card {
            background: white;
            border-radius: 10px;
            margin-bottom: 20px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            border-left: 5px solid #3498db;
        }
        .dc-header {
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            padding: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .dc-name {
            font-size: 20px;
            font-weight: 600;
        }
        .dc-role {
            background: rgba(255,255,255,0.2);
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 12px;
        }
        .dc-content {
            padding: 20px;
        }
        .test-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        .test-item {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 4px solid #ddd;
        }
        .test-item.healthy { border-left-color: #2ecc71; background: #e8f6f3; }
        .test-item.warning { border-left-color: #f39c12; background: #fef9e7; }
        .test-item.critical { border-left-color: #e74c3c; background: #fdedec; }
        .test-info { flex-grow: 1; }
        .test-name { font-weight: 600; margin-bottom: 5px; }
        .test-details { font-size: 12px; color: #7f8c8d; }
        .test-status {
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
        }
        .status-healthy { background: #2ecc71; color: white; }
        .status-warning { background: #f39c12; color: white; }
        .status-critical { background: #e74c3c; color: white; }
        .status-unknown { background: #95a5a6; color: white; }
        .event-item {
            background: #fff;
            border-left: 4px solid #e74c3c;
            padding: 15px;
            margin-bottom: 10px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        }
        .event-time { font-size: 12px; color: #7f8c8d; margin-bottom: 5px; }
        .event-id { 
            display: inline-block;
            background: #e74c3c;
            color: white;
            padding: 2px 8px;
            border-radius: 10px;
            font-size: 11px;
            margin-right: 10px;
        }
        .resource-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 10px;
        }
        .resource-item {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 12px;
            text-align: center;
        }
        .resource-value {
            font-size: 18px;
            font-weight: bold;
            margin: 5px 0;
        }
        .resource-label {
            font-size: 12px;
            color: #7f8c8d;
        }
        .footer {
            text-align: center;
            padding: 30px;
            background: #2c3e50;
            color: white;
            margin-top: 30px;
        }
        .footer a {
            color: #3498db;
            text-decoration: none;
        }
        .timestamp {
            background: #34495e;
            padding: 10px;
            border-radius: 5px;
            display: inline-block;
            margin-top: 10px;
        }
        .legend {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 20px;
            flex-wrap: wrap;
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
        .healthy-color { background: #2ecc71; }
        .warning-color { background: #f39c12; }
        .critical-color { background: #e74c3c; }
        .unknown-color { background: #95a5a6; }
        .accordion {
            cursor: pointer;
            padding: 10px;
            background: #ecf0f1;
            border: none;
            text-align: left;
            font-size: 16px;
            transition: 0.4s;
            margin-top: 10px;
            border-radius: 5px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .accordion:hover {
            background: #bdc3c7;
        }
        .accordion:after {
            content: '\002B';
            font-size: 20px;
        }
        .active:after {
            content: '\2212';
        }
        .panel {
            padding: 0 18px;
            background-color: white;
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.2s ease-out;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ Active Directory Health Report</h1>
            <div class="header-info">
                <div><strong>Домен:</strong> $domain</div>
                <div><strong>Время отчета:</strong> $reportTime</div>
                <div><strong>Контроллеров:</strong> $($HealthData.Summary.TotalDCs)</div>
            </div>
            <div class="legend">
                <div class="legend-item">
                    <div class="legend-color healthy-color"></div>
                    <span>Здорово ($($HealthData.Summary.HealthyDCs))</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color warning-color"></div>
                    <span>Предупреждение ($($HealthData.Summary.WarningDCs))</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color critical-color"></div>
                    <span>Критично ($($HealthData.Summary.CriticalDCs))</span>
                </div>
            </div>
        </div>
        
        <div class="status-summary">
            <div class="status-card $($HealthData.OverallStatus.ToLower())">
                <div class="status-icon">📊</div>
                <div class="status-count">$($HealthData.OverallStatus)</div>
                <div class="status-label">Общий статус</div>
            </div>
            <div class="status-card healthy">
                <div class="status-icon">✅</div>
                <div class="status-count">$($HealthData.Summary.HealthyDCs)</div>
                <div class="status-label">Здоровые DC</div>
            </div>
            <div class="status-card warning">
                <div class="status-icon">⚠️</div>
                <div class="status-count">$($HealthData.Summary.WarningDCs)</div>
                <div class="status-label">Предупреждения</div>
            </div>
            <div class="status-card critical">
                <div class="status-icon">🚨</div>
                <div class="status-count">$($HealthData.Summary.CriticalDCs)</div>
                <div class="status-label">Критические</div>
            </div>
        </div>
"@

    # Секция для каждого контроллера домена
    foreach ($dc in $HealthData.DomainControllers) {
        $dcStatus = 'healthy'
        if ($dc.CriticalIssues -gt 0) { $dcStatus = 'critical' }
        elseif ($dc.WarningIssues -gt 0) { $dcStatus = 'warning' }
        
        $html += @"
        <div class="dc-section">
            <div class="section-title">🏢 $($dc.HostName) 
                <span class="test-status status-$dcStatus">$($dc.OverallStatus)</span>
                <span style="font-size: 14px; color: #7f8c8d; margin-left: auto;">
                    🕒 Uptime: $($dc.UpTime) | 📍 Site: $($dc.Site)
                </span>
            </div>
            
            <div class="dc-card">
                <div class="dc-header">
                    <div class="dc-name">$($dc.ComputerName)</div>
                    <div class="dc-role">$($dc.OS) | $($dc.Roles -join ', ')</div>
                </div>
                
                <div class="dc-content">
"@

        # Репликация
        if ($dc.ReplicationStatus) {
            $html += @"
                    <button class="accordion">🔄 Репликация ($($dc.ReplicationStatus.Count) проверок)</button>
                    <div class="panel">
                        <div class="test-grid">
"@
            foreach ($rep in $dc.ReplicationStatus) {
                $html += @"
                            <div class="test-item $($rep.Status.ToLower())">
                                <div class="test-info">
                                    <div class="test-name">$($rep.Test)</div>
                                    <div class="test-details">$($rep.Details)</div>
                                </div>
                                <div class="test-status status-$($rep.Status.ToLower())">$($rep.Status)</div>
                            </div>
"@
            }
            $html += @"
                        </div>
                    </div>
"@
        }

        # Службы
        if ($dc.ServiceStatus) {
            $html += @"
                    <button class="accordion">⚙️ Службы ($($dc.ServiceStatus.Count))</button>
                    <div class="panel">
                        <div class="test-grid">
"@
            foreach ($service in $dc.ServiceStatus) {
                $html += @"
                            <div class="test-item $($service.Status.ToLower())">
                                <div class="test-info">
                                    <div class="test-name">$($service.Name)</div>
                                    <div class="test-details">$($service.DisplayName)</div>
                                </div>
                                <div class="test-status status-$($service.Status.ToLower())">$($service.Status)</div>
                            </div>
"@
            }
            $html += @"
                        </div>
                    </div>
"@
        }

        # Системные ресурсы
        if ($dc.SystemResources) {
            $html += @"
                    <button class="accordion">📊 Ресурсы ($($dc.SystemResources.Count))</button>
                    <div class="panel">
                        <div class="resource-grid">
"@
            foreach ($resource in $dc.SystemResources) {
                $html += @"
                            <div class="resource-item">
                                <div class="resource-label">$($resource.Resource)</div>
                                <div class="resource-value">$($resource.Value)</div>
                                <div class="test-status status-$($resource.Status.ToLower())">$($resource.Status)</div>
                            </div>
"@
            }
            $html += @"
                        </div>
                    </div>
"@
        }

        # События
        if ($dc.EventLogChecks) {
            $html += @"
                    <button class="accordion">📝 События ($($dc.EventLogChecks.Count))</button>
                    <div class="panel">
"@
            foreach ($event in $dc.EventLogChecks) {
                $html += @"
                        <div class="event-item">
                            <div class="event-time">$($event.TimeCreated.ToString('dd.MM.yyyy HH:mm'))</div>
                            <span class="event-id">ID: $($event.ID)</span>
                            <span class="event-level">$($event.Level)</span><br>
                            <strong>$($event.Description)</strong><br>
                            <small>$($event.Message)</small>
                        </div>
"@
            }
            $html += @"
                    </div>
"@
        }

        $html += @"
                </div>
            </div>
        </div>
"@
    }

    # Сводка по проблемам
    $html += @"
        <div class="dc-section">
            <div class="section-title">📋 Сводка проблем</div>
            <div style="background: #f8f9fa; padding: 20px; border-radius: 10px;">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
                    <div style="text-align: center;">
                        <div style="font-size: 36px; font-weight: bold; color: #2c3e50;">
                            $($HealthData.Summary.TotalIssues)
                        </div>
                        <div style="font-size: 14px; color: #7f8c8d;">Всего проблем</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 36px; font-weight: bold; color: #e74c3c;">
                            $($HealthData.Summary.CriticalIssues)
                        </div>
                        <div style="font-size: 14px; color: #7f8c8d;">Критических</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 36px; font-weight: bold; color: #f39c12;">
                            $($HealthData.Summary.WarningIssues)
                        </div>
                        <div style="font-size: 14px; color: #7f8c8d;">Предупреждений</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 36px; font-weight: bold; color: #2ecc71;">
                            $($HealthData.Summary.HealthyDCs)
                        </div>
                        <div style="font-size: 14px; color: #7f8c8d;">Здоровых DC</div>
                    </div>
                </div>
                
                <div style="margin-top: 20px; padding: 15px; background: white; border-radius: 8px;">
                    <h4 style="margin-top: 0;">Рекомендации:</h4>
                    <ul>
"@

    if ($HealthData.Summary.CriticalIssues -gt 0) {
        $html += "<li>🔴 <strong>НЕМЕДЛЕННО УСТРАНИТЬ:</strong> Критические проблемы требуют срочного вмешательства</li>"
    }
    if ($HealthData.Summary.WarningIssues -gt 0) {
        $html += "<li>🟡 <strong>ПЛАНИРОВАТЬ УСТРАНЕНИЕ:</strong> Предупреждения могут стать критическими</li>"
    }
    if ($HealthData.Summary.CriticalIssues -eq 0 -and $HealthData.Summary.WarningIssues -eq 0) {
        $html += "<li>🟢 <strong>ВСЁ В НОРМЕ:</strong> Все системы функционируют стабильно</li>"
    }

    $html += @"
                    </ul>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <div style="font-size: 14px; margin-bottom: 10px;">
                Отчет сгенерирован автоматически системой мониторинга Active Directory
            </div>
            <div class="timestamp">
                Время выполнения: $executionTime минут | $reportTime
            </div>
            <div style="margin-top: 20px; font-size: 12px; opacity: 0.8;">
                Системный администратор | $domain | Версия скрипта 4.5
            </div>
        </div>
    </div>
    
    <script>
        var acc = document.getElementsByClassName("accordion");
        for (var i = 0; i < acc.length; i++) {
            acc[i].addEventListener("click", function() {
                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                    panel.style.maxHeight = null;
                } else {
                    panel.style.maxHeight = panel.scrollHeight + "px";
                } 
            });
        }
    </script>
</body>
</html>
"@

    $html | Out-File -FilePath $FilePath -Encoding UTF8
    Write-Host "HTML отчет сохранен: $FilePath" -ForegroundColor Green
}


# Основная логика
try {
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "🚀 ЗАПУСК КОМПЛЕКСНОЙ ПРОВЕРКИ ACTIVE DIRECTORY" -ForegroundColor Cyan
    Write-Host "Домен: $Domain" -ForegroundColor White
    Write-Host "Время начала: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "="*80 -ForegroundColor Cyan
    
    # Получение информации о домене
    Write-Host "`n[1/8] 📁 Получение информации о домене..." -ForegroundColor Yellow
    try {
        $HealthReport.DomainInfo = Get-ADDomain @ADParams -Server $Domain
        Write-Host "  ✓ Домен: $($HealthReport.DomainInfo.DNSRoot)" -ForegroundColor Green
        Write-Host "  ✓ Лес: $($HealthReport.DomainInfo.Forest)" -ForegroundColor Green
    }
    catch {
        Write-Error "Ошибка получения информации о домене: $_"
        return
    }
    
    # Получение списка контроллеров домена
    Write-Host "`n[2/8] 🖥️ Поиск контроллеров домена..." -ForegroundColor Yellow
    $domainControllers = @()
    
    if ($DCHostnames) {
        foreach ($dc in $DCHostnames) {
            try {
                $dcInfo = Get-ADDomainController -Identity $dc @ADParams
                $domainControllers += $dcInfo
                Write-Host "  ✓ $dc" -ForegroundColor Green
            }
            catch {
                Write-Warning "  ✗ Не удалось найти DC: $dc"
            }
        }
    }
    else {
        try {
            $domainControllers = Get-ADDomainController -Filter * @ADParams
            Write-Host "  ✓ Найдено контроллеров: $($domainControllers.Count)" -ForegroundColor Green
            $domainControllers | ForEach-Object {
                Write-Host "    - $($_.HostName) ($($_.Site))" -ForegroundColor Gray
            }
        }
        catch {
            Write-Error "Ошибка получения списка DC: $_"
            return
        }
    }
    
    if ($domainControllers.Count -eq 0) {
        Write-Error "Контроллеры домена не найдены"
        return
    }
    
    $HealthReport.Summary.TotalDCs = $domainControllers.Count
    
    # Проверка каждого контроллера домена
    $dcCounter = 0
    foreach ($dc in $domainControllers) {
        $dcCounter++
        Write-Host "`n[3/8] 🔍 Проверка контроллера: $($dc.HostName) [$dcCounter/$($domainControllers.Count)]" -ForegroundColor Magenta
        
        $dcHealth = @{
            ComputerName = $dc.HostName
            IPAddress = $dc.IPv4Address
            Site = $dc.Site
            OS = $dc.OperatingSystem
            Roles = @($dc.OperationMasterRoles | ForEach-Object { $_.ToString() })
            HostName = $dc.Name
            UpTime = 'Unknown'
            OverallStatus = 'Healthy'
            CriticalIssues = 0
            WarningIssues = 0
            ReplicationStatus = @()
            ServiceStatus = @()
            EventLogChecks = @()
            SystemResources = @()
        }
        
        # Проверка доступности
        Write-Host "  [1/7] 🔗 Проверка доступности..." -ForegroundColor Gray
        try {
            $pingTest = Test-Connection -ComputerName $dc.HostName -Count 2 -Quiet
            if ($pingTest) {
                $dcHealth.ReplicationStatus += [PSCustomObject]@{
                    Test = 'Ping Test'
                    Status = 'Healthy'
                    Details = 'Доступен по сети'
                }
            }
            else {
                $dcHealth.ReplicationStatus += [PSCustomObject]@{
                    Test = 'Ping Test'
                    Status = 'Critical'
                    Details = 'Не отвечает на ping'
                }
                $dcHealth.CriticalIssues++
                Write-Host "    ✗ Не отвечает на ping" -ForegroundColor Red
                continue  # Пропускаем дальнейшие проверки для недоступного DC
            }
        }
        catch {
            $dcHealth.ReplicationStatus += [PSCustomObject]@{
                Test = 'Ping Test'
                Status = 'Critical'
                Details = "Ошибка: $_"
            }
            $dcHealth.CriticalIssues++
            continue
        }
        
        # Получение uptime
        try {
            $osInfo = Get-WmiObject -ComputerName $dc.HostName -Class Win32_OperatingSystem -ErrorAction Stop
            $lastBoot = $osInfo.ConvertToDateTime($osInfo.LastBootUpTime)
            $uptime = (Get-Date) - $lastBoot
            $dcHealth.UpTime = "$([Math]::Floor($uptime.TotalDays)) дн. $($uptime.Hours) ч."
        }
        catch {
            $dcHealth.UpTime = 'Unknown'
        }
        
        # Проверка репликации
        if ($CheckDC) {
            Write-Host "  [2/7] 🔄 Проверка репликации..." -ForegroundColor Gray
            try {
                $replPartners = Get-ADReplicationPartnerMetadata -Target $dc.HostName @ADParams -ErrorAction Stop
                
                foreach ($partner in $replPartners) {
                    $status = 'Healthy'
                    if ($partner.LastReplicationResult -ne 0) { $status = 'Critical' }
                    elseif ($partner.LastReplicationSuccess -lt (Get-Date).AddHours(-24)) { $status = 'Warning' }
                    
                    $dcHealth.ReplicationStatus += [PSCustomObject]@{
                        Test = "Replication to $($partner.Partner.Split(',')[0].Replace('CN=',''))"
                        Status = $status
                        Details = "Последняя успешная: $($partner.LastReplicationSuccess)"
                    }
                    
                    if ($status -eq 'Critical') { $dcHealth.CriticalIssues++ }
                    elseif ($status -eq 'Warning') { $dcHealth.WarningIssues++ }
                }
                
                Write-Host "    ✓ Проверено $($replPartners.Count) партнеров" -ForegroundColor Green
            }
            catch {
                $dcHealth.ReplicationStatus += [PSCustomObject]@{
                    Test = 'Replication Check'
                    Status = 'Critical'
                    Details = "Ошибка: $_"
                }
                $dcHealth.CriticalIssues++
                Write-Host "    ✗ Ошибка проверки репликации" -ForegroundColor Red
            }
        }
        
        # Проверка служб
        if ($CheckServices) {
            Write-Host "  [3/7] ⚙️ Проверка служб..." -ForegroundColor Gray
            foreach ($serviceName in $CriticalServices) {
                try {
                    $service = Get-Service -ComputerName $dc.HostName -Name $serviceName -ErrorAction SilentlyContinue
                    
                    if ($service) {
                        $status = if ($service.Status -eq 'Running') { 'Healthy' } else { 'Critical' }
                        
                        $dcHealth.ServiceStatus += [PSCustomObject]@{
                            Name = $serviceName
                            DisplayName = $service.DisplayName
                            Status = $status
                            State = $service.Status
                            StartType = $service.StartType
                        }
                        
                        if ($status -eq 'Critical') { $dcHealth.CriticalIssues++ }
                    }
                }
                catch {
                    # Служба не найдена - это нормально для некоторых DC
                }
            }
            Write-Host "    ✓ Проверено $($dcHealth.ServiceStatus.Count) служб" -ForegroundColor Green
        }
        
        # Проверка DNS
        if ($CheckDNS) {
            Write-Host "  [4/7] 🌐 Проверка DNS..." -ForegroundColor Gray
            $dnsTests = Test-DNSService -ComputerName $dc.HostName
            $dcHealth.ReplicationStatus += $dnsTests
            
            foreach ($test in $dnsTests) {
                if ($test.Status -eq 'Critical') { $dcHealth.CriticalIssues++ }
                elseif ($test.Status -eq 'Warning') { $dcHealth.WarningIssues++ }
            }
            Write-Host "    ✓ Проверка DNS завершена" -ForegroundColor Green
        }
        
        # Проверка SYSVOL
        if ($CheckSysvol) {
            Write-Host "  [5/7] 📂 Проверка SYSVOL..." -ForegroundColor Gray
            $sysvolTests = Test-SysvolReplication -ComputerName $dc.HostName
            $dcHealth.ReplicationStatus += $sysvolTests
            
            foreach ($test in $sysvolTests) {
                if ($test.Status -eq 'Critical') { $dcHealth.CriticalIssues++ }
                elseif ($test.Status -eq 'Warning') { $dcHealth.WarningIssues++ }
            }
            Write-Host "    ✓ Проверка SYSVOL завершена" -ForegroundColor Green
        }
        
        # Проверка системных ресурсов
        Write-Host "  [6/7] 📊 Проверка системных ресурсов..." -ForegroundColor Gray
        $resources = Get-SystemResources -ComputerName $dc.HostName
        $dcHealth.SystemResources = $resources
        
        foreach ($resource in $resources) {
            if ($resource.Status -eq 'Warning') { $dcHealth.WarningIssues++ }
        }
        Write-Host "    ✓ Проверка ресурсов завершена" -ForegroundColor Green
        
        # Проверка журналов событий
        if ($CheckEventLogs) {
            Write-Host "  [7/7] 📝 Проверка журналов событий..." -ForegroundColor Gray
            $events = Get-EventLogSummary -ComputerName $dc.HostName -HoursBack ($DaysBack * 24)
            $dcHealth.EventLogChecks = $events
            
            if ($events.Count -gt 0) {
                $dcHealth.WarningIssues += $events.Count
                Write-Host "    ⚠ Найдено $($events.Count) критических событий" -ForegroundColor Yellow
            }
            else {
                Write-Host "    ✓ Критических событий не найдено" -ForegroundColor Green
            }
        }
        
        # Определение общего статуса DC
        if ($dcHealth.CriticalIssues -gt 0) {
            $dcHealth.OverallStatus = 'Critical'
            $HealthReport.Summary.CriticalDCs++
        }
        elseif ($dcHealth.WarningIssues -gt 0) {
            $dcHealth.OverallStatus = 'Warning'
            $HealthReport.Summary.WarningDCs++
        }
        else {
            $dcHealth.OverallStatus = 'Healthy'
            $HealthReport.Summary.HealthyDCs++
        }
        
        $HealthReport.DomainControllers += $dcHealth
        $HealthReport.Summary.TotalIssues += ($dcHealth.CriticalIssues + $dcHealth.WarningIssues)
        $HealthReport.Summary.CriticalIssues += $dcHealth.CriticalIssues
        $HealthReport.Summary.WarningIssues += $dcHealth.WarningIssues
        
        Write-Host "  📊 Итог: $($dcHealth.OverallStatus) (Критично: $($dcHealth.CriticalIssues), Предупреждений: $($dcHealth.WarningIssues))" -ForegroundColor $(if($dcHealth.OverallStatus -eq 'Critical'){'Red'}elseif($dcHealth.OverallStatus -eq 'Warning'){'Yellow'}else{'Green'})
    }
    
    # Определение общего статуса отчета
    if ($HealthReport.Summary.CriticalDCs -gt 0) {
        $HealthReport.OverallStatus = 'Critical'
    }
    elseif ($HealthReport.Summary.WarningDCs -gt 0) {
        $HealthReport.OverallStatus = 'Warning'
    }
    else {
        $HealthReport.OverallStatus = 'Healthy'
    }
    
    # Генерация отчета
    Write-Host "`n[4/8] 📄 Генерация HTML отчета..." -ForegroundColor Yellow
    $reportFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    New-HealthHTMLReport -HealthData $HealthReport -FilePath $reportFullPath
    
    # Отправка email
    if ($SendEmail -and $EmailTo) {
        Write-Host "`n[5/8] 📧 Отправка отчета по email..." -ForegroundColor Yellow
        if (-not $SmtpServer) {
            $SmtpServer = $domainControllers[0].HostName
        }
        
        $emailSubject = "AD Health Report: $($HealthReport.DomainInfo.DNSRoot) - $($HealthReport.OverallStatus) - $(Get-Date -Format 'dd.MM.yyyy')"
        $emailBody = @"
<html>
<body>
    <h2>Отчет о состоянии Active Directory</h2>
    <p><strong>Домен:</strong> $($HealthReport.DomainInfo.DNSRoot)</p>
    <p><strong>Общий статус:</strong> <span style="color: $(Get-StatusColor -Status $HealthReport.OverallStatus)">$($HealthReport.OverallStatus)</span></p>
    <p><strong>Контроллеров:</strong> $($HealthReport.Summary.TotalDCs) (Здоровых: $($HealthReport.Summary.HealthyDCs), Предупреждений: $($HealthReport.Summary.WarningDCs), Критических: $($HealthReport.Summary.CriticalDCs))</p>
    <p><strong>Проблем:</strong> $($HealthReport.Summary.TotalIssues) (Критических: $($HealthReport.Summary.CriticalIssues), Предупреждений: $($HealthReport.Summary.WarningIssues))</p>
    <p><strong>Время генерации:</strong> $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
    <p>Детальный отчет находится во вложении.</p>
    <hr>
    <p><small>Отчет сгенерирован автоматически</small></p>
</body>
</html>
"@
        
        try {
            Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $emailSubject `
                -Body $emailBody -BodyAsHtml -Attachments $reportFullPath `
                -SmtpServer $SmtpServer -Encoding UTF8 -ErrorAction Stop
            
            Write-Host "  ✓ Отчет отправлен: $($EmailTo -join ', ')" -ForegroundColor Green
        }
        catch {
            Write-Warning "  ✗ Не удалось отправить email: $_"
        }
    }
    
    # Финальная сводка
    $endTime = Get-Date
    $executionTime = New-TimeSpan -Start $StartTime -End $endTime
    
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "🎉 ПРОВЕРКА ЗАВЕРШЕНА" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    Write-Host "`n📊 СВОДНЫЕ РЕЗУЛЬТАТЫ:" -ForegroundColor White
    Write-Host "  • Домен: $($HealthReport.DomainInfo.DNSRoot)" -ForegroundColor Gray
    Write-Host "  • Контроллеров: $($HealthReport.Summary.TotalDCs)" -ForegroundColor Gray
    Write-Host "  • Здоровых: $($HealthReport.Summary.HealthyDCs)" -ForegroundColor Green
    Write-Host "  • С предупреждениями: $($HealthReport.Summary.WarningDCs)" -ForegroundColor Yellow
    Write-Host "  • Критических: $($HealthReport.Summary.CriticalDCs)" -ForegroundColor Red
    Write-Host "  • Всего проблем: $($HealthReport.Summary.TotalIssues)" -ForegroundColor Gray
    Write-Host "  • Общий статус: $($HealthReport.OverallStatus)" -ForegroundColor $(Get-StatusColor -Status $HealthReport.OverallStatus)
    Write-Host "  • Отчет: $reportFullPath" -ForegroundColor Cyan
    Write-Host "  • Время выполнения: $($executionTime.TotalMinutes.ToString('0.00')) минут" -ForegroundColor Gray
    
    Write-Host "`n⚠ РЕКОМЕНДАЦИИ:" -ForegroundColor White
    if ($HealthReport.Summary.CriticalIssues -gt 0) {
        Write-Host "  • Срочно устраните критические проблемы!" -ForegroundColor Red
    }
    if ($HealthReport.Summary.WarningIssues -gt 0) {
        Write-Host "  • Запланируйте устранение предупреждений" -ForegroundColor Yellow
    }
    if ($HealthReport.Summary.CriticalIssues -eq 0 -and $HealthReport.Summary.WarningIssues -eq 0) {
        Write-Host "  • Все системы функционируют нормально" -ForegroundColor Green
    }
    
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    
}
catch {
    Write-Error "Критическая ошибка при выполнении проверки: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    Write-Host "`nЗавершение работы скрипта..." -ForegroundColor Gray
}
