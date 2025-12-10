<#
.SYNOPSIS
    Получает полное рекурсивное членство пользователя или группы в группах Active Directory.
.DESCRIPTION
    Функция определяет все группы Active Directory (включая вложенные), в которые входит указанный субъект.
    Поддерживает глубокую рекурсию, может работать с доменными и локальными группами (при указании компьютера).
    Формирует детальный отчет с иерархией вложения.
.PARAMETER Identity
SamAccountName, DistinguishedName или SID пользователя/группы для анализа.
.PARAMETER Server
        Контроллер домена для выполнения запросов (опционально).
.PARAMETER OutputType
    Формат вывода результатов: List (простой список), Tree (древовидная структура), Detailed (объекты с полными свойствами).
.PARAMETER IncludeLocal
    Включает проверку членства в локальных группах указанного компьютера (требуются права администратора).
.PARAMETER ComputerName
    Имя компьютера для проверки локальных групп (только с параметром -IncludeLocal).
.EXAMPLE
    Find-UserGroupMembershipRecursive -Identity "SG-Finance" -OutputType Tree | Format-List
    Отображает древовидную структуру вложенности для группы SG-Finance
    Find-UserGroupMembershipRecursive -Identity "petrova.a" -IncludeLocal -ComputerName "SRV-FILE01"
    Проверяет членство пользователя в группах AD и локальных группах на сервере SRV-FILE01

    # Пример 1: Поиск по имени пользователя
    .\Find-UserGroupMembershipRecursive.ps1 -Identity "ivanov.i"

    # Пример 2: Поиск по имени группы  
    .\Find-UserGroupMembershipRecursive.ps1 -Identity "SG-Finance"

    # Пример 3: Поиск по полному DN
    .\Find-UserGroupMembershipRecursive.ps1 -Identity "CN=Ivanov Ivan,OU=Users,DC=company,DC=local"

.NOTES
    Author: Dmitry Plotinsky
    Version: 2.1
    Date: 2024-03-15
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Identity,
    
    [Parameter(Mandatory=$false)]
    [string]$Server,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("List", "Tree", "Detailed")]
    [string]$OutputType = "List",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeLocal,
    
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME
)

# Инициализация
# Проверка и импорт необходимого модуля
if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Verbose "Модуль ActiveDirectory успешно загружен"
    }
    catch {
        Write-Error "Не удалось загрузить модуль ActiveDirectory. Установите RSAT или запустите на контроллере домена."
        return
    }
}

# Хэш-таблицы для хранения результатов и предотвращения циклов
$global:AllGroups = @{}  # Все найденные группы: ключ=DN, значение=объект группы
$global:MembershipTree = @{}  # Дерево членства: ключ=группа, значение=родительские группы
$global:Visited = @{}  # Посещенные группы для избежания зацикливания
$global:GroupStack = New-Object System.Collections.Stack  # Стек для рекурсии
#endregion

#region Вспомогательные функции
function Get-ADObjectRecursive {
    param(
        [string]$ObjectDN,
        [string]$DC,
        [string]$ParentGroup = $null
    )
    
    # Проверка на циклические ссылки
    if ($global:Visited.ContainsKey($ObjectDN)) {
        Write-Verbose "Обнаружен цикл или уже посещенная группа: $ObjectDN"
        return
    }
    $global:Visited[$ObjectDN] = $true
    
    try {
        # Получаем все группы, в которых состоит текущий объект
        $memberOf = Get-ADObject -Identity $ObjectDN -Server $DC -Properties memberOf -ErrorAction Stop
        
        if ($memberOf.memberOf) {
            foreach ($groupDN in $memberOf.memberOf) {
                # Добавляем связь в дерево
                if (-not $global:MembershipTree.ContainsKey($groupDN)) {
                    $global:MembershipTree[$groupDN] = New-Object System.Collections.ArrayList
                }
                $null = $global:MembershipTree[$groupDN].Add($ParentGroup)
                
                # Получаем полные данные о группе (если еще не получали)
                if (-not $global:AllGroups.ContainsKey($groupDN)) {
                    try {
                        $group = Get-ADGroup -Identity $groupDN -Server $DC -Properties SamAccountName, Name, GroupCategory, GroupScope, DistinguishedName, SID
                        $global:AllGroups[$groupDN] = $group
                        Write-Verbose "Найдена группа: $($group.SamAccountName)"
                        
                        # Рекурсивно проверяем членство этой группы в других группах
                        Get-ADObjectRecursive -ObjectDN $groupDN -DC $DC -ParentGroup $groupDN
                    }
                    catch {
                        Write-Warning "Не удалось получить информацию о группе: $groupDN. Ошибка: $_"
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Ошибка при обработке объекта $ObjectDN : $_"
    }
}

function Build-Tree {
    param(
        [string]$RootDN,
        [int]$Level = 0,
        [hashtable]$Tree
    )
    
    $indent = "  " * $Level
    $result = @()
    
    # Находим все группы, для которых текущая группа является родителем
    $childGroups = $Tree.Keys | Where-Object { $Tree[$_] -contains $RootDN }
    
    if ($global:AllGroups.ContainsKey($RootDN)) {
        $group = $global:AllGroups[$RootDN]
        $result += "$indent├─ $($group.SamAccountName) [$($group.Name)]"
    }
    else {
        $result += "$indent├─ $RootDN"
    }
    
    foreach ($child in $childGroups) {
        $result += Build-Tree -RootDN $child -Level ($Level + 1) -Tree $Tree
    }
    
    return $result
}

function Get-LocalGroupMembership {
    param(
        [string]$UserName,
        [string]$TargetComputer
    )
    
    $localGroups = @()
    
    try {
        # Используем WMI для получения локальных групп
        $computer = [ADSI]"WinNT://$TargetComputer,computer"
        $groups = $computer.Children | Where-Object { $_.SchemaClassName -eq "group" }
        
        foreach ($group in $groups) {
            $groupName = $group.Name[0]
            $groupPath = "WinNT://$TargetComputer/$groupName"
            $groupObj = [ADSI]$groupPath
            
            # Проверяем, является ли пользователь членом группы
            $members = $groupObj.Invoke("Members")
            foreach ($member in $members) {
                $memberPath = $member.GetType().InvokeMember("Adspath", 'GetProperty', $null, $member, $null)
                if ($memberPath -like "*/$UserName" -or $memberPath -like "*/$UserName,*") {
                    $localGroups += [PSCustomObject]@{
                        GroupName = $groupName
                        Computer = $TargetComputer
                        GroupType = "Local"
                        DistinguishedName = "CN=$groupName,CN=LocalGroups,$TargetComputer"
                        SamAccountName = "$TargetComputer\$groupName"
                    }
                    break
                }
            }
        }
    }
    catch {
        Write-Warning "Ошибка при проверке локальных групп на $TargetComputer : $_"
    }
    
    return $localGroups
}

# Основная логика
try {
    Write-Host "`n=== Анализ членства в группах ===" -ForegroundColor Cyan
    Write-Host "Субъект: $Identity" -ForegroundColor White
    Write-Host "Время начала: $(Get-Date -Format 'HH:mm:ss')`n" -ForegroundColor Gray
    
    # Определяем параметры для подключения к AD
    $adParams = @{}
    if ($Server) { $adParams.Server = $Server }
    
    # Получаем информацию об исходном объекте
    Write-Verbose "Поиск объекта в Active Directory..."
    try {
        $targetObject = Get-ADObject -Filter { 
            (SamAccountName -eq $Identity) -or 
            (DistinguishedName -eq $Identity) -or 
            (SID -eq $Identity)
        } @adParams -Properties SamAccountName, DistinguishedName, ObjectClass, Name -ErrorAction Stop
        
        if (-not $targetObject) {
            Write-Error "Объект '$Identity' не найден в Active Directory."
            return
        }
    }
    catch {
        Write-Error "Ошибка при поиске объекта: $_"
        return
    }
    
    Write-Host "Найден объект: $($targetObject.Name) [$($targetObject.ObjectClass)]" -ForegroundColor Green
    
    # Запускаем рекурсивный поиск
    Write-Verbose "Выполняется рекурсивный поиск групп..."
    Get-ADObjectRecursive -ObjectDN $targetObject.DistinguishedName -DC $Server -ParentGroup "SOURCE"
    
    # Обработка локальных групп (если запрошено)
    $localMembership = @()
    if ($IncludeLocal) {
        Write-Host "`nПроверка локальных групп на компьютере: $ComputerName" -ForegroundColor Yellow
        $localMembership = Get-LocalGroupMembership -UserName $targetObject.SamAccountName -TargetComputer $ComputerName
    }

# Формирование результатов
    Write-Host "`n=== РЕЗУЛЬТАТЫ ===" -ForegroundColor Cyan
    Write-Host "Всего найдено групп Active Directory: $($global:AllGroups.Count)" -ForegroundColor White
    
    switch ($OutputType) {
        "List" {
            # Простой отсортированный список
            $global:AllGroups.Values | Sort-Object SamAccountName | Select-Object `
                @{N="Тип";E={$_.ObjectClass}},
                SamAccountName,
                Name,
                @{N="Область";E={$_.GroupScope}},
                @{N="Категория";E={$_.GroupCategory}},
                DistinguishedName
        }
        
        "Tree" {
            # Древовидное представление
            $treeOutput = Build-Tree -RootDN "SOURCE" -Tree $global:MembershipTree
            
            # Удаляем маркер SOURCE из вывода
            $treeOutput = $treeOutput | ForEach-Object {
                $_ -replace 'SOURCE', "$($targetObject.SamAccountName) [Исходный объект]"
            }
            
            # Выводим дерево
            $treeOutput | ForEach-Object {
                if ($_ -match "\[Исходный объект\]") {
                    Write-Host $_ -ForegroundColor Magenta
                } elseif ($_ -match "^\s+├─") {
                    Write-Host $_ -ForegroundColor Green
                } else {
                    Write-Host $_
                }
            }
            
            # Таблица соответствий DN -> SamAccountName для удобства
            Write-Host "`n--- Справочник групп ---" -ForegroundColor DarkGray
            $global:AllGroups.Values | Sort-Object SamAccountName | ForEach-Object {
                Write-Host "$($_.SamAccountName.PadRight(30)) : $($_.Name)" -ForegroundColor DarkGray
            }
        }
        
        "Detailed" {
            # Подробные объекты с дополнительными свойствами
            $results = foreach ($group in $global:AllGroups.Values) {
                [PSCustomObject]@{
                    SamAccountName = $group.SamAccountName
                    Name = $group.Name
                    DistinguishedName = $group.DistinguishedName
                    SID = $group.SID.Value
                    GroupScope = $group.GroupScope
                    GroupCategory = $group.GroupCategory
                    ParentGroups = if ($global:MembershipTree.ContainsKey($group.DistinguishedName)) {
                        $global:MembershipTree[$group.DistinguishedName] | ForEach-Object {
                            if ($global:AllGroups.ContainsKey($_)) {
                                $global:AllGroups[$_].SamAccountName
                            } else { $_ }
                        }
                    }
                    ObjectClass = $group.ObjectClass
                }
            }
            
            # Добавляем локальные группы к результатам
            if ($localMembership) {
                $results += $localMembership | ForEach-Object {
                    [PSCustomObject]@{
                        SamAccountName = $_.SamAccountName
                        Name = $_.GroupName
                        DistinguishedName = $_.DistinguishedName
                        SID = "N/A"
                        GroupScope = "Local"
                        GroupCategory = "Security"
                        ParentGroups = @()
                        ObjectClass = "Group"
                    }
                }
            }
            
            $results | Sort-Object GroupScope, SamAccountName
        }
    }
    
    # Вывод информации о локальных группах отдельно, если не Detailed формат
    if ($localMembership -and $OutputType -ne "Detailed") {
        Write-Host "`n--- Локальные группы ($ComputerName) ---" -ForegroundColor Yellow
        $localMembership | Select-Object GroupName, Computer | Format-Table -AutoSize
    }
    
    # Сводная статистика
    Write-Host "`n--- Статистика ---" -ForegroundColor Cyan
    $groupStats = $global:AllGroups.Values | Group-Object GroupScope | Select-Object Name, Count
    foreach ($stat in $groupStats) {
        Write-Host "$($stat.Name): $($stat.Count)" -ForegroundColor White
    }
    
    if ($localMembership) {
        Write-Host "Локальные: $($localMembership.Count)" -ForegroundColor White
    }
    
    $endTime = Get-Date
    Write-Host "`nВремя выполнения: $((New-TimeSpan -Start $startTime -End $endTime).TotalSeconds.ToString('0.00')) сек." -ForegroundColor Gray
    
}
catch {
    Write-Error "Критическая ошибка в скрипте: $_"
    Write-Error $_.ScriptStackTrace
}
finally {
    # Очистка глобальных переменных
    Remove-Variable -Name AllGroups, MembershipTree, Visited, GroupStack -Scope Global -ErrorAction SilentlyContinue
}
