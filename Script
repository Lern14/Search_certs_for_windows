param( [string] $auditlist)

Function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>.

<meta name="save" content="history">

<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 8px; COLOR: #ffffff; FONT-FAMILY: Arial; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
.dspcont { display:none; BORDER-RIGHT: #B1BABF 1px solid; BORDER-TOP: #B1BABF 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; WIDTH: 95%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #CC0000}
a.dsphead1 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead2 {BORDER-RIGHT: #B1BABF 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #B1BABF 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #B1BABF 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #B1BABF 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}

#___________________________________________________________________________________________________________________Управление HTML форматом(строки\столбцы)
td {PADDING-RIGHT: 1px; VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma; WORD-WRAP: break-word;}
th {VERTICAL-ALIGN: TOP; COLOR: #CC0000; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt}
TABLE {
    TABLE-LAYOUT: fixed; 
    FONT-SIZE: 100%; 
    WIDTH: 100%;
    margin-bottom:  20px; /* Добавленная строка */
}
</style>


<script type="text/javascript">
function dsp(loc){
   if(document.getElementById){
      var foc=loc.firstChild;
      foc=loc.firstChild.innerHTML?
         loc.firstChild:
         loc.firstChild.nextSibling;
      foc.innerHTML=foc.innerHTML==''?'':'';
      foc=loc.parentNode.nextSibling.style?
         loc.parentNode.nextSibling:
         loc.parentNode.nextSibling.nextSibling;
      foc.style.display=foc.style.display=='block'?'none':'block';}}  

if(!document.getElementById)
   document.write('<style type="text/css">\n'+'.dspcont{display:block;}\n'+ '</style>');
</script>

</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#CC0000">
<font face="Arial" size="1"><b>VBANK190</b></font><br>
<font face="Arial" size="1">Отчет создан\обновлен $(Get-Date)</font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
Return $Report
}

Function Get-CustomHeader0 ($Title){
$Report = @"
        <h1><a class="dsphead0">$($Title)</a></h1>
    <div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader ($Num, $Title){
$Report = @"
    <h2><a href="javascript:void(0)" class="dsphead$($Num)" onclick="dsp(this)">
    <span class="expando"></span>$($Title)</a></h2>
    <div class="dspcont">
"@
Return $Report
}

Function Get-CustomHeaderClose{

    $Report = @"
        </DIV>
        <div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader0Close{

    $Report = @"
</DIV>
"@
Return $Report
}

Function Get-CustomHTMLClose{

    $Report = @"
</div>

</body>
</html>
"@
Return $Report
}

Function Get-HTMLTable{
    param([array]$Content)
    $HTMLTable = $Content | ConvertTo-Html
    $HTMLTable = $HTMLTable -replace '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">', ""
    $HTMLTable = $HTMLTable -replace '<html xmlns="http://www.w3.org/1999/xhtml">', ""
    $HTMLTable = $HTMLTable -replace '<head>', ""
    $HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
    $HTMLTable = $HTMLTable -replace '</head><body>', ""
    $HTMLTable = $HTMLTable -replace '</body></html>', ""
    Return $HTMLTable
}


Function Get-HTMLDetail ($Heading, $Detail){
$Report = @"
<TABLE>
    <tr>
    <th width='35%'><b>$Heading</b></font></th>
    <td width='75%'>$($Detail)</td>
    </tr>
</TABLE>
"@
Return $Report
}

if ($auditlist -eq ""){
    Write-Host "No list specified, using $env:computername"
    $targets = $env:computername
}
else
{
    if ((Test-Path $auditlist) -eq $false)
    {
        Write-Host "Invalid audit path specified: $auditlist"
        exit
    }
    else
    {
        Write-Host "Using Audit list: $auditlist"
        $Targets = Get-Content $auditlist
    }
}
Function ExtractDetail($subject, $key) {
    if ($subject -match "$key\s*=\s*([^,]+)") {
        return $matches[1].Trim()
    } else {
        return $null
    }
}

#_______________________________Присваевыем пользователю Дату и ID

Function Get-ReportID {
    return "$(Get-Date -Format "yyyyMMdd")-$env:USERNAME"
}

Foreach ($Target in $Targets){

Write-Output "Collating Detail for $Target"
    $ComputerSystem = Get-WmiObject -computername $Target Win32_ComputerSystem
    switch ($ComputerSystem.DomainRole){
        0 { $ComputerRole = "Standalone Workstation" }
        1 { $ComputerRole = "Member Workstation" }
        2 { $ComputerRole = "Standalone Server" }
        3 { $ComputerRole = "Member Server" }
        4 { $ComputerRole = "Domain Controller" }
        5 { $ComputerRole = "Domain Controller" }
        default { $ComputerRole = "Information not available" }
    }
$CurrentUserStorePath = "Cert:\CurrentUser\My" (Хранилище сертификатов)
$CurrentUserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

#____________________ Получение сертификатов пользователя
$certificates = Get-ChildItem -Path $CurrentUserStorePath


$reportPath = "C:\!ReportCRT\report.html" (Файл отчёта)
$reportID = Get-ReportID

# ____________________Если отчет уже существует, проверяем его на наличие текущего ID отчета
if (Test-Path $reportPath) {
    $currentReport = Get-Content -Path $reportPath -Raw

    if ($currentReport -like "*$CurrentUserName*" ) {
        Write-Output "Информация по $CurrentUserName уже содержится в отчете."
        continue
        }
        # Удаляем старое закрытие отчета
        $baseReport = $currentReport -replace '<div class="filler"></div>', ''
        
        # Вставляем новые данные перед закрывающим тегом
        $HTMLReport = $baseReport

} else {
    $HTMLReport = Get-CustomHTML "Отчет по сертификатам пользователей"
}

$HTMLReport += Get-CustomHeader 1 "Список сертификатов $CurrentUserName"


$certificatesArray = @()

# ____________________Перебор всех сертификатов и добавление информации в массив
foreach ($cert in $certificates) {

    $keyUsage = ($cert.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.15" }).KeyUsages

    $SNValue = ExtractDetail $cert.Subject "SN"
    $GValue = ExtractDetail $cert.Subject "G"
    $CNValue = ExtractDetail $cert.Subject "CN"
    $CNVOrgan = ExtractDetail $cert.Issuer "CN"
    

    $certName = "$SNValue $GValue"
    $certNameIssuer = $CNValue
    $certOrgan = $CNVOrgan

    $certObj = [PSCustomObject]@{
        'Владелец сертификата' = $certName
        'Серийный номер сертификата' = $cert.SerialNumber
        'Отпечаток сертификата' = $cert.Thumbprint
        'Назначение ключа' = $keyUsage
        'Компания' = $certNameIssuer
        'Издатель сертификата' = $CNVOrgan
        'Дата начала действия сертификата' = $cert.NotBefore
        'Дата завершения действия сертификата' = $cert.NotAfter
    }
    $certificatesArray += $certObj
}
#_______________________ Добавляем таблицу с сертификатами в отчет
$HTMLTable = Get-HTMLTable -Content $certificatesArray
$HTMLReport += $HTMLTable

$HTMLReport += Get-CustomHeaderClose
$HTMLReport += Get-CustomHTMLClose


$HTMLReport |  Out-File -FilePath $reportPath
