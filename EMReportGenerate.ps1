param (
    [Parameter()]
    $IsDebugRunning = $true
)

function Get-ObjNumber([object]$objItem) {
    if ($null -eq $objItem) {
        return 0;
    }
    if ($objItem -isnot [System.Array]) {
        return 1;
    }
    if ($objItem -is [system.array]) {
        return $objItem.Count
    }
    if ($objItem -is [system.Object]) {
        return 1;
    }
}

function SendReportByMail {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [string[]]$receiver, 
        [string]$senderUser, 
        [string]$subject, 

        [Parameter(Mandatory = $false)] 
        [string]$mailBody,
        [string[]]$ccTo,
        [string[]]$bccTo,
        [string]$attachment,
        [string[]]$attachments,
        [string]$mailServer = 'mx.microchip.com',
        [bool]$isHTML = $false
    )

    $testResult = ""
    
    try {
        $mailCommand = "Send-MailMessage -To $('$receiver') $(if ($ccTo) {'-Cc $ccTo'}) $(if ($bccTo) {'-Bcc $bccTo'}) -From '$senderUser' -Subject '$subject' $(if ($mailBody) {'-Body $mailBody -Encoding ([System.Text.Encoding]::UTF8)'}) $(if ($attachment) {'-Attachments $attachment'}) $(if ($attachments) {'-Attachments $attachments'}) $(if ($isHTML) {'-BodyAsHtml'}) -SmtpServer '$mailServer'"
        Invoke-Expression $mailCommand -ErrorAction Stop;

        $testResult = Test-Connection $mailServer -Count 1
        if ($testResult) {
            Write-Host("Mail server: {0}, IP addr: {1}, Status: {2}" -f $testResult.Address, $testResult.IPV4Address.IPAddressToString, $testResult.StatusCode)
            Write-ToLogFile -LogContent ("Mail server: {0}, IP addr: {1}, Status: {2}" -f $testResult.Address, $testResult.IPV4Address.IPAddressToString, $testResult.StatusCode)
        }
    }
    catch {
        Write-Host ("Get error from when send report by mail !!!, detail: {0}" -f $error[0])
        Write-ToLogFile -LogContent ("Get error from when send report by mail, detail: {0}" -f $error[0])
    }
}

function Write-ToLogFile {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 

        [Parameter(Mandatory = $false)] 
        [Alias('LogPath')] 
        [string]$Path = $LogFile
    )
    Process { 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        if (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path.";
            New-Item $Path -Force -ItemType File ;
        } 
        else {
            # Nothing to see here yet. 
            if ($Message -eq $PDTWINITSTATUS) {
                New-Item $Path -Force -ItemType File ;
            }
        } 

        # Write log entry to $Path 
        "$Message" | Out-File -FilePath $Path -Append;
    } 
    End {
    }
}

function DoFileArchive {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("ArchiveFolder")] 
        [string]$archiveFolderName, 
        [Alias("FileLists")] 
        [object]$archiveFileLists
    )

    Write-Host ("Move Excel report files to: {0}" -f $archiveFolderName)
    Write-ToLogFile -LogContent ("Move Excel report files to: {0}" -f $archiveFolderName)
    if ((Test-Path -Path $archiveFolderName) -eq $false) {
        New-Item -Path $archiveFolderName -ItemType Directory
    }

    Foreach ($fileName in $archiveFileLists) {
        if (Test-Path -Path $fileName) {
            Move-Item $fileName -Destination $archiveFolderName -Force
        }
    }
}

function GenerateReportAMER {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("SourceXlsxData")] 
        [object]$sourceData, 
        [Alias("ExportXlssxFilename")] 
        [string]$exportFileName, 

        [Parameter(Mandatory = $false)]
        [bool]$isTesting = $false
    )

    if ($isTesting) {
        [string[]]$mailTo = @("Jimmy.Sha@microchip.com")
        [string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$mailTo = @("Nitin.Zhao@microchip.com", "Lenard.Tai@microchip.com", "Kayle.Liu@microchip.com", "Jimmy.Sha@microchip.com", "Roxie.Lee@microchip.com")
        #[string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$testMailTo = @("Jimmy.Sha@microchip.com")
        #[string[]]$testailCC = @("jsha@sst.com")
    }
    else {
        [string[]]$mailTo = @("Sameer.Ebadi@microchip.com");
        [string[]]$mailCC = @("Jason.So@microchip.com")
    }

    $mailSender = "EM AMER Report<emreport_amer@microchip.com>";
    $mailSubject = "EM AMER Report - " + $TodayDate;

    Write-Host ("Data(AMER) filtering ...")
    Write-ToLogFile -LogContent ("Data(AMER) filtering ...")

    $report_Object = New-Object -TypeName PSObject
    $amer_Devices = $sourceData | Where-Object { $_.Location.ToLower().StartsWith("us") -or `
            $_.Location.ToLower().StartsWith("ca") -and `
        ($_.Location.ToLower().Contains("tempe") -eq $false) -and `
        ($_.Location.ToLower().Contains("colorado") -eq $false) -and `
        ($_.Location.ToLower().Contains("gresham") -eq $false) -and `
        ($_.Location.ToLower().Contains("boulder") -eq $false) -and `
        ($_.Location.ToLower().Contains("subcon") -eq $false) -and `
        ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in America, exclude Subcon devices

    #######################################################
    ### Problem devices in America
    #######################################################
    #$amer_probDevices = $amer_Devices | where {$_.ProblemDevice -eq $true}
    [array]$amer_probDevices = $amer_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$amer_cbProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$amer_cbcProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    [array]$amer_sepProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$amer_patchProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    [array]$amer_mgntByProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$amer_adProbDevices = $amer_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #$amer_StateAntivirusDevices = $amer_sepProbDevices | Where-Object { "non-corp av".Equals($_.StateAntivirus.Trim().ToLower()) -or "not applicable".Equals($_.StateAntivirus.Trim().ToLower()) }
    #$amer_StatePatchDevices = $amer_patchProbDevices | Where-Object { "agent not found".Equals($_.StatePatching.Trim().ToLower()) -or "needs os version".Equals($_.StatePatching.Trim().ToLower()) -or "not applicable".Equals($_.StatePatching.Trim().ToLower()) }
    #$amer_StateCarbonBlack = $amer_cbProbDevices | Where-Object { "not connecting".Equals($_.StateCarbonBlack.Trim().ToLower()) }

    #Write-Host ("Generate Excel file: {0}" -f $exportFileName_1)
    #Write-ToLogFile -LogContent ("Generate Excel file: {0}" -f $exportFileName_1)
    #######################################################
    ### Create source data file with all America Devices
    #######################################################
    #$excel1 = Export-Excel -Path $exportFileName_1 -InputObject $sourceData -WorksheetName ($getDate.ToString("yyyyMMdd").ToString()) -ClearSheet -AutoFilter -FreezePane @(2,2)
    #$excel1 = Export-Excel -Path $exportFileName_1 -InputObject $amer_Devices -WorksheetName 'US' -AutoSize -AutoFilter -FreezePane @(2,2) -PassThru
    #$excel1.Save();
    #$excel1.Dispose();

    #$list = $amer_probDevices | Where-Object { (IsNewSystem -objItem $_) -eq $false }
    Write-Host ("Generate Excel file(AMER): {0}" -f $exportFileName)
    Write-ToLogFile -LogContent ("Generate Excel file(AMER): {0}" -f $exportFileName)
    ##############################################################
    ### Create EPM report file with all America problem Devices
    ##############################################################
    if ($null -ne $amer_adProbDevices -and $amer_adProbDevices.Count -ge 0) {
        $epmExcel = $amer_adProbDevices | Export-Excel -Path $exportFileName -WorksheetName "AD" -ClearSheet -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $amer_cbProbDevices -and $amer_cbProbDevices.Count -ge 0) {
        $epmExcel = $amer_cbProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CB" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $amer_cbcProbDevices -and $amer_cbcProbDevices.Count -ge 0) {
        $epmExcel = $amer_cbcProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CBC" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    #if ($null -ne $amer_sepProbDevices -and $amer_sepProbDevices.Count -ge 0) {
    #    $epmExcel = $amer_sepProbDevices | Export-Excel -Path $exportFileName -WorksheetName "SEP" -AutoSize -AutoFilter -FreezePane @(2, 2)
    #}
    if ($null -ne $amer_patchProbDevices -and $amer_patchProbDevices.Count -ge 0) {
        $epmExcel = $amer_patchProbDevices | Export-Excel -Path $exportFileName -WorksheetName "Patch" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $amer_mgntByProbDevices -and $amer_mgntByProbDevices.Count -ge 0) {
        $epmExcel = $amer_mgntByProbDevices | Export-Excel -Path $exportFileName -WorksheetName "MgmtBy" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $amer_probDevices -and $amer_probDevices.Count -ge 0) {
        $epmExcel = $amer_probDevices | Export-Excel -Path $exportFileName -WorksheetName ($getDate.ToString("yyyyMMdd").ToString()) -AutoSize -AutoFilter -FreezePane @(2, 2) -PassThru
    }
    $epmExcel.Save();
    $epmExcel.Dispose();
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | select {($_.Address).Replace('J','')}
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | % {$epmExcel.Workbook.Worksheets["SEP"].Row(($_.Address).Replace('J','')).Hidden = $true}

    Write-Host "Generate Email report(AMER) ..."
    Write-ToLogFile -LogContent ("Generate Email report(AMER) ...")

    $amer_Report = "AMER - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $amer_probDevices ), `
    (Get-ObjNumber $amer_adProbDevices ), `
    (Get-ObjNumber $amer_cbProbDevices ), `
    (Get-ObjNumber $amer_cbcProbDevices ), `
    (Get-ObjNumber $amer_patchProbDevices );

    $epmReport = "<H2>EM AMER Report on $TodayDate</H2>$amer_Report</br></br>"

    $sepProbDeviceshtmlReport = ""
    $patchProbDeviceshtmlReport = ""
    $cbProbDeviceshtmlReport = ""
    $fullHtmlReport = ""
    $simpleHtmlReport = ""

    if ($null -ne $amer_StateAntivirusDevices -and $amer_StateAntivirusDevices.Count -ge 0) {
        $sepProbDeviceshtmlReport += '<H3>SEP Problem Devices (' + (Get-ObjNumber $amer_StateAntivirusDevices ) + ') - Need L2 to verify</H3>';
        $sepProbDeviceshtmlReport += $amer_StateAntivirusDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateAntivirus, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $amer_StatePatchDevices -and $amer_StatePatchDevices.Count -ge 0) {
        $patchProbDeviceshtmlReport += '<H3>Patch Problem Devices (' + (Get-ObjNumber $amer_StatePatchDevices ) + ') - Need L2 to verify</H3>';
        $patchProbDeviceshtmlReport += $amer_StatePatchDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StatePatching, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $amer_StateCarbonBlack -and $amer_StateCarbonBlack.Count -ge 0) {
        $cbProbDeviceshtmlReport += '<H3>CB Problem Devices (' + (Get-ObjNumber $amer_StateCarbonBlack ) + ') - Need L2 to verify</H3>';
        $cbProbDeviceshtmlReport += $amer_StateCarbonBlack | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateCarbonBlack, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if (($sepProbDeviceshtmlReport -and "".Equals($sepProbDeviceshtmlReport) -ne $true) -or `
        ($patchProbDeviceshtmlReport -and "".Equals($patchProbDeviceshtmlReport) -ne $true) -or `
        ($cbProbDeviceshtmlReport -and "".Equals($cbProbDeviceshtmlReport) -ne $true)) {
        $fullHtmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport $sepProbDeviceshtmlReport $patchProbDeviceshtmlReport $cbProbDeviceshtmlReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }
    else {
        $simpleHtmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }

    #$htmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";

    Write-Host "`r`n`r`n"
    Write-Host "$amer_Report"
    Write-Host "`r`n`r`n"

    Write-ToLogFile -LogContent ("EM Report - AMER")
    Write-ToLogFile -LogContent ("$amer_Report")

    $report_Object | Add-Member -NotePropertyName 'id' -NotePropertyValue "AMER"
    $report_Object | Add-Member -NotePropertyName 'title' -NotePropertyValue "EM Report - AMER"
    $report_Object | Add-Member -NotePropertyName 'subject' -NotePropertyValue $mailSubject
    $report_Object | Add-Member -NotePropertyName 'sender' -NotePropertyValue $mailSender
    $report_Object | Add-Member -NotePropertyName 'mailTo' -NotePropertyValue $mailTo
    $report_Object | Add-Member -NotePropertyName 'mailCC' -NotePropertyValue $mailCC
    $report_Object | Add-Member -NotePropertyName 'probDevices' -NotePropertyValue $amer_probDevices
    $report_Object | Add-Member -NotePropertyName 'cbProbDevices' -NotePropertyValue $amer_cbProbDevices
    $report_Object | Add-Member -NotePropertyName 'cbcProbDevices' -NotePropertyValue $amer_cbcProbDevices
    $report_Object | Add-Member -NotePropertyName 'sepProbDevices' -NotePropertyValue $amer_sepProbDevices
    $report_Object | Add-Member -NotePropertyName 'patchProbDevices' -NotePropertyValue $amer_patchProbDevices
    $report_Object | Add-Member -NotePropertyName 'mgntByProbDevices' -NotePropertyValue $amer_mgntByProbDevices
    $report_Object | Add-Member -NotePropertyName 'adProbDevices' -NotePropertyValue $amer_adProbDevices
    $report_Object | Add-Member -NotePropertyName 'reportSummary' -NotePropertyValue $amer_Report
    $report_Object | Add-Member -NotePropertyName 'reportTotal' -NotePropertyValue $amer_Report
    $report_Object | Add-Member -NotePropertyName 'fullHtmlReport' -NotePropertyValue $fullHtmlReport
    $report_Object | Add-Member -NotePropertyName 'simpleHtmlReport' -NotePropertyValue $simpleHtmlReport
    $report_Object | Add-Member -NotePropertyName 'exportReportFile' -NotePropertyValue $exportFileName

    return $report_Object

    #Write-Host ("Sending Email report(AMER) ...")
    #Write-ToLogFile -LogContent ("Sending Email report(AMER) ...")

    #SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $testMailTo -ccTo $testMailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML exportFileName$true -attachment $

}

function GenerateReportEMEA {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("SourceXlsxData")] 
        [object]$sourceData, 
        [Alias("ExportXlssxFilename")] 
        [string]$exportFileName, 

        [Parameter(Mandatory = $false)]
        [bool]$isTesting = $false
    )

    if ($isTesting) {
        [string[]]$mailTo = @("Jimmy.Sha@microchip.com")
        [string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$mailTo = @("Nitin.Zhao@microchip.com", "Lenard.Tai@microchip.com", "Kayle.Liu@microchip.com", "Jimmy.Sha@microchip.com", "Roxie.Lee@microchip.com")
        #[string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$testMailTo = @("Jimmy.Sha@microchip.com")
        #[string[]]$testailCC = @("jsha@sst.com")
    }
    else {
        [string[]]$mailTo = @("Peter.Dickenson@microchip.com");
        [string[]]$mailCC = @("Jason.So@microchip.com")
    }

    $mailSender = "EM EMEA Report<emreport_emea@microchip.com>";
    $mailSubject = "EM EMEA Report - " + $TodayDate;

    Write-Host ("Data(EMEA) filtering ...")
    Write-ToLogFile -LogContent ("Data(EMEA) filtering ...")

    $report_Object = New-Object -TypeName PSObject
    $emea_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("AT") -or `
            $_.Location.ToUpper().StartsWith("BE") -or `
            $_.Location.ToUpper().StartsWith("CH") -or `
            $_.Location.ToUpper().StartsWith("DE") -or `
            $_.Location.ToUpper().StartsWith("DK") -or `
            $_.Location.ToUpper().StartsWith("ES") -or `
            $_.Location.ToUpper().StartsWith("FI") -or `
            $_.Location.ToUpper().StartsWith("FR") -or `
            $_.Location.ToUpper().StartsWith("GB") -or `
            $_.Location.ToUpper().StartsWith("IE") -or `
            $_.Location.ToUpper().StartsWith("IT") -or `
            $_.Location.ToUpper().StartsWith("IL") -or `
            $_.Location.ToUpper().StartsWith("NL") -or `
            $_.Location.ToUpper().StartsWith("NO") -or `
            $_.Location.ToUpper().StartsWith("RO") -or `
            $_.Location.ToUpper().StartsWith("SE") -and `
        ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in Europe, exclude Subcon devices

    #######################################################
    ### Problem devices in Europe
    #######################################################
    #$emea_probDevices = $emea_Devices | where {$_.ProblemDevice -eq $true}
    [array]$emea_probDevices = $emea_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$emea_cbProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$emea_cbcProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    [array]$emea_sepProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$emea_patchProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    [array]$emea_mgntByProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$emea_adProbDevices = $emea_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #$emea_StateAntivirusDevices = $emea_sepProbDevices | Where-Object { "non-corp av".Equals($_.StateAntivirus.Trim().ToLower()) -or "not applicable".Equals($_.StateAntivirus.Trim().ToLower()) }
    #$emea_StatePatchDevices = $emea_patchProbDevices | Where-Object { "agent not found".Equals($_.StatePatching.Trim().ToLower()) -or "needs os version".Equals($_.StatePatching.Trim().ToLower()) -or "not applicable".Equals($_.StatePatching.Trim().ToLower()) }
    #$emea_StateCarbonBlack = $emea_cbProbDevices | Where-Object { "not connecting".Equals($_.StateCarbonBlack.Trim().ToLower()) }

    Write-Host ("Generate Excel file(EMEA): {0}" -f $exportFileName)
    Write-ToLogFile -LogContent ("Generate Excel file(EMEA): {0}" -f $exportFileName)

    ##############################################################
    ### Create EPM report file with all Europe problem Devices
    ##############################################################
    if ($null -ne $emea_adProbDevices -and $emea_adProbDevices.Count -ge 0) {
        $epmExcel = $emea_adProbDevices | Export-Excel -Path $exportFileName -WorksheetName "AD" -ClearSheet -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $emea_cbProbDevices -and $emea_cbProbDevices.Count -ge 0) {
        $epmExcel = $emea_cbProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CB" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $emea_cbcProbDevices -and $emea_cbcProbDevices.Count -ge 0) {
        $epmExcel = $emea_cbcProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CBC" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    #if ($null -ne $emea_sepProbDevices -and $emea_sepProbDevices.Count -ge 0) {
    #    $epmExcel = $emea_sepProbDevices | Export-Excel -Path $exportFileName -WorksheetName "SEP" -AutoSize -AutoFilter -FreezePane @(2, 2)
    #}
    if ($null -ne $emea_patchProbDevices -and $emea_patchProbDevices.Count -ge 0) {
        $epmExcel = $emea_patchProbDevices | Export-Excel -Path $exportFileName -WorksheetName "Patch" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $emea_mgntByProbDevices -and $emea_mgntByProbDevices.Count -ge 0) {
        $epmExcel = $emea_mgntByProbDevices | Export-Excel -Path $exportFileName -WorksheetName "MgmtBy" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $emea_probDevices -and $emea_probDevices.Count -ge 0) {
        $epmExcel = $emea_probDevices | Export-Excel -Path $exportFileName -WorksheetName ($getDate.ToString("yyyyMMdd").ToString()) -AutoSize -AutoFilter -FreezePane @(2, 2) -PassThru
    }
    $epmExcel.Save();
    $epmExcel.Dispose();
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | select {($_.Address).Replace('J','')}
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | % {$epmExcel.Workbook.Worksheets["SEP"].Row(($_.Address).Replace('J','')).Hidden = $true}

    Write-Host "Generate Email report(EMEA) ..."
    Write-ToLogFile -LogContent ("Generate Email report(EMEA) ...")

    $emea_Report = "EMEA - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $emea_probDevices ), `
    (Get-ObjNumber $emea_adProbDevices ), `
    (Get-ObjNumber $emea_cbProbDevices ), `
    (Get-ObjNumber $emea_cbcProbDevices ), `
    (Get-ObjNumber $emea_patchProbDevices );

    $epmReport = "<H2>EM EMEA Report on $TodayDate</H2>$emea_Report</br></br>"

    $sepProbDeviceshtmlReport = ""
    $patchProbDeviceshtmlReport = ""
    $cbProbDeviceshtmlReport = ""
    $fullHtmlReport = ""
    $simpleHtmlReport = ""

    if ($null -ne $emea_StateAntivirusDevices -and $emea_StateAntivirusDevices.Count -ge 0) {
        $sepProbDeviceshtmlReport += '<H3>SEP Problem Devices (' + (Get-ObjNumber $emea_StateAntivirusDevices ) + ') - Need L2 to verify</H3>';
        $sepProbDeviceshtmlReport += $emea_StateAntivirusDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateAntivirus, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $emea_StatePatchDevices -and $emea_StatePatchDevices.Count -ge 0) {
        $patchProbDeviceshtmlReport += '<H3>Patch Problem Devices (' + (Get-ObjNumber $emea_StatePatchDevices ) + ') - Need L2 to verify</H3>';
        $patchProbDeviceshtmlReport += $emea_StatePatchDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StatePatching, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $emea_StateCarbonBlack -and $emea_StateCarbonBlack.Count -ge 0) {
        $cbProbDeviceshtmlReport += '<H3>CB Problem Devices (' + (Get-ObjNumber $emea_StateCarbonBlack ) + ') - Need L2 to verify</H3>';
        $cbProbDeviceshtmlReport += $emea_StateCarbonBlack | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateCarbonBlack, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if (($sepProbDeviceshtmlReport -and "".Equals($sepProbDeviceshtmlReport) -ne $true) -or `
        ($patchProbDeviceshtmlReport -and "".Equals($patchProbDeviceshtmlReport) -ne $true) -or `
        ($cbProbDeviceshtmlReport -and "".Equals($cbProbDeviceshtmlReport) -ne $true)) {
        $fullHtmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport $sepProbDeviceshtmlReport $patchProbDeviceshtmlReport $cbProbDeviceshtmlReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }
    else {
        $simpleHtmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }

    #$htmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";

    Write-Host "`r`n`r`n"
    Write-Host "$emea_Report"
    Write-Host "`r`n`r`n"

    Write-ToLogFile -LogContent ("EM Report - EMEA")
    Write-ToLogFile -LogContent ("$emea_Report")

    $report_Object | Add-Member -NotePropertyName 'id' -NotePropertyValue "EMEA"
    $report_Object | Add-Member -NotePropertyName 'title' -NotePropertyValue "EPM Report - EMEA"
    $report_Object | Add-Member -NotePropertyName 'subject' -NotePropertyValue $mailSubject
    $report_Object | Add-Member -NotePropertyName 'sender' -NotePropertyValue $mailSender
    $report_Object | Add-Member -NotePropertyName 'mailTo' -NotePropertyValue $mailTo
    $report_Object | Add-Member -NotePropertyName 'mailCC' -NotePropertyValue $mailCC
    $report_Object | Add-Member -NotePropertyName 'probDevices' -NotePropertyValue $emea_probDevices
    $report_Object | Add-Member -NotePropertyName 'cbProbDevices' -NotePropertyValue $emea_cbProbDevices
    $report_Object | Add-Member -NotePropertyName 'cbcProbDevices' -NotePropertyValue $emea_cbcProbDevices
    $report_Object | Add-Member -NotePropertyName 'sepProbDevices' -NotePropertyValue $emea_sepProbDevices
    $report_Object | Add-Member -NotePropertyName 'patchProbDevices' -NotePropertyValue $emea_patchProbDevices
    $report_Object | Add-Member -NotePropertyName 'mgntByProbDevices' -NotePropertyValue $emea_mgntByProbDevices
    $report_Object | Add-Member -NotePropertyName 'adProbDevices' -NotePropertyValue $emea_adProbDevices
    $report_Object | Add-Member -NotePropertyName 'reportSummary' -NotePropertyValue $emea_Report
    $report_Object | Add-Member -NotePropertyName 'reportTotal' -NotePropertyValue $emea_Report
    $report_Object | Add-Member -NotePropertyName 'fullHtmlReport' -NotePropertyValue $fullHtmlReport
    $report_Object | Add-Member -NotePropertyName 'simpleHtmlReport' -NotePropertyValue $simpleHtmlReport
    $report_Object | Add-Member -NotePropertyName 'exportReportFile' -NotePropertyValue $exportFileName

    return $report_Object

    #Write-Host ("Sending Email report(EMEA) ...")
    #Write-ToLogFile -LogContent ("Sending Email report(EMEA) ...")

    #SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $testMailTo -ccTo $testMailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName

}

function GenerateReportAPAC {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("SourceXlsxData")] 
        [object]$sourceData, 
        [Alias("ExportXlssxFilename")] 
        [string]$exportFileName, 

        [Parameter(Mandatory = $false)]
        [bool]$isTesting = $false
    )

    if ($isTesting) {
        [string[]]$mailTo = @("Jimmy.Sha@microchip.com")
        [string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$mailToCN_HK = @("jsha@sst.com")
        #[string[]]$mailCCToCN_HK = @("Jimmy.Sha@microchip.com")
        #[string[]]$mailTo = @("Nitin.Zhao@microchip.com", "Lenard.Tai@microchip.com", "Kayle.Liu@microchip.com", "Jimmy.Sha@microchip.com", "Roxie.Lee@microchip.com")
        #[string[]]$mailCC = @("jsha@sst.com")
        #[string[]]$testMailTo = @("Jimmy.Sha@microchip.com")
        #[string[]]$testailCC = @("jsha@sst.com")
        
    }
    else {
        [string[]]$mailTo = @("Jimmy.Sha@microchip.com", 
            "Ryan.Lin@microchip.com", 
            "Julian.Tseng@microchip.com", 
            "ShairaMai.Marcelo@microchip.com", 
            "Nitin.Zhao@microchip.com", 
            "Lenard.Tai@microchip.com", 
            "Ian.Lai@microchip.com", 
            "Eric.Chen@microchip.com", 
            "dana.hu@microchip.com", 
            "Roxie.Lee@microchip.com", 
            "sebastian.zhu@microchip.com", 
            "Joanne.Chan@microchip.com", 
            "WaiShong.Lee@microchip.com", 
            "KokKien.Ng@microchip.com", 
            "CarlAngelo.Nievarez@microchip.com",
            "kevin.yeap@microchip.com", 
            "christopher.huang@microchip.com");
        [string[]]$mailCC = @("Navakarti.Satiyah@microchip.com")

        #[string[]]$mailToCN_HK = @("Nitin.Zhao@microchip.com", "Lenard.Tai@microchip.com", "Ian.Lai@microchip.com", "Eric.Chen@microchip.com", "Roxie.Lee@microchip.com")
        #[string[]]$mailCCToCN_HK = @("Jason.So@microchip.com", "Navakarti.Satiyah@microchip.com", "Jimmy.Sha@microchip.com")
    }

    $mailSender = "EM APAC Report<emreport_apac@microchip.com>";
    $mailSubject = "EM APAC Report - " + $TodayDate;

    Write-Host ("Data(APAC) filtering ...")
    Write-ToLogFile -LogContent ("Data(APAC) filtering ...")

    $report_Object = New-Object -TypeName PSObject
    $cn_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("CN") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in China, exclude Subcon devices
    $hk_Devices = $sourceData | Where-Object { ($_.Location.ToUpper().StartsWith("HK") -or $_.Location.ToLower().StartsWith("macau")) -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in Hong Kong and Macau
    $jp_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("JP") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Japan, exclude Subcon devices
    $kr_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("KR") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Korea, exclude Subcon devices
    $my_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("MY") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Malaysia, exclude Subcon devices
    $ptc_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("PH") -and $_.Location.ToUpper().Contains("PTC") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Philippines, exclude Subcon devices
    $sg_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("SG") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Singapore, exclude Subcon devices
    $tw_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("TW") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Taiwan, exclude Subcon devices
    $vn_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("VN") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location   ### Devices in Vietnam, exclude Subcon devices
    $au_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("AU") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in Australia, exclude Subcon devices
    $nz_Devices = $sourceData | Where-Object { $_.Location.ToUpper().StartsWith("NZ") -and ($_.Location.ToLower().Contains("subcon") -eq $false) -and ($_.Location.ToLower().Contains("dimerco") -eq $false) } | Sort-Object Location  ### Devices in Tekron, exclude Subcon devices
    $asub_Devices = $sourceData | Where-Object { $_.Location.ToLower().Contains("subcon") -or $_.Location.ToLower().Contains("dimerco") } ### Devices in Subcon

    #######################################################
    ### Problem devices in China
    #######################################################
    [array]$cn_probDevices = $cn_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$cn_cbProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$cn_cbcProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$cn_sepProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$cn_patchProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$cn_mgntByProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$cn_adProbDevices = $cn_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Hong kong
    #######################################################
    [array]$hk_probDevices = $hk_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$hk_cbProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$hk_cbcProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$hk_sepProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$hk_patchProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$hk_mgntByProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$hk_adProbDevices = $hk_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Japan
    #######################################################
    [array]$jp_probDevices = $jp_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$jp_cbProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$jp_cbcProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$jp_sepProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$jp_patchProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$jp_mgntByProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$jp_adProbDevices = $jp_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Korea
    #######################################################
    [array]$kr_probDevices = $kr_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$kr_cbProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$kr_cbcProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$kr_sepProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$kr_patchProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$kr_mgntByProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$kr_adProbDevices = $kr_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Malaysia
    #######################################################
    [array]$my_probDevices = $my_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$my_cbProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$my_cbcProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$my_sepProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$my_patchProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$my_mgntByProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$my_adProbDevices = $my_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Philippines
    #######################################################
    [array]$ptc_probDevices = $ptc_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$ptc_cbProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$ptc_cbcProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$ptc_sepProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$ptc_patchProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$ptc_mgntByProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$ptc_adProbDevices = $ptc_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Singapore
    #######################################################
    [array]$sg_probDevices = $sg_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$sg_cbProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$sg_cbcProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$sg_sepProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$sg_patchProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$sg_mgntByProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$sg_adProbDevices = $sg_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Taiwan
    #######################################################
    [array]$tw_probDevices = $tw_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$tw_cbProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$tw_cbcProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$tw_sepProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$tw_patchProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$tw_mgntByProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$tw_adProbDevices = $tw_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Vietnam
    #######################################################
    [array]$vn_probDevices = $vn_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$vn_cbProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$vn_cbcProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #array]$vn_sepProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$vn_patchProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$vn_mgntByProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$vn_adProbDevices = $vn_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Australia
    #######################################################
    [array]$au_probDevices = $au_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    #[array]$au_cbProbDevices = $au_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    #[array]$au_sepProbDevices = $au_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    #[array]$au_patchProbDevices = $au_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$au_mgntByProbDevices = $au_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    #[array]$au_adProbDevices = $au_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in Tekron
    #######################################################
    [array]$nz_probDevices = $nz_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    #[array]$nz_cbProbDevices = $nz_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    #[array]$nz_sepProbDevices = $nz_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    #[array]$nz_patchProbDevices = $nz_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$nz_mgntByProbDevices = $nz_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    #[array]$nz_adProbDevices = $nz_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    #######################################################
    ### Problem devices in APAC Subcon
    #######################################################
    [array]$asub_probDevices = $asub_Devices | Where-Object { $_.ProblemDevice -eq $true -and ("windows pc".Equals($_.DeviceType.Trim().ToLower()) -or "apple mac".Equals($_.DeviceType.Trim().ToLower())) } | Where-Object { (IsNewSystem -objItem $_) -eq $false } | Where-Object { $_ -ne $null }
    [array]$asub_cbProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$asub_cbcProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    #[array]$asub_sepProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$asub_patchProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    #[array]$asub_mgntByProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$asub_adProbDevices = $asub_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    
    #######################################################
    ### Total devices in CN_HK
    #######################################################
    #$cn_hk_Devices = ($cn_Devices + $hk_Devices)

    #######################################################
    ### Total problem devices in CN_HK
    #######################################################
    $cn_hk_probDevices = ($cn_probDevices + $hk_probDevices) | Where-Object { $_ -ne $null }
    Write-ToLogFile -LogContent ("cn_probDevices {0}" -f (Get-ObjNumber $cn_probDevices))
    Write-ToLogFile -LogContent ("cn_probDevices: {0}" -f (($cn_probDevices | Select-Object -ExpandProperty DeviceName) -join ", "))
    Write-ToLogFile -LogContent ("hk_probDevices {0}" -f (Get-ObjNumber $hk_probDevices))
    Write-ToLogFile -LogContent ("hk_probDevices: {0}" -f (($hk_probDevices | Select-Object -ExpandProperty DeviceName) -join ", "))
    Write-Host (Get-ObjNumber $cn_hk_probDevices)
    Write-ToLogFile -LogContent ("cn_hk_probDevices {0}" -f (Get-ObjNumber $cn_hk_probDevices))
    Write-ToLogFile -LogContent ("cn_hk_probDevices: {0}" -f (($cn_hk_probDevices | Select-Object -ExpandProperty DeviceName) -join ", "))

    #######################################################
    ### Total devices in KR_JP_TW
    #######################################################
    #$kr_jp_tw_Devices = ($kr_Devices + $jp_Devices + $tw_Devices)

    #######################################################
    ### Total problem devices in KR_JP_TW
    #######################################################
    $kr_jp_tw_probDevices = ($kr_probDevices + $jp_probDevices + $tw_probDevices) | Where-Object { $_ -ne $null }
    Write-Host (Get-ObjNumber $kr_jp_tw_probDevices)
    Write-ToLogFile -LogContent ("kr_jp_tw_probDevices {0}" -f $kr_jp_tw_probDevices.Length)
    Write-ToLogFile -LogContent ("{0}" -f (($kr_jp_tw_probDevices | Select-Object -ExpandProperty DeviceName) -join ", "))

    #######################################################
    ### Total devices in MY_SG_VN
    #######################################################
    #$my_sg_vn_Devices = ($my_Devices + $sg_Devices + $vn_Devices)

    #######################################################
    ### Total problem devices in MY_SG_VN
    #######################################################
    $my_sg_vn_probDevices = ($my_probDevices + $sg_probDevices + $vn_probDevices) | Where-Object { $_ -ne $null }
    Write-Host (Get-ObjNumber $my_sg_vn_probDevices)
    Write-ToLogFile -LogContent ("my_sg_vn_probDevices {0}" -f $my_sg_vn_probDevices.Length)
    # Write-ToLogFile -LogContent ("{0}" -f ($my_sg_vn_probDevices | Select-Object DeviceName))
    Write-ToLogFile -LogContent ("{0}" -f (($my_sg_vn_probDevices | Select-Object -ExpandProperty DeviceName) -join ", "))

    #######################################################
    ### Problem devices in whole APAC
    #######################################################
    $apac_probDevices = ($cn_hk_probDevices + $kr_jp_tw_probDevices + $my_sg_vn_probDevices + $ptc_probDevices + $au_probDevices + $nz_probDevices + $asub_probDevices) | Where-Object { $_ -ne $null } ;
    [array]$apac_cbProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCarbonBlack -eq $true }
    [array]$apac_cbcProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemCBCloud -eq $true }
    [array]$apac_sepProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemAntivirus -eq $true }
    [array]$apac_patchProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { $_.ProblemPatching -eq $true }
    [array]$apac_mgntByProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { $null -eq $_.ManagedBy -or "".Equals($_.ManagedBy.Trim()) }
    [array]$apac_adProbDevices = $apac_probDevices | Where-Object { $_ -ne $null } | Where-Object { "--not bound--".Equals($_.ADDomain.Trim().ToLower()) }

    if ($null -ne $apac_sepProbDevices -and $apac_sepProbDevices.Count -ge 0) {
        [array]$apac_StateAntivirusDevices = $apac_sepProbDevices | Where-Object { "non-corp av".Equals($_.StateAntivirus.Trim().ToLower()) -or "not applicable".Equals($_.StateAntivirus.Trim().ToLower()) }
    }
    if ($null -ne $apac_patchProbDevices -and $apac_patchProbDevices.Count -ge 0) {
        [array]$apac_StatePatchDevices = $apac_patchProbDevices | Where-Object { "agent not found".Equals($_.StatePatching.Trim().ToLower()) -or "needs os version".Equals($_.StatePatching.Trim().ToLower()) -or "not applicable".Equals($_.StatePatching.Trim().ToLower()) }
    }
    if ($null -ne $apac_cbProbDevices -and $apac_cbProbDevices.Count -ge 0) {
        [array]$apac_StateCarbonBlack = $apac_cbProbDevices | Where-Object { "not connecting".Equals($_.StateCarbonBlack.Trim().ToLower()) }
    }
    <#
    Write-Host ("Generate Excel file(CN_HK): {0}" -f $exportFileName_2)
    Write-ToLogFile -LogContent ("Generate Excel file(CN_HK): {0}" -f $exportFileName_2)
    ##############################################################
    ### Create EPM report file with all CN_HK problem Devices
    ##############################################################
    $epmExcel = ($cn_adProbDevices + $hk_adProbDevices) | Export-Excel -Path $exportFileName_2 -WorksheetName "AD" -ClearSheet -AutoSize -AutoFilter -FreezePane @(2,2)
    $epmExcel = ($cn_cbProbDevices + $hk_cbProbDevices) | Export-Excel -Path $exportFileName_2 -WorksheetName "CB" -AutoSize -AutoFilter -FreezePane @(2,2)
    $epmExcel = ($cn_sepProbDevices + $hk_sepProbDevices) | Export-Excel -Path $exportFileName_2 -WorksheetName "SEP" -AutoSize -AutoFilter -FreezePane @(2,2)
    $epmExcel = ($cn_patchProbDevices + $hk_patchProbDevices) | Export-Excel -Path $exportFileName_2 -WorksheetName "Patch" -AutoSize -AutoFilter -FreezePane @(2,2)
    $epmExcel = ($cn_mgntByProbDevices + $hk_mgntByProbDevices) | Export-Excel -Path $exportFileName_2 -WorksheetName "MgmtBy" -AutoSize -AutoFilter -FreezePane @(2,2)
    $epmExcel = $cn_hk_probDevices | Export-Excel -Path $exportFileName_2 -WorksheetName ($TodayDate.ToString()) -AutoSize -AutoFilter -FreezePane @(2,2) -PassThru
    $epmExcel.Save();
    $epmExcel.Dispose();
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | select {($_.Address).Replace('J','')}
    #$epmExcel.Workbook.Worksheets["SEP"].Cells.Item(1, 10, $epmExcel.Workbook.Worksheets["SEP"].AutoFilterAddress.Rows, 10) | where{ $_.Value.Trim().Equals("Restart Required")} | % {$epmExcel.Workbook.Worksheets["SEP"].Row(($_.Address).Replace('J','')).Hidden = $true}
    #>

    Write-Host ("Generate Excel file(APAC): {0}" -f $exportFileName)
    Write-ToLogFile -LogContent ("Generate Excel file(APAC): {0}" -f $exportFileName)
    ##############################################################
    ### Create EPM report file with all APAC problem Devices
    ##############################################################
    if ($null -ne $apac_adProbDevices -and $apac_adProbDevices.Count -ge 0) {
        $epmExcel = $apac_adProbDevices | Export-Excel -Path $exportFileName -WorksheetName "AD" -ClearSheet -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $apac_cbProbDevices -and $apac_cbProbDevices.Count -ge 0) {
        $epmExcel = $apac_cbProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CB" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $apac_cbcProbDevices -and $apac_cbcProbDevices.Count -ge 0) {
        $epmExcel = $apac_cbcProbDevices | Export-Excel -Path $exportFileName -WorksheetName "CBC" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    #if ($null -ne $apac_sepProbDevices -and $apac_sepProbDevices.Count -ge 0) {
    #    $epmExcel = $apac_sepProbDevices | Export-Excel -Path $exportFileName -WorksheetName "SEP" -AutoSize -AutoFilter -FreezePane @(2, 2)
    #}
    if ($null -ne $apac_patchProbDevices -and $apac_patchProbDevices.Count -ge 0) {
        $epmExcel = $apac_patchProbDevices | Export-Excel -Path $exportFileName -WorksheetName "Patch" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $apac_mgntByProbDevices -and $apac_mgntByProbDevices.Count -ge 0) {
        $epmExcel = $apac_mgntByProbDevices | Export-Excel -Path $exportFileName -WorksheetName "MgmtBy" -AutoSize -AutoFilter -FreezePane @(2, 2)
    }
    if ($null -ne $apac_probDevices -and $apac_probDevices.Count -ge 0) {
        $epmExcel = $apac_probDevices | Export-Excel -Path $exportFileName -WorksheetName ($TodayDate.ToString()) -AutoSize -AutoFilter -FreezePane @(2, 2) -PassThru
    }
    $epmExcel.Save();
    $epmExcel.Dispose();

    Write-Host "Generate Email report(APAC) ..."
    Write-ToLogFile -LogContent ("Generate Email report(APAC) ...")

    $cn_hk_Report = "Hongkong & Mainland China - Total: {0} ({1}/{2}), AD: {3}/{4}, CB: {5}/{6}, CBC: {7}/{8}, Patch: {9}/{10}" -f `
    (Get-ObjNumber $cn_hk_probDevices), `
    (Get-ObjNumber $hk_probDevices), (Get-ObjNumber $cn_probDevices), `
    (Get-ObjNumber $hk_adProbDevices), (Get-ObjNumber $cn_adProbDevices), `
    (Get-ObjNumber $hk_cbProbDevices), (Get-ObjNumber $cn_cbProbDevices), `
    (Get-ObjNumber $hk_cbcProbDevices), (Get-ObjNumber $cn_cbcProbDevices), `
    (Get-ObjNumber $hk_patchProbDevices), (Get-ObjNumber $cn_patchProbDevices);

    $kr_jp_tw_Report = "Korea, Japan & Taiwan - Total: {0} ({1}/{2}/{3}), AD: {4}/{5}/{6}, CB: {7}/{8}/{9}, CBC: {10}/{11}/{12}, Patch: {13}/{14}/{15}" -f `
    (Get-ObjNumber $kr_jp_tw_probDevices), `
    (Get-ObjNumber $kr_probDevices), (Get-ObjNumber $jp_probDevices), (Get-ObjNumber $tw_probDevices), `
    (Get-ObjNumber $kr_adProbDevices), (Get-ObjNumber $jp_adProbDevices), (Get-ObjNumber $tw_adProbDevices), `
    (Get-ObjNumber $kr_cbProbDevices), (Get-ObjNumber $jp_cbProbDevices), (Get-ObjNumber $tw_cbProbDevices), `
    (Get-ObjNumber $kr_cbcProbDevices), (Get-ObjNumber $jp_cbcProbDevices), (Get-ObjNumber $tw_cbcProbDevices), `
    (Get-ObjNumber $kr_patchProbDevices), (Get-ObjNumber $jp_patchProbDevices), (Get-ObjNumber $tw_patchProbDevices);

    $my_sg_vn_Report = "Malaysia, Singapore & Vietnam - Total: {0} ({1}/{2}/{3}), AD: {4}/{5}/{6}, CB: {7}/{8}/{9}, CBC: {10}/{11}/{12}, Patch: {13}/{14}/{15}" -f `
    (Get-ObjNumber $my_sg_vn_probDevices), `
    (Get-ObjNumber $my_probDevices), (Get-ObjNumber $sg_probDevices), (Get-ObjNumber $vn_probDevices), `
    (Get-ObjNumber $my_adProbDevices), (Get-ObjNumber $sg_adProbDevices), (Get-ObjNumber $vn_adProbDevices), `
    (Get-ObjNumber $my_cbProbDevices), (Get-ObjNumber $sg_cbProbDevices), (Get-ObjNumber $vn_cbProbDevices), `
    (Get-ObjNumber $my_cbcProbDevices), (Get-ObjNumber $sg_cbcProbDevices), (Get-ObjNumber $vn_cbcProbDevices), `
    (Get-ObjNumber $my_patchProbDevices), (Get-ObjNumber $sg_patchProbDevices), (Get-ObjNumber $vn_patchProbDevices);

    $ptc_Report = "Philippines PTC - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $ptc_probDevices ), `
    (Get-ObjNumber $ptc_adProbDevices ), `
    (Get-ObjNumber $ptc_cbProbDevices ), `
    (Get-ObjNumber $ptc_cbcProbDevices ), `
    (Get-ObjNumber $ptc_patchProbDevices ); 

    $asub_Report = "APAC Subcon - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $asub_probDevices  ), `
    (Get-ObjNumber $asub_adProbDevices ), `
    (Get-ObjNumber $asub_cbProbDevices ), `
    (Get-ObjNumber $asub_cbcProbDevices ), `
    (Get-ObjNumber $asub_patchProbDevices );

    $apac_Report = "All APAC - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $apac_probDevices  ), `
    (Get-ObjNumber $apac_adProbDevices ), `
    (Get-ObjNumber $apac_cbProbDevices ), `
    (Get-ObjNumber $apac_cbcProbDevices ), `
    (Get-ObjNumber $apac_patchProbDevices );

    $apac_Total_Report = "APAC - Total: {0}, AD: {1}, CB: {2}, CBC: {3}, Patching: {4}" -f `
    (Get-ObjNumber $apac_probDevices  ), `
    (Get-ObjNumber $apac_adProbDevices ), `
    (Get-ObjNumber $apac_cbProbDevices ), `
    (Get-ObjNumber $apac_cbcProbDevices ), `
    (Get-ObjNumber $apac_patchProbDevices );
    
    <#
    $apac_Total_Report = "APAC - Total: {0}, CB: {1}, SEP: {2}, Patching: {3}, AD: {4}, [CBC: {5}]" -f `
    (Get-ObjNumber $apac_probDevices  ), `
    (Get-ObjNumber $apac_cbProbDevices ), `
    (Get-ObjNumber $apac_sepProbDevices ), `
    (Get-ObjNumber $apac_patchProbDevices ), `
    (Get-ObjNumber $apac_adProbDevices ), `
    (Get-ObjNumber $apac_cbcProbDevices );
    #>

    $epmReport = "<H2>EM APAC Report on $TodayDate</H2>$cn_hk_Report</br>$kr_jp_tw_Report</br>$my_sg_vn_Report</br>$ptc_Report</br>$asub_Report</br></br>$apac_Report</br></br>"

    $sepProbDeviceshtmlReport = ""
    $patchProbDeviceshtmlReport = ""
    $cbProbDeviceshtmlReport = ""
    #$htmlReport = ""
    $fullHtmlReport = ""
    $simpleHtmlReport = ""

    if ($null -ne $apac_StateAntivirusDevices -and $apac_StateAntivirusDevices.Count -ge 0) {
        $sepProbDeviceshtmlReport += '<H3>SEP Problem Devices (' + (Get-ObjNumber $apac_StateAntivirusDevices ) + ') - Need L2 to verify</H3>';
        $sepProbDeviceshtmlReport += $apac_StateAntivirusDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateAntivirus, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $apac_StatePatchDevices -and $apac_StatePatchDevices.Count -ge 0) {
        $patchProbDeviceshtmlReport += '<H3>Patch Problem Devices (' + (Get-ObjNumber $apac_StatePatchDevices ) + ') - Need L2 to verify</H3>';
        $patchProbDeviceshtmlReport += $apac_StatePatchDevices | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StatePatching, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if ($null -ne $apac_StateCarbonBlack -and $apac_StateCarbonBlack.Count -ge 0) {
        $cbProbDeviceshtmlReport += '<H3>CB Problem Devices (' + (Get-ObjNumber $apac_StateCarbonBlack ) + ') - Need L2 to verify</H3>';
        $cbProbDeviceshtmlReport += $apac_StateCarbonBlack | Select-Object DeviceName, HostName, ADDomain, Location, ManagedBy, StateCarbonBlack, LastSeen | ConvertTo-Html -Fragment | Out-string;
    }

    if (($sepProbDeviceshtmlReport -and "".Equals($sepProbDeviceshtmlReport) -ne $true) -or `
        ($patchProbDeviceshtmlReport -and "".Equals($patchProbDeviceshtmlReport) -ne $true) -or `
        ($cbProbDeviceshtmlReport -and "".Equals($cbProbDeviceshtmlReport) -ne $true)) {
        $fullHtmlReport = (ConvertTo-Html -Head $ReportHead -Body "$epmReport $sepProbDeviceshtmlReport $patchProbDeviceshtmlReport $cbProbDeviceshtmlReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }
    else {
        $simpleHtmlReport = (ConvertTo-Html -Head $ReportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";
    }

    Write-Host "`r`n`r`n"
    Write-Host "$cn_hk_Report `r`n$kr_jp_tw_Report`r`n$my_sg_vn_Report`r`n$ptc_Report`r`n$asub_Report`r`n$apac_Report"
    Write-Host "`r`n`r`n"

    Write-ToLogFile -LogContent ("EM Report - APAC")
    Write-ToLogFile -LogContent ("$cn_hk_Report `r`n$kr_jp_tw_Report`r`n$my_sg_vn_Report`r`n$ptc_Report`r`n$asub_Report`r`n`r`n$apac_Report")

    $report_Object | Add-Member -NotePropertyName 'id' -NotePropertyValue "APAC"
    $report_Object | Add-Member -NotePropertyName 'title' -NotePropertyValue "EPM Report - APAC"
    $report_Object | Add-Member -NotePropertyName 'subject' -NotePropertyValue $mailSubject
    $report_Object | Add-Member -NotePropertyName 'sender' -NotePropertyValue $mailSender
    $report_Object | Add-Member -NotePropertyName 'mailTo' -NotePropertyValue $mailTo
    $report_Object | Add-Member -NotePropertyName 'mailCC' -NotePropertyValue $mailCC
    $report_Object | Add-Member -NotePropertyName 'probDevices' -NotePropertyValue $apac_probDevices
    $report_Object | Add-Member -NotePropertyName 'cbProbDevices' -NotePropertyValue $apac_cbProbDevices
    $report_Object | Add-Member -NotePropertyName 'cbcProbDevices' -NotePropertyValue $apac_cbcProbDevices
    $report_Object | Add-Member -NotePropertyName 'sepProbDevices' -NotePropertyValue $apac_sepProbDevices
    $report_Object | Add-Member -NotePropertyName 'patchProbDevices' -NotePropertyValue $apac_patchProbDevices
    $report_Object | Add-Member -NotePropertyName 'mgntByProbDevices' -NotePropertyValue $apac_mgntByProbDevices
    $report_Object | Add-Member -NotePropertyName 'adProbDevices' -NotePropertyValue $apac_adProbDevices
    $report_Object | Add-Member -NotePropertyName 'reportSummary' -NotePropertyValue ("$cn_hk_Report `r`n$kr_jp_tw_Report`r`n$my_sg_vn_Report`r`n$ptc_Report`r`n$asub_Report`r`n`r`n$apac_Report")
    $report_Object | Add-Member -NotePropertyName 'reportTotal' -NotePropertyValue $apac_Total_Report
    $report_Object | Add-Member -NotePropertyName 'fullHtmlReport' -NotePropertyValue $fullHtmlReport
    $report_Object | Add-Member -NotePropertyName 'simpleHtmlReport' -NotePropertyValue $simpleHtmlReport
    $report_Object | Add-Member -NotePropertyName 'exportReportFile' -NotePropertyValue $exportFileName

    return $report_Object

    #Write-Host ("Sending Email report(APAC) ...")
    #Write-ToLogFile -LogContent ("Sending Email report(APAC) ...")

    #SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $mailToCN_HK -ccTo $mailCCToCN_HK -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName
    #SendReportByMail -receiver $testMailTo -ccTo $testMailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachment $exportFileName

}

function SendRegionsReport([object]$report_Obj) {
    if ($report_Obj -is [System.Object]) {
        $htmlReport = ""

        if (($null -ne $report_Obj.fullHtmlReport) -and ("".Equals($report_Obj.fullHtmlReport) -ne $true)) {
            $htmlReport = $report_Obj.fullHtmlReport
        }
        else {
            $htmlReport = $report_Obj.simpleHtmlReport
        }

        Write-Host ("Sending Email report({0}) ..." -f $report_Obj.id)
        Write-ToLogFile -LogContent ("Sending Email report({0}) ..." -f $report_Obj.id)
        SendReportByMail -receiver $report_Obj.mailTo -ccTo $report_Obj.mailCC -sender $report_Obj.sender -subject $report_Obj.subject -mailBody $htmlReport -isHTML $true -attachment $report_Obj.exportReportFile
    }
    else {
        Write-Host ("Error when send regional report !!!, detail: {0}" -f $error[0])
        Write-ToLogFile -LogContent ("Error when send regional report !!!, detail: {0}" -f $error[0])
    }
}

function SendGlobalReport {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("ReportObjects")] 
        [Object]$reportObjs, 

        [Parameter(Mandatory = $false)]
        [bool]$isTesting = $false
    )
    if ($reportObjs -is [System.Array]) {
        $epmReport = "<H2>EM Report on $TodayDate</H2>"
        $globalReport = ""
        [string[]]$mailAttachments = @()
        foreach ($itemObj in $reportObjs) {
            $epmReport += ("" + $itemObj.reportTotal + "</br></br>") 
            $globalReport += ("" + $itemObj.reportTotal + "`r`n`r`n")
            if (($null -ne $itemObj.exportReportFile) -and ("".Equals($itemObj.exportReportFile) -ne $true)) {
                $mailAttachments += $itemObj.exportReportFile
            }
        }
        $htmlReport = (ConvertTo-Html -Head $reportHead -Body "$epmReport" | Out-String) -replace "(?sm)<table>\s+</table>";

        Write-Host "`r`n`r`n"
        Write-Host "$globalReport"
        Write-Host "`r`n`r`n"

        Write-ToLogFile -LogContent ("EM Report - Global")
        Write-ToLogFile -LogContent ("$globalReport")

        if ($isTesting) {
            [string[]]$mailTo = @("Jimmy.Sha@microchip.com")
            [string[]]$mailCC = @("jsha@sst.com")
            
        }
        else {
            [string[]]$mailTo = @("Ariel.Crespo@microchip.com", 
                "Corbin.Marginson@microchip.com", 
                "Jimmy.Sha@microchip.com", 
                "Julian.Tseng@microchip.com", 
                "Lam.Tran@microchip.com", 
                "Markus.Bernhart@microchip.com", 
                "Martin.Denning@microchip.com", 
                "Navakarti.Satiyah@microchip.com", 
                "Sameer.Ebadi@microchip.com", 
                "WaiShong.Lee@microchip.com");
            [string[]]$mailCC = @("Jason.So@microchip.com", "Peter.Khoo@microchip.com", "Emmanuel.Saindon@microchip.com")
    
        }

        $mailSender = "EM Global Report<emreport_global@microchip.com>";
        $mailSubject = "EM Report - " + $TodayDate;

        Write-Host ("Sending Email report(Global) ...")
        Write-ToLogFile -LogContent ("Sending Email report(Global) ...")
        SendReportByMail -receiver $mailTo -ccTo $mailCC -sender $mailSender -subject $mailSubject -mailBody $htmlReport -isHTML $true -attachments $mailAttachments
    }
    else {
        Write-Host ("Error when send global report !!!, detail: {0}" -f $error[0])
        Write-ToLogFile -LogContent ("Error when send global report !!!, detail: {0}" -f $error[0])
    }
}

function UpdateReportToTeams() {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()] 
        [Alias("ReportObjects")] 
        [Object]$reportObjs, 

        [Parameter(Mandatory = $false)]
        [bool]$isTesting = $false
    )

    if (($reportObjs -is [System.Array]) -and ($reportObjs.Count -eq 3)) {

        $emReport = ("EM APAC Report on {0}`r`n{1}`r`n`r`nGlobal EM Report on {2}`r`n{3}`r`n{4}`r`n{5}" -f `
                $TodayDate, `
                $reportObjs[1].reportSummary, `
                $TodayDate, `
                $reportObjs[0].reportTotal, `
                $reportObjs[1].reportTotal, `
                $reportObjs[2].reportTotal)

        $taskId = "BvCuJ4o0tEu9TvODJHEfe2UAIOCE"

        if ($isTesting -ne $true) {
            $taskId = "REIRV9MKLk2JZoitlaWwDWUAJ_6s"
        }

        $titleUrl = "https://graph.microsoft.com/v1.0/planner/tasks/$taskid"
        $detailUrl = "https://graph.microsoft.com/v1.0/planner/tasks/$taskid/details"

        Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

        $result = Invoke-MgGraphRequest -Method GET $titleUrl
        if (($result -ne 0) -and (($result.title.toString()).ToLower().StartsWith("at risk in em"))) {

            $result = Invoke-MgGraphRequest -Method GET $detailUrl

            if ($result -ne 0) {
                Write-Host ("Sending report to Teams ...")
                Write-ToLogFile -LogContent ("Sending report to Teams ...")
                
                $headers = @{}
                $headers.Add("If-Match", $result["@odata.etag"])
        
                $newConetent = @{
                    description = $emReport
                }
        
                $contentJson = $newConetent | ConvertTo-Json
                Invoke-MgGraphRequest -Headers  $headers -Uri $detailUrl -Method 'PATCH' -ContentType 'application/json' -Body $contentJson
                $result = $null

                Write-Host ("Tasks ID: $taskId")
                Write-ToLogFile -LogContent ("Tasks ID: $taskId")
            }
        }

        Disconnect-MgGraph
    }
    else {
        Write-Host ("Error when update report to teams tasks !!!, detail: {0}" -f $error[0])
        Write-ToLogFile -LogContent ("Error when update report to teams tasks !!!, detail: {0}" -f $error[0])
    }
}

function IsNewSystem([object]$objItem) {
    if ($null -eq $objItem) {
        return $false;
    }
    else {
        if ([bool]($objItem.PSObject.Properties.name -match "FirstSeen")) {
            $result = $GetDate - (Get-Date $objItem.FirstSeen)
            if ($result.TotalDays -le 3) {
                return $true
            }
            else {
                return $false
            }
        }
        else {
            return $false
        }
    }
}

$stopWatch = [System.Diagnostics.Stopwatch]::StartNew();

if (Get-Module -ListAvailable -Name ImportExcel) {
    #Write-Host "Module exists"
}
else {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

if (Get-Module -ListAvailable -Name Microsoft.Graph) {
    #Write-Host "Module exists"
}
else {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
    Install-Module -Name Microsoft.Graph -MinimumVersion 1.1.0 -Scope CurrentUser -Force
}
#try {Import-Module $PSScriptRoot\..\ImportExcel} catch {throw ; return}

Import-Module ImportExcel
Import-Module Microsoft.Graph.Planner

$GetDate = Get-Date; # get current date.
$TodayDate = $GetDate.ToString("yyyyMMdd");
$ScriptFolderPath = $PSScriptRoot;
$XlsxFileName = ($ScriptFolderPath + '\data.xlsx');
$LogFile = ($ScriptFolderPath + '\logs\EMReportGenerate_' + $TodayDate + ".log");
[boolean]$OnTestingStatus = [System.Convert]::ToBoolean($IsDebugRunning)

Write-Host "Script running on debug mode: $OnTestingStatus"

$ReportHead = @"
    <style>
    @charset "UTF-8";
    
    body
    {
        font-size:11pt;
        webkit-text-size-adjust:none; 
        width:100% !important;
    }
    table
    {
        margin:0in;
        font-size:11.0pt;
        font-family:"Calibri",sans-serif;
        border-collapse:collapse;
    }
    td
    {
        font-size:10.5pt;
        border:1px solid #4F81BD;
        padding:2px 2px 2px 2px;
    }
    th
    {
        font-size:11.5pt;
        text-align:center;
        padding-top:5px;
        padding-bottom:5px;
        padding-right:5px;
        padding-left:5px;
        background-color:#4F81BD;
        color:#ffffff;
    }
    name tr
    {
        color:#F00000;
        background-color:#EAF2D3;
    }
    </style>

"@

if (Test-Path -Path $XlsxFileName) {
    #$exportFileName_2 = $scriptFolderPath + '\EPM_Problem_CN_HK_' + $todayDate +'.xlsx';

    $ExportFileName_1 = $ScriptFolderPath + '\EM_Problem_AMER_' + $TodayDate + '.xlsx';
    $ExportFileName_2 = $ScriptFolderPath + '\EM_Problem_EMEA_' + $TodayDate + '.xlsx';
    $ExportFileName_3 = $ScriptFolderPath + '\EM_Problem_APAC_' + $TodayDate + '.xlsx';
    $ExportFiles = @($ExportFileName_1, $ExportFileName_2, $ExportFileName_3, $XlsxFileName)
    $ArchiveFolderName = $ScriptFolderPath + '\' + $TodayDate;

    Write-Host ("Load data from Excel file: {0}" -f $XlsxFileName)
    Write-ToLogFile -LogContent ("Load data from Excel file: {0}" -f $XlsxFileName)
    $sourceData = Import-Excel -Path $XlsxFileName -StartRow 3;

    $amer_ReportObj = GenerateReportAMER -sourceData $sourceData -exportFileName $ExportFiles[0] -isTesting $OnTestingStatus
    $emea_ReportObj = GenerateReportEMEA -sourceData $sourceData -exportFileName $ExportFiles[1] -isTesting $OnTestingStatus
    $apac_ReportObj = GenerateReportAPAC -sourceData $sourceData -exportFileName $ExportFiles[2] -isTesting $OnTestingStatus

    SendGlobalReport -reportObjs @($amer_ReportObj, $apac_ReportObj, $emea_ReportObj) -isTesting $OnTestingStatus
    #SendRegionsReport -report_Obj $amer_ReportObj
    #SendRegionsReport -report_Obj $emea_ReportObj
    SendRegionsReport -report_Obj $apac_ReportObj

    UpdateReportToTeams -reportObjs @($amer_ReportObj, $apac_ReportObj, $emea_ReportObj) -isTesting $OnTestingStatus

    DoFileArchive -ArchiveFolder $ArchiveFolderName -FileLists $ExportFiles
}
else {
    Write-Host "Source file(data.xlsx) not exist !!!!";
    Write-ToLogFile -LogContent ("Source data(data.xlsx) not exist !!!!")
}

$stopWatch.Stop();
Write-Host -ForegroundColor yellow ('Total Runnning time: ' + $stopWatch.Elapsed.TotalMinutes + ' minutes');