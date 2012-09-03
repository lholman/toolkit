#-----------------------------------------------------------------------------------------------------------------------------

function global:Open-Excel($excel, [string]$excelfilename, [string]$excelfilenew)
{
    if (!$(test-path $excelfilename))
    {
        write-host "File doesn't exist..."
        return $null
    }
    
    $excelfile = $excel.Workbooks.Open($excelfilename)
    if($excelfilenew -ne "") {
        $excelfile.saveas($excelfilenew)
    }
    
    return $excelfile
}

#-----------------------------------------------------------------------------------------------------------------------------

function copyYesterdaysFigures
{
    
    $sheet1 = $excel.Worksheets.Item(1)
    $salescol = [int](GetConfig "PARAMS" "salescol")
    $inventorycol = [int](GetConfig "PARAMS" "inventorycol")
    $startrow = [int](GetConfig "PARAMS" "startrow")
    $endrow = [int](GetConfig "PARAMS" "endrow")
    $salesstatuscol = [int](GetConfig "PARAMS" "salesstatuscol")
    $inventorystatuscol = [int](GetConfig "PARAMS" "inventorystatuscol")
    
    copyCells $startrow $endrow $salescol $sheet1
    copyCells $startrow $endrow $inventorycol $sheet1
    clearStatus $startrow $endrow $salesstatuscol $sheet1
    clearStatus $startrow $endrow $inventorystatuscol $sheet1
    return $sheet1
    
    }

#-----------------------------------------------------------------------------------------------------------------------------

function clearStatus([int]$row, [int]$endrow, [int]$col, $sheet1)
{
    do {
        $sheet1.Cells.Item($row, $col) = ""
        $row++
    }
    while ($row -ne $endrow)
}

#-----------------------------------------------------------------------------------------------------------------------------

function copyCells([int]$row, [int]$endrow, [int]$col, $sheet1)
{
    do {
        $sheet1.Cells.Item($row, $col-1) = $sheet1.Cells.Item($row, $col).text
        $row++
    }
    while ($row -ne $endrow)
}

#-----------------------------------------------------------------------------------------------------------------------------

function getlatestspreadsheet
{
    $dayoftheweek   = (Get-Date -format dddd)
    if([String]::Compare($dayoftheweek, "Monday", $False) -eq 0)
    {
        return ((get-date).AddDays(-3)).ToString("ddMMyyyy")
    }
    return ((get-date).AddDays(-1)).ToString("ddMMyyyy")
}

#-----------------------------------------------------------------------------------------------------------------------------

function GetConfig([string]$nodename, [string]$value) {

    [System.Xml.XmlDocument] $xd = new-object System.Xml.XmlDocument
    $file = resolve-path("c:\skylight\config.xml") #"$pwd\$global:configfile")
    $xd.load($file)

    $nodelist = $xd.selectnodes("/Config/$nodename")
    
    foreach ($Node in $nodelist) {
      $inputsNode = $Node.GetAttribute($value)
    }
    
    return $inputsNode
    clear-variable -name xd 
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function ignoreSSLCerts
{
    try {
        $netAssembly = [Reflection.Assembly]::GetAssembly([System.Net.Configuration.SettingsSection])
         
        if($netAssembly)
        {
            $bindingFlags = [Reflection.BindingFlags] "Static,GetProperty,NonPublic"
            $settingsType = $netAssembly.GetType("System.Net.Configuration.SettingsSectionInternal")
         
            $instance = $settingsType.InvokeMember("Section", $bindingFlags, $null, $null, @())
         
            if($instance)
            {
                $bindingFlags = "NonPublic","Instance"
                $useUnsafeHeaderParsingField = $settingsType.GetField("useUnsafeHeaderParsing", $bindingFlags)
         
                if($useUnsafeHeaderParsingField)
                {
                  $useUnsafeHeaderParsingField.SetValue($instance, $true)
                }
            }
        }  
    } catch {
        write-host " An error was encountered: $error[0]"
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function bindExchangeService([string]$mailboxname, [string]$mailboxpwd)
{
    try {
        ## Set Exchange Version  
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
        $creds = New-Object System.Net.NetworkCredential($mailboxname, $mailboxpwd)   
        $service.Credentials = $creds  

        # Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
        
        #CAS URL Option 1 Autodiscover  
        $service.AutodiscoverUrl($mailboxname,{$true})  
        write-host "Using CAS Server : " + $Service.url    
        return $service
    } catch {
        write-host "An error was encountered: $error[0]"
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function bindExchangeInbox([object]$service, [string]$mailboxname)
{
    try {
        #Bind to Inbox    
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailboxname)     
        $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderid)
        return $inbox
    } catch {
        write-host " An error was encountered: $error[0]"
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function getemails([object]$inbox, [string]$sheetname, [string]$datatype, [string]$from, [string]$downloaddir, $sheet1)
{
    try {
        #Define the properties to get  
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)      
            
        #Define AQS Search String
        $AQSString = "attachment:~=$sheetname AND received:today"    #AND fromaddress:~=$from 
        #Define ItemView to retrieve just 1000 Items    
        $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
        $fiItems = $null    
        do{    
            $fiItems = $service.FindItems($inbox.id, $AQSString, $ivItemView) 

            [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
            foreach($Item in $fiItems.Items){                       
                foreach($attachment in $Item.Attachments){
                    $attachment.load()
                    $attachmentname = $attachment.Name
                    $fiFile = new-object System.IO.FileStream(($downloaddir + “\” + $attachmentname), [System.IO.FileMode]::Create)
                    $fiFile.Write($attachment.Content, 0, $attachment.Content.Length)
                    $fiFile.Close()
                    $filename = "$downloaddir$attachmentname"
                    write-host "Downloaded Attachment : $downloaddir$attachmentname"
                    extractData $filename $datatype $sheet1
                    remove-item $filename -Force -Recurse
                }  
            }    
            $ivItemView.Offset += $fiItems.Items.Count    
        }while($fiItems.MoreAvailable -eq $true)    
    } catch {
        write-host " An error was encountered: $error[0]"
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function UpdateSalesStatus($sheet1)
{
    #$sheet1 = $excelfile.Worksheets.Item(1)  
    
    $today = ((get-date).dayofweek.tostring()).substring(0,3)
    $todaydate = (get-date -format MM/dd/yyyy)
    
    $salescol       = [int](GetConfig "PARAMS" "salescol")
    $salestimecol   = [int](GetConfig "PARAMS" "salestimescol")
    $salesstatuscol = [int](GetConfig "PARAMS" "salesstatuscol")
    $salesdayscol   = [int](GetConfig "PARAMS" "salesdayscol")
    $startrow       = [int](GetConfig "PARAMS" "startrow")
    $endrow         = [int](GetConfig "PARAMS" "endrow")
    $namecol        = [int](GetConfig "PARAMS" "namecol")
    
    #check todays day to see if a file was supposed to have arrived
    do {
        $scheduledays = $sheet1.Cells.Item($startrow, $salesdayscol).text
        if($arrivaltime -ne "") {
            if($scheduledays -ne "") {
                if($scheduledays.contains($today)) {
                    if($sheet1.Cells.Item($startrow, $salescol).text -ne "") {
                        $scheduletime = ([datetime]($sheet1.Cells.Item($startrow, $salestimecol).text)).toshorttimestring()
                        $arrivaltime  = ([datetime]($sheet1.Cells.Item($startrow, $salescol).text)).toshorttimestring()
                        $arrivaldate  = ([datetime]($sheet1.Cells.Item($startrow, $salescol).text)).toshortdatestring()
                        $customer     = $sheet1.Cells.Item($startrow, $namecol).text
                        if($arrivaltime -le $scheduletime)
                        {
                            $status = "Ok"
                            $sheet1.Cells.Item($startrow, $salesstatuscol) = $status
                        } else {
                            if($arrivaldate -le $todaydate)
                            {
                                $status = "Ok"
                                $sheet1.Cells.Item($startrow, $salesstatuscol) = $status
                            } else {
                                $status = "Not Ok"
                                $sheet1.Cells.Item($startrow, $salesstatuscol) = $status
                            }
                        }
                    } else {
                        $status = "Not Ok"
                        $sheet1.Cells.Item($startrow, $salesstatuscol) = $status
                    }
                } else {
                    $status = "Not Expected"    
                    $sheet1.Cells.Item($startrow, $salesstatuscol) = $status
                }
            } else {
                    $status = "Not Listed"    
                    $sheet1.Cells.Item($startrow, $salesstatuscol) = $status                            
            }
        }
        $startrow++
    }     
    while ($startrow -lt $endrow)
  
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function UpdateInvStatus($sheet1)
{
    #$sheet1 = $excelfile.Worksheets.Item(1)  
    
    $today = ((get-date).dayofweek.tostring()).substring(0,3)
    $todaydate = (get-date -format MM/dd/yyyy)
    
    $invcol       = [int](GetConfig "PARAMS" "inventorycol")
    $invtimecol   = [int](GetConfig "PARAMS" "inventorytimescol")
    $invstatuscol = [int](GetConfig "PARAMS" "inventorystatuscol")
    $invdayscol   = [int](GetConfig "PARAMS" "invdays")
    $startrow       = [int](GetConfig "PARAMS" "startrow")
    $endrow         = [int](GetConfig "PARAMS" "endrow")
    $namecol        = [int](GetConfig "PARAMS" "namecol")
    
    #check todays day to see if a file was supposed to have arrived
    do {
        $scheduledays = $sheet1.Cells.Item($startrow, $invdayscol).text
        if($sheet1.Cells.Item($startrow, $invcol).text -ne "") {
            if($arrivaltime -ne "") {
                if($scheduledays -ne "") {
                    if($scheduledays.contains($today)) {
                        if($sheet1.Cells.Item($startrow, $invcol).text -ne "") {
                            $scheduletime = ([datetime]($sheet1.Cells.Item($startrow, $invtimecol).text)).toshorttimestring()
                            $arrivaltime  = ([datetime]($sheet1.Cells.Item($startrow, $invcol).text)).toshorttimestring()
                            $arrivaldate  = ([datetime]($sheet1.Cells.Item($startrow, $invcol).text)).toshortdatestring()
                            $customer     = $sheet1.Cells.Item($startrow, $namecol).text
                            if($arrivaltime -le $scheduletime)
                            {
                                $status = "Ok"
                                $sheet1.Cells.Item($startrow, $invstatuscol) = $status
                            } else {
                                if($arrivaldate -le $todaydate)
                                {
                                    $status = "Ok"
                                    $sheet1.Cells.Item($startrow, $invstatuscol) = $status
                                } else {
                                    $status = "Not Ok"
                                    $sheet1.Cells.Item($startrow, $invstatuscol) = $status
                                }
                            }
                        } else {
                            $status = "Not Ok"
                            $sheet1.Cells.Item($startrow, $invstatuscol) = $status  
                        }
                    } else {
                        $status = "Not Expected"    
                        $sheet1.Cells.Item($startrow, $invstatuscol) = $status
                    }
                }
            } else {
                $status = "Not Listed"    
                $sheet1.Cells.Item($startrow, $salesstatuscol) = $status  
            }
        }
        $startrow++
    }     
    while ($startrow -lt $endrow)
  
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function ReconcileStatusData($sheet1)
{
    #$sheet1 = $excelfile.Worksheets.Item(1)  
    
    $today      = ((get-date).dayofweek.tostring()).substring(0,3)
    $todaydate = ((get-date).toshortdatestring())
    
    $salesstatuscol = [int](GetConfig "PARAMS" "salesstatuscol")
    $invstatuscol = [int](GetConfig "PARAMS" "inventorystatuscol")
    $startrow       = [int](GetConfig "PARAMS" "startrow")
    $endrow         = [int](GetConfig "PARAMS" "endrow")
    $namecol        = [int](GetConfig "PARAMS" "namecol")
    $invdayscol     = [int](GetConfig "PARAMS" "invdays")
    $salesdayscol   = [int](GetConfig "PARAMS" "salesdayscol")
    
    if($invdayscol -eq "") { $sheet1.Cells.Item($startrow, $invstatuscol) = "" }
    if($salesdayscol -eq "") { $sheet1.Cells.Item($startrow, $salesstatuscol) = "" }
    
    #check todays day to see if a file was supposed to have arrived
    do {
        $invstatus      = $sheet1.Cells.Item($startrow, $invstatuscol).text
        $salesstatus    = $sheet1.Cells.Item($startrow, $salesstatuscol).text
        $customer       = $sheet1.Cells.Item($startrow, $namecol).text   
        
        switch($true)
        {
            (($invstatus -eq "Ok") -and ($salesstatus -eq "Ok"))
            {
                $retaildata.add($customer, "Ok")
                break 
            }
            (($invstatus -eq "Ok") -and ($salesstatus -eq ""))
            {
                $retaildata.add($customer, "Ok")
                break
            }
            (($invstatus -eq "") -and ($salesstatus -eq "Ok"))
            {
                $retaildata.add($customer, "Ok")
                break               
            }    
            (($invstatus -eq "Ok") -and ($salesstatus -eq "Not Expected"))
            {
                $retaildata.add($customer, "Ok")
                break               
            }     
            (($invstatus -eq "Not Expected") -and ($salesstatus -eq "Not Expected"))
            {
                $retaildata.add($customer, "Not Expected")
                break               
            }            
            (($invstatus -eq "Not Expected") -and ($salesstatus -eq "Ok"))
            {
                $retaildata.add($customer, "Ok")
                break               
            }  
            (($invstatus -eq "") -and ($salesstatus -eq "Not Expected"))
            {
                $retaildata.add($customer, "Ok")
                break               
            }     
            (($invstatus -eq "Not Expected") -and ($salesstatus -eq ""))
            {
                $retaildata.add($customer, "Ok")
                break               
            }
            (($invstatus -eq "Not Listed") -or ($salesstatus -eq "Not Listed"))
            {
                $retaildata.add($customer, "Not Ok")
                break               
            }             
            ($invstatus -eq "Not Ok")
            {
                $retaildata.add($customer, "Not Ok")
                break               
            }             
            ($salesstatus -eq "Not Ok")
            {
                $retaildata.add($customer, "Not Ok")
                break               
            }             
        }                      
        $startrow++
    }     
    while ($startrow -lt $endrow)
  
}
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function extractData([string]$filename, [string]$datatype)
{
    $excel1 = new-object -comobject Excel.Application
    $excel1.visible = $True
    
    $exceltmp = $excel1.Workbooks.Open($filename)
    $sheettemp = $exceltmp.Worksheets.Item(1)
    
    $sheet1 = $excelfile.Worksheets.Item(1)    
    
    $startrow = [int](GetConfig "PARAMS" "startrow")
    $endrow = [int](GetConfig "PARAMS" "endrow")
        
    switch($datatype)
    {
        "SalesData"
        {
            $codecol    = [int](GetConfig "PARAMS" "codecol")
            $salescol   = [int](GetConfig "PARAMS" "salescol")
            $row = 2
            $col = 1
            $valcol   = (GetConfig "PARAMS" "salesdatacol")
            
            do {
                $code = $sheet1.Cells.Item($startrow, $codecol).text
                $retval = (LocateLatestFigure $code $sheettemp $valcol)
                $sheet1.Cells.Item($startrow, $salescol) = $retval
                $startrow++
            }     
            while ($startrow -lt $endrow)
            $excel1.workbooks.close()
            $excel1.quit()
            break
        }
        "InvData"
        {
            $codecol    = [int](GetConfig "PARAMS" "codecol")
            $invcol     = [int](GetConfig "PARAMS" "inventorycol")
            $row        = 2
            $col        = 1
            $valcol     = [int](GetConfig "PARAMS" "invdatacol")
            
            do {
                $code = $sheet1.Cells.Item($startrow, $codecol).text
                $retval = (LocateLatestFigure $code $sheettemp $valcol)
                $sheet1.Cells.Item($startrow, $invcol) = $retval
                $startrow++
            }     
            while ($startrow -lt $endrow)
            $excel1.workbooks.close()
            $excel1.quit()
            break
        }
        "SonyInterchange"
        {
            $namecol        = [int](GetConfig "PARAMS" "namecol")
            $sonyintcol     = [int](GetConfig "PARAMS" "sonyinterchangecol")
            $firstdatarow   = [int](GetConfig "PARAMS" "sonytransstartrow")
            $lastdatarow    = [int](GetConfig "PARAMS" "sonytransendrow")
            $valcol         = [int](GetConfig "PARAMS" "sonyintvalcol")
            
            do {
                $name = $sheet1.Cells.Item($startrow, $namecol).text
                $retval = (LocateLatestsonyInterchangeFigures $name $sheettemp $valcol $firstdatarow $lastdatarow)
                $sheet1.Cells.Item($startrow, $sonyintcol) = $retval
                $startrow++
            }     
            while ($startrow -lt $endrow)
            $excel1.workbooks.close()
            $excel1.quit()
            break            
        }
        "AS2Interchange"
        {
            $codecol        = [int](GetConfig "PARAMS" "codecol")
            $as2intcol      = [int](GetConfig "PARAMS" "as2interchangecol")
            $firstdatarow   = [int](GetConfig "PARAMS" "as2transstartrow")
            $lastdatarow    = [int](GetConfig "PARAMS" "as2transendrow")
            $valcol         = [int](GetConfig "PARAMS" "as2intvalcol")
            
            do {
                $code = $sheet1.Cells.Item($startrow, $codecol).text
                $retval = (LocateLatestas2InterchangeFigures $code $sheettemp $valcol $firstdatarow $lastdatarow)
                $sheet1.Cells.Item($startrow, $as2intcol) = $retval
                $startrow++
            }     
            while ($startrow -lt $endrow)
            $excel1.workbooks.close()
            $excel1.quit()
            break
        }
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function LocateLatestFigure([string]$code, $sheettemp, [int]$valcol)
{
    $row = 2
    $col = 1
    do {
        $currentcode = $sheettemp.Cells.Item($row, $col).text
        #Check to see if the the code is the one we're looking for
        if($currentcode -eq $code) {
            do {
                $currentcode = $sheettemp.Cells.Item($row, $col).text
                #Now start looping through to see when the code doesn't match
                if($currentcode -ne $code) {
                    #no check if the code cell is empty and check the last entry
                    if($currentcode -eq "") {
                        if(($sheettemp.Cells.Item($row-1, $col).text) -eq $code) {
                            return ($sheettemp.Cells.Item($row-1, $valcol).text)
                        } else {
                            break
                        }
                    }
                    return ($sheettemp.Cells.Item($row-1, $valcol).text)
                }
                $row++
            }     
            while ($currentcode -ne "")
        }  
        $row++
    }     
    while ($currentcode -ne "")
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function LocateLatestas2InterchangeFigures([string]$code, $sheettemp, [int]$valcol, [int]$firstdatarow, [int]$lastdatarow)
{
    
    $row = $firstdatarow
    $col = 1
    do {
        $currentcode = $sheettemp.Cells.Item($row, $col).text
        #Check to see if the the code is the one we're looking for
        if($currentcode -eq $code) {
            return ($sheettemp.Cells.Item($row, $valcol).text) 
        }  
        $row++
    }     
    while ($row -le $lastdatarow)
    return ""
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function LocateLatestsonyInterchangeFigures([string]$name, $sheettemp, [int]$valcol, [int]$firstdatarow, [int]$lastdatarow)
{
    $name = ($name -replace ' \(.*\)', "")
    $row = $firstdatarow
    $col = 1
    do {
        $currentname = $sheettemp.Cells.Item($row, $col).text
        #Check to see if the the code is the one we're looking for
        if($currentname -eq $name) {
            return ($sheettemp.Cells.Item($row, $valcol).text) 
        }  
        $row++
    }     
    while ($row -le $lastdatarow)
    return ""
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

function checkTrailingBackslash([string]$path)  {

    try {
        if($path.substring($path.length -1, 1) -ne "\") {
            $path = $path+"\"
        }
        return $path
    } catch {
        write-host "There was a parsing the path $path. [$Error[0]]"
        exit
    }
}

#-----------------------------------------------------------------------------------------------------------------------------

function GenerateHTML([string]$emailfile, [string]$bodytext, [string]$subject, [string]$datetime, [boolean]$error) {
    PROCESS {
        out-file -encoding ASCII $emailfile -input "<html>"
        out-file -encoding ASCII $emailfile -input "<head>" -append
        out-file -encoding ASCII $emailfile -input "<title></title>" -append
        out-file -encoding ASCII $emailfile -input "<style type='text/css'>" -append
        out-file -encoding ASCII $emailfile -input ".smallboldleft" -append
        out-file -encoding ASCII $emailfile -input "{"  -append
        out-file -encoding ASCII $emailfile -input "text-align: left;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 10pt;"  -append
        out-file -encoding ASCII $emailfile -input "font-weight: bold;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input ".smallnormalleftok" -append
        out-file -encoding ASCII $emailfile -input "{" -append
        out-file -encoding ASCII $emailfile -input "text-align: left;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 10pt;" -append
        out-file -encoding ASCII $emailfile -input "color: white;" -append
        out-file -encoding ASCII $emailfile -input "background: green;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input ".smallnormalleftnotok" -append
        out-file -encoding ASCII $emailfile -input "{" -append
        out-file -encoding ASCII $emailfile -input "text-align: left;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 10pt;" -append
        out-file -encoding ASCII $emailfile -input "color: white;" -append
        out-file -encoding ASCII $emailfile -input "background: red;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input ".smallnormalleftnotexpected" -append
        out-file -encoding ASCII $emailfile -input "{" -append
        out-file -encoding ASCII $emailfile -input "text-align: left;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 10pt;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input ".tabletitleleft" -append
        out-file -encoding ASCII $emailfile -input "{" -append
        out-file -encoding ASCII $emailfile -input "text-align: left;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 11pt;" -append
        out-file -encoding ASCII $emailfile -input "font-weight: bold;" -append
        out-file -encoding ASCII $emailfile -input "width: 200px;" -append
        out-file -encoding ASCII $emailfile -input "border-bottom:solid #00355F 1.0pt;" -append
        out-file -encoding ASCII $emailfile -input "border-top:solid #00355F 1.0pt;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input ".tableentry" -append
        out-file -encoding ASCII $emailfile -input "{" -append
        out-file -encoding ASCII $emailfile -input "text-align: center;" -append
        out-file -encoding ASCII $emailfile -input "font-family: Calibri;" -append
        out-file -encoding ASCII $emailfile -input "font-size: 10pt;" -append
        out-file -encoding ASCII $emailfile -input "width: 600px;" -append
        out-file -encoding ASCII $emailfile -input "}" -append
        out-file -encoding ASCII $emailfile -input "</style>" -append
        out-file -encoding ASCII $emailfile -input "</head>" -append
        out-file -encoding ASCII $emailfile -input "<body>" -append
        out-file -encoding ASCII $emailfile -input "<table border=0 cellspacing=0 cellpadding=0 width='100%' style='width:100.0%'>" -append
        out-file -encoding ASCII $emailfile -input "<tr style='height:49.5pt'>" -append
        out-file -encoding ASCII $emailfile -input "<td width='1%' style='width:1.0%;padding:0in 21.75pt 0in 0in;height:49.5pt'>" -append
        out-file -encoding ASCII $emailfile -input "<p><img width=207 height=72 src='cid:imgSRG'alt='SRG'/></p>" -append
        out-file -encoding ASCII $emailfile -input "</td>" -append
        out-file -encoding ASCII $emailfile -input "</tr>" -append
        out-file -encoding ASCII $emailfile -input "</table>" -append
        out-file -encoding ASCII $emailfile -input "<table width=100% style='width:100%;border-collapse:collapse'>" -append
        out-file -encoding ASCII $emailfile -input "<tr>" -append
        out-file -encoding ASCII $emailfile -input "<td style='border:none;border-bottom:solid #00355F 1.0pt; padding:10.0pt 22.5pt 3.75pt 0cm'>" -append
        out-file -encoding ASCII $emailfile -input "<p><span style='font-size:10.0pt;font-family:Calibri'><b>$datetime<br/>$subject</b></span></p>" -append
        out-file -encoding ASCII $emailfile -input "</td>" -append
        out-file -encoding ASCII $emailfile -input "</tr>" -append
        out-file -encoding ASCII $emailfile -input "</table>" -append
        out-file -encoding ASCII $emailfile -input "<table width=100% style='width:100%;border-collapse:collapse'>" -append
        out-file -encoding ASCII $emailfile -input "<tr>" -append
        out-file -encoding ASCII $emailfile -input "<td colspan=2 style='padding:10.0pt 22.5pt 0cm 0cm'>" -append
        out-file -encoding ASCII $emailfile -input "<p>&nbsp;</p>" -append
        out-file -encoding ASCII $emailfile -input "</td></tr>" -append
        out-file -encoding ASCII $emailfile -input "<tr>" -append
        out-file -encoding ASCII $emailfile -input "<td style='padding:9.0pt 22.5pt 0cm 0cm'>" -append
        out-file -encoding ASCII $emailfile -input "<p><span style='font-size:9.0pt;font-family:Calibri'>" -append
        out-file -encoding ASCII $emailfile -input $bodytext -append
        out-file -encoding ASCII $emailfile -input "<p>&nbsp;</p>" -append
        out-file -encoding ASCII $emailfile -input "</td></tr></table>" -append
        out-file -encoding ASCII $emailfile -input "<p>" -append
        out-file -encoding ASCII $emailfile -input "<table width=100% style='width:100%;border-collapse:collapse'>" -append
        out-file -encoding ASCII $emailfile -input "<tr>" -append
        out-file -encoding ASCII $emailfile -input "<td class='tabletitleleft'>Retailer</td>" -append
        out-file -encoding ASCII $emailfile -input "<td class='tabletitleleft'>Skylight Status</td>" -append
        out-file -encoding ASCII $emailfile -input "<td class='tabletitleleft'></td>" -append
        out-file -encoding ASCII $emailfile -input "<td class='tabletitleleft'>&nbsp;</td>" -append
        out-file -encoding ASCII $emailfile -input "</tr>" -append
        $Retaildata.GetEnumerator() | Sort-Object Name | Foreach-Object {$_}| buildTables
        out-file -encoding ASCII $emailfile -input "</table>" -append        
        out-file -encoding ASCII $emailfile -input "</p>" -append
        out-file -encoding ASCII $emailfile -input "</div>" -append
        out-file -encoding ASCII $emailfile -input "</body>" -append
        out-file -encoding ASCII $emailfile -input "</html>" -append
    }
}

#-----------------------------------------------------------------------------------------------------------------------------

function BuildEmailBody([string]$outfile, [string]$excelfile) {
    PROCESS {
        $to       = (GetConfig "Mail" "To")
        $subject  = (GetConfig "Mail" "Subject")
        $sender   = (GetConfig "Mail" "Sender") 
        $bodytext = (GetConfig "Mail" "Body")
        $cc       = (GetConfig "Mail" "cc")
        $datetime = Get-Date
        $datetime = $datetime.ToLongDateString()
        
        GenerateHTML $outfile $bodytext $subject $datetime $true

        SendEmail $to $outfile $subject $sender $cc $excelfile    
    }
}

#-----------------------------------------------------------------------------------------------------------------------------

function SendEmail([string]$to,[string]$outfile, [string]$subject, [string]$sender, [string]$cc, [string]$excelfile) 
{
    PROCESS {
        $from = $sender
        $body = Get-Content $outfile
        $msg  = New-object System.Net.Mail.MailMessage $from, $to, $subject, $body
        if($cc -ne "") {$msg.cc.add($cc)}
        $msg.attachments.add($excelfile)
        
        $msg.IsBodyHTML = $true
        $ImagePath1 = "c:\skylight\srg-logo.png"
        $LinkedResource1 = New-Object Net.Mail.LinkedResource($ImagePath1, "image/jpeg")
        $LinkedResource1.ContentId = "imgSRG"; 
        
        $HtmlView = [Net.Mail.AlternateView]::CreateAlternateViewFromString($body, "text/html")
        $HtmlView.LinkedResources.Add($LinkedResource1)

        $msg.AlternateViews.Add($HtmlView)
              
        $server = (GetConfig "Mail" "SMTPServer")
        $client = new-object system.net.mail.smtpclient $server  
         
        write-host "Sending an e-mail message to $to by using SMTP host $server"
        try {  
           $client.Send($msg)  
           write-host "Message to: $to, from: $from has been successfully sent"
        }  
        catch {  
          write-host "Exception caught in Message: $Error[0]"
          $error.clear() 
        }
    }
}

#-----------------------------------------------------------------------------------------------------------------------------

function buildTables
{
    PROCESS {
        $retailer = $_.key
        $result   = $_.value

        switch ($result)
        {
            "Ok"
                {$class = "smallnormalleftok"}
            "Not Ok"
                {$class = "smallnormalleftnotok"}
            "Not Expected"
                {$class = "smallnormalleftok"}
        }
        out-file -encoding ASCII $emailfile -input "<tr>" -append
        out-file -encoding ASCII $emailfile -input "<td class='$class'>$retailer</td>" -append            
        out-file -encoding ASCII $emailfile -input "<td class='$class'>$result</td>" -append  
        out-file -encoding ASCII $emailfile -input "<td class='$class'>&nbsp;</td>" -append
        out-file -encoding ASCII $emailfile -input "<td class='$class'>&nbsp;</td>" -append
        out-file -encoding ASCII $emailfile -input "</tr>" -append
    }
}

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files (x86)\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"  

$Retaildata   = @{}

$downloaddir = (checkTrailingBackslash (GetConfig "PARAMS" "skylightpath"))
$mailboxname = (GetConfig "PARAMS" "mailboxname")
$mailboxpwd  = (GetConfig "PARAMS" "mailboxpassword")

$htmlemailout = (GetConfig "PARAMS" "htmloutput")
$htmlemailout = "c:\skylight\$htmlemailout"
out-file -encoding ASCII $htmlemailout

$excel = new-object -comobject Excel.Application
$excel.visible = $True

$latest = getlatestspreadsheet

$strPath     = (checkTrailingBackslash (GetConfig "PARAMS" "skylightpath"))
$strfilename = (GetConfig "PARAMS" "spreadsheetname")+"$latest.xlsx"
$strfilenew  = (GetConfig "PARAMS" "spreadsheetname")+((get-date).ToString("ddMMyyyy"))

$excelfile = (Open-Excel $excel $strPath$strfilename $strPath$strfilenew)
$sheet1 = copyYesterdaysFigures 

$arrfiles = (GetConfig "PARAMS" "inputfiles").split(",")
$arrfilestype = (GetConfig "PARAMS" "inputfilestype").split(",")
$arrfilesfrom = (GetConfig "PARAMS" "inputfilesfrom").split(",")

ignoreSSLCerts
$service = bindExchangeService $mailboxname $mailboxpwd
$inbox = bindExchangeInbox $service $mailboxname

for($i=0; $i -le $arrfiles.getupperbound(0); $i++)
{
    getemails $inbox $arrfiles[$i] $arrfilestype[$i] $arrfilesfrom[$i] $downloaddir $sheet1
}

UpdateSalesStatus $sheet1
UpdateInvStatus $sheet1

ReconcileStatusData $sheet1

$excel.ActiveWorkbook.save()
$excel.quit()

BuildEmailBody $htmlemailout "$strPath$strfilenew.xlsx"



