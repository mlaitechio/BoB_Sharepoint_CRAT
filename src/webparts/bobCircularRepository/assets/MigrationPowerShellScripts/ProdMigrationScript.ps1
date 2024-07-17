# ======================  SharePointPnPPowerShellOnline

if ( (Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable | ? Version -eq "3.29.2101.0").Name.Length -eq 0 ) {
    try {
        Write-Host "PowerShell module 'SharePointPnPPowerShellOnline' version 3.29.2101.0  not found!"
        Write-Host "The module is available at https://github.com/SharePoint/PnP-PowerShell/releases/tag/3.29.2101.0"  
        Write-Host "The script will try to install this module, once installed, the script will continue.`n"

        Write-Host "- Trying to install PowerShell module 'SharePointPnPPowerShellOnline' version 3.29.2101.0  ... `n " -ForegroundColor Yellow
             
        Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.29.2101.0 -SkipPublisherCheck -AllowClobber -Force        
        Write-Host "- Module succesfully intalled! `n " -ForegroundColor Yellow

        Write-Host "- Loading module ... `n " -ForegroundColor Yellow
        Import-Module SharePointPnPPowerShellOnline -RequiredVersion 3.29.2101.0 -DisableNameChecking -ErrorAction Stop

        Write-Host "- Trying to install PowerShell module 'Import Excel' ... `n " -ForegroundColor Yellow

        Install-Module -Name ImportExcel -Scope CurrentUser

        Write-Host "- Loading excel module ... `n " -ForegroundColor Yellow
        Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop
    }
    catch {
        Write-Host "Error while trying to install PowerShell module 'SharePointPnPPowerShellOnline'!"
        Write-Host "The module is also available at https://github.com/SharePoint/PnP-PowerShell/releases/tag/2.23.1802.0  `n";
        $global:errorLogs += WriteErrorLogs -circularNumber "N/A" -message "$($_.Exception.Message)"; 
        exit
    }    
}
else {
    try {
        Import-Module SharePointPnPPowerShellOnline -RequiredVersion 3.29.2101.0 -DisableNameChecking -ErrorAction Stop;
        Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop;
    }
    catch {
        Write-Host "`nPowerShell module 'SharePointPnPPowerShellOnline' not loaded!"
        Write-Host "Check if your computer has the correct PowerShell module properly installed before running the script."
        Write-Host "The script requires SharePointPnPPowerShellOnline version 3.29.2101.0 "  
        Write-Host "The module is available at https://github.com/SharePoint/PnP-PowerShell/releases/tag/3.29.2101.0  `n";
        $global:errorLogs += WriteErrorLogs -circularNumber "N/A" -message "$($_.Exception.Message)";   
        exit;
    }
}



Function ConvertDate([string]$uploadDate) {
    $monthDateYearString = $null;
    try {
        if ($null -ne $uploadDate -and $uploadDate -ne "") {
            $dateString = $uploadDate.Split('-');
            $monthDateYearString = $dateString[1] + "-" + $dateString[0] + "-" + $dateString[2];
        }
    }
    
    catch {
        $global:errorLogs += WriteErrorLogs -circularNumber "N/A" -message "$($_.Exception.Message)";
        $monthDateYearString = $null;
    }

    return $monthDateYearString
}


#Function to Add all files as a attachment to list item
Function AddAttachmentsFromFolder($ListItem, $circularNumber, $docPath, $isAdd) {
   
    $FullFilePath = $FolderPath + $docPath;
    # $ChildFile = Get-ChildItem -Path $FullFolderPath;

    # foreach ($j in $ChildFile) {
    # $FullFilePath = $FullFolderPath + "\" + $j.Name
        
    try {
        $ErrorActionPreference = "Stop";
        
        # Add-PnPListItemAttachment -List $list -Identity $ListItem -Path $FullFilePath;
        if ($null -ne $docPath -and $isAdd) {
            $FileStream = New-Object IO.FileStream($FullFilePath, [System.IO.FileMode]::Open)
            $AttachmentInfo = New-Object -TypeName Microsoft.SharePoint.Client.AttachmentCreationInformation
            $AttachmentInfo.FileName = Split-Path $FullFilePath -Leaf
            $AttachmentInfo.ContentStream = $FileStream
            $AttachedFile = $ListItem.AttachmentFiles.Add($AttachmentInfo);
            Invoke-PnPQuery
            $FileStream.Close();

            Write-Host "Attachment Added to ListItem" -ForegroundColor Green;
            $global:logs += WriteLogs -circularNumber $circularNumber -isMigrated "Yes" -status "Attachment" -message "Attachment Added Successfully";
        }
        else {
            if ($null -ne $docPath) {
                $newFileName = $docPath.Split('\')[1];
                $listAttachmentFolderPath = "Lists/$($listName)/Attachments/$($ListItem.Id)";
                Add-PnPFile -Path $FullFilePath -Folder $listAttachmentFolderPath -NewFileName $newFileName -ErrorAction SilentlyContinue;
                Write-Host "Attachment Updated to ListItem $($ListItem.Id)" -ForegroundColor Green;
                $global:logs += WriteLogs -circularNumber $circularNumber -isMigrated "Yes" -status "Attachment" -message "Attachment Updated Successfully";
            } 
            else {
                $global:logs += WriteLogs -circularNumber $circularNumber -isMigrated "Yes" -status "Attachment" -message "Doc Path Missing"; 
            }
        }

       
    }
    catch {
        Write-host "$FullFilePath : Error :"$_.Exception.Message -f Red
        $global:errorLogs += WriteErrorLogs -circularNumber $circularNumber -message "$($_.Exception.Message)";  
    }
    finally {
        $ErrorActionPreference = "Continue";
    }
}

Function WriteLogs($circularNumber, $isMigrated, $status, $message) {
    <# Logs file to be added for Audit Purposes #>

    $currentlogs += [PSCustomObject]@{
        TimeStamp      = (Get-Date -Format "s") + "Z";
        CircularNumber = $circularNumber;
        IsMigrated     = $isMigrated;
        Status         = $status;
        Message        = $message;  
        Classification = $circularClassification;      
    }

    return $currentlogs;

    
}

Function WriteErrorLogs($circularNumber, $message) {
    $currentErrorLogs += [PSCustomObject]@{
        TimeStamp      = (Get-Date -Format "s") + "Z";
        Status         = "Error";
        Message        = $message;
        CircularNumber = $circularNumber;
        Classification = $circularClassification;        
    };

    return $currentErrorLogs;
}


Function MasterCircularMigration() {

    $masterCircularData = Import-Excel -Path $masterCircularPath ;
    $circularNumber = "";
    
    try {
        Connect-PnPOnline -Url $siteUrl -UseWebLogin;
    }

    catch {
        Write-Host "Error : $($_.Exception.Message)" -ForegroundColor Red;
        $global:errorLogs += WriteErrorLogs -circularNumber "N/A" -message "$($_.Exception.Message)";  
        continue;
    }

    #$departmentListItem = Get-PnPListItem -List $departmentList -PageSize 1000;

    ForEach ($record in $masterCircularData) {
            
        $circularNumber = $record.'CIRCULAR NO';
            
        #$department = $departmentListItem | Where-Object { $_["Title"] -eq $record.DEPTNAME }
        
        #if ($null -ne $department) {
        # Parse the input date string into a DateTime object 
        $convertedDate = ConvertDate -uploadDate $record.'CIR DATE';
            
        $date = $null;
        # Format the DateTime object into a string suitable for SharePoint list item
        $sharePointDateTime = $null;
            
        if ($null -ne $convertedDate) {
                
            $date = Get-Date ($convertedDate) -Format "s" -ErrorAction Continue;
            $sharePointDateTime = $date + "Z";
        }
        
    
        $currentItem = GetCircularListItem -circularNumber $circularNumber;

        #Update List Items - Map with Internal Names of the Fields!.Below you define a custom object @ {}for each record and update the SharePoint list
        if ($null -ne $currentItem) {

            Write-Host "Updating record for $($record.'CIRCULAR NO')" -ForegroundColor Green;
            try {

                $ErrorActionPreference = "Stop";
                $listItemUpdate = Set-PnPListItem -List $listName -Identity $currentItem -Values @{
                    "Subject"            = $($record.'SUBJECT');
                    "MigratedDepartment" = $($record.'DEPTNAME');
                    #"Department"         = $department.Id;
                    "Classification"     = $circularClassification;
                    "CircularNumber"     = $($record.'CIRCULAR NO');
                    "CircularStatus"     = "Published";
                    "MigratedOriginator" = $($record.'ORIGINATOR');
                    "MigratedIssuedFor"  = $($record.'ISSUED FOR');
                    "MigratedRefNumber"  = $($record.'REF NO');
                    "MigratedDocPath"    = $($record.'DOC PATH');
                    "IsMigrated"         = "Yes"
                    "PublishedDate"      = $sharePointDateTime;
    
                } ;
            


                Write-Host "Updated Item" -ForegroundColor Green;
                $global:logs += WriteLogs -circularNumber $record.'CIRCULAR NO' -isMigrated "Yes" -status "Updated" -message "Item Updated Successfully";

                AddAttachmentsFromFolder -ListItem $listItemUpdate -circularNumber $circularNumber -docPath $record.'DOC PATH' -isAdd $false;

            }

            catch {
                $global:errorLogs += WriteErrorLogs -circularNumber $circularNumber -message "$($_.Exception.Message)";  
            }

            finally {
                $ErrorActionPreference = "Continue";
            }
 
        }
                
        else {   
            #Add List Items - Map with Internal Names of the Fields!.Below you define a custom object @ {}for each record and update the SharePoint list.
            try {

                $ErrorActionPreference = "Stop";
                Write-Host "Adding record for $($record.'CIRCULAR NO')" -ForegroundColor Green;
                Add-PnPListItem -List $listName -Values @{
                    "Subject"            = $($record.'SUBJECT');
                    "MigratedDepartment" = $($record.'DEPTNAME');
                    "Classification"     = $circularClassification;
                    "CircularNumber"     = $($record.'CIRCULAR NO');
                    "CircularStatus"     = "Published";
                    "MigratedOriginator" = $($record.'ORIGINATOR');
                    "MigratedIssuedFor"  = $($record.'ISSUED FOR');
                    "MigratedRefNumber"  = $($record.'REF NO');
                    "MigratedDocPath"    = $($record.'DOC PATH');
                    "IsMigrated"         = "Yes"
                    "PublishedDate"      = $sharePointDateTime;

                };

                Write-Host "Added Item" -ForegroundColor Green;
                $global:logs += WriteLogs -circularNumber $record.'CIRCULAR NO' -isMigrated "Yes" -status "Created" -message "Item Added Successfully";

                AddAttachmentsFromFolder -ListItem $listItemUpdate -circularNumber $circularNumber -docPath $record.'DOC PATH' -isAdd $true;
            }
            catch {
                $global:errorLogs += WriteErrorLogs -circularNumber "$($circularNumber)" -message "$($_.Exception.Message)";  
            }

            finally {
                $ErrorActionPreference = "Continue";
            }
        }

        Write-Progress -PercentComplete (($i / $masterCircularData.Count) * 100) -Status "Migrating master circulars to Circular Repository SharePoint List" -Activity "Item $($i) of $($masterCircularData.Count)";
        $i++;
    }
    

    
}

# Function CircularMigration() {

#     $circularData = Import-Excel -Path $circularPath ;
#     $i = 0

#     try {

#         Connect-PnPOnline -Url $siteUrl -UseWebLogin;
#     }

#     catch {

#         Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red;
#         $global:errorLogs += WriteErrorLogs -circularNumber "N/A" -message "$($_.Exception.Message)";
#     }

#     #$departmentListItem = Get-PnPListItem -List $departmentList -PageSize 1000;

#     ForEach ($record in $circularData) {
           
#         $currentItem = GetCircularListItem -circularNumber $($record.'CIRCULAR_NO');

#         $circularNumber = $record.'CIRCULAR_NO';

#         # Parse the input date string into a DateTime object 
#         $convertedDate = ConvertDate -uploadDate $record.'CIR DATE';
            
#         $date = $null;
#         # Format the DateTime object into a string suitable for SharePoint list item
#         $sharePointDateTime = $null;
            
#         if ($null -ne $convertedDate) {
                
#             $date = Get-Date ($convertedDate) -Format "s" -ErrorAction SilentlyContinue ;
#             $sharePointDateTime = $date + "Z";
#         }

#         #Update List Items - Map with Internal Names of the Fields!.Below you define a custom object @ {}for each record and update the SharePoint list
#         if ($null -ne $currentItem) {

#             try {
#                 Write-Host "Updating record for $($record.'CIRCULAR_NO')" -ForegroundColor Green;
                
#                 $ErrorActionPreference = "Stop";    
#                 Set-PnPListItem -List $listName -Identity $currentItem -Values @{
#                     "Subject"            = $($record.'SUBJECT');
#                     "MigratedDepartment" = $($record.'DEPT_NAME');
#                     "Classification"     = $circularClassification;
#                     "CircularNumber"     = $($record.'CIRCULAR_NO');
#                     "CircularStatus"     = "Published";
#                     "MigratedOriginator" = $($record.'ORIGINATOR');
#                     "MigratedIssuedFor"  = $($record.'ISSUED_FOR');
#                     #"MigratedRefNumber"  = $($record.'REF NO');
#                     "MigratedDocPath"    = $($record.'DOC_PATH');
#                     "IsMigrated"         = "Yes"
#                     "PublishedDate"      = $sharePointDateTime;
#                     "MigratedSubFileNo"  = $($record.'SUBFILE_NO');
    
#                 } -ErrorAction SilentlyContinue;

#                 Write-Host "Updated Item" -ForegroundColor Green;
#                 $global:logs += WriteLogs -circularNumber $circularNumber -isMigrated "Yes" -status "Updated" -message "Item Updated Successfully";
                
#             }
#             catch {
#                 $global:errorLogs += WriteErrorLogs -circularNumber $circularNumber -message "$($_.Exception.Message)";
#             }
#             finally {
#                 $ErrorActionPreference = "Continue";
#             }
#         }

#         else {

#             try {

#                 $ErrorActionPreference = "Stop";
#                 Write-Host "Adding record for $($record.'CIRCULAR_NO')" -ForegroundColor Green;

#                 #Add List Items - Map with Internal Names of the Fields!.Below you define a custom object @ {}for each record and update the SharePoint list.
#                 Add-PnPListItem -List $listName -Values @{
#                     "Subject"            = $($record.'SUBJECT');
#                     "MigratedDepartment" = $($record.'DEPT_NAME');
#                     "Classification"     = $circularClassification;
#                     "CircularNumber"     = $($record.'CIRCULAR_NO');
#                     "CircularStatus"     = "Published";
#                     "MigratedOriginator" = $($record.'ORIGINATOR');
#                     "MigratedIssuedFor"  = $($record.'ISSUED_FOR');
#                     #"MigratedRefNumber"  = $($record.'REF NO');
#                     "MigratedDocPath"    = $($record.'DOC_PATH');
#                     "IsMigrated"         = "Yes"
#                     "PublishedDate"      = $sharePointDateTime;
#                     "MigratedSubFileNo"  = $($record.'SUBFILE_NO');

#                 } -ErrorAction SilentlyContinue;

#                 Write-Host "Items Added" -ForegroundColor Green;
#                 $global:logs += WriteLogs -circularNumber $circularNumber -isMigrated "Yes" -status "Created" -message "Item Created Successfully";
#             }
#             catch {
#                 $global:errorLogs += WriteErrorLogs -circularNumber $circularNumber -message "$($_.Exception.Message)";
#             }

#             finally {
#                 $ErrorActionPreference = "Continue";
#             }
#         }

#         Write-Progress -PercentComplete (($i / $circularData.Count) * 100) -Status "Migrating circulars to Circular Repository SharePoint List" -Activity "Item $($i) of $($circularData.Count)"
#         $i++; 
#     }
# }

Function GetCircularListItem($circularNumber) {
    $listItem = $null;
    try {
       
        $listItem = Get-PnPListItem  -List $listName -Query "@<View>
        <ViewFields>
            <FieldRef Name='Title'/>
            <FieldRef Name='CircularNumber'/>                                                                                                                                       
        </ViewFields>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='CircularNumber'/>
                        <Value Type='Text'>$($circularNumber)</Value>
                 </Eq>
            </Where>
        </Query>
    </View>";

        if ($null -eq $listItem) {
            $listItem = $null;
            
        }

    }
    catch {
        $global:errorLogs += WriteErrorLogs -circularNumber "$($record.'CIRCULAR NO')" -message "$($_.Exception.Message)"; ;
    }

    return $listItem;
}



if ($circularClassification -eq "Master") {
    MasterCircularMigration;
    Write-Host "Master Circular migration completed" -ForegroundColor Green
}
# if ($circularClassification -eq "Circular") {
#     CircularMigration;
#     Write-Host "Circular migration completed" -ForegroundColor Green
# }

if ($global:logs.Count -gt 0) {
    $logPath = $OutLogPath + "$($circularClassification)_Logs_$((Get-Date -Format "s").Replace(':','')).csv";
    $global:logs | Select-Object TimeStamp, CircularNumber, IsMigrated, Status, Message, Classification | Export-Csv -Path $logPath  -Encoding UTF8 
}

if ($global:errorLogs.Count -gt 0) {
    $errorLogPath = $ErrorOutLogPath + "$($circularClassification)_ErrorLogs_$((Get-Date -Format "s").Replace(':','')).csv";
    $global:errorLogs | Select-Object TimeStamp, Status, Message, CircularNumber, Classification   | Export-Csv -Path $errorLogPath -Encoding UTF8 
}
