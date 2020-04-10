# Credit for this script goes out to following authors from whose scripts I have borrowed components
# Pat Richard (http://www.ehloworld.com/) who wrote New-FileDownload function
# Tom Arbuthnot (http://lyncdup.com) who wrote Lync session download script, giving me idea for this script
# Falak Mahmood (http://falakmahmood.blogspot.se/) who wrote SharePoint Conference 2014 download script

Function Global:Convert-ToFriendlyName
{
    param ($Text)
    # Unwanted characters (includes spaces and '-') converted to a regex:
    $SpecChars = '!', '£', '$', '%', '&', '^', '*', '(', ')', '@', '=', '+', '¬', '`', '\', '<', '>', '.', '?', '/', ':', ';', '#', '~', "'", '-', '"', '|', '&#39;s'
    $remspecchars = [string]::join('|', ($SpecChars | % {[regex]::escape($_)}))

    # Convert the text given to correct naming format (Uppercase)
    $name = (Get-Culture).textinfo.totitlecase("$Text".tolower())

    # Remove unwanted characters
    $name = $name -replace $remspecchars, " "
    $name
}

function Global:Set-ModuleStatus { 
	[CmdletBinding(SupportsShouldProcess = $True)]
	param	(
		[parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No module name specified!")] 
		[string]$name
	)
	if(!(Get-Module -name "$name")) { 
		if(Get-Module -ListAvailable | ? {$_.name -eq "$name"}) { 
			Import-Module -Name "$name" 
			# module was imported
			return $true
		} else {
			# module was not available
			return $false
		}
	}else {
		# module was already imported
		# Write-Host "$name module already imported"
		# return $true
	}
} # end function Set-ModuleStatus

function Global:New-FileDownload {
	[CmdletBinding(SupportsShouldProcess = $True)]
	param(
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[ValidateNotNullOrEmpty()]
		[string]$SourceFile,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFolder,
		[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true, Mandatory = $false)] 
		[string]$DestFile
	)
	
    [bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)

	if (!($DestFolder)){$DestFolder = $TargetFolder}
	# Write-Host "Checking for BitsModule"
    Set-ModuleStatus -name BitsTransfer
	
    if (!($DestFile)){[string] $DestFile = $SourceFile.Substring($SourceFile.LastIndexOf("/") + 1)}
	if (Test-Path $DestFolder){
		Write-Verbose "Folder: `"$DestFolder`" exists."
	} else{
		Write-Host "Folder: `"$DestFolder`" does not exist, creating..."
		New-Item $DestFolder -type Directory | Out-Null
		Write-Host "Done! " -ForegroundColor Green
	}
	if (Test-Path "$DestFolder\$DestFile"){
		Write-Host -ForegroundColor Yellow "File: $DestFile already exists."
        #write finish result to global
        $Global:NewFileDownloadResult = $?
	}else{
		if ($HasInternetAccess){
			Write-Host "File: $DestFile does not exist, downloading..." 
			
            Try {
                # Forcing the error output to  a custom variable, as it was the only way to catch the non-terminating error

                # clear down error
                $bitserror = $null
                Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction Continue
                # Write-Host "Done! " -ForegroundColor Green

                # This sends the result variable of the last command run to the global scope
                # $? is true if command ran successfully and false if it didn't
                
                # Show-ErrorDetails $bitserror
                
                # Write if this was successful or not to a global variable
                $Global:NewFileDownloadResult = $?
                $Global:NewFileDownloadError = $bitserror
                
                $bitserror

                If ($BitsError)
                            {
                                # loop three times
                                    While ($loop -lt "2" -and $Global:NewFileDownloadResult -eq $false) 
                                    {
                                    $loop = $loop + 1
                                    Write-Host "Download Retry attempt $loop of 2"
                                    $bitserror = $null
                                    Write-Host "Trying for Session: $url"
                                    Write-Host "Trying URL: $SourceFile"
                                    Start-BitsTransfer -Source "$SourceFile" -Destination "$DestFolder\$DestFile" -RetryInterval 60 -RetryTimeout 600 -ErrorVariable BitsError -ErrorAction Continue 
                                    $Global:NewFileDownloadResult = $?
                                    Start-Sleep -Seconds 5
                                    
                                    
                                    
                                    }
                             

                            } # Close If error 400


                }
	       catch
                {
                Write-Host "Hit Generic Catch on New-FileDownload"
                Write-Host $bitserror
                
                }
		


			
		}else{
			Write-Host "Internet access not detected. Please resolve and try again." -ForegroundColor red
		}
	}
} # end function New-FileDownload

function Global:Download-MECContent
{
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Filetype can be either All, Slides or Recordings
        [Parameter(Mandatory=$true)]
        [ValidateSet("All", "Slides", "Recordings")]
        $FileType,

        # Location to store downloaded files to. Script only tested against local storage.
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $TargetFolder,

        # Keynote recording is 3GB, no need to download unless explicitly specified.
        [switch]$Keynote
    )

    if (-not (Test-Path $TargetFolder)) 
    {
        Write-Host "Folder does not exist. Creating folder..."
        New-Item $TargetFolder -type directory
    }
    
    cd $TargetFolder

    If ($FileType -eq "All" -or $FileType -eq "Slides")
    {
        $pptx = ([xml](new-object net.webclient).downloadstring("http://channel9.msdn.com/Events/MEC/2014/RSS/slides"))
    
        foreach ($i in $pptx.rss.channel.item)
        {  
            $url = New-Object System.Uri($i.enclosure.url)
            $file = $url.Segments[-1]
	        $file = Convert-ToFriendlyName $i.title
            $file = $file.TrimEnd(" ")
	        $file = $file + ".pptx"
            
            New-FileDownload -SourceFile $url -DestFolder $TargetFolder -DestFile "$file"
        }
    }

    If ($FileType -eq "All" -or $FileType -eq "Recordings")
    {
        $mp4 = ([xml](new-object net.webclient).downloadstring("http://channel9.msdn.com/Events/MEC/2014/RSS/mp4high"))
    
        foreach ($i in $mp4.rss.channel.item)
        {  
            $url = New-Object System.Uri($i.enclosure.url)
            $file = $url.Segments[-1]
	        $file = Convert-ToFriendlyName $i.title
            $file = $file.TrimEnd(" ")
	        $file = $file + ".mp4"
            
            if ($file -match "Keynote")
            {
                if ($Keynote)
                {
                    New-FileDownload -SourceFile $url -DestFolder $TargetFolder -DestFile "$file"
                }
            }
            else
            {
                    New-FileDownload -SourceFile $url -DestFolder $TargetFolder -DestFile "$file"
            }
        }
    }
}


Write-Host "Ready to download. Just run one of the following to start the download..."  -ForegroundColor Green
Write-Host "Replace C:\MEC2014 with folder of your choice."  -ForegroundColor Green
Write-Host "Keynote recording is almost 3GB so you must specify -Keynote if you want to download it, see last example..."  -ForegroundColor Green
Write-Host "   Download-MECContent -Filetype All -TargetFolder C:\MEC2014" -ForegroundColor Yellow
Write-Host "   Download-MECContent -Filetype Slides -TargetFolder C:\MEC2014" -ForegroundColor Yellow
Write-Host "   Download-MECContent -Filetype Recordings -TargetFolder C:\MEC2014" -ForegroundColor Yellow
Write-Host "   Download-MECContent -Filetype Recordings -TargetFolder C:\MEC2014 -KeyNote" -ForegroundColor Yellow