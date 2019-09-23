param(
    $Extension = ".jpg",
    $ExtFilter = "*$Extension",
    $SourcePath = @("C:\Local\Photo Archiving\Photo Archiver\Photo Archiver\Archive Photos (New)\Source Photos", "C:\Local\Photo Archiving\Photo Archiver\Photo Archiver\Archive Photos (New)\Source Photos Additional"),
    $RootArchiveFolder = "C:\Local\Test\Archive Images New",
    $UndatedArchiveFolder = "$RootArchiveFolder\Undated",
    $ArchiveFolderPattern = "",
    $ArchiveFilePattern = "",
    [switch]$ByCamera
)

$Script:FilesCopied = 0
$Script:FilesRenamed = 0
$Script:FilesSkipped = 0

$DupNameIndex = 0

######
$SourceImagesWithMetadata = @()
######

#region Functions
function Get-xDateTimeString {
    Param (
        [string]$Pattern = "%d/%m/%Y (%H:%M:%S)",
        $DateTime = $false
    )
    if ($DateTime -eq $false) { Return (get-date -uformat $Pattern) }
    else { Return (Get-Date $DateTime -uformat $Pattern) }
}
function Test-FilesAreEqual {
    param (
        [System.IO.FileInfo]$First,
        [System.IO.FileInfo]$Second
    )

    $BYTES_TO_READ = 65536 #32768

    if ($First.Length -ne $Second.Length) {
        Return $false
    }

    $Iterations = [System.Math]::Ceiling($First.Length / $BYTES_TO_READ)

    $File1 = $First.OpenRead()
    $File2 = $Second.OpenRead()

    $one = New-Object byte[] $BYTES_TO_READ
    $two = New-Object byte[] $BYTES_TO_READ

    for ($i = 0; $i -lt $Iterations; $i++) {
        $File1.Read($one, 0, $BYTES_TO_READ) | Out-Null
        $File2.Read($two, 0, $BYTES_TO_READ) | Out-Null

        for ($x = 0; $x -lt $BYTES_TO_READ; $x += 8) {
            if ([System.BitConverter]::ToInt64($one, $x) -ne [System.BitConverter]::ToInt64($two, $x)) {
                $File1.Close()
                $File2.Close()
                Return $false
            }
        }
    }
    
    $File1.Close()
    $File2.Close()

    Return $true

}
function Get-xStringNumber {
    Param ($String)
	
    $CharacterHash = @{
        "a" = 1;  "b" = 2;  "c" = 3;  "d" = 4;  "e" = 5
        "f" = 6;  "g" = 7;  "h" = 8;  "i" = 9;  "j" = 10
        "k" = 11; "l" = 12; "m" = 13; "n" = 14; "o" = 15
        "p" = 16; "q" = 17; "r" = 18; "s" = 19; "t" = 20
        "u" = 21; "v" = 22; "w" = 23; "x" = 24; "y" = 25
        "z" = 26
        " " = 0
        ""  = 0
        "0" = 0
        "1" = 101
        "2" = 102
        "3" = 103
        "4" = 104
        "5" = 105
        "6" = 106
        "7" = 107
        "8" = 108
        "9" = 109
    }
	
    [int]$StringNumber = 0
    if ($null -ne $String) {
        if($String -ne ""){
            $CharacterArray = $String.ToCharArray()
            foreach ($Character in $CharacterArray) {
                $StringNumber += [int]($CharacterHash[[string]$Character])
            }
        }
    }
    Return $StringNumber
}
function Get-xImageMetadataHashtable {
    param (
        $SourceImage = $CurrentFileObject,
        $RootArchiveFolder,
        $UndatedArchiveFolder,
        [switch]$ByCamera
    )
	
    $Continue = $false
	
    try {
        $ImageStream = New-Object System.IO.FileStream($SourceImage.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::ReadWrite) -ErrorAction Stop
        $Decoder = New-Object System.Windows.Media.Imaging.JpegBitmapDecoder($ImageStream, [Windows.Media.Imaging.BitmapCreateOptions]::None, [Windows.Media.Imaging.BitmapCacheOption]::None) -ErrorAction Stop
        $Continue = $true
    } catch {
        if ($ImageStream) { $ImageStream.Dispose() }
        $Continue = $false
    }
	
    if ($Continue -eq $true) {
        $Metadata = $Decoder.Frames[0].Metadata
        Remove-Variable -Name Decoder
		
        $DateTimeOriginal = [String]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=36867}")
        If ($DateTimeOriginal -lt 1) { $DateTimeOriginal = $false }
        if ([string]::IsNullOrWhiteSpace($DateTimeOriginal)){ $DateTimeOriginal = $false }
		
        $SubSecTimeOriginal = [String]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37521}")
        try { $Title = [string]$Metadata.GetQuery("/xmp/dc:title/x-default") }
        catch { $Title = "" }
		
        $CameraFocalLength = [string]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37386}")
        $CameraUserComment = [string]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37510}")
        $ImageDescription = [string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=270}")
		
        $CameraMake = [string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=271}")
        $CameraModel = ([string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=272}")).TrimEnd()
        if ($CameraModel.length -lt 1) { $CameraModel = "Unknown" }
		
        $ImageStream.Dispose()
        $Revision = "{0:D2}" -f 0 # "00"  
        $Missing = "{0:D3}" -f 0
        $Separator = "-"
        $Extension = $SourceImage.Extension
		
        # Generate CameraID
        $CameraIDMake = Get-xStringNumber -String $CameraMake
        $CameraIDModel = Get-xStringNumber -String $CameraModel
        $CameraIDUserComment = Get-xStringNumber -String $CameraUserComment # 0
	
        $CameraID = "{0:D3}" -f ($CameraIDMake + $CameraIDModel + $CameraIDUserComment)

		
        if ($DateTimeOriginal -ne $false) {
            # false may be sufficient
            $Year = $DateTimeOriginal.Substring(0, 4)
            $Month = $DateTimeOriginal.Substring(5, 2)
            $MonthName = (get-date -month $Month -format MMM)
            $Day = $DateTimeOriginal.Substring(8, 2)
            $DayName = (get-date -year $Year -month $Month -day $Day).dayofweek
            $MonthFolderName = ($Year + "-" + $Month + "-" + (get-date -month $Month -format MMMM))
			
            # Convert Text milliseconds to Integer
            If ([string]::IsNullOrWhiteSpace($SubSecTimeOriginal) -eq $false) {
                $SubSecondString = "{0:D3}" -f [int]([float]("." + $SubSecTimeOriginal) * 1000)
            }
            Else {
                $SubSecondString = $Missing
            }
			
            $DateString = "$Year.$Month.$Day"
            $TimeString = ($DateTimeOriginal.Substring(11, 8)).Replace(":", "") + "." + $SubSecondString
						
            if ($ByCamera -eq $true) {
                $ArchiveFolder = $RootArchiveFolder + "\" + $CameraModel + "\" + $Year + "\" + $MonthFolderName
            } else {
                $ArchiveFolder = $RootArchiveFolder + "\" + $Year + "\" + $MonthFolderName
            }
			
            $ArchiveBaseName = $DateString + $Separator + $TimeString + $Separator + $CameraID + $Separator
            $ArchiveBaseNameShort = $DateString + $Separator + $TimeString + $Separator
            $ArchiveName = $DateString + $Separator + $TimeString + $Separator + $CameraID + $Separator + $Revision + $Extension
            $ArchiveBasePath = $ArchiveFolder + "\" + $DateString + $Separator + $TimeString
			
        } else {
            if ($ByCamera -eq $true) {
                $ArchiveFolder = $UndatedArchiveFolder + "\" + $CameraModel
            } else {
                $ArchiveFolder = $UndatedArchiveFolder
            }
            $ArchivedPattern = "-[0-9][0-9][0-9]-[0-9][0-9](\.|$)" # Regex patterrn for *-000-00.*
            if ($SourceImage.Name -match $ArchivedPattern) {
                $MatchLength = $Matches[0].Length
                $ArchiveBaseName = $SourceImage.Name.Substring(0, ($SourceImage.Name.Length - ($MatchLength + ($Extension.Length -1)))) + $Separator + $CameraID + $Separator
                $ArchiveBaseNameShort = $SourceImage.Name.Substring(0, ($SourceImage.Name.Length - ($MatchLength + ($Extension.Length -1)))) + $Separator
                $ArchiveName = $SourceImage.Name.Substring(0, ($SourceImage.Name.Length - ($MatchLength + ($Extension.Length -1)))) + $Separator + $CameraID + $Separator + $Revision + $Extension
                $ArchiveBasePath = $ArchiveFolder + "\" + $SourceImage.Name.Substring(0, ($SourceImage.Name.Length - ($MatchLength + ($Extension.Length -1)))) # + $Separator #+ $CameraID + $Separator
            } else {
                $ArchiveBaseName = $SourceImage.BaseName + $Separator + $CameraID + $Separator
                $ArchiveBaseNameShort = $SourceImage.BaseName + $Separator
                $ArchiveName = $SourceImage.BaseName + $Separator + $CameraID + $Separator + $Revision + $Extension
                $ArchiveBasePath = $ArchiveFolder + "\" + $SourceImage.BaseName
            }
        }
        # Create Hashtable
        $Hashtable = [ordered]@{
            "DateTimeOriginal"     = $DateTimeOriginal
            "SubSecTimeOriginal"   = $SubSecTimeOriginal
            "Model"                = $CameraModel
            "Make"                 = $CameraMake
            "Title"                = $Title
            "ImageDescription"     = $ImageDescription
            "ImageDescriptionNew"  = $false
            "UserComment"          = $CameraUserComment
            "FocalLength"          = $CameraFocalLength
            "ArchiveBaseName"      = $ArchiveBaseName
            "ArchiveBaseNameShort" = $ArchiveBaseNameShort
            "ArchiveBasePath"      = $ArchiveBasePath
            "ArchiveFolder"        = $ArchiveFolder
            "ArchiveAction"        = $false
            "ArchiveName"          = $ArchiveName
            "CameraID"             = $CameraID
            "Revision"             = $Revision
            "NoMetadata"           = $false
        }
    } else {
        $Hashtable = [ordered]@{
            "DateTimeOriginal"     = $false
            "SubSecTimeOriginal"   = $false
            "Model"                = $false
            "Make"                 = $false
            "Title"                = $false
            "ImageDescription"     = $false
            "ImageDescriptionNew"  = $false
            "UserComment"          = $false
            "FocalLength"          = $false
            "ArchiveBaseName"      = $false
            "ArchiveBaseNameShort" = $false
            "ArchiveBasePath"      = $false
            "ArchiveFolder"        = $false
            "ArchiveAction"        = $false
            "ArchiveName"          = $false
            "CameraID"             = $false
            "Revision"             = $false
            "NoMetadata"           = $true
        }
    }
    Return $Hashtable
}
function Rename-xFiles {
    param($Files)

    foreach ($File in $Files) {
        if ((Test-Path -path $File.FullArchivePath) -eq $false) {
            #Write-Host "Renaming $($File.FullName) to $($File.FullArchivePath)" -ForegroundColor Green
            Write-Host "R" -ForegroundColor Yellow -NoNewline
            rename-item -LiteralPath $File.FullName -NewName $File.PhotoMetadata.ArchiveName
            $File.PhotoMetadata.ArchiveAction = $false
            $Script:FilesRenamed++
        }
        if ($File.FullArchivePath -eq $File.FullName) {
            #Write-Host "New and old name match - No action required $($File.FullName)" -ForegroundColor Cyan
            $File.PhotoMetadata.ArchiveAction = $false
        }
    }

    $RemainingFiles = @($Files | ? { $_.PhotoMetadata.ArchiveAction -eq $true })
    #Write-Host "Remaining Files $RemainingFiles" | FT -auto
    #Write-Host "Archive Actions $($Files.PhotoMetadata.ArchiveAction)"

    if (($null -eq $RemainingFiles) -or ($RemainingFiles.Count -eq 0)) {
        #Write-Host "All files renamed successfully" -ForegroundColor Green
        return $null
    }
    else {
        #Write-Host "Unable to rename some files - Initiating Recursion ($($RemainingFiles.Count) remaining files)"
        Write-Host "^" -ForegroundColor Magenta -NoNewline
        Rename-xFiles -Files $RemainingFiles
    }

}

#endregion Functions

$ScriptStart = Get-Date
$RunTime = Get-xDateTimeString -DateTime $ScriptStart -Pattern "%Y%m%d-%H.%M.%S"
$swThresh = 1000

# Prerequisites
[System.Reflection.Assembly]::LoadWithPartialName("PresentationCore") | Out-Null

#Find All Source Photos (JPG)
Write-Host "Searching Source Path/s for Images (.jpg)..." -ForegroundColor Yellow
$SourceImagePaths = ($SourcePath | %{cmd /c dir "$_\*.jpg" /b /s /a-d-h-s})

Write-Host "$($SourceImagePaths.Count) Images found." -ForegroundColor Green

# Progress Bar Setup
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$i = 1;$t = $SourceImagePaths.Count / 100; $tc = $SourceImagePaths.Count

foreach ($SourceImagePath in $SourceImagePaths) {

    # Get Source JPG metadata
    $CurrentFileObject = (Get-Item -LiteralPath $SourceImagePath | Select-Object *, 
        @{Label = "PhotoMetaData"; Exp = { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder } },
        @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
        @{Label = "FullArchivePath"; Exp = { "" } }
    )

    #####
    $SourceImagesWithMetadata += $CurrentFileObject
    #####

    # Check Destination for existing photos with same primary path (sans version info)
    if (Test-Path -LiteralPath $CurrentFileObject.PhotoMetaData.ArchiveFolder) {
    
        $SearchFilter = $CurrentFileObject.PhotoMetaData.ArchiveBasePath + "*"
        try { $MatchingDestinationFiles = Get-ChildItem $SearchFilter }
        catch { $MatchingDestinationFiles = $null }
    
        if ($MatchingDestinationFiles) {
            # Perform binary compare to look for duplicates
            $DuplicateFileFound = $false
            foreach ($MatchingDestinationFile in $MatchingDestinationFiles) {
                if ((Test-FilesAreEqual -First $CurrentFileObject.FullName -Second $MatchingDestinationFile.FullName) -eq $true) {
                    $DuplicateFileFound = $true
                    $Script:FilesSkipped++
                    #Write-Host "Duplicate Found $($CurrentFileObject.FullName) = $($MatchingDestinationFile.FullName)" -ForegroundColor Cyan
                    #Write-Host "." -ForegroundColor White -NoNewline
                    Break
                    # Duplicate found in destination folder - do not copy
                }
            }
            if ($DuplicateFileFound -eq $false) {
            
                # Get Metadata from matching destination photos
                $MatchingDestinationFiles = ($MatchingDestinationFiles | Select-Object *, 
                    @{Label = "PhotoMetaData"; Exp = { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder } },
                    @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
                    @{Label = "FullArchivePath"; Exp = { "" } }
                )
            
                $RenameMatchArrayGrouped = @($MatchingDestinationFiles | % { $_ }; $CurrentFileObject) | Sort LastWriteTime | Group LastWriteTime
                $BaseRevisionNum = 0
                $SubRevisionArray = ("a".."z")
                foreach ($PhotoGroup in $RenameMatchArrayGrouped) {
                    $BaseRevision = "{0:D2}" -f $BaseRevisionNum
                    if ($PhotoGroup.Count -gt 1) {
                        $SubRevisionIndex = 0
                        foreach ($Photo in $PhotoGroup.Group) {
                            $Photo.Revision = "$BaseRevision-$($SubRevisionArray[$SubRevisionIndex])"
                            $Photo.FullArchivePath = $Photo.PhotoMetaData.ArchiveBasePath + "-" + $Photo.PhotoMetaData.CameraID + "-" + $Photo.Revision + $Photo.Extension
                            $Photo.PhotoMetaData.ArchiveName = $Photo.PhotoMetaData.ArchiveBaseName + $Photo.Revision + $Photo.Extension
                            $Photo.PhotoMetaData.ArchiveAction = $true
                            $SubRevisionIndex++
                        }
                    }
                    else {
                        foreach ($Photo in $PhotoGroup.Group) {
                            $Photo.Revision = $BaseRevision
                            $Photo.FullArchivePath = $Photo.PhotoMetaData.ArchiveBasePath + "-" + $Photo.PhotoMetaData.CameraID + "-" + $Photo.Revision + $Photo.Extension
                            $Photo.PhotoMetaData.ArchiveName = $Photo.PhotoMetaData.ArchiveBaseName + $Photo.Revision + $Photo.Extension
                            $Photo.PhotoMetaData.ArchiveAction = $true
                        }
                    }
                    $BaseRevisionNum++
                }

                $RenameMatchArray = $RenameMatchArrayGrouped.Group | ? { $_.FullName -ne $CurrentFileObject.FullName }

                Rename-xFiles -Files $RenameMatchArray
                #Write-Host "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" -ForegroundColor Yellow
                Write-Host "+" -ForegroundColor Green -NoNewline
                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                $Script:FilesCopied++
            
            }
        }
        else {
            #Copy File
            try {
                #Write-Host "Copy Photo(Existing Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName -ForegroundColor Cyan
                #Write-Host "Copying new File to Archive (Existing Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" -ForegroundColor Yellow
                Write-Host "A" -ForegroundColor Green -NoNewline
                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination ($CurrentFileObject.PhotoMetaData.ArchiveFolder + "\" + $CurrentFileObject.PhotoMetaData.ArchiveName)
                $Script:FilesCopied++
            }
            catch { } # Error copying file to Destination
        }
    }
    else {
        New-Item -Path $CurrentFileObject.PhotoMetaData.ArchiveFolder -ItemType Directory | Out-Null
        #Write-Host "Copy Photo(New Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName
        #Write-Host "Copying new File to Archive (New Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" -ForegroundColor Yellow
        Write-Host "+" -ForegroundColor Green -NoNewline
        Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination ($CurrentFileObject.PhotoMetaData.ArchiveFolder + "\" + $CurrentFileObject.PhotoMetaData.ArchiveName)
        $Script:FilesCopied++
    }

    if($sw.Elapsed.TotalMilliseconds -ge $swThresh){Write-Progress -Activity "Archiving Files" -PercentComplete ([int]$i/$t) -Status "$i of $tc";$sw.Reset();$sw.Start()};$i++

}

Write-Progress -Activity "Archiving Files" -Completed

$ScriptEnd = Get-Date

$ScriptDuration = $ScriptEnd - $ScriptStart

Write-Host
Write-Host

Write-Host "Script Duration`:`t$($ScriptDuration.TotalSeconds) Seconds" -ForegroundColor Cyan

Write-Host "Files Copied  = $($Script:FilesCopied)"  -ForegroundColor Yellow
Write-Host "Files Renamed = $($Script:FilesRenamed)" -ForegroundColor Yellow
Write-Host "Files Skipped = $($Script:FilesSkipped)" -ForegroundColor Yellow