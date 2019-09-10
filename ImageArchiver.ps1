param(
    $Extension = ".jpg",
    $ExtFilter = "*$Extension",
    $SourceFolder = "C:\Local\Photo Archiving\Photo Archiver\Photo Archiver\Archive Photos (New)\Source Photos",
    #$SourceFolder = "C:\Local\Photo Archiving\Photo Archiver\Photo Archiver\Archive Photos (New)\Source Photos Additional",
    #$RootArchiveFolder = "C:\Local\Test\Archive Images testX",
    $RootArchiveFolder = "C:\Local\Test\Archive Images New",
    $UndatedArchiveFolder = "$RootArchiveFolder\Undated"
)

$Script:FilesCopied = 0
$Script:FilesRenamed = 0
$Script:FilesSkipped = 0

$WorkingFolder = "C:\Local\Photo Archiving\Photo Archiver\Photo Archiver\Archive Photos (New)" ### To update for all runtimes

# Functions
. "$WorkingFolder\ImageArchiverFunctions.ps1"

#$ScriptStart = Get-Date
#$RunTime = Get-xDateTimeString -DateTime $ScriptStart -Pattern "%Y%m%d-%H.%M.%S"
#$swThresh = 1000

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

    #$RemainingFiles = $Files.PhotoMetadata.ArchiveAction | ? { $_ -eq $true }
    $RemainingFiles = @($Files | ? { $_.PhotoMetadata.ArchiveAction -eq $true })
    Write-Host "Remaining Files $RemainingFiles" | FT -auto
    Write-Host "Archive Actions $($Files.PhotoMetadata.ArchiveAction)"

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

# Prerequisites
[System.Reflection.Assembly]::LoadWithPartialName("PresentationCore") | Out-Null

#Find All Source Photos (JPG)
$SourceImagePaths = (cmd /c dir "$SourceFolder\*.jpg" /b /s /a-d-h-s)

Write-Host "$($SourceImagePaths.Count) Images found."

#$sw = [System.Diagnostics.Stopwatch]::StartNew()
#$i = 1;$t = $SourceImagePaths.Count / 100; $tc = $SourceImagePaths.Count
foreach ($SourceImagePath in $SourceImagePaths) {

    # Get Source JPG metadata
    $CurrentFileObject = (Get-Item -LiteralPath $SourceImagePath | Select-Object *, 
        @{Label = "PhotoMetaData"; Exp = { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder } },
        @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
        @{Label = "FullArchivePath"; Exp = {""} }
    )

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
                    @{Label = "FullArchivePath"; Exp = {""} }
                )
            
                $RenameMatchArray = @($MatchingDestinationFiles | % { $_ }; $CurrentFileObject) | Sort LastWriteTime # Assending
                $rev = 0
                foreach ($Photo in $RenameMatchArray) {
                    $Photo.Revision = "{0:D2}" -f $rev
                    $Photo.FullArchivePath = $Photo.PhotoMetaData.ArchiveBasePath + "-" + $Photo.PhotoMetaData.CameraID + "-" + $Photo.Revision + $Photo.Extension
                    $Photo.PhotoMetaData.ArchiveName = $Photo.PhotoMetaData.ArchiveBaseName + $Photo.Revision + $Photo.Extension
                    $Photo.PhotoMetaData.ArchiveAction = $true
                    $rev++
                }

                $RenameMatchArray = $RenameMatchArray | ?{$_.FullName -ne $CurrentFileObject.FullName}

                #$RenameMatchArray | Select Name, LastWriteTime, Revision | ft
                ######

                Rename-xFiles -Files $RenameMatchArray
                #Write-Host "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" -ForegroundColor Yellow
                Write-Host "{+}" -ForegroundColor Green -NoNewline
                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                $Script:FilesCopied++
            
            }
        }
        else {
            #Copy File
            try {
                #Write-Host "Copy Photo(Existing Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName -ForegroundColor Cyan
                #Write-Host "Copying new File to Archive (Existing Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" -ForegroundColor Yellow
                Write-Host "(+)" -ForegroundColor Green -NoNewline
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

    #if($sw.Elapsed.TotalMilliseconds -ge $swThresh){Write-Progress -Activity "Archiving Files" -PercentComplete ([int]$i/$t) -Status "$i of $tc";$sw.Reset();$sw.Start()};$i++

}

#Write-Progress -Activity "Archiving Files" -Completed

Write-Host
Write-Host

Write-Host "Files Copied  = $($Script:FilesCopied)"  -ForegroundColor Yellow
Write-Host "Files Renamed = $($Script:FilesRenamed)" -ForegroundColor Yellow
Write-Host "Files Skipped = $($Script:FilesSkipped)" -ForegroundColor Yellow


