Add-Type -AssemblyName System.Drawing

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
function Get-xCharacterNumber {
    param ([string]$Character)
    
    $Character = ($Character.ToCharArray())[0]
    $CharacterHash = @{
        "a" = 1
        "b" = 2
        "c" = 3
        "d" = 4
        "e" = 5
        "f" = 6
        "g" = 7
        "h" = 8
        "i" = 9
        "j" = 10
        "k" = 11
        "l" = 12
        "m" = 13
        "n" = 14
        "o" = 15
        "p" = 16
        "q" = 17
        "r" = 18
        "s" = 19
        "t" = 20
        "u" = 21
        "v" = 22
        "w" = 23
        "x" = 24
        "y" = 25
        "z" = 26
        " " = 0
        "0" = 0
        "1" = 1
        "2" = 2
        "3" = 3
        "4" = 4
        "5" = 5
        "6" = 6
        "7" = 7
        "8" = 8
        "9" = 9
    }
    $Number = $CharacterHash[$Character]
    Return $Number
}
function Get-xStringNumber {
    Param ($String)
    
    $CharacterHash = @{
        "a" = 1
        "b" = 2
        "c" = 3
        "d" = 4
        "e" = 5
        "f" = 6
        "g" = 7
        "h" = 8
        "i" = 9
        "j" = 10
        "k" = 11
        "l" = 12
        "m" = 13
        "n" = 14
        "o" = 15
        "p" = 16
        "q" = 17
        "r" = 18
        "s" = 19
        "t" = 20
        "u" = 21
        "v" = 22
        "w" = 23
        "x" = 24
        "y" = 25
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
function Get-xDateTimeString {
    Param (
        [string]$Pattern = "%d/%m/%Y (%H:%M:%S)",
        $DateTime = $false
    )
    if ($DateTime -eq $false) { Return (get-date -uformat $Pattern) }
    else { Return (Get-Date $DateTime -uformat $Pattern) }
}
function Get-xCameraID_Old {
    Param ($SourceImage)
    
    $CameraIDMake = Get-xCharacterNumber -Character $SourceImage.Make
    $CameraIDModel = Get-xStringNumber -String $SourceImage.Model
    #$CameraIDUserComment = Get-xStringNumber -String $SourceImage.UserComment
    
    #if ($SourceImage.Make -like "*apple*") { $CameraIDFocalLength = Get-xStringNumber -String $SourceImage.FocalLength }
    #else { $CameraIDFocalLength = 0 }
    
    #Return $CameraIDMake + $CameraIDModel + $CameraIDFocalLength + $CameraIDUserComment
    Return $CameraIDMake + $CameraIDModel
}
function Get-xCameraID {
    Param ($SourceImage)
    
    $CameraIDMake = Get-xStringNumber -String $SourceImage.Make
    $CameraIDModel = Get-xStringNumber -String $SourceImage.Model
    $CameraIDUserComment = Get-xStringNumber -String $SourceImage.UserComment
    
    Return $CameraIDMake + $CameraIDModel + $CameraIDUserComment
}
function Get-xDateTimeOriginal {
    param (
        $SourceImage = $FilePaths[0]
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

        $ImageStream.Dispose()

    }

    Return $DateTimeOriginal
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
        $CameraIDUserComment = 0 #Get-xStringNumber -String $CameraUserComment
    
        $CameraID = "{0:D3}" -f ($CameraIDMake + $CameraIDModel + $CameraIDUserComment)

        
        if ([string]::IsNullOrWhiteSpace($DateTimeOriginal) -eq $false) {
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
            $ArchivedPattern = "[0-9][0-9][0-9]-[0-9][0-9]." # Regex patterrn for *-000-0-00.*
            if ($SourceImage.BaseName -match $ArchivedPattern) {
                $ArchiveBaseName = $SourceImage.BaseName.Substring(0, ($SourceImage.BaseName.Length - (9 + $Extension.Length))) + $Separator + $CameraID + $Separator
                $ArchiveBaseNameShort = $SourceImage.BaseName.Substring(0, ($SourceImage.BaseName.Length - (9 + $Extension.Length))) + $Separator
                $ArchiveName = $SourceImage.BaseName.Substring(0, ($SourceImage.BaseName.Length - (9 + $Extension.Length))) + $Separator + $CameraID + $Separator + $Revision + $Extension
                $ArchiveBasePath = $ArchiveFolder + "\" + $SourceImage.BaseName.Substring(0, ($SourceImage.BaseName.Length - (9 + $Extension.Length))) #+ $Separator #+ $CameraID + $Separator
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
function PSUsing {
    param(
        [IDisposable] $disposable,
        [ScriptBlock] $scriptBlock
    )
 
    try {
        & $scriptBlock
    }
    finally {
        if ($null -ne $disposable) {
            $disposable.Dispose()
        }
    }
}
function Get-ExifProperty {
    param
    (
        [string] $ImagePath,
        [int] $ExifTagCode
    )

    $fullPath = (Resolve-Path $ImagePath).Path

    PSUsing ($fs = [System.IO.File]::OpenRead($fullPath)) {
        PSUsing ($image = [System.Drawing.Image]::FromStream($fs, $false, $false)) {
            if (-not $image.PropertyIdList.Contains($ExifTagCode)) {
                return $null
            }

            $propertyItem = $image.GetPropertyItem($ExifTagCode)
            $valueBytes = $propertyItem.Value
            $value = [System.Text.Encoding]::ASCII.GetString($valueBytes) -replace "`0$"
            return $value
        }
    }
}
function Get-TagName {
    Param($TagCode)

    switch($TagCode){
        36867    {$TagName = "DateTimeOriginal"}
        270      {$TagName = "ImageDescription"}
        37521    {$TagName = "SubSecTimeOriginal"}
        271      {$TagName = "Make"}
        272      {$TagName = "Model"}
        305      {$TagName = "Software"}
        315      {$TagName = "Artist"}
        33432    {$TagName = "Copyright"}
        1        {$TagName = "GPSLatRef"}
        3        {$TagName = "GPSLongRef"}
        5        {$TagName = "GPSAltRef"}
        40091    {$TagName = "Title"}
        40093    {$TagName = "Author"}
        40094    {$TagName = "Keywords"}
        40095    {$TagName = "Subject"}
        37510    {$TagName = "UserComment"}
        40961    {$TagName = "ColorSpace"}
        41992    {$TagName = "Contrast"}
        41986    {$TagName = "ExposureMode"}
        34850    {$TagName = "ExposureProgram"}
        37385    {$TagName = "Flash"}
        37384    {$TagName = "LightSource"}
        37383    {$TagName = "MeteringMode"}
        41993    {$TagName = "Saturation"}
        41990    {$TagName = "SceneCaptutreMode"}
        41994    {$TagName = "Sharpness"}
        41996    {$TagName = "SubjectRange"}
        41987    {$TagName = "WhiteBalance"}
        33434    {$TagName = "Exposuretime"}
        2        {$TagName = "GPSLattitude"}
        4        {$TagName = "GPSLongitude"}
        41988    {$TagName = "DigitalZoomRatio"}
        37380    {$TagName = "Expbias"}
        33437    {$TagName = "FNumber"}
        37386    {$TagName = "FocalLength"}
        41989    {$TagName = "FocalLengthIn35mmFormat"}
        40963    {$TagName = "Height"}
        34855    {$TagName = "ISO"}
        37381    {$TagName = "MaxApperture"}
        40962    {$TagName = "Width"}
        6        {$TagName = "GPSAltitude"}
        274      {$TagName = "Orientation"}
        18246    {$TagName = "Rating"}
        41728    {$TagName = "FileSource"}
        37500    {$TagName = "MakerNote"}
        0        {$TagName = "GPSVer"}
        default  {$TagName = "Unknown"}
    }

    Return $TagName
}
function Get-ExifProperties {
    param([string] $ImagePath)

    $fullPath = (Resolve-Path $ImagePath).Path
    $ExifValues = New-Object System.Collections.ArrayList

    PSUsing ($fs = [System.IO.File]::OpenRead($fullPath)) {
        PSUsing ($image = [System.Drawing.Image]::FromStream($fs, $false, $false)) {
            $ExifTagCodes = $image.PropertyIdList
            foreach($ExifTagCode in $ExifTagCodes){
                #write-host "Exif Tag Code = $ExifTagCode" -ForegroundColor Green
                $valueBytes = ($image.GetPropertyItem($ExifTagCode)).Value
                if($null -ne $valueBytes){$value = [System.Text.Encoding]::ASCII.GetString($valueBytes) -replace "`0$"}
                else{$Value = "Error! Null Value"}
                $ExifValues.Add(
                    [PSCustomObject]@{
                        TagCode = $ExifTagCode
                        TagName = Get-TagName -TagCode $ExifTagCode
                        Value   = $Value
                    }
                ) | Out-Null
            }
        }
    }
    Return $ExifValues
}
