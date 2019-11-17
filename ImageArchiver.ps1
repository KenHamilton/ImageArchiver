<#
.NOTES
	===========================================================================
	 Created on:   	16/10/2019
	 Created by:   	Ken Hamilton
	 Organization: 	
	 Filename:     	ImageArchiver.ps1
	 Version:       0.1.0
	===========================================================================

.SYNOPSIS
		Image archiving script.

.DESCRIPTION
		This script copies images from one or more source locations to a designated archive destination while filtering out duplicates and managing file names using a convention that allows for revisions of the original image.
        The process uses existing file metadata to generate file names based on Date/Time information as well as camera attributes to provide uniqueness.
        
        An example name is as follows:

        2014.01.39-122059.200-000-0-00.jpg
          │   │  │    │    │   │  │  │  └──────────────────────────────────────────────────────────┐
          │   │  │    │    │   │  │  └────────────────────────────────────────────────────┐        │
          │   │  │    │    │   │  └──────────────────────────────────────────┐            │        │
          │   │  │    │    │   └──────────────────────────────────┐          │            │        │ 
          │   │  └──┐ └──┐ └────────────────┐                     │          │            │        │ 
        Year.Month.Day-HourMinuteSecond.SubSecondTimeOriginal-CameraID-LatestRevision-Revision.FileExtension

        Where existing metadata may be absent, such as "SubSecondTimeOriginal" a place holder of zeros is substituted "000".
        
        The CameraID is made from a combination of properties and converted into a number between 0 and 999. This property assists with resolving conflicts from differnet cameras where shots were taken at "exactly" the same time.
        This is probably rare but may be an issue where the same popular smart phone model is being used at a wedding or other event where photos are shared.

        The "LatestRevision" property indictes the most recent version/edit of an image with either a "1" or "0" - "1" being most recent.
        "Revision" is simply the chronilogical version number of the same image.
        Both of these proprties are updated where required when additional images are added to the archive.


.PARAMETER  SourceFolder
One or more Folders/Directories or Drives where images to be archived are currently stored.

.PARAMETER  RootArchiveFolder
Top level Folder where Images will be archived to.
        
.PARAMETER  UndatedArchiveFolder
Images lacking Date/Time taken metadata will be archived here.
        
.PARAMETER  LogfilePath
Folder path for this script's Log files - if enabled

.PARAMETER  SaveMetadata
Switch to enable the saving of metadata to XML that can be re-imported into a PowerShell object
        
.PARAMETER  ImageMetadataPath
Top level folder for saving metadata in XML format that can be re-imported into a PowerShell object
        
.PARAMETER  LogEnabled
Switch to enabled the script's log file

.PARAMETER  Test
DEV: Switch to enable parameter values defined within the script - ### To be removed ###
        
.PARAMETER  Verbose
Switch to enable the display of progress bar and information entries to the Console

.PARAMETER  Mode
Archiving Mode. The following archiving methods are availabale with "All" being the default:
    [OriginalOnly] - Archive only the earliest version of a given image. If there are pre-existing archived images, these will be replaced if earlier versions are found in the source/import folder. This will also delete any additional, newer versions in the archive.

    [OriginalAdd]	  - Archive only the earliest version of a given image while retaining existing(newer) images in the archive.
                           Increment the revision number of the new files.

    [All]     - (Default) Archive all unique versions of an image sorted via an incremental revision number based on Modification Date

.PARAMETER  ArchiveByCamera
Switch to enable the use of a parent folder based on the respective image's Camera Model metadata if present. "Unknown" is used where this is unavailable and the usual chronological folders are used beneath.
        
.PARAMETER  Whatif
Switch to disable any rename or copy operations

.INPUTS
None. You cannot pipe objects to Archive-Images.ps1.

.OUTPUTS
NA

.EXAMPLE
    Archive-Images.ps1 -SourceFolder C:\Local\PhotosImportFolder -RootArchiveFolder C:\Local\Photos -Verbose
    [Note] "All" Mode implied.

.EXAMPLE  
    Archive-Images.ps1 -SourceFolder C:\Local\PhotosImportFolder -RootArchiveFolder C:\Local\Photos -Mode Replace -Verbose -LogEnabled -WhatIf
.EXAMPLE
    Archive-Images.ps1 -SourceFolder C:\Local\PhotosImportFolder -RootArchiveFolder C:\Local\Photos -Mode Add | select ArchiveName,ArchiveFolder,LatestRevision,CameraID,KeyWords,Model,FileName,BitmapHash,FileHash,DateTimeOriginal,FileCreateDate,FileModifyDate,SourceFile | out-gridview
#>

### To Do

# [X] Full Logging
# [ ] Fix "By Camera"
# [ ] Clean up output
# [ ] Metadata Export from all Source Images
# [ ] Merge Title/Descriptions?

# $USrcHashes = $SourceImagePaths | %{get-FileHash -LiteralPath $_ | Group Hash | % { $_.group[0] } }

function Archive-Photos {

    [CmdletBinding()]param(
        $Extension = ".jpg",
        $ExtFilter = "*$Extension",
        $SourcePath = @("C:\Local\Test\Photo Archive"),
        #$SourcePath = @("H:\Photo Archive\iPhone SE\2019\2019-05-May"),
        #$RootArchiveFolder = "H:\Photo Archive 3",
        $RootArchiveFolder = "c:\Local\Test\Photo Archive AllByCamera 2",
        $UndatedArchiveFolder = "$RootArchiveFolder\Undated",
        $ExcludedFolders = @("$($ENV:SystemDrive)\`$Recycle", "$($ENV:SystemDrive)\Windows", "$RootArchiveFolder"),
        $ArchiveFolderPattern = "",
        $ArchiveFilePattern = "",
        [switch]$ByCamera,
        $LogfilePath = $false,
        [switch]$LogEnabled,
        $Mode = "All", # All, OriginalOnly, OriginalAdd
        [switch]$WriteHost
    )


    # Prerequisites
    [System.Reflection.Assembly]::LoadWithPartialName("PresentationCore") | Out-Null


    $Script:FilesCopied = 0
    $Script:FilesDeleted = 0
    $Script:FilesRenamed = 0
    $Script:FilesSkipped = 0

    $DupNameIndex = 0

    ######
    $SourceImagesWithMetadata = @()
    $SourceImagesWithOutMetadata = @()
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
            [System.IO.FileInfo]$Second,
            $Method = "Hash" # or "Hash"
        )

        if ($Method -eq "Binary") {

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
        elseIf ($Method -eq "Hash") {

            $HashPair = $First, $Second | Get-xFileHash -Algorithm SHA512 | Select -ExpandProperty Hash
            if ($HashPair[0] -eq $HashPair[1]) { 
                Return $true
            }
            else {
                Return $false
            }

        }
    }
    function Get-xStringNumber {
        Param ($String)
	
        $CharacterHash = @{
            "a" = 1; "b" = 2; "c" = 3; "d" = 4; "e" = 5
            "f" = 6; "g" = 7; "h" = 8; "i" = 9; "j" = 10
            "k" = 11; "l" = 12; "m" = 13; "n" = 14; "o" = 15
            "p" = 16; "q" = 17; "r" = 18; "s" = 19; "t" = 20
            "u" = 21; "v" = 22; "w" = 23; "x" = 24; "y" = 25
            "z" = 26
            " " = 0
            "" = 0
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
            if ($String -ne "") {
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
            $SourceImage,
            $RootArchiveFolder,
            $UndatedArchiveFolder,
            [switch]$ByCamera
        )
	
        $Continue = $true
	
        try {
            $ImageStream = New-Object System.IO.FileStream($SourceImage.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read) -ErrorAction Stop -ErrorVariable FileStreamError
            $Decoder = New-Object System.Windows.Media.Imaging.JpegBitmapDecoder($ImageStream, [Windows.Media.Imaging.BitmapCreateOptions]::None, [Windows.Media.Imaging.BitmapCacheOption]::None) -ErrorAction Stop
            $Continue = $true
        }
        catch {
            if ($ImageStream) { $ImageStream.Dispose() }
            $Continue = $true
            $Decoder = $false
        }
	
        if ($Continue -eq $true) {
            if ($Decoder -ne $false) {
                $Metadata = $Decoder.Frames[0].Metadata
            }
            else {
                $Metadata = [pscustomobject] @{"Name" = "NonMetadata" }
                $Metadata | Add-Member -MemberType ScriptMethod -Name "GetQuery" -Value { Return $null }
            }
            Remove-Variable -Name Decoder
		
            $ExifDateTimeOriginal = [string]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=36867}")
            $SubSecTimeOriginal = [String]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37521}")
            # Convert Text milliseconds to Integer
            If ([string]::IsNullOrWhiteSpace($SubSecTimeOriginal) -eq $false) {
                $SubSecondString = "{0:D3}" -f [int]([float]("." + $SubSecTimeOriginal) * 1000)
            }
            Else {
                $SubSecondString = "{0:D3}" -f 0
            }

            try { $Title = [string]$Metadata.GetQuery("/xmp/dc:title/x-default") }
            catch { $Title = "" }
		
            $CameraFocalLength = [string]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37386}")
            $CameraUserComment = [string]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=37510}")
            $ImageDescription = [string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=270}")
		
            $CameraMake = [string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=271}")
            $CameraModel = ([string]$Metadata.GetQuery("/app1/Ifd/exIf:{uint=272}")).TrimEnd() #### adjust code to remove illegal characters, arrays etc.
            if ($CameraModel.length -lt 1) { $CameraModel = "Unknown Camera" }
		
            $ImageStream.Dispose()
            $Revision = "{0:D2}" -f 0 # "00"  
            $Separator = "-"
            $Extension = $SourceImage.Extension
		
            # Generate CameraID
            $CameraIDMake = Get-xStringNumber -String $CameraMake
            $CameraIDModel = Get-xStringNumber -String $CameraModel
            $CameraIDUserComment = Get-xStringNumber -String $CameraUserComment # 0
	
            $CameraID = "{0:D3}" -f ($CameraIDMake + $CameraIDModel + $CameraIDUserComment)
            ##############

            if (($ExifDateTimeOriginal -lt 1) -or ([string]::IsNullOrWhiteSpace($ExifDateTimeOriginal))) {
                # No Exif DateTime - Try IPTC
                try {
                    $IPTCDate = [string]$Metadata.GetQuery("/app13/irb/8bimiptc/iptc/date created") 
                    $IPTCTime = [string]$Metadata.GetQuery("/app13/irb/8bimiptc/iptc/time created")
                }
                catch { $IPTCDate = $null }
                if ([string]::IsNullOrWhiteSpace($IPTCDate)) {
    
                    # Other - XMP
                    #$DT3 = $Metadata.GetQuery("/xmp/xmp:CreateDate")
                    #$DT4 = $Metadata.GetQuery("/app1/ifd/exif/{ushort=36868}")
                    #$DT5 = $Metadata.GetQuery("/xmp/exif:DateTimeOriginal")
    
                    # No Date Time Object - Use Modified Date - LastWriteTime
                    # Change SubSecond to a 3 character Hex string based on File Size to add uniqueness
                    [string]$HexSize = '{0:X3}' -f $SourceImage.Length
                    $SubSecondString = $HexSize[($HexSize.Length - 1)..($HexSize.Length - 3)] -join ""
                    $Year = "{0:D4}" -f $SourceImage.LastWriteTime.Year
                    $Month = "{0:D2}" -f $SourceImage.LastWriteTime.Month
                    $Day = "{0:D2}" -f $SourceImage.LastWriteTime.Day
                    $DateTimeTaken = [PSCustomObject]@{
                        NoDate          = $true
                        Year            = $Year
                        Month           = $Month
                        Day             = $Day
                        MonthName       = (get-date -month $Month -format MMM)
                        DayName         = (get-date -year $Year -month $Month -day $Day).dayofweek
                        MonthFolderName = ($Year + "-" + $Month + "-" + (get-date -month $Month -format MMMM))
                        DateString      = "$Year.$Month.$Day"
                        TimeString      = ("$($SourceImage.LastWriteTime.Hour)$($SourceImage.LastWriteTime.Minute)$($SourceImage.LastWriteTime.Second)" + "." + $SubSecondString)
                        TimeZone        = "+0000" 
                    }
                }
                else {
                    # IPTC DateTime Taken Object
                    $Year = $IPTCDate.Substring(0, 4)
                    $Month = $IPTCDate.Substring(4, 2)
                    $Day = $IPTCDate.Substring(6, 2)
                    $DateTimeTaken = [PSCustomObject]@{
                        NoDate          = $false
                        Year            = $Year
                        Month           = $Month
                        Day             = $Day
                        MonthName       = (get-date -month $Month -format MMM)
                        DayName         = (get-date -year $Year -month $Month -day $Day).dayofweek
                        MonthFolderName = ($Year + "-" + $Month + "-" + (get-date -month $Month -format MMMM))
                        DateString      = "$Year.$Month.$Day"
                        TimeString      = ($IPTCTime.Substring(0, 6)) + "." + $SubSecondString
                        TimeZone        = ($IPTCTime.Substring(6, 5))
                    }
                }
            }
            else {
                # Exif DateTime Taken Object
                $Year = $ExifDateTimeOriginal.Substring(0, 4)
                $Month = $ExifDateTimeOriginal.Substring(5, 2)
                $Day = $ExifDateTimeOriginal.Substring(8, 2)
                try { [string]$ExifTimeZone = [String]$Metadata.GetQuery("/app1/Ifd/exIf/subIfd:{uint=36881}") }
                catch { $ExifTimeZone = "Unavailable" }
                $DateTimeTaken = [PSCustomObject]@{
                    NoDate          = $false
                    Year            = $Year
                    Month           = $Month
                    Day             = $Day
                    MonthName       = (get-date -month $Month -format MMM)
                    DayName         = (get-date -year $Year -month $Month -day $Day).dayofweek
                    MonthFolderName = ($Year + "-" + $Month + "-" + (get-date -month $Month -format MMMM))
                    DateString      = "$Year.$Month.$Day"
                    TimeString      = ($ExifDateTimeOriginal.Substring(11, 8)).Replace(":", "") + "." + $SubSecondString
                    TimeZone        = $ExifTimeZone
                }
            }

            if ($ByCamera -eq $true) {
                if ($DateTimeTaken.NoDate -eq $false) {
                    $ArchiveFolder = $RootArchiveFolder + "\" + $CameraModel + "\" + $DateTimeTaken.Year + "\" + $DateTimeTaken.MonthFolderName
                }
                else {
                    $ArchiveFolder = $UndatedArchiveFolder + "\" + $CameraModel + "\" + $DateTimeTaken.Year + "\" + $DateTimeTaken.MonthFolderName
                }
            }
            else {
                if ($DateTimeTaken.NoDate -eq $false) {
                    $ArchiveFolder = $RootArchiveFolder + "\" + $DateTimeTaken.Year + "\" + $DateTimeTaken.MonthFolderName
                }
                else {
                    $ArchiveFolder = $UndatedArchiveFolder + "\" + $DateTimeTaken.Year + "\" + $DateTimeTaken.MonthFolderName
                }
            }

            $ArchiveBaseName = $DateTimeTaken.DateString + $Separator + $DateTimeTaken.TimeString + $Separator + $CameraID + $Separator
            $ArchiveBaseNameShort = $DateTimeTaken.DateString + $Separator + $DateTimeTaken.TimeString + $Separator
            $ArchiveName = $DateTimeTaken.DateString + $Separator + $DateTimeTaken.TimeString + $Separator + $CameraID + $Separator + $Revision + $Extension
            $ArchiveBasePath = $ArchiveFolder + "\" + $DateTimeTaken.DateString + $Separator + $DateTimeTaken.TimeString
        
            # Create Hashtable
            $Hashtable = [ordered]@{
                "DateTimeOriginal"     = $DateTimeTaken
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
        }
        else {
            $Hashtable = [ordered]@{
                "DateTimeOriginal"     = $false
                "SubSecTimeOriginal"   = $false
                "Model"                = $false
                "Make"                 = $false
                "Title"                = $Error[0].Exception.Message
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
        param(
            $Files,
            $LogfilePath = "",
            $LogEnabled = $false,
            $SourceFile
        )

        foreach ($File in $Files) {
            if ((Test-Path -path $File.FullArchivePath) -eq $false) {
                Write-Verbose "Renaming $($File.FullName) to $($File.FullArchivePath)" #-ForegroundColor Green
                Write-Verbose "r" #-ForegroundColor Gray -NoNewline
                if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Renaming $($File.FullName) to $($File.FullArchivePath)" -LogFile $LogFile -Status "Completed" }
                rename-item -LiteralPath $File.FullName -NewName $File.PhotoMetadata.ArchiveName
                $File.PhotoMetadata.ArchiveAction = $false
                $Script:FilesRenamed++
            }
            if ($File.FullArchivePath -eq $File.FullName) {
                Write-Verbose "New and old name match - No action required $($File.FullName)" #-ForegroundColor Cyan
                if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "New and old name match - No action required $($File.FullName)" -LogFile $LogFile -Status "Skipped" }
                Write-Verbose "_" #-ForegroundColor Cyan -NoNewline
                $File.PhotoMetadata.ArchiveAction = $false
            }
        }

        $RemainingFiles = @($Files | ? { $_.PhotoMetadata.ArchiveAction -eq $true })

        if (($null -eq $RemainingFiles) -or ($RemainingFiles.Count -eq 0)) {
            Write-Verbose "All files renamed successfully" #-ForegroundColor Green
            if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "All files renamed successfully" -LogFile $LogFile -Status "Completed" }
            return $null
        }
        else {
            Write-Verbose "Remaining Files $($RemainingFiles | FT -auto)"
            Write-Verbose "Archive Actions $($Files.PhotoMetadata.ArchiveAction | FL)"
            Write-Verbose "($($RemainingFiles.Count))" # -NoNewline
            Write-Verbose "^" #-ForegroundColor Magenta -NoNewline
            Write-Verbose "Source File`: $($SourceFile.FullName)"
            if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Recursion Initiated - ($($RemainingFiles.Count)) Remaining Files" -LogFile $LogFile -Status "Initiated" }
            Rename-xFiles -Files $RemainingFiles -LogfilePath $LogfilePath -LogEnabled $LogEnabled -SourceFile $SourceFile
        }

    }

    <#
Invoke-FastFind and associated CSharp Code ($FileExtensionsCS):

The MIT License (MIT)

Copyright (c) 2015 Øyvind Kallstad

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

    $FileExtensionsCS = @"
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Communary
{
    internal sealed class Win32Native
	{
		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		public static extern SafeFindHandle FindFirstFileExW(
			string lpFileName,
			FINDEX_INFO_LEVELS fInfoLevelId,
			out WIN32_FIND_DATAW lpFindFileData,
			FINDEX_SEARCH_OPS fSearchOp,
			IntPtr lpSearchFilter,
			FINDEX_ADDITIONAL_FLAGS dwAdditionalFlags);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		public static extern bool FindNextFile(SafeFindHandle hFindFile, out WIN32_FIND_DATAW lpFindFileData);

		[DllImport("kernel32.dll")]
		public static extern bool FindClose(IntPtr hFindFile);

		[DllImport("shlwapi.dll", CharSet = CharSet.Auto)]
		public static extern bool PathMatchSpec([In] String pszFileParam, [In] String pszSpec);

		[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		public static extern bool GetDiskFreeSpace(
			string lpRootPathName,
			out uint lpSectorsPerCluster,
			out uint lpBytesPerSector,
			out uint lpNumberOfFreeClusters,
			out uint lpTotalNumberOfClusters);

		[DllImport("kernel32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool DeleteFileW([MarshalAs(UnmanagedType.LPWStr)]string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool RemoveDirectoryW(string lpPathName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool SetFileAttributesW(
             string lpFileName,
             [MarshalAs(UnmanagedType.U4)] FileAttributes dwFileAttributes);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetFileAttributesW(string lpFileName);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern uint GetSecurityInfo(
            IntPtr hFindFile,
            SE_OBJECT_TYPE ObjectType,
            SECURITY_INFORMATION SecurityInfo,
            out IntPtr pSidOwner,
            out IntPtr pSidGroup,
            out IntPtr pDacl,
            out IntPtr pSacl,
            out IntPtr pSecurityDescriptor);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
        public static extern uint GetNamedSecurityInfoW(
            string pObjectName,
            SE_OBJECT_TYPE ObjectType,
            SECURITY_INFORMATION SecurityInfo,
            out IntPtr pSidOwner,
            out IntPtr pSidGroup,
            out IntPtr pDacl,
            out IntPtr pSacl,
            out IntPtr pSecurityDescriptor);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint LookupAccountSid(
            string lpSystemName,
            IntPtr psid,
            StringBuilder lpName,
            ref uint cchName,
            [Out] StringBuilder lpReferencedDomainName,
            ref uint cchReferencedDomainName,
            out uint peUse);

        [DllImport("advapi32", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool ConvertSidToStringSid(
        IntPtr sid,
        out IntPtr sidString);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LocalFree(
            IntPtr handle
        );

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateFileW(
            [MarshalAs(UnmanagedType.LPWStr)] string filename,
            [MarshalAs(UnmanagedType.U4)] FileAccess access,
            [MarshalAs(UnmanagedType.U4)] FileShare share,
            IntPtr securityAttributes,
            [MarshalAs(UnmanagedType.U4)] FileMode creationDisposition,
            [MarshalAs(UnmanagedType.U4)] FileAttributes flagsAndAttributes,
            IntPtr templateFile);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
		public struct WIN32_FIND_DATAW
		{
			public FileAttributes dwFileAttributes;
			internal System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
			internal System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
			internal System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
			public uint nFileSizeHigh;
			public uint nFileSizeLow;
			public uint dwReserved0;
			public uint dwReserved1;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			public string cFileName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
			public string cAlternateFileName;
		}

		public enum FINDEX_INFO_LEVELS
		{
			FindExInfoStandard,             // Return a standard set of attribute information.
			FindExInfoBasic,                // Does not return the short file name, improving overall enumeration speed. cAlternateFileName is always a NULL string.
			FindExInfoMaxInfoLevel          // This value is used for validation. Supported values are less than this value.
		}

		public enum FINDEX_SEARCH_OPS
		{
			FindExSearchNameMatch,          // The search for a file that matches a specified file name.
			FindExSearchLimitToDirectories, // This is an advisory flag. If the file system supports directory filtering, the function searches for a file that matches the specified name and is also a directory. If the file system does not support directory filtering, this flag is silently ignored.
			FindExSearchLimitToDevices      // This filtering type is not available.
		}

		[Flags]
		public enum FINDEX_ADDITIONAL_FLAGS
		{
			FindFirstExCaseSensitive,
			FindFirstExLargeFetch
		}

        public enum SE_OBJECT_TYPE
        {
            SE_UNKNOWN_OBJECT_TYPE,
            SE_FILE_OBJECT,
            SE_SERVICE,
            SE_PRINTER,
            SE_REGISTRY_KEY,
            SE_LMSHARE,
            SE_KERNEL_OBJECT,
            SE_WINDOW_OBJECT,
            SE_DS_OBJECT,
            SE_DS_OBJECT_ALL,
            SE_PROVIDER_DEFINED_OBJECT,
            SE_WMIGUID_OBJECT,
            SE_REGISTRY_WOW64_32KEY
        }

        public enum SECURITY_INFORMATION
        {
            OWNER_SECURITY_INFORMATION = 1,     // The owner identifier of the object is being referenced. Right required to query: READ_CONTROL. Right required to set: WRITE_OWNER.
            GROUP_SECURITY_INFORMATION = 2,     // The primary group identifier of the object is being referenced. Right required to query: READ_CONTROL. Right required to set: WRITE_OWNER.
            DACL_SECURITY_INFORMATION = 4,      // The DACL of the object is being referenced. Right required to query: READ_CONTROL. Right required to set: WRITE_DAC.
            SACL_SECURITY_INFORMATION = 8,      // The SACL of the object is being referenced. Right required to query: ACCESS_SYSTEM_SECURITY. Right required to set: ACCESS_SYSTEM_SECURITY.
        }
    }

	[SecurityCritical]
	internal class SafeFindHandle : SafeHandleZeroOrMinusOneIsInvalid
	{
		[SecurityCritical]
		public SafeFindHandle() : base(true)
		{ }

		[SecurityCritical]
		protected override bool ReleaseHandle()
		{
			return Win32Native.FindClose(base.handle);
		}
	}

	public static class FILETIMEExtensions
	{
		public static DateTime ToDateTime(this System.Runtime.InteropServices.ComTypes.FILETIME time)
		{
			ulong high = (ulong)time.dwHighDateTime;
			ulong low = (ulong)time.dwLowDateTime;
			long fileTime = (long)((high << 32) + low);
			return DateTime.FromFileTimeUtc(fileTime);
		}
	}

	public static class FileExtensions
	{
		// prefix for long path support
		private const string normalPrefix = @"\\?\";
		private const string uncPrefix = @"\\?\UNC\";

		public static uint GetSectorSize(string path)
		{
			// add prefix to allow for maximum path of up to 32,767 characters
			string prefixedPath;
			if (path.StartsWith(@"\\"))
			{
				prefixedPath = path.Replace(@"\\", uncPrefix);
			}
			else
			{
				prefixedPath = normalPrefix + path;
			}

			uint lpSectorsPerCluster;
			uint lpBytesPerSector;
			uint lpNumberOfFreeClusters;
			uint lpTotalNumberOfClusters;

			string pathRoot = Path.GetPathRoot(path);
			if (!pathRoot.EndsWith(@"\"))
			{
				pathRoot = pathRoot + @"\";
			}

			bool result = Win32Native.GetDiskFreeSpace(pathRoot, out lpSectorsPerCluster, out lpBytesPerSector, out lpNumberOfFreeClusters, out lpTotalNumberOfClusters);
			if (result)
			{
				uint clusterSize = lpSectorsPerCluster * lpBytesPerSector;
				return clusterSize;
			}
			else
			{
				return 0;
			}
		}

		public static void DeleteFile(string path)
		{
			string prefixedPath;
			if (path.StartsWith(@"\\"))
			{
				prefixedPath = path.Replace(@"\\", uncPrefix);
			}
			else
			{
				prefixedPath = normalPrefix + path;
			}

            bool success = Win32Native.DeleteFileW(prefixedPath);
            if (!success)
            {
                int lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(lastError);
            }
		}

        public static void DeleteDirectory(string path)
        {
            string prefixedPath;
            if (path.StartsWith(@"\\"))
            {
                prefixedPath = path.Replace(@"\\", uncPrefix);
            }
            else
            {
                prefixedPath = normalPrefix + path;
            }

            bool success = Win32Native.RemoveDirectoryW(prefixedPath);
            if (!success)
            {
                int lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(lastError);
            }
        }

        public static void AddFileAttributes(string path, FileAttributes fileAttributes)
        {
            string prefixedPath;
            if (path.StartsWith(@"\\"))
            {
                prefixedPath = path.Replace(@"\\", uncPrefix);
            }
            else
            {
                prefixedPath = normalPrefix + path;
            }

            bool success = Win32Native.SetFileAttributesW(prefixedPath, fileAttributes);
            if (!success)
            {
                int lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(lastError);
            }
        }

        public static uint GetFileAttributes(string path)
        {
            string prefixedPath;
            if (path.StartsWith(@"\\"))
            {
                prefixedPath = path.Replace(@"\\", uncPrefix);
            }
            else
            {
                prefixedPath = normalPrefix + path;
            }

            return (Win32Native.GetFileAttributesW(prefixedPath));
        }

        public static ArrayList ReadLastLines(string path, int last)
        {
            ArrayList output = new ArrayList();
            int count = 0;
            foreach (var line in File.ReadLines(path).Reverse())
            {
                count++;
                output.Add(line);
                if (count == last)
                {
                    break;
                }
            }
            output.Reverse();
            return output;
        }

        public static string GetFileOwner(string path)
        {

            string prefixedPath;
            if (path.StartsWith(@"\\"))
            {
                prefixedPath = path.Replace(@"\\", uncPrefix);
            }
            else
            {
                prefixedPath = normalPrefix + path;
            }

            IntPtr NA = IntPtr.Zero;
            IntPtr sidOwner;

            var errorCode = Win32Native.GetNamedSecurityInfoW(prefixedPath, Win32Native.SE_OBJECT_TYPE.SE_FILE_OBJECT, Win32Native.SECURITY_INFORMATION.OWNER_SECURITY_INFORMATION, out sidOwner, out NA, out NA, out NA, out NA);
            if (errorCode == 0)
            {
                const uint bufferLength = 64;
                StringBuilder fileOwner = new StringBuilder();
                var accountLength = bufferLength;
                var domainLength = bufferLength;
                StringBuilder ownerAccount = new StringBuilder((int)bufferLength);
                StringBuilder ownerDomain = new StringBuilder((int)bufferLength);
                uint peUse;

                errorCode = Win32Native.LookupAccountSid(null, sidOwner, ownerAccount, ref accountLength, ownerDomain, ref domainLength, out peUse);
                if (errorCode != 0)
                {
                    fileOwner.Append(ownerDomain);
                    fileOwner.Append(@"\");
                    fileOwner.Append(ownerAccount);
                    return fileOwner.ToString();
                }
                else
                {
                    IntPtr sidString = IntPtr.Zero;
                    if (Win32Native.ConvertSidToStringSid(sidOwner, out sidString))
                    {
                        //string account = new System.Security.Principal.SecurityIdentifier(sidOwner).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
                        //Console.WriteLine(account);
                        return Marshal.PtrToStringAuto(sidString);
                    }
                    else
                    {
                        int lastError = Marshal.GetLastWin32Error();
                        throw new Win32Exception(lastError);
                    }
                }
            }
            else
            {
                int lastError = Marshal.GetLastWin32Error();
                throw new Win32Exception(lastError);
            }
        }

		public static List<FileInformation> FastFind(string path, string searchPattern, bool getFile, bool getDirectory, bool recurse, int? depth, bool parallel, bool suppressErrors, bool largeFetch, bool getHidden, bool getSystem, bool getReadOnly, bool getCompressed, bool getArchive, bool getReparsePoint, string filterMode)
		{
			object resultListLock = new object();
			Win32Native.WIN32_FIND_DATAW lpFindFileData;
			Win32Native.FINDEX_ADDITIONAL_FLAGS additionalFlags = 0;
			if (largeFetch)
			{
				additionalFlags = Win32Native.FINDEX_ADDITIONAL_FLAGS.FindFirstExLargeFetch;
			}

			// add prefix to allow for maximum path of up to 32,767 characters
			string prefixedPath;
			if (path.StartsWith(@"\\"))
			{
				prefixedPath = path.Replace(@"\\", uncPrefix);
			}
			else
			{
				prefixedPath = normalPrefix + path;
			}

			var handle = Win32Native.FindFirstFileExW(prefixedPath + @"\*", Win32Native.FINDEX_INFO_LEVELS.FindExInfoBasic, out lpFindFileData, Win32Native.FINDEX_SEARCH_OPS.FindExSearchNameMatch, IntPtr.Zero, additionalFlags);

			List<FileInformation> resultList = new List<FileInformation>();
			List<FileInformation> subDirectoryList = new List<FileInformation>();

			if (!handle.IsInvalid)
			{
				do
				{
					// skip "." and ".."
					if (lpFindFileData.cFileName != "." && lpFindFileData.cFileName != "..")
					{
						// if directory...
						if ((lpFindFileData.dwFileAttributes & FileAttributes.Directory) == FileAttributes.Directory)
						{
							// ...and if we are performing a recursive search...
							if (recurse)
							{
								// ... populate the subdirectory list
								string fullName = Path.Combine(path, lpFindFileData.cFileName);
								subDirectoryList.Add(new FileInformation { Path = fullName });
							}
						}

						// skip folders if only the getFile parameter is used
						if (getFile && !getDirectory)
						{
							if ((lpFindFileData.dwFileAttributes & FileAttributes.Directory) == FileAttributes.Directory)
							{
								continue;
							}
						}

						// if file matches search pattern and attribute filter, add it to the result list
						if (MatchesFilter(lpFindFileData.dwFileAttributes, lpFindFileData.cFileName, searchPattern, getFile, getDirectory, getHidden, getSystem, getReadOnly, getCompressed, getArchive, getReparsePoint, filterMode))
						{
							string fullName = Path.Combine(path, lpFindFileData.cFileName);
							long? thisFileSize = null;
                            if ((lpFindFileData.dwFileAttributes & FileAttributes.Directory) != FileAttributes.Directory)
                            {
                                thisFileSize = (lpFindFileData.nFileSizeHigh * (2 ^ 32) + lpFindFileData.nFileSizeLow);
                            }
							resultList.Add(new FileInformation { Name = lpFindFileData.cFileName, Path = Path.Combine(path, lpFindFileData.cFileName), Parent = path, Attributes = lpFindFileData.dwFileAttributes, FileSize = thisFileSize, CreationTime = lpFindFileData.ftCreationTime.ToDateTime(), LastAccessTime = lpFindFileData.ftLastAccessTime.ToDateTime(), LastWriteTime = lpFindFileData.ftLastWriteTime.ToDateTime() });

						}
					}
				}
				while (Win32Native.FindNextFile(handle, out lpFindFileData));

				// close the file handle
				handle.Dispose();

				// handle recursive search
				if (recurse)
				{
					// handle depth of recursion
					if (depth > 0)
					{
						if (parallel)
						{
							subDirectoryList.AsParallel().ForAll(x =>
							{
								List<FileInformation> resultSubDirectory = new List<FileInformation>();
								resultSubDirectory = FastFind(x.Path, searchPattern, getFile, getDirectory, recurse, (depth - 1), false, suppressErrors, largeFetch, getHidden, getSystem, getReadOnly, getCompressed, getArchive, getReparsePoint, filterMode);
								lock (resultListLock)
								{
									resultList.AddRange(resultSubDirectory);
								}
							});
						}

						else
						{
							foreach (FileInformation directory in subDirectoryList)
							{
								foreach (FileInformation result in FastFind(directory.Path, searchPattern, getFile, getDirectory, recurse, (depth - 1), false, suppressErrors, largeFetch, getHidden, getSystem, getReadOnly, getCompressed, getArchive, getReparsePoint, filterMode))
								{
									resultList.Add(result);
								}
							}
						}
					}

					// if no depth are specified
					else if (depth == null)
					{
						if (parallel)
						{
							subDirectoryList.AsParallel().ForAll(x =>
							{
								List<FileInformation> resultSubDirectory = new List<FileInformation>();
								resultSubDirectory = FastFind(x.Path, searchPattern, getFile, getDirectory, recurse, null, false, suppressErrors, largeFetch, getHidden, getSystem, getReadOnly, getCompressed, getArchive, getReparsePoint, filterMode);
								lock (resultListLock)
								{
									resultList.AddRange(resultSubDirectory);
								}
							});
						}

						else
						{
							foreach (FileInformation directory in subDirectoryList)
							{
								foreach (FileInformation result in FastFind(directory.Path, searchPattern, getFile, getDirectory, recurse, null, false, suppressErrors, largeFetch, getHidden, getSystem, getReadOnly, getCompressed, getArchive, getReparsePoint, filterMode))
								{
									resultList.Add(result);
								}
							}
						}
					}
				}
			}

			// error handling
			else if (handle.IsInvalid && !suppressErrors)
			{
				int hr = Marshal.GetLastWin32Error();
				if (hr != 2 && hr != 0x12)
				{
					//throw new Win32Exception(hr);
					Console.WriteLine("{0}:  {1}", path, (new Win32Exception(hr)).Message);
				}
			}

			return resultList;
		}

		private static bool MatchesFilter(FileAttributes fileAttributes, string name, string searchPattern, bool aFile, bool aDirectory, bool aHidden, bool aSystem, bool aReadOnly, bool aCompressed, bool aArchive, bool aReparsePoint, string filterMode)
		{
			// first make sure that the name matches the search pattern
			if (Win32Native.PathMatchSpec(name, searchPattern))
			{
				// then we build our filter attributes enumeration
				FileAttributes filterAttributes = new FileAttributes();

				if (aDirectory)
				{
					filterAttributes |= FileAttributes.Directory;
				}

				if (aHidden)
				{
					filterAttributes |= FileAttributes.Hidden;
				}

				if (aSystem)
				{
					filterAttributes |= FileAttributes.System;
				}

				if (aReadOnly)
				{
					filterAttributes |= FileAttributes.ReadOnly;
				}

				if (aCompressed)
				{
					filterAttributes |= FileAttributes.Compressed;
				}

				if (aReparsePoint)
				{
					filterAttributes |= FileAttributes.ReparsePoint;
				}

				if (aArchive)
				{
					filterAttributes |= FileAttributes.Archive;
				}

				// based on the filtermode, we match the file with our filter attributes a bit differently
				switch (filterMode)
				{
					case "Include":
						if ((fileAttributes & filterAttributes) == filterAttributes)
						{
							return true;
						}
						else
						{
							return false;
						}
					case "Exclude":
						if ((fileAttributes & filterAttributes) != filterAttributes)
						{
							return true;
						}
						else
						{
							return false;
						}
					case "Strict":
						if (fileAttributes == filterAttributes)
						{
							return true;
						}
						else
						{
							return false;
						}
				}
				return false;
			}
			else
			{
				return false;
			}
		}

		[Serializable]
		public class FileInformation
		{
			public string Name;
			public string Path;
			public string Parent;
			public FileAttributes Attributes;
			public long? FileSize;
			public DateTime CreationTime;
			public DateTime LastAccessTime;
			public DateTime LastWriteTime;
		}

        [Serializable]
        public class SecurityInformation
        {
            public string Path;
            public string Owner;
            public string Access;
        }
	}
}
"@
    Add-Type -TypeDefinition $FileExtensionsCS
    function Invoke-FastFind {
        <#
        .SYNOPSIS
            Search for files and folders.
        .DESCRIPTION
            This function uses WIN32 API to perform faster searching for files and folders. It also supports large paths.
        .EXAMPLE
            Invoke-FastFind
            Will list all files and folders in the current directory.
        .EXAMPLE
            Invoke-FastFind c:\
            Will list all files and folders in c:\
        .EXAMPLE
            Invoke-FastFind c:\ prog*
            Will list all files and folders in c:\ that starts with 'prog'.
        .EXAMPLE
            Invoke-FastFind c:\ -Directory
            Will list all directories in c:\
        .EXAMPLE
            Invoke-FastFind c:\ -Directory -Hidden
            Will list all hidden directories in c:\
        .EXAMPLE
            Invoke-FastFind c:\ -System -Hidden -AttributeFilterMode Exclude
            Will list all files and folders in c:\ that don't have the System and Hidden attributes set.
        .EXAMPLE
            Invoke-FastFind c:\ -Hidden -System -Archive -AttributeFilterMode Strict
            Will list all files and folders in c:\ that only have the Hidden, System and Archive attributes set.
        .NOTES
            Author: Øyvind Kallstad
            Date: 13.11.2015
            Version: 1.0
        .LINK
            https://communary.wordpress.com/
            https://github.com/gravejester/Communary.FileExtensions
    #>
        [CmdletBinding()]
        param (
            # Path where search starts from. The default value is the current directory.
            [Parameter(Position = 1)]
            [ValidateNotNullOrEmpty()]
            [string[]] $Path = ((Get-Location).Path),

            # Search filter. Accepts wildcards; * and ?. The default value is '*'.
            [Parameter(Position = 2)]
            [string] $Filter = '*',

            [Parameter()]
            [Alias('f')]
            [switch] $File,

            [Parameter()]
            [Alias('d')]
            [switch] $Directory,

            [Parameter()]
            [switch] $Hidden,

            [Parameter()]
            [switch] $System,

            [Parameter()]
            [switch] $ReadOnly,

            [Parameter()]
            [switch] $Compressed,

            [Parameter()]
            [switch] $Archive,

            [Parameter()]
            [switch] $ReparsePoint,

            # Choose the filter mode for attribute filtering. Valid choices are 'Include', 'Exclude' and 'Strict'. Default is 'Include'.
            [Parameter()]
            [ValidateSet('Include', 'Exclude', 'Strict')]
            [string] $AttributeFilterMode = 'Include',

            # Perform recursive search.
            [Parameter()]
            [Alias('r')]
            [switch] $Recurse,

            # Depth of recursive search. Default is null (unlimited recursion).
            [Parameter()]
            [nullable[int]] $Depth = $null,

            # Use a larger buffer for the search, which *can* increase performance. Not supported for operating systems older than Windows Server 2008 R2 and Windows 7.
            [Parameter()]
            [switch] $LargeFetch
        )

        if ($PSBoundParameters['File'] -and $PSBoundParameters['Directory']) {
            $File = $false
            $Directory = $false
        }

        foreach ($thisPath in $Path) {
            #if (Test-Path -Path $thisPath) {

            # adds support for relative paths
            #$resolvedPath = (Resolve-Path -Path $thisPath).Path
            #$resolvedPath = $resolvedPath.Replace('Microsoft.PowerShell.Core\FileSystem::','')
            $resolvedPath = $thisPath

            # handle a quirk where \ at the end of a non-UNC, non-root path failes
            if (-not ($resolvedPath.ToString().StartsWith('\\'))) {
                if ($resolvedPath.ToString().EndsWith('\')) {
                    if (-not($resolvedPath -eq ([System.IO.Path]::GetPathRoot($resolvedPath)))) {
                        $resolvedPath = $resolvedPath.ToString().TrimEnd('\')
                    }
                }
            }

            # call FastFind to perform search
            [Communary.FileExtensions]::FastFind($resolvedPath, $Filter, $File, $Directory, $Recurse, $Depth, $true, $true, $LargeFetch, $Hidden, $System, $ReadOnly, $Compressed, $Archive, $ReparsePoint, $AttributeFilterMode)
            #}
            #else {
            #    Write-Warning "$thisPath - Invalid path"
            #}
        }
    }
    function Write-xLog {
        Param(
            $LogFile,
            [string]$DateTime = [string](Get-xDateTimeString -Pattern "%Y%m%d-%H.%M.%S"),
            [string]$Action = "",
            [string]$Status = "", # (Started, Completed, Running, Error, Warning)
            [string]$SourcePath = "",
            [string]$ArchivePath = "",
            [string]$Message = ""
        )
        $LogFile.WriteLine("$DateTime,$Action,$Status,$SourcePath,$ArchivePath,$Message")
    }
    Function Get-xFileHash {
        [CmdletBinding()]
        Param(
            [Parameter(ValueFromPipeline)]
            $Path,
            [string]$Encryption = "SHA256",
            $Count,
            [switch]$Progress
        )
        begin {
            if ($Progress) {
                $swThresh = 1000
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $i = 1; $t = $Count / 100; $tc = $Count
            }
        }
        process {
    
            $ImageStream = ([IO.StreamReader]$Path).BaseStream
            $FormatedHash = New-Object System.Text.StringBuilder
            $RawHash = [Security.Cryptography.HashAlgorithm]::Create($Encryption).ComputeHash($ImageStream)
            foreach ($HashValue in $RawHash) { $null = $FormatedHash.Append("{0:X2}" -f $HashValue) }
            $Hash = $FormatedHash.ToString()
            #Remove-Variable FormatedHash
            $ImageStream.Dispose()
            if ($Progress) { if ($sw.Elapsed.TotalMilliseconds -ge $swThresh) { Write-Progress -Activity "Generating Hashes" -PercentComplete ([int]$i / $t) -Status "$i of $tc"; $sw.Reset(); $sw.Start() }; $i++ }
            Return [PSCustomObject]@{
                "Algorithm" = $Encryption
                "Hash"      = $Hash
                "Path"      = $Path
            }
        }
        end {
            if ($Progress) { Write-Progress -Activity "Generating Hashes" -Completed }
        }
    }

    #endregion Functions

    $ScriptStart = Get-Date
    $RunTime = Get-xDateTimeString -DateTime $ScriptStart -Pattern "%Y%m%d-%H.%M.%S"
    $swThresh = 1000

    # Log File
    if (-not $LogfilePath) {
        $LogFolder = "$RootArchiveFolder\Logs"
        $LogfilePath = "$LogFolder\ArchivePhotosScriptLog($Runtime).csv"
    }
    if ($LogEnabled -eq $true) {
        If ((Test-Path -literalPath $LogFolder) -eq $false) { New-Item $LogFolder -ItemType Directory > $null }
        $LogFile = New-Object System.IO.StreamWriter($LogfilePath)
        #Header
        $LogFile.WriteLine("DateTime,Action,Status,SourcePath,ArchivePath,Message")   
    }

    if ($LogEnabled -eq $true) { Write-xLog -Action "Start" -Status "Running" -Message "Mode`:$Mode, Test`:$Test" -LogFile $LogFile }

    #Find All Source Photos (JPG)
    # Super Fast Win32 Method
    Write-Verbose "Searching Source Path/s for Images (.jpg)..." #-ForegroundColor Yellow
    $Duration_SeachSourceForImages = (Measure-Command { $AllSourceImagePaths = Invoke-FastFind -Recurse -File -Filter $ExtFilter -Path $SourcePath -LargeFetch -System -Hidden -AttributeFilterMode Exclude | Select -ExpandProperty Path }).ToString()
    if ($LogEnabled -eq $true) { Write-xLog -Action "Collect Image File Paths & Hashes ($($AllSourceImagePaths.Count))" -Message "Duration`: $Duration_SeachSourceForImages" -LogFile $LogFile -Status "Completed" }

    # DOS - Issue with Text Encoding
    #Write-Verbose "[B] Searching Source Path/s for Images (.jpg)..." #-ForegroundColor Yellow
    #(Measure-Command { $AllSourceImagePaths = ($SourcePath | % { cmd /c dir "$_\*.jpg" /b /s /a-d-h-s }) }).TotalSeconds

    $SourceImagePaths = $AllSourceImagePaths
    foreach ($ExcludedFolder in $ExcludedFolders) {
        $ExcludedFolder
        #$SourceImagePaths = $SourceImagePaths | ? { $_ -notlike "$ExcludedFolder*" }
        $SourceImagePaths = @($SourceImagePaths | ? { $_ -notlike "$ExcludedFolder*" } | ? { $_.Substring($_.LastIndexOf("\") + 1, 2) -ne "._" })
    }

    <#
    ###############################################
    # Remove all duplicates from Source including pre-archived matches

    Write-Verbose "Group Source Images by Hash"
    $Duration_GroupSourceImagesByHash = (measure-command {
            $USrcHashes = $SourceImagePaths | get-xFileHash -Count $SourceImagePaths.Count -Progress | Group Hash | % { $_.group[0] } 
            $USrcHashes | Add-Member -MemberType NoteProperty -Name "Location" -Value "Source"
        }).TotalSeconds
    if ($LogEnabled -eq $true) { Write-xLog -Action "Group Source Images by Hash " -Message "Duration`: $Duration_GroupSourceImagesByHash" -LogFile $LogFile -Status "Completed" }

    Write-Verbose "Get All Archived Hashes"
    (measure-command { $DestImagePaths = Invoke-FastFind -Recurse -File -Filter $ExtFilter -Path $RootArchiveFolder -LargeFetch -System -Hidden -AttributeFilterMode Exclude | Select -ExpandProperty Path }).TotalSecond
    if ($null -ne $DestImagePaths) {
        (measure-command {
                #$UDestHashes = get-FileHash -LiteralPath $DestImagePaths #bug Check for null
                $UDestHashes = $DestImagePaths | Get-xFileHash -Count $DestImagePaths.Count -Progress
                $UDestHashes | Add-Member -MemberType NoteProperty -Name "Location" -Value "Archive"
            }).TotalSeconds

        Write-Verbose "Combine Hashes"
        (measure-command { $AllHashes = $USrcHashes + $UDestHashes | Group Hash }).TotalSeconds

    }
    else {
        Write-Verbose "Combine Hashes"
        (measure-command { $AllHashes = $USrcHashes + $UDestHashes | Group Hash }).TotalSeconds
    }


    Write-Verbose "Eliminate Duplicates from Source"
    (measure-command { $FinalImportPaths = $AllHashes | ? { $_.Count -eq 1 } | % { $_.Group | ? { $_.Location -eq "Source" } } }).TotalSeconds

    Write-Verbose "FinalImportPaths Count"
    $FinalImportPaths.Count


    ###############################################

    Write-Verbose "$($AllSourceImagePaths.Count) Images found. (Total)" #-ForegroundColor Green
    Write-Verbose "$($SourceImagePaths.Count) Images found. (Filtered)" #-ForegroundColor Green
    Write-Verbose "$($FinalImportPaths.Count) Images found. (Deduped)" #-ForegroundColor Green


    # Progress Bar Setup
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $i = 1; $t = $FinalImportPaths.Count / 100; $tc = $FinalImportPaths.Count

    foreach ($SourceImagePath in $FinalImportPaths.Path) {

        # Get Source JPG metadata
        if ($LogEnabled -eq $true) { Write-xLog -Action "Get Metadata" -Message "Reading Metadata from $SourceImagePath" -LogFile $LogFile -Status "Initiated" }
        $CurrentFileObject = (Get-Item -LiteralPath $SourceImagePath | Select-Object *, 
            @{Label = "PhotoMetaData"; Exp = { 
                    if ($ByCamera -eq $true) { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder -ByCamera }
                    else { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder }
                }
            },
            @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
            @{Label = "FullArchivePath"; Exp = { "" } }
        )

        #####
        if ($CurrentFileObject.PhotoMetaData.NoMetadata -eq $false) {
            $SourceImagesWithMetadata += $CurrentFileObject
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Successfully Read Metadata" -SourcePath $SourceImagePath -LogFile $LogFile -Status "Completed" }
        }
        else {
            $SourceImagesWithOutMetadata += $CurrentFileObject
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Failed Reading Metadata" -SourcePath $SourceImagePath -LogFile $LogFile -Status "Error" }
        }
        #####

        # Check Destination for existing photos with same primary path (sans version info)
        if ($CurrentFileObject.PhotoMetaData.NoMetadata -eq $false) {
            if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Check For Matching Image Versions in Archive" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
            if (Test-Path -LiteralPath $CurrentFileObject.PhotoMetaData.ArchiveFolder) {
    
                $SearchFilter = $CurrentFileObject.PhotoMetaData.ArchiveBasePath + "*"
                try { $MatchingDestinationFiles = @(Get-ChildItem $SearchFilter) }
                catch { $MatchingDestinationFiles = $null }
    
                if ($MatchingDestinationFiles) {
                    if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Matching Versions Found" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             
                    # Get Metadata from matching destination photos
                    $MatchingDestinationFiles = ($MatchingDestinationFiles | Select-Object *, 
                        @{Label = "PhotoMetaData"; Exp = { 
                                if ($ByCamera -eq $true) { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder -ByCamera }
                                else { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder } 
                            }
                        },
                        @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
                        @{Label = "FullArchivePath"; Exp = { "" } }
                    )
            
                    $RenameMatchArrayGrouped = @($MatchingDestinationFiles | % { $_ }; $CurrentFileObject) | Sort LastWriteTime | Group LastWriteTime
                    $BaseRevisionNum = 0
                    $SubRevisionArray = ("a".."z")
                    foreach ($PhotoGroup in $RenameMatchArrayGrouped) {
                        $BaseRevision = "{0:D2}" -f $BaseRevisionNum
                        # Check for Matching Image that is not a duplicate but has a matching Modified Time Stamp
                        #if (($PhotoGroup.Count -gt 1) -and ($PhotoGroup.Group.FullName -contains $CurrentFileObject.FullName)) {
                        if ($PhotoGroup.Count -gt 1)  {
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Non-Duplicate Versions Found with Matching Modification Date" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                            $SubRevisionIndex = 0
                            foreach ($Photo in $PhotoGroup.Group) {
                                $Photo.Revision = "$BaseRevision-$($SubRevisionArray[$SubRevisionIndex])"
                                $Photo.FullArchivePath = $Photo.PhotoMetaData.ArchiveBasePath + "-" + $Photo.PhotoMetaData.CameraID + "-" + $Photo.Revision + $Photo.Extension
                                $Photo.PhotoMetaData.ArchiveName = $Photo.PhotoMetaData.ArchiveBaseName + $Photo.Revision + $Photo.Extension
                                $Photo.PhotoMetaData.ArchiveAction = $true
                                $SubRevisionIndex++

                                if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Photo Revision = $($Photo.Revision)" -SourcePath $Photo.FullName -ArchivePath $Photo.PhotoMetaData.ArchiveName -LogFile $LogFile -Status "Completed" }

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

                    switch ($Mode) {
                        "All" {
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Rename Surrounding Files in Archive(Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                            Rename-xFiles -Files $RenameMatchArray -SourceFile $CurrentFileObject -LogfilePath $LogfilePath -LogEnabled $LogEnabled
                            #Write-Verbose "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                            #Write-Verbose "A" #-ForegroundColor Yellow -NoNewline
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                            Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                            $Script:FilesCopied++
                        }
                        "OriginalOnly" {
                            # Check if earliest revision
                            if ($CurrentFileObject.Revision -eq ("{0:D2}" -f 0)) {
                                # 00
                                # Deleting any existing/matching files in Archive
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Delete" -Message "Delete Older Files from Archive (Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                if ($RenameMatchArray) {
                                    #Write-Verbose "`#" #-ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    #Write-Host "`#" -ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    $RenameMatchArray | % { 
                                        Remove-Item -LiteralPath $_.FullName -Force 
                                        $Script:FilesDeleted++
                                    }
                                }

                                #Write-Verbose "O" #-ForegroundColor Yellow -NoNewline
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                                $Script:FilesCopied++
                            }
                            else {
                                # Don't copy as it isn't earlier than existing files in archive
                            }
                        }
                        "OriginalAdd" {
                            if ($CurrentFileObject.Revision -eq ("{0:D2}" -f 0)) {
                                # 00
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Rename Newer Files in Archive (Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                Rename-xFiles -Files $RenameMatchArray
                                #Write-Verbose "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                                #Write-Verbose "+" #-ForegroundColor Yellow -NoNewline
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                                $Script:FilesCopied++
                            }
                            else {
                                # Don't copy as it isn't earlier than existing files in archive
                            }
                        }
                        default { }
                    }
            
                
                }
                else {
                    #Copy File
                    try {
                        #Write-Verbose "Copy Photo(Existing Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName #-ForegroundColor Cyan
                        #Write-Verbose "Copying new File to Archive (Existing Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                        #Write-Verbose "+" #-ForegroundColor DarkGreen -NoNewline
                        if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode) (No Pre-existing Files in Archive)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                        Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination ($CurrentFileObject.PhotoMetaData.ArchiveFolder + "\" + $CurrentFileObject.PhotoMetaData.ArchiveName)
                        $Script:FilesCopied++
                    }
                    catch { } # Error copying file to Destination
                }
            }
            else {
                New-Item -Path $CurrentFileObject.PhotoMetaData.ArchiveFolder -ItemType Directory | Out-Null
                #Write-Verbose "Copy Photo(New Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName
                #Write-Verbose "Copying new File to Archive (New Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                #Write-Verbose "+" #-ForegroundColor Green -NoNewline
                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode) (No Pre-existing Files in Archive  - Create Folder)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination ($CurrentFileObject.PhotoMetaData.ArchiveFolder + "\" + $CurrentFileObject.PhotoMetaData.ArchiveName)
                $Script:FilesCopied++
            }
        }
        else {
            Write-Verbose "Unable to Retrieve Metadata from`: $($CurrentFileObject.FullName) $($CurrentFileObject.PhotoMetaData.Title)"
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Unable to Retrieve Metadata from`: $($CurrentFileObject.FullName)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
        }

        if ($sw.Elapsed.TotalMilliseconds -ge $swThresh) { Write-Progress -Activity "Archiving Files" -PercentComplete ([int]$i / $t) -Status "$i of $tc"; $sw.Reset(); $sw.Start() }; $i++

    }

    #>

    # Remove all duplicates from Source including pre-archived matches

    #region New Code

    Write-Verbose "Group Source Images by Hash and Remove Duplicates"
    $Duration_GroupSourceImagesByHash = (measure-command {
            $FinalImportPaths = $SourceImagePaths | get-xFileHash -Count $SourceImagePaths.Count -Progress | Group Hash | % { $_.group[0] } 
        }).TotalSeconds
    if ($LogEnabled -eq $true) { Write-xLog -Action "Group Source Images by Hash " -Message "Duration`: $Duration_GroupSourceImagesByHash" -LogFile $LogFile -Status "Completed" }

    # Progress Bar Setup
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $i = 1; $t = $FinalImportPaths.Count / 100; $tc = $FinalImportPaths.Count

    foreach ($SourceImage in $FinalImportPaths) {

        # Get Source JPG metadata
        if ($LogEnabled -eq $true) { Write-xLog -Action "Get Metadata" -Message "Reading Metadata from $SourceImage.Path" -LogFile $LogFile -Status "Initiated" }
        $CurrentFileObject = (Get-Item -LiteralPath $SourceImage.Path | Select-Object *, 
            @{Label = "PhotoMetaData"; Exp = { 
                    if ($ByCamera -eq $true) { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder -ByCamera }
                    else { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder }
                }
            },
            @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
            @{Label = "FullArchivePath"; Exp = { "" } }
        )

        if ($CurrentFileObject.PhotoMetaData.NoMetadata -eq $false) {
            $SourceImagesWithMetadata += $CurrentFileObject
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Successfully Read Metadata" -SourcePath $SourceImage.Path -LogFile $LogFile -Status "Completed" }
        }
        else {
            $SourceImagesWithOutMetadata += $CurrentFileObject
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Failed Reading Metadata" -SourcePath $SourceImage.Path -LogFile $LogFile -Status "Error" }
        }

        if ($CurrentFileObject.PhotoMetaData.NoMetadata -eq $false) {

            # Search for existing files in Archive path 
            $SearchFilter = $CurrentFileObject.PhotoMetaData.ArchiveBasePath + "*"
            try { 
                try{$MatchingDestinationFiles = @(Get-ChildItem $SearchFilter -ErrorAction Stop | Select *, @{Label = "Hash"; Exp = { (Get-xFileHash -path $_.FullName).Hash } })}
                catch{$MatchingDestinationFiles = $false}
                if ($MatchingDestinationFiles) {
                    if ($MatchingDestinationFiles.Hash -contains $SourceImage.Hash) {
                        $HasDuplicateinArchive = $true # Drop Source Image
                    }
                    else {
                        $HasDuplicateinArchive = $false
                    }
                }
                else {
                    $HasDuplicateinArchive = $false
                }
            } 
            catch { 
                $MatchingDestinationFiles = $false
                $HasDuplicateinArchive = $false 
            }

   
            if ($MatchingDestinationFiles -ne $false) {
                if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Matching Versions Found" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                if ($HasDuplicateinArchive -ne $true) {                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
                    # Get Metadata from matching destination photos
                    $MatchingDestinationFiles = ($MatchingDestinationFiles | Select-Object *, 
                        @{Label = "PhotoMetaData"; Exp = { 
                                if ($ByCamera -eq $true) { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder -ByCamera }
                                else { Get-xImageMetadataHashtable -SourceImage $_ -RootArchiveFolder $RootArchiveFolder -UndatedArchiveFolder $UndatedArchiveFolder } 
                            }
                        },
                        @{Label = "Revision"; Exp = { "{0:D2}" -f 0 } },
                        @{Label = "FullArchivePath"; Exp = { "" } }
                    )

                    $RenameMatchArrayGrouped = @($MatchingDestinationFiles | % { $_ }; $CurrentFileObject) | Sort LastWriteTime | Group LastWriteTime
                    $BaseRevisionNum = 0
                    $SubRevisionArray = ("a".."z")
                    foreach ($PhotoGroup in $RenameMatchArrayGrouped) {
                        $BaseRevision = "{0:D2}" -f $BaseRevisionNum
                        # Check for Matching Image that is not a duplicate but has a matching Modified Time Stamp
                        if ($PhotoGroup.Count -gt 1) {
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Versioning Check" -Message "Non-Duplicate Versions Found with Matching Modification Date" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
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

                    switch ($Mode) {
                        "All" {
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Rename Surrounding Files in Archive(Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                            Rename-xFiles -Files $RenameMatchArray
                            #Write-Verbose "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                            #Write-Verbose "A" #-ForegroundColor Yellow -NoNewline
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                            Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                            $Script:FilesCopied++
                        }
                        "OriginalOnly" {
                            # Check if earliest revision
                            if ($CurrentFileObject.Revision -eq ("{0:D2}" -f 0)) {
                                # 00
                                # Deleting any existing/matching files in Archive
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Delete" -Message "Delete Older Files from Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                if ($RenameMatchArray) {
                                    #Write-Verbose "`#" #-ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    Write-Host "`#" -ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    $RenameMatchArray | % { 
                                        Remove-Item -LiteralPath $_.FullName -Force 
                                        $Script:FilesDeleted++
                                    }
                                }

                                #Write-Verbose "O" #-ForegroundColor Yellow -NoNewline
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                                $Script:FilesCopied++
                            }
                            else {
                                # Don't copy as it isn't earlier than existing files in archive
                            }
                        }
                        "OriginalOnlyND" {
                            if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Rename Surrounding Files in Archive(Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                            Rename-xFiles -Files $RenameMatchArray
                            # Check if earliest revision (including non-dups of same mod date)
                            if ($CurrentFileObject.Revision -like (("{0:D2}" -f 0) + "*")) {
                                # 00
                                # Deleting any existing/matching files in Archive
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Delete" -Message "Delete Older Files from Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                
                                if ($RenameMatchArray) {
                                    #Write-Verbose "`#" #-ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    Write-Host "`#" -ForegroundColor Red -BackgroundColor Yellow -NoNewline
                                    $RenameMatchArray | ?{$_.Revision -notlike (("{0:D2}" -f 0) + "*")} | % { 
                                        Remove-Item -LiteralPath $_.FullArchivePath -Force 
                                        $Script:FilesDeleted++
                                    }
                                }
                        
                                #Write-Verbose "O" #-ForegroundColor Yellow -NoNewline
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                                $Script:FilesCopied++
                            }
                            else {
                                # Don't copy as it isn't earlier than existing files in archive
                            }
                        }
                        "OriginalAdd" {
                            if ($CurrentFileObject.Revision -eq ("{0:D2}" -f 0)) {
                                # 00
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Rename" -Message "Rename Newer Files in Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchiveFolder -LogFile $LogFile -Status "Completed" }
                                Rename-xFiles -Files $RenameMatchArray
                                #Write-Verbose "Copying new File to Archive $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                                #Write-Verbose "+" #-ForegroundColor Yellow -NoNewline
                                if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                                Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination $CurrentFileObject.FullArchivePath
                                $Script:FilesCopied++
                            }
                            else {
                                # Don't copy as it isn't earlier than existing files in archive
                            }
                        }
                        default { }
          
                    }
                }
                else {
                    # Has Duplicate - Drop Source Image
                }
            }
            else {
                #Copy File
                try {
                    if((Test-Path -LiteralPath $CurrentFileObject.PhotoMetaData.ArchiveFolder) -ne $true){ New-Item -Path $CurrentFileObject.PhotoMetaData.ArchiveFolder -ItemType Directory | Out-Null }
                    #Write-Verbose "Copy Photo(Existing Folder): " $CurrentFileObject.PhotoMetaData.ArchiveName #-ForegroundColor Cyan
                    #Write-Verbose "Copying new File to Archive (Existing Folder) $($CurrentFileObject.FullName) to $($CurrentFileObject.FullArchivePath)" #-ForegroundColor Yellow
                    #Write-Verbose "+" #-ForegroundColor DarkGreen -NoNewline
                    if ($LogEnabled -eq $true) { Write-xLog -Action "Archive" -Message "Copy Source File to Archive (Mode`:$Mode) (No Pre-existing Files in Archive)" -SourcePath $SourceImage.Path -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
                    Copy-Item -LiteralPath $CurrentFileObject.FullName -Destination ($CurrentFileObject.PhotoMetaData.ArchiveFolder + "\" + $CurrentFileObject.PhotoMetaData.ArchiveName)
                    $Script:FilesCopied++
                }
                catch { 
                    Write-Host $Error[0]
                } # Error copying file to Destination
            }
        }
        else {
            Write-Verbose "Unable to Retrieve Metadata from`: $($CurrentFileObject.FullName) $($CurrentFileObject.PhotoMetaData.Title)"
            if ($LogEnabled -eq $true) { Write-xLog -Action "Metadata" -Message "Unable to Retrieve Metadata from`: $($CurrentFileObject.FullName)" -SourcePath $SourceImagePath -ArchivePath $CurrentFileObject.PhotoMetaData.ArchivePath -LogFile $LogFile -Status "Completed" }
        }

        if ($sw.Elapsed.TotalMilliseconds -ge $swThresh) { Write-Progress -Activity "Archiving Files" -PercentComplete ([int]$i / $t) -Status "$i of $tc"; $sw.Reset(); $sw.Start() }; $i++
    }

    #endregion New Code

    Write-Progress -Activity "Archiving Files" -Completed

    $ScriptEnd = Get-Date

    $ScriptDuration = $ScriptEnd - $ScriptStart

    #Write-Verbose
    #Write-Verbose

    Write-Verbose "Script Duration`:`t$($ScriptDuration.TotalSeconds) Seconds" #-ForegroundColor Cyan

    Write-Verbose "Files Copied  = $($Script:FilesCopied)"  #-ForegroundColor Yellow
    Write-Verbose "Files Deleted = $($Script:FilesDeleted)"  #-ForegroundColor Yellow
    Write-Verbose "Files Renamed = $($Script:FilesRenamed)" #-ForegroundColor Yellow
    Write-Verbose "Files Skipped = $($Script:FilesSkipped)" #-ForegroundColor Yellow

    if ($LogEnabled -eq $true) { Write-xLog -Action "Script" -Message "Script Completed (Duration`:$($ScriptDuration.TotalSeconds))" -LogFile $LogFile -Status "Completed" }

    # Close Log
    $LogFile.Close()


}

Archive-Photos -LogEnabled -Mode OriginalOnlyND -ByCamera -Verbose