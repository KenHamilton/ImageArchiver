# ImageArchiver
Script to search for all images across multiple sources and build/maintain a canonical archive

20190910 - Initial Upload. Code fragments only. Not fully functional.
20190920 - Scripts Consolidated. First working draft
20191001 - Rewrote handeling of jpgs without datetime taken metadata
20191004 - Replaced CMD \C Dir Search Method with Invoke-FastFind Copyright (c) 2015 Øyvind Kallstad (MIT)
20191018 - Added copy mode (All, OriginalOnly, OriginalAdd), Duplicate removal is now done before any copy operation
20191118 - Bug fixes, Change File Hash calculation for Archived files to inline with search for matching source files, Added OriginalOnlyND option to keep non-duplicate files with the same modification date