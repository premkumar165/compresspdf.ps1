<#
.SYNOPSIS
This is a simple PowerShell script that allows you to compress PDF files
.DESCRIPTION
The script uses GhostScript to compress PDF files. It uses community recommendations for the various compression 
.EXAMPLE
./Compress-PDF -File C:\example.pdf -CompressionLevel "ebook"
.EXAMPLE
./Compress-PDF -File C:\example.pdf -CompressionLevel "ebook" -CompatibilityLevel "1.4" -AdvancedCompress
.LINK
https://perplexity.nl
https://www.ghostscript.com/download/gsdnld.html
https://stackoverflow.com/questions/46195795/ghostscript-pdf-batch-compression/46196373
#>
[CmdletBinding()]
param
(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][System.IO.FileInfo]$File,
    [ValidateSet('screen', 'ebook', 'printer', 'prepress', 'default')][string]$CompressionLevel = "ebook",
    [ValidateSet('1.7','1.6','1.5', '1.4', '1.3', '1.2','1.1', '1.0')][string]$CompatibilityLevel = "1.4",
    [switch]$Overwrite,
    [string]$GhostScript,
    [switch]$AdvancedCompress
)
Begin
{
    if (-Not ($GhostScript))
    {
        $GhostScript = ((Get-ChildItem 'C:\Program Files\gs\' -Directory | Sort-Object LastWriteTime -Descending)[0].Fullname) + "\bin\gswin64.exe"
    }
    if (-Not(Test-Path $GhostScript))
    {
        Write-Error "The GhostScript installation path file you selected does not exist, please (re)install and try again"
        Exit
    }
    Write-Verbose $GhostScript
}
Process
{
    Write-Verbose $File.FullName
    $Destination = $File.FullName -replace ".pdf",  "-Converted.pdf"
    Write-Verbose $Destination
    if (-Not ($Overwrite))
    {
        if (Test-Path $Destination)
        {
            Write-Error "$Destination already exists, please use the Overwrite switch to force overwriting the destination file"
            Exit
        }
    }
    else
    {
        Remove-Item $Destination -Force -ErrorAction SilentlyContinue
    }
    $Arguments = '-sDEVICE=pdfwrite -dCompatibilityLevel=' + $CompatibilityLevel + ' -dPDFSETTINGS=/' + $CompressionLevel + ' '
    
    if ($AdvancedCompress)
    {
        $Arguments = $Arguments + '-dEmbedAllFonts=true -dSubsetFonts=true -dAutoRotatePages=/None -dColorImageDownsampleType=/Bicubic -dColorImageResolution=300 -dGrayImageDownsampleType=/Bicubic -dGrayImageResolution=300 -dMonoImageDownsampleType=/Bicubic -dMonoImageResolution=300 '
    }
    $Arguments = $Arguments + '-dNOPAUSE -dQUIET -dBATCH -sOutputFile="' + $Destination + '" "' + $File.FullName + '"'
   
    Write-Verbose $Arguments
    Start-Process $GhostScript -ArgumentList $Arguments -Wait -WindowStyle hidden
}
