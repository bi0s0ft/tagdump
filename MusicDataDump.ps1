[CmdLetBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

$Global:shell = New-Object -ComObject Shell.Application
$Global:FileProperties = @{
    Name = 0;
    Size = 1;
    ModifedAt = 3;
    CreatedAt = 4;
    AccessAd = 5;
    Artists = 13;
    Album = 14;
    Year = 15;
    Genre = 16;
    Authors = 20;
    Title = 21;
    Track = 26;
    Length = 27;
    BitRate = 28
}

class FileMetaData {
    [string]$Path;
    [String]$Name;
    [String]$Size;
    [string]$ModifedAt;
    [string]$CreatedAt;
    [string]$AccessedAt;
    [String]$Artists;
    [String]$Album;
    [String]$Year;
    [String]$Genre;
    [String]$Authors;
    [String]$Title;
    [string]$Track;
    [String]$Length;
    [String]$BitRate;

    FileMetaData([hashtable]$metaData)
    {
        foreach ($key in $metaData.Keys) {
            if ($this.PSObject.Properties.Match($key).Count -gt 0) {
                $this.$key = $metaData[$key]
            }
        }
    }
}

function Get-Metadata 
{
    Param($path)
    
    $filesMetaData = @()
    $i = 0
    
    
    $files = $(ls -Recurse -File $path -Filter "*.mp3")
    write-host $files.Length " mp3s found"
    $lastDir = ""
    $folder = $Global:shell.NameSpace($path)
    foreach($file in $files) {
        if($file.DirectoryName -ne $lastDir) {
            $lastDir = $file.DirectoryName
            $folder = $Global:shell.NameSpace($file.DirectoryName)
        }
        
        $item = $folder.ParseName($file.Name)
        $metaData = @{}
        $metaData["Path"] = $file.DirectoryName
        foreach($property in $Global:FileProperties.Keys) {
            $data = $folder.GetDetailsOf($item, $Global:FileProperties[$property])
            $metaData[$property] = $data;
        }
        $filesMetaData += [FileMetaData]::New($metaData)
        $i++
        write-progress -activity "Reading file metadata ... " -status "$($file.FullName)" -PercentComplete (($i / $files.Length) * 100)
    }
    
    return $filesMetaData
}

Get-Metadata $SourceDirectory | export-csv $OutputPath