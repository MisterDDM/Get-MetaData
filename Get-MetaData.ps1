Function Get-MetaData
{
    [CmdletBinding()]
    Param
    (
        [Parameter( Mandatory = $true, 
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true, 
                    ValueFromRemainingArguments = $false, 
                    Position=0 )]
        [string[]]$Path,

        [Parameter( Mandatory = $false, 
                    ValueFromPipeline = $true,
                    ValueFromPipelineByPropertyName = $true, 
                    ValueFromRemainingArguments = $false, 
                    Position=1 )]  
        [string[]]$FileExtension = ('.avi', '.iso', '.mkv', '.m2ts', '.m4v', '.mp4')
    )

    Begin
    {}
    Process
    {
        if ( $PSBoundParameters['Path'] )
        {
            $Path | ForEach-Object {
    
                $objShell = New-Object -ComObject Shell.Application 
                $objFolder = $objShell.namespace($_) 
    
                $Files = $FileExtension | ForEach-Object { 
                    $extension = $_ 
                    $objFolder.Items() | Where-Object { $_.Path -match ([Regex]::Escape($extension)) }
                }

                $Files | ForEach-Object {
                    $File = $_
                    $Hash = @{}
                    0..400 | ForEach-Object {
                        if ($($objFolder.GetDetailsOf($objFolder.Items, $_)))
                        {
                            $PropertyName = $($objFolder.GetDetailsOf($objFolder.Items, $_)) 

                            if ($($objFolder.GetDetailsOf($File, $_)))
                            {
                                $PropertyValue = $($objFolder.GetDetailsOf($File, $_))
                                $Hash[$PropertyName] = $PropertyValue
                            }
                        }
                    }
                }
                $FileMetaData = New-Object psobject -Property $Hash
                $Hash.Clear()
                Write-Output $FileMetaData
            }
        }
    }
    End
    {}
}