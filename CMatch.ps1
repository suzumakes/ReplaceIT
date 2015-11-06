Param (
    [string]$File,
    [string]$Find,
    [string]$Replace,
    [switch]$AllMatches
)

Function LogIT {

    Param (
        [bool]$Log,
        $Text
    )

    $LogFileName = ( Get-Variable -Name LogFile -Scope Global ).Value

    If ( $Log ) {
        $DateTime = Get-Date -Format "MM/dd/yyyy HH:mm"
        $Output = $DateTime + "`t" + $Text
        $Output | Out-File $LogFileName -Append
    }
}

$Log = ( Get-Variable -name LogIT -scope Global ).value
$TimeStamp = Get-Date -Format "dd_HH"

$CurDir = $PWD.ToString()
$LogFileName = $CurDir + "\ReplaceIT_" + $TimeStamp + ".log"

Set-Variable -Name LogFile -Value $LogFileName -Scope Global

$FileContents = get-content $File -Erroraction silentlycontinue -ErrorVariable NotFound

$SearchString = "$Find"

if ( !( $AllMatches ) ) {

    $FileContents = ( [string]::join( "`r`n",$FileContents ) )

    $Found = $FileContents -cmatch "$SearchString"

    $NewFile = $File + ".bak"

    If ( $Found ) {

        $ReplaceWith = $FileContents -replace $SearchString, $Replace

        if ( !( test-path $NewFile ) ) {
            copy-item $File $NewFile
            write-host "saved backup to $NewFile"
        }

        $ReplaceWith|set-content $File
        LogIT -Log $Log -Text "replaced all $Find with $Replace"

    } Else {
        LogIT -Log $Log -Text "$Find not found"
    }
}
