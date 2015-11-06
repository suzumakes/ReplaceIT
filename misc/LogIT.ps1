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

LogIT -Log $Log -Text "Replaced $Item --> $ReplaceWith"
