Param (
    [string]$File,
    [string]$Find,
    [string]$Replace,
    [switch]$AllMatches,
    [string]$Pattern,
    [string]$Start,
    [string]$End,
    [string]$NewStart,
    [string]$NewEnd
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

Function SuperScript {

    Param (
        $FileContents,
        $Pattern,
        $Start,
        $End,
        $NewStart,
        $NewEnd,
        $File
    )

    $MyMatches = $FileContents | Select-String -Pattern $Pattern -AllMatches

    $Update = $false

    ForEach ( $Item in $MyMatches ) {

        $Update = $true
        $ThisMatch = $Item -match $Pattern

        If ( $matches.Count -gt 0 ) {

            $ReplaceWith = $NewStart + $Matches[1] + $NewEnd
            $FileContents = $FileContents -replace [regex]::escape( $Item ), $ReplaceWith

            LogIT -Log $Log -Text "replaced $Item --> $ReplaceWith"
        }
    }

    If ( $Update ) {
        $FileContents|set-content $File
    }
}

$FileContents = get-content $File -Erroraction silentlycontinue -ErrorVariable NotFound

$SearchString = "$Find"

if ( !( $AllMatches ) ) {

    $FileContents = ( [string]::join( "`r`n",$FileContents ) )

    $Found = $FileContents -match "$SearchString"

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

Else {
    LogIT -Log $Log -Text "start of replacing super/subscripts"
    SuperScript -FileContents $FileContents -Pattern $Pattern -Start $Start -End $End -NewStart $NewStart -NewEnd $NewEnd -File $File -Log $Log
}
