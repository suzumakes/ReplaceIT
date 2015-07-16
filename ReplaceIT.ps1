<#

This is a Find and Replace PowerShell Script for cleaning Word-Filtered HTML.

Many Thanks for Michael Clark.

====================
TO DO

1. Limit recursion to 1 level
====================

#>

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

Function LogIt {

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

  $Log = ( Get-Variable -name LogIt -scope global ).value
  $TimeStamp = Get-Date -Format "MMdd_HHmm"

  $CurDir = $PWD.ToString()
  $LogFileName = $CurDir + "\ReplaceIt_" + $TimeStamp + ".log"

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
    $ThisMatch = $Item -Match $Pattern

    If ( $matches.Count -gt 0 ) {

      $ReplaceWith = $NewStart + $Matches[1] + $NewEnd
      $FileContents = $FileContents -Replace [regex]::escape( $Item ), $ReplaceWith

      LogIt -Log $Log -Text "Replaced $Item --> $ReplaceWith"

      If ( $Update ) {
        $FileContents|set-content $file
      }
    }
  }
}

$FileContents = get-content $file -Erroraction silentlycontinue -ErrorVariable NotFound

$SearchString = "$Find"

if ( !( $AllMatches ) ) {

  $FileContents = ( [string]::join( "`r`n",$filecontents ) )

  $Found = $fileContents -match "$searchstring"

  $newfile = $file + ".bak"

  If ( $Found ) {

    $replacewith = $FileContents -replace $SearchString, $Replace

    if ( !( test-path $newfile ) ) {
      copy-item $file $newfile 
      write-host "Saved backup to $newfile"
    }

    $replacewith|set-content $file
    
    LogIt -Log $Log -Text "Replaced all $find with $Replace"
    
  } Else {

    LogIt -Log $Log -Text "$find not found"

  }
}

Else {
  LogIt -Log $Log -Text "Replacing Superscripts and Subscripts"
  SuperScript -FileContents $FileContents -Pattern $Pattern -Start $Start -End $End -NewStart $NewStart -NewEnd $NewEnd -File $File -Log $Log
}

