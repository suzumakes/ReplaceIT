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
  [string]$NewEnd,
  [switch]$Vb,
  [switch]$Log
)

Function SuperScript {
  Param (
    $FileContents,
    $Pattern,
    $Start,
    $End,
    $NewStart,
    $NewEnd,
    $File,
    $Vb
  )

  $MyMatches = $FileContents | Select-String -Pattern $Pattern -AllMatches

  $Update = $false

  ForEach ( $Item in $MyMatches ) {
    $Update = $true
    $ThisMatch = $Item -Match $Pattern

    $ReplaceWith = $NewStart + $Matches[1] + $NewEnd
    $FileContents = $FileContents -Replace [regex]::escape( $Item ), $ReplaceWith

    If ( $Vb ) {
      write-host "Replacing " $Item  " --> "  $ReplaceWith
    }

    If ( $Update ) {
      $FileContents|set-content $file
    }
  }
}
  
$FileContents = get-content $file -Erroraction silentlycontinue -ErrorVariable NotFound

If ( $NotFound.Count -eq "1" ) {
  write-host "$file not found -- Please check location and filename"
  break
}

$SearchString = "$Find"

if ( !( $AllMatches ) ) {

  $FileContents = ( [string]::join( "`r`n",$filecontents ) )

  $Found = $fileContents -match "$searchstring"

  If ( $Log ) {
    $logfile = ".\ReplaceITLog-" + $timestamp + ".log"
  }

  $newfile = $file + ".bak"

  If ( $Found ) {

    $replacewith = $FileContents -replace $SearchString,$Replace

    if ( !( test-path $newfile ) ) {
      copy-item $file $newfile 
      write-host "Saved backup to $newfile"
    }

    $replacewith|set-content $file
    
    If ( $Log ) {

      If ( $Vb ) {
        write-host "Replaced all $find with $Replace"
        "Replaced all $find with $Replace" | out-file $logfile -append
      } Else {
        "Replaced all $find with $Replace" | out-file $logfile -append
      }

    }

  } Else {

    If ( $Log ) {

       If ( $Vb ) {
        write-host ">$find not found"
        "$find not found" | out-file $logfile -append
       } Else {
        "$find not found" | out-file $logfile -append
       }

    }

  }
}

Else {
  write-host "Replacing Superscripts and Subscripts"
  SuperScript -FileContents $FileContents -Pattern $Pattern -Start $Start -End $End -NewStart $NewStart -NewEnd $NewEnd -File $File
}

