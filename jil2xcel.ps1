# jil2xcel.ps1
# Version 1.1
# @hanson0x89

param (
  [Parameter(Mandatory=$true)][string]$jilFile
)
 
function main
{
  $oldVerbose = $VerbosePreference
  $VerbosePreference = "continue"

  If (!(Test-Path $jilFile)) 
  {
    Write-Verbose "Input file not found: $jilFile. Exiting..."
    exit(127)
  }
 
    # Get JIL and strip comment lines
  $jilDB = Get-Content $jilFile | select-string -pattern "`/`* -------" -notmatch | Out-String

    # Create array of jobs
  $jobs = $JilDB -split 'insert_job: '

    # Create Excel file 
  $ExcelObject = new-Object -comobject Excel.Application 
  $ExcelObject.visible = $false 
  $ExcelObject.DisplayAlerts =$false

  $ActiveWorkbook = $ExcelObject.Workbooks.Add() 
  $ActiveWorksheet = $ActiveWorkbook.Worksheets.Item(1)
 
    # Write each job to Excel file
  $Row = 1
  for ($i=0; $i -lt $jobs.Length; $i++) {
    $jobOneLiner =""
    $jobs[$i].split("`r`n") | ForEach-Object {
    $jobOneLiner += "$_ " 
    }
    $ActiveWorksheet.Cells.Item($Row, 1) = $jobOneLiner
    $Row++
  }
    # Save Excel file and clean up
  $scriptDir = (Get-Location).Path
  $date= get-date -format "yyyyMMddHHss"
  $ActiveWorkbook.SaveAs("$scriptDir\jobs_$date.xlsx")
  $ExcelObject.Quit()
  $ExcelObject = $Null
} # main

# Entry point
main