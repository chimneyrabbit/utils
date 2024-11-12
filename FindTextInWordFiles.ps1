#######################################################################################
# Microsoft Word search utility
# Searches for text in multiple Word documents.
#
# Syntax:
#
# .\FindTextInWordFiles.ps1 -FindText "<text to search>" -Path "<folder to search>" 
# [-Recurse] [-MatchCase] [-MatchWholeWords] [-MatchWildcards]
#
# where:
#
# * -FindText "<text to search>" - the text you want to search for, enclosed in double quotes.
# * -Path "<folder to search>" - the folder you want to search, enclosed in double quotes.
# * -Recurse - search all subfolders.
# * -MatchCase - case sensitive search.
# * -MatchWholeWords - search only whole words instead of partial words.
# * -MatchWildcards - use wildcards like * in the search. Wildcard searches are automatically 
#   case sensitive.
#
# For example:
# .\FindTextInWordFiles.ps1 -FindText "smart card" -Path "C:\Work\Subversion\Patches\Trunk" 
# -MatchCase -Recurse
#
# You must specify at least the -FindText and the -Path.
#
#######################################################################################

# Set up command line parameters
param (
    [string]$FindText = '',           # The text for which to search
    [string]$Path = '',               # The folder to search
    [switch]$Recurse=$false,          # Search subfolders switch
    [switch]$MatchCase=$false,        # Case-sensive switch
    [switch]$MatchWholeWord=$false,   # Whole words only switch
    [switch]$MatchWildcards=$false    # Use wildcards switch
)

# "Press any key to continue" function with variable colours
function Wait-For-Key {
    param (
        $status = "Message"
    )
    switch ($status){
        "Message" {$colour = "Yellow"}
        "OK" {$colour = "Green"}
        "Error" {$colour = "Red"}
    }
    Write-Host "`r`nPress any key to continue..." -ForegroundColor $colour
    [void][System.Console]::ReadKey($true)
}

# If no text or folder is specified, display the help screen
if (($FindText -eq '') -or ($Path -eq '')){
    $helpText = @"
Microsoft Word search utility
Searches for text in multiple Word documents.

Syntax:

.\FindTextInWordFiles.ps1 -FindText "<text to search>" -Path "<folder to search>" [-MatchCase] [-Recurse] [-MatchWholeWord] [-MatchWildcards]

where:

* -FindText "<text to search>" - the text you want to search for, enclosed in double quotes.
* -Path "<folder to search>" - the folder you want to search, enclosed in double quotes.
* -MatchCase - case sensitive search.
* -Recurse - search all subfolders.
* -MatchWholeWord - search only whole words instead of partial words.
* -MatchWildcards - use wildcards like * in the search. Wildcard searches are automatically case sensitive.

For example:
.\FindTextInWordFiles.ps1 -text "smart card" -folder "C:\Work\Subversion\Patches\Trunk" -MatchCase -Recurse

You must specify at least the -FindText and the -Path.
"@
    Write-Host "`r`n$helpText" -ForegroundColor Yellow
    Wait-For-Key -status "Message"
    Exit
}

# Set parameters for the file search
if ($Recurse.IsPresent){$recurseSwitch = " -Recurse"} else {$recurseSwitch = ''}

# Trim trailing slash from path
if ($Path.Substring($Path.Length-1,1) -eq '\') {$Path = $Path.Substring(0,$Path.Length-1)}

# Check that the path to search exists
if (-Not (Test-Path -Path $Path)){
    Write-Host "The path $Path does not exist" -ForegroundColor Red
    Wait-For-Key -status "Error"
    Exit
}

# Create a Word object. Microsoft Word must be installed.
try { $word = New-Object -ComObject Word.Application }
catch {
    Write-Host "`r`nAn error occurred creating a Word object." -ForegroundColor Red
    Wait-For-Key -status "Error"
    Exit
}

# Provide feedback to the user on what the script is doing
Write-Host "`r`nChecking $Path for the text `"$FindText`"..." -ForegroundColor Yellow
if ($MatchCase.IsPresent -or $MatchWildcards.IsPresent) {Write-Host "* Case sensitive search." -ForegroundColor Yellow}
else {Write-Host "* Case insensitive search." -ForegroundColor Yellow}
if ($MatchWholeWord.IsPresent) {Write-Host "* Searching whole words." -ForegroundColor Yellow}
else {Write-Host "* Searching partial words." -ForegroundColor Yellow}
if ($MatchWildcards.IsPresent) {Write-Host "* Using wildcards." -ForegroundColor Yellow}
else {Write-Host "* No wildcards." -ForegroundColor Yellow}
if ($recurse.IsPresent) {Write-Host "* Searching subfolders." -ForegroundColor Yellow}
else {Write-Host "* Searching current folder only." -ForegroundColor Yellow}

# Set the start time
$starttime = Get-Date

# Get a list of all files to process
Write-Host "`r`nGetting a list of files to search..." -ForegroundColor Yellow
$getChild = "Get-ChildItem -Path `"$Path\*`"$recurseSwitch -Include *.doc* -Exclude ~*.doc"
$docs = Invoke-Expression $getChild
$count = $docs.Count
$total = 0

# Display how many files the script is going to search
$searchingText = "Searching " + $count + " files."
if ($count -eq 1){$searchingText = $searchingText.Replace('files','file')}
Write-Host $searchingText -ForegroundColor Yellow

# Search each document in turn
$i=1
foreach ($doc in $docs)
{
    # Display a progress bar
    Write-Progress -Activity "Searching files" -status "Checking $($doc.FullName)" -PercentComplete ($i/$count * 100)
    # Open the document read-only and search it
    # The syntax is: Documents.Open(FileName,ConfirmConversions,ReadOnly).Content.Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards)
    if ($word.Documents.Open($doc.FullName,$false,$true).Content.Find.Execute($FindText, $MatchCase.IsPresent, $MatchWholeWord.IsPresent, $MatchWildcards.IsPresent))
    {
        # We've found the text in the document
        Write-Host "$doc contains `"$FindText`"" -ForegroundColor Green
        $total++
    }
    # Close the document once we're done with it
    $word.Application.ActiveDocument.Close()
    $i++
}

# Display the total number of files that contain the search text
$foundText = "`r`nFound `"" + $FindText+ "`" in " + $total + " files."
if ($total -eq 1){$foundText = $foundText.Replace('files','file')}
Write-Host $foundText -ForegroundColor Yellow

# Close the Word object to clean up
$word.Quit()

# Display how long the search took
$endtime = Get-Date
$duration = $endtime - $starttime 
$durationText = "Search completed in " + $duration.Hours + " hours, " + $duration.Minutes + " minutes, and " + $duration.Seconds + " seconds."
if ($duration.Hours -eq 1){$durationText = $durationText.Replace('hours','hour')}
if ($duration.Minutes -eq 1){$durationText = $durationText.Replace('minutes','minute')}
if ($duration.Seconds -eq 1){$durationText = $durationText.Replace('seconds','second')}
Write-Host $durationText -Foregroundcolor Green

# And we're done
Wait-For-Key -status "Message"