#######################################################################################
# Markdown readme utility
# Copies and edits a template for a patch
#
# Syntax:
#
# .\CreateMarkdown.ps1 -Patch <patch name> [-Template <patch type>]
#
# Run the script with no parameters to view help and examples.
#
# The script carries out the following:
# 
# 1. Checks that the patch folder exists.
# 2. Checks if the readme HTML has changes in svn from anyone other than the patch 
#    system.
# 3. Deletes the readme HTML if it's a legacy HTML version.
# 3. Checks that the Markdown file doesn't already exist.
# 4. Copies all the files from the appropriate template folder to the patch folder.
# 5. Renames the template file to be the same as the patch.
# 6. Rewrites the readme markdown with the patch number replacing the template 
#    placeholder.
#
#######################################################################################


param (
    [string]$Patch = '',   # Patch name
    [string]$Template = '' # Patch type
)

# Set the error action to Stop - file operations generate non-terminating errors by default
$ErrorActionPreference = "Stop"
# Set the root of the Subversion repository
$svnRoot = "C:\Work\Subversion\Patches\Trunk\"
# Set debug output
$debug = $true

if ($Patch -eq '') {
    $helpText = @"
Markdown readme template script

Syntax:
CreateMarkdown -Patch <patch name> [-Template <patch type>]

For example:
CreateMarkdown -Patch HOTFIX-12.11.0.9
CreateMarkdown -Patch ACUANT-1.0.3 -Template MODULE
CreateMarkdown -Patch CAMIGRATE-1.0.5 -Template U

The -Template parameter is not required if the patch is a HOTFIX, CONFIG, or UMC.

This script supports the following types of patch:

* HOTFIX - MyID server hotfix; for example, HOTFIX-12.11.0.9.
* CONFIG - MyID server configuration update; for example, CONFIG-3021.1.3.
* MODULE - MyID additional module; for example, ACUANT-1.0.3.
* MSIMODULE - MyID additional module that uses the .msi installer instead of PowerShell; for example, ADJUDICATION-2.2.0.
* U - MyID update with manually-applied updates; for example CAMIGRATE-1.0.5.
* UMC - MyID Client Components; for example, UMC-49.0.1000.1.

"@
    Write-Host "`r`n$helpText" -ForegroundColor Yellow
    Write-Host "Press any key to continue..." -ForegroundColor Yellow
    [void][System.Console]::ReadKey($true)
    exit
}

function Show-Error {
    Write-Host "An unexpected error occurred.`r`n" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    Write-Host "`r`nExiting.`r`n" -ForegroundColor Red
    Write-Host "Press any key to continue..." -ForegroundColor Red
    [void][System.Console]::ReadKey($true)
    exit
}
# Function to call svn and return its results
function Read-Subversion($svnArguments){
    # Set the options for the process
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "svn"
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = $svnArguments
    # Start the process
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    # Create a custom object to return
    $svnResults = "" | Select-Object -Property Output,Error,ExitCode
    $svnResults.Output = $p.StandardOutput.ReadToEnd()
    $svnResults.Error = $p.StandardError.ReadToEnd()
    $svnResults.ExitCode = $p.ExitCode
    return $svnResults
}

if ($Template -eq ''){
    $Template = $Patch.Substring(0,[math]::max($Patch.IndexOf('-'),0))
}
$Template = $Template.ToUpper()
$Patch = $Patch.ToUpper()

Write-Host "`r`nMarkdown readme template utility`r`n" -ForegroundColor Yellow

switch ($Template)
{
    "HOTFIX" {
        Write-Host "Using the HOTFIX template.`r`n" -ForegroundColor Yellow
        $TemplateX = "HOTFIX-x.x.x.x"
    }
    "CONFIG" {
        Write-Host "Using the CONFIG template.`r`n" -ForegroundColor Yellow
        $TemplateX = "CONFIG-xxxx.x.x"
    }
    "MODULE" {
        Write-Host "Using the MODULE template.`r`n" -ForegroundColor Yellow
        $TemplateX = "MODULE-x.x.x"
    }
    "MSIMODULE" {
        Write-Host "Using the MSIMODULE template.`r`n" -ForegroundColor Yellow
        $TemplateX = "MSIMODULE-x.x.x"
    }
    "U" {
        Write-Host "Using the U template.`r`n" -ForegroundColor Yellow
        $TemplateX = "Uxxxxxxxx"
    }
    "UMC" {
        Write-Host "Using the UMC template.`r`n" -ForegroundColor Yellow
        $TemplateX = "UMC-x.x.x.x"
    }
    Default {
        Write-Host "Template type $Template is not recognized.`r`nIf the patch is not a HOTFIX, CONFIG, or UMC, you must specify the -Template parameter.`r`n" -ForegroundColor Red
        Write-Host "Press any key to continue..." -ForegroundColor Red
        [void][System.Console]::ReadKey($true)
        exit
    }
}

# Generate paths from the parameters
$TemplateFolder = $svnRoot + "_Templates\" + $Template
$PatchFolder = $svnRoot + $Patch
$PatchMarkdown = $PatchFolder + "\" + $Patch + ".md"
$PatchReadme = $PatchFolder + "\release\" + $Patch + "_readme.html"

if ($debug){
    Write-Host "Patch: $Patch `r`nTemplate: $Template`r`nTemplate placeholder: $TemplateX`r`nPatch folder: $PatchFolder`r`nTemplate folder: $TemplateFolder`r`nReadme: $PatchReadme`r`nMarkdown: $PatchMarkdown" -ForegroundColor Blue
}

# Check that the patch folder exists
Write-Host "Checking that the patch folder exists..." -ForegroundColor Yellow
if (Test-Path -Path $PatchFolder) {
    Write-Host "$PatchFolder exists.`r`n" -ForegroundColor Green
}
else {
    Write-Host "The patch folder $PatchFolder does not exist.`r`n" -ForegroundColor Red
    Write-Host "Exiting.`r`n" -ForegroundColor Red
    Write-Host "Press any key to continue..." -ForegroundColor Red
    [void][System.Console]::ReadKey($true)
    exit
}

# Check that the readme exists but does not have any changes
Write-Host "Checking for an existing readme..." -ForegroundColor Yellow
if (Test-Path -Path $PatchReadme) {
    # Get the most recent Subversion log for the readme
    $svnResults = Read-Subversion("log --limit 1 $PatchReadme")
    if ($debug){
        Write-Host "svn output:" $svnResults.Output "`r`nsvn error:" $svnResults.Error "`r`nsvn exit code:" $svnResults.ExitCode -ForegroundColor Blue
    }

    # If the svn process resulted in an error, quit
    if ($svnResults.ExitCode -ne 0){
        Write-Host "An error occurred reading the Subversion log." -ForegroundColor Red
        Write-Host $svnResults.Error -ForegroundColor Red
        Write-Host "Exiting.`r`n" -ForegroundColor Red
        Write-Host "Press any key to continue..." -ForegroundColor Red
        [void][System.Console]::ReadKey($true)
        exit
    }

    # If the most recent commit was by Build1389SVC with a comment of "Patch Approval" or "Jenkins Auto Commit" you can safely delete it
    if (($svnResults.Output.IndexOf("Build1389SVC") -ne -1) -and (($svnResults.Output.IndexOf("Patch Approval") -ne -1) -or ($svnResults.Output.IndexOf("Jenkins Auto Commit") -ne -1))){
        Write-Host "Default obsolete HTML readme detected. Deleting..." -ForegroundColor Yellow
        try {
            Remove-Item $PatchReadme
        }
        catch {
            Show-Error
        }
        Write-Host "Old readme deleted.`r`n" -ForegroundColor Green
    }
    else {
        # Otherwise someone or something else was the last to commit the readme, so it's not safe to continue
        Write-Host "An HTML readme with changes has been detected. Cannot continue.`r`n" -ForegroundColor Red
        Write-Host "Exiting.`r`n" -ForegroundColor Red
        Write-Host "Press any key to continue..." -ForegroundColor Red
        [void][System.Console]::ReadKey($true)
        exit
    }
}
else {
    # If there isn't a readme it's safe to continue
    Write-Host "The patch readme $PatchReadme does not exist.`r`n" -ForegroundColor Green
}

# Check if the Markdown file already exists, and if so, quit.
Write-Host "Checking whether the Markdown readme already exists..." -ForegroundColor Yellow
if (Test-Path -Path $PatchMarkdown) {
    Write-Host "$PatchMarkdown already exists.`r`n" -ForegroundColor Red
    Write-Host "Exiting.`r`n" -ForegroundColor Red
    Write-Host "Press any key to continue..." -ForegroundColor Red
    [void][System.Console]::ReadKey($true)
    exit
}
else {
    # Safe to continue
    Write-Host "The Markdown readme $PatchMarkdown does not exist.`r`n" -ForegroundColor Green
}

# Copy all the files from the appropriate template folder to the patch folder
Write-Host "Copying the template files to the patch folder..." -ForegroundColor Yellow

try {
    Copy-Item -Path $TemplateFolder\* -Destination $PatchFolder -Recurse -Force
}
catch {
    Show-Error
}

Write-Host "Template files copied.`r`n" -ForegroundColor Green

# Rename the template file to be the same as the patch
Write-Host "Renaming the template file..." -ForegroundColor Yellow

try {
    Rename-Item -Path "$PatchFolder\$TemplateX.md" -NewName "$Patch.md"
}
catch {
    Show-Error
}
Write-Host "Template file renamed.`r`n" -ForegroundColor Green

# Rewrite the readme markdown with the patch number replacing the template placeholder
Write-Host "Editing the patch number in the readme markdown..." -ForegroundColor Yellow

try {
    (Get-Content -Path "$PatchFolder\$Patch.md") -replace $TemplateX,$Patch | Set-Content -Path "$PatchFolder\$Patch.md"
}
catch {
    Show-Error
}
Write-Host "Patch number updated.`r`n" -ForegroundColor Green
Write-Host "Press any key to continue..." -ForegroundColor Green
[void][System.Console]::ReadKey($true)
