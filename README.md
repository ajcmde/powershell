# Snippets for Powershell

# tags 
#POWERSHELL, #MICROSOFTOFFICE, #COM, #AUTOMATION

# description
provides snippets to make life easier with PowerScript

# COMReflection
TBD

## example:

    $wb = [COMReflection]::InvokeNamedParameter($xcl.Application.Workbooks, "Open", @{
        "Filename" = "$pathexcel"; 
        "ReadOnly" = -1;
    })

# COMOffice
TBD

## example:
 
    $wrd = New-Object COMOffice("word")
    $document = $wrd.Application.Documents.Add()
    $wrd.Application.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $document.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue   
