# Snippets for PowerShell

# tags 
#POWERSHELL, #MICROSOFTOFFICE, #COM, #AUTOMATION

# description
provides snippets to make life easier with PowerShell

# COMReflection
The class provides the following static functions

    static [void] SetNamedProperty([System.Object]$object, [String] $propertyName, [System.Object]$propertyValue)

    static [System.Object]GetNamedProperty([System.Object]$object, [String]$propertyName) 

    static [System.Object]InvokeNamedParameter([System.Object]$object, [String]$method, [Hashtable]$parameter) 

    static [Void]SetNamedProperties([System.Object]$object, [Hashtable]$keys_values)

## example:

    $wb = [COMReflection]::InvokeNamedParameter($xcl.Application.Workbooks, "Open", @{
        "Filename" = "$pathexcel"; 
        "ReadOnly" = -1;
    })

# COMOffice
The constructor adds the object model for the selected office component ("excel", "word", "powerpoint", "outlook" and "office") to PowerShell and add all enumerations values to the object's Enum property.

    COMOffice([string]$component)

The class provides the following properties

    [System.Object] Application
    [System.Collections.Hashtable] Enum

If the Apllication property is read the first time, the application object of the instantiated Microsoft Office application will be created. With Excel an existing instance (which acts as a server) will be leveraged as application object. The component "office" has no Application object.  
The application object will be closed when $null is assigned. An excepetion will be thrown if any other value than $null is written.

The Enum property will allow access to any enumarations of the instantiated object model through the hashtable with the name of the value as key.  

## example:
 
    $wrd = New-Object COMOffice("word")
    $document = $wrd.Application.Documents.Add()
    $wrd.Application.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $document.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue  

    ...

    $wrd.Application = $null
