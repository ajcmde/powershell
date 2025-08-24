# Snippets for PowerShell

# tags 
#POWERSHELL, #MICROSOFTOFFICE, #COM, #AUTOMATION

# description
provides snippets to make life easier with PowerShell and Office

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

## example (outlook: export contacts as XML for Fritz!Box import)
The output is written to the script folder. Note: This example needs to be adopted to your outlook configuration. 

    Import-Module "$PSScriptRoot/comreflection.ps1"
    Import-Module "$PSScriptRoot/comoffice.ps1"

    function XML_NUMBER {
        param (
            [Parameter(Mandatory=$true)]
            [object]$xml_telephony, 
            [Parameter(Mandatory=$true)]
            [string]$type,
            [Parameter(Mandatory=$true)]
            [string]$number
        )
        $xml_number = $xml_telephony.OwnerDocument.CreateElement("number")
        $xml_number.SetAttribute("type", $type) | Out-Null
        $xml_number.SetAttribute("prio", $xml_telephony.ChildNodes.Count ? "0" : "1") | Out-Null
        $xml_number.SetAttribute("id", $xml_telephony.ChildNodes.Count.ToString()) | Out-Null 
        $xml_number.InnerText = $number
        $xml_telephony.AppendChild($xml_number) | Out-Null
    }
    
    function XML_CONTACT {
        param (
            [Parameter(Mandatory=$true)]
            [object]$xml_phonebook,
            [Parameter(Mandatory=$true)]
            [object]$contact
        )
        $xml_contact = $xml_phonebook.OwnerDocument.CreateElement("contact")
        $xml_phonebook.AppendChild($xml_contact) | Out-Null
        $xml_category = $xml.CreateElement("category")
        $xml_contact.AppendChild($xml_category) | Out-Null
        $xml_category.InnerText = "0"
    
        $xml_person = $xml.CreateElement("person")
        $xml_contact.AppendChild($xml_person) | Out-Null
        $xml_realname = $xml.CreateElement("realName")
        $xml_person.AppendChild($xml_realname) | Out-Null
        $xml_realname.InnerText = $contact.FileAs
    
        $xml_telephony = $xml.CreateElement("telephony")
        $xml_contact.AppendChild($xml_telephony) | Out-Null
    
        if ($contact.PrimaryTelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "home" -number $contact.PrimaryTelephoneNumber
        }
        if ($contact.HomeTelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "home" -number $contact.HomeTelephoneNumber
        }
        if ($contact.Home2TelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "home" -number $contact.Home2TelephoneNumber
        }
        if ($contact.MobileTelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "mobile" -number $contact.MobileTelephoneNumber
        }
        if ($contact.BusinessTelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "work" -number $contact.BusinessTelephoneNumber
        }
        if ($contact.Business2TelephoneNumber) {
            XML_NUMBER -xml_telephony $xml_telephony -type "work" -number $contact.Business2TelephoneNumber
        }
        $xml_telephony.SetAttribute("nid", $xml_telephony.ChildNodes.Count.ToString()) | Out-Null

        # services
        $xml_setup = $xml.CreateElement("setup")
        $xml_contact.AppendChild($xml_setup) | Out-Null
    
        $xml_features = $xml.CreateElement("features")
        $xml_features.SetAttribute("doorphone", "0") | Out-Null
        $xml_contact.AppendChild($xml_features) | Out-Null
    
        $lastupdate = Get-Date($contact.LastModificationTime).toUniversalTime() -UFormat %s
        $xml_modtime = $xml.CreateElement("modtime")
        $xml_modtime.InnerText = $lastupdate
        $xml_contact.AppendChild($xml_modtime) | Out-Null
    
        $xml_uniqueid = $xml.CreateElement("uniqueid")
        $xml_uniqueid.InnerText = $contact.ConversationID
        $xml_contact.AppendChild($xml_uniqueid) | Out-Null
    }
    
    
    # main script
    $outlook = New-Object COMOffice("outlook")
    $namespace = $outlook.Application.GetNameSpace("MAPI")
    
    $xml = New-Object System.Xml.XmlDocument
    $xml_Declaration = $xml.CreateXmlDeclaration("1.0", "utf-8", $null)
    $xml.AppendChild($xml_Declaration) | Out-Null
    
    $xml_phonebooks = $xml.CreateElement("phonebooks") 
    $xml.AppendChild($xml_phonebooks) | Out-Null
    $xml_phonebook = $xml.CreateElement("phonebook")
    $xml_phonebooks.AppendChild($xml_phonebook) | Out-Null
    
    # adopt to your outlook folder structure (email address and contact folder name)
    $namespace.Folders("emailname@domain.com").Folders("Contakte").Items |
        ForEach-Object -Process {
            XML_CONTACT -xml_phonebook $xml_phonebook -contact $_
        }
    
    $xml.OuterXml > "$PSScriptRoot\contacts.xml" 
    $outlook = $null
    
    "working directory: " + $PSScriptRoot 
