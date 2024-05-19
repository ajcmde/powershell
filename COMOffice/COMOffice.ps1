class COMOffice {
    [System.Collections.Hashtable]$Enum
    [String]$Component
    hidden [System.Object]$_Application = $null

    static COMOffice() {
        Update-TypeData -Force -TypeName $([COMOffice].Name) -MemberType ScriptProperty  -MemberName 'Application' -Value { 
                # property get _Application
                if($null -eq $this._Application) {
                    #create new application object
                    if("excel" -eq $this.Component) {
                        try {
                            $this._Application = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application') 
                        }
                        catch { }
                    }
                    if($null -eq $this._Application) {
                        $this._Application = New-Object -ComObject $($this.Component + '.application') 
                    }
                } 
                $this._Application
            } -SecondValue { # property set _Application
                if($null -ne $args[0]) {
                    throw "illegal value"
                }
                if($null -ne $this._Application) {
                    if("excel" -eq $This.Component) {
                        if($this._Application.Workbooks.Count() -eq 0) {
                            $this._Application.Quit()
                        }
                    }
                    else {
                        $this._Application.Quit()
                    }
                }                
                $this._Application = $null
                $this._Application = $null
                [System.GC]::Collect()
            }
    }

    COMOffice([string]$component) {
        $this.Enum = @{}
        $this.Component = $component
        if("office" -eq $component) {
            $path = "$env:WINDIR\assembly\GAC_MSIL\office\*\office.dll"
        }
        else {
            $path = "$env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.$component\*\Microsoft.Office.Interop.$component.dll"
        }
        Add-Type -Path $path -PassThru |
        Where-Object { $_.BaseType -eq [System.Enum] } | 
        ForEach-Object -Process { $_.GetEnumValues() } | ForEach-Object -Process {
            if([string]::IsNullOrEmpty($this.Enum[[string]$_])) { 
                $this.Enum.Add([string]$_, [int]$_) 
            } elseif ($this.Enum[[string]$_] -ne [int]$_) {
                Write-Warning "$([COMOffice].Name).Enum duplicate hashes: $([string]$_): $($this.Enum[[string]$_]) >< $([int]$_)"
            }
        }                   
    }
}
