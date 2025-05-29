#
# This is free and unencumbered software released into the public domain.
#
# Anyone is free to copy, modify, publish, use, compile, sell, or
# distribute this software, either in source code form or as a compiled
# binary, for any purpose, commercial or non-commercial, and by any
# means.
#
# In jurisdictions that recognize copyright laws, the author or authors
# of this software dedicate any and all copyright interest in the
# software to the public domain. We make this dedication for the benefit
# of the public at large and to the detriment of our heirs and
# successors. We intend this dedication to be an overt act of
# relinquishment in perpetuity of all present and future rights to this
# software under copyright law.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
# OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
# ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.
#
# For more information, please refer to <https://unlicense.org>
#

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
        if("outlook" -eq $component) {
            COMOffice("office") | Out-Null # outlook needs office interop
        }
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
