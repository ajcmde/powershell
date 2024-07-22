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

class COMReflection {
    static [Void]SetNamedProperty([System.Object]$object, [String] $propertyName, [System.Object]$propertyValue) {
        #[Void]$object.GetType().InvokeMember($propertyName, "SetProperty", $NULL, $object, $propertyValue)
        [Void]$object.GetType().InvokeMember($propertyName, [System.Reflection.BindingFlags]::SetProperty, $null, $object, $propertyValue)    
    }
    static [System.Object]GetNamedProperty([System.Object]$object, [String]$propertyName) {
        return $object.GetType().InvokeMember($propertyName, [System.Reflection.BindingFlags]::GetProperty, $null, $object, $null) 
    }
    static [System.Object]InvokeNamedParameter([System.Object]$object, [String]$method, [Hashtable]$parameter) {
        return $object.GetType().InvokeMember($method, [System.Reflection.BindingFlags]::InvokeMethod,
            $null,  # Binder
            $Object,  # Target
            ([Object[]]($parameter.Values)),  # Args
            $null,  # Modifiers
            $null,  # Culture
            ([String[]]($parameter.Keys))  # NamedParameters
            )
    }
    static [Void]SetNamedProperties([System.Object]$object, [Hashtable]$keys_values) {
        $keys_values.keys | ForEach-Object -Process {
            $key = $_
            $value = $keys_values[$key]
            try { 
                $expression = "`$Object.$key=`$value"
                Invoke-Expression  $expression
            }
            catch {
                Write-Warning "set_property failed: [$key] = [$value]: $_"
            }
        }
    }
} 
