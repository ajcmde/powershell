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
                #@@@ callstack
                Write-Warning "${callstack}: set_property failed: [$key] = [$value]: $_"
            }
        }
    }
} 
