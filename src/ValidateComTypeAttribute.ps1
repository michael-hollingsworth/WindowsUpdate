#using namespace Microsoft.VisualBasic
Add-Type -AssemblyName Microsoft.VisualBasic

<#
.SYNOPSIS
    Validates that a COM object is of a specific type.
.EXAMPLE
    ```PowerShell
    function Test-ComType {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true, Position = 0)]
            [ValidateComType(Type = ('IUpdateSession', 'IUpdateSession2', 'IUpdateSession3'))]
            [__ComObject]$ComObject
        )
    }
    ```

    Pass:
    ```PowerShell
    Test-ComType -ComObject (New-Object -ComObject Microsoft.Update.Session)
    ```

    Fail:
    ```PowerShell
    Test-ComType -ComObject (New-Object -ComObject Microsoft.Update.Searcher)
    ```
.NOTES
    This attribute can be defined as either `[ValidateComType('<TYPE_NAME>')]`, `[ValidateComType('<TYPE_NAME_1>', '<TYPE_NAME_2>', ...)]`, or `[ValidateCimClass(Type = ('<TYPE_NAME>', '<TYPE_NAME_2>', ...))]`.
.NOTES
    Author: Michael Hollingsworth
#>
class ValidateComTypeAttribute : System.Management.Automation.ValidateArgumentsAttribute {
    [ValidateNotNullOrEmpty()]
    [String[]]$Type

    ValidateComTypeAttribute() {
    }

    ValidateComTypeAttribute([String]$Type) {
        if ([String]::IsNullOrWhiteSpace($Type)) {
            throw [System.Management.Automation.ErrorRecord]::new(
                [System.Management.Automation.PSArgumentNullException]::new('Type'),
                'ArgumentIsNullOrWhiteSpace',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Type
            )
        }

        $this.Type = @(, $Type)
    }

    ValidateComTypeAttribute([String[]]$Type) {
        foreach ($string in $Type) {
            if ([String]::IsNullOrWhiteSpace($string)) {
                throw [System.Management.Automation.ErrorRecord]::new(
                    [System.Management.Automation.PSArgumentNullException]::new('Type'),
                    'ArgumentIsNullOrWhiteSpace',
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $string
                )
            }
        }

        $this.Type = $Type
    }

    [Void] Validate([Object]$arguments, [System.Management.Automation.EngineIntrinsics]$engineIntrinsics) {
        [String]$argumentType = [Microsoft.VisualBasic.Information]::TypeName($arguments)

        if (($this.Type.Count -eq 1) -and ($argumentType -eq $this.Type)) {
            return
        } elseif ($argumentType -in $this.Type) {
            return
        }

        throw [System.Management.Automation.ValidationMetadataException]::new(
            "Argument '$arguments' must be a valid '$($this.Type)' COM object.",
            [System.Management.Automation.PSArgumentException]::new("Argument '$arguments' must be a valid '$($this.Type)' COM object.")
        )
    }

    [String] ToString() {
        if ($this.Type.Count -eq 1) {
            return "[ValidateComTypeAttribute(Type = '$($this.Type)')]"
        }

        return "[ValidateComTypeAttribute(Type = ('$($this.Type -join "', '")'))]"
    }
}