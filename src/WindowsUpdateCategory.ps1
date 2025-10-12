enum WindowsUpdateType {
    None
    Product
    UpdateClassification
}

class WindowsUpdateCategory {
    [String]$Name
    [WindowsUpdateType]$Type
    [String]$Description
    [Guid]$CategoryID

    WindowsUpdateCategory() {
        $this.Name = $null
        $this.Type = [WindowsUpdateType]::None
        $this.Description = $null
        $this.CategoryID = [Guid]::Empty
    }

    WindowsUpdateCategory([__ComObject]$UpdateCategory) {
        [ValidateComTypeAttribute('ICategory')]
        [__ComObject]$UpdateCategory = $UpdateCategory

        $this.Name = $UpdateCategory.name
        $this.Type = [Enum]::Parse([WindowsUpdateType], $UpdateCategory.Type)
        $this.Description = $UpdateCategory.Description
        $this.CategoryID = [Guid]::Parse($UpdateCategory.CategoryID)
    }

    WindowsUpdateCategory([String]$Name, [String]$Description, [WindowsUpdateType]$Type, [Guid]$CategoryID) {
        $this.Name = $Name
        $this.Type = $Type
        $this.Description = $Description
        $this.CategoryID = $CategoryID
    }
}