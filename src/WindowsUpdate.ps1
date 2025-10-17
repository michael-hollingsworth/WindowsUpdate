class WindowsUpdate {
    [WindowsUpdateCategory[]]$Category
    [ValidateComTypeAttribute(Type = ('IUpdate', 'IUpdate2', 'IUpdate3', 'IUpdate4', 'IUpdate5'))]
    hidden [__ComObject]$_updateObject

    WindowsUpdate([__ComObject]$Update) {
        $this._updateObject = $Update
        $this.Category = foreach ($category in $this._updateObject.Categories) {
            [WindowsUpdateCategory]::new($category)
            $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($category)
        }
    }

    [Void] AcceptEula() {
        if ($null -eq $this._updateObject) {
            return
        }

        if ($this._updateObject.EulaAccepted) {
            return
        }

        Write-Verbose -Message "Accepting the EULA for the update [$($this.Title)] with the following terms and conditions: `r`n$($this.EulaText)"

        $this._updateObject.AcceptEula()
        return
    }

    [Void] Download() {
        if ($null -eq $this._updateObject) {
            return
        }

        if ($this._updateObject.IsDownloaded) {
            Write-Verbose -Message "this update is already downloaded"
            return
        }

        try {
            [__ComObject]$private:updateSession = New-Object -ComObject Microsoft.Update.Session
            [__ComObject]$updateColl = New-Object -ComObject Microsoft.Update.UpdateColl
            [__ComObject]$updateDownloader = $updateSession.CreateUpdateDownloader()

            $null = $updateColl.Add($this._updateObject)

            $updateDownloader.Updates = $updateColl
            [__ComObject]$downloadResults = $updateDownloader.Download()
        } finally {
            if ($null -ne $updateSession) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateSession)
            }
            if ($null -ne $updateColl) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateColl)
            }
            if ($null -ne $updateDownloader) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateDownloader)
            }
            if ($null -ne $downloadResults) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($downloadResults)
            }
        }

        return
    }

    [Void] Install() {
        $this.Install($false)
        return
    }

    [Void] Install([Boolean]$Force) {
        [WindowsUpdate]::Install($this, $Force)
        return
    }

    [String] ToString() {
        return $this._updateObject.Identity.UpdateID
    }

    static [WindowsUpdate[]] Get() {
        return ([WindowsUpdate]::Search('IsInstalled = 0 AND IsHidden = 0'))
    }

    static [WindowsUpdate[]] GetInstalledUpdates() {
        #TODO: Look into the other ways of searching for installed updates
        return ([WindowsUpdate]::Search('Installed = 1'))
    }

    static [WindowsUpdate[]] Search([String]$SearchQuery) {
        if ([String]::IsNullOrWhiteSpace($SearchQuery)) {
            throw [System.Management.Automation.ErrorRecord]::new(
                [System.Management.Automation.PSArgumentNullException]::new('SearchQuery'),
                'ArgumentIsNullOrWhiteSpace',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $SearchQuery
            )
        }

        try {
            [__ComObject]$private:updateSession = New-Object -ComObject Microsoft.Update.Session
            [__ComObject]$updateSearcher = $updateSession.CreateUpdateSearcher()

            Write-Verbose -Message "Searching for Windows Updates with the search query [$SearchQuery]."

            [__ComObject]$searchResults = $updateSearcher.Search($SearchQuery)

            return $(foreach ($update in $searchResults.Updates) {
                [WindowsUpdate]::new($update)
            })
        } finally {
            if ($null -ne $updateSession) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateSession)
            }
            if ($null -ne $updateSearcher) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateSearcher)
            }
            if ($null -ne $searchResults) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($searchResults)
            }
        }
    }

    static [Void] Install([WindowsUpdate[]]$WindowsUpdate) {
        [WindowsUpdate]::Install($WindowsUpdate, $false)
    }

    static [Void] Install([WindowsUpdate[]]$WindowsUpdate, [Boolean]$Force) {
        if ($null -eq $WindowsUpdate) {
            return
        }

        try {
            [__ComObject]$private:updateSession = New-Object -ComObject Microsoft.Update.Session
            [__ComObject]$updateColl = New-Object -ComObject Microsoft.Update.UpdateColl

            foreach ($update in $WindowsUpdate) {
                if ($update.IsInstalled) {
                    Write-Verbose "Skipping update [$($update.Name)] because it is already installed."
                    continue
                }

                if (-not $update.IsEulaAccepted) {
                    if (-not $Force) {
                        throw "this could be problematic. Typically, only Feature updates require the EULA to be accepted"
                    }

                    $update.AcceptEula()
                }

                Write-Verbose -Message "Downloading and installing update [$($update.Name)]"

                $null = $updateColl.Add($update._updateObject)
            }

            [__ComObject]$updateDownloader = $updateSession.CreateUpdateDownloader()
            $updateDownloader.Updates = $updateColl

            Write-Verbose -Message "Downloading [$($updateColl.Updates.Count)] updates."
            [__ComObject]$downloadResults = $updateDownloader.Download()

            [__ComObject]$updateInstaller = $updateSession.CreateUpdateInstaller()
            $updateInstaller.Updates = $updateColl

            Write-Verbose -Message "Installing [$($updateColl.Updates.Count)] updates."
            [__ComObject]$installResults = $updateInstaller.Install()
        } finally {
            if ($null -ne $updateSession) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateSession)
            }
            if ($null -ne $updateColl) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateColl)
            }
            if ($null -ne $updateDownloader) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateDownloader)
            }
            if ($null -ne $downloadResults) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($downloadResults)
            }
            if ($null -ne $updateInstaller) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateInstaller)
            }
            if ($null -ne $installResults) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($installResults)
            }
        }

        return
    }

    static [Void] Commit() {
        try {
            [__ComObject]$private:updateSession = New-Object -ComObject Microsoft.Update.Session
            [__ComObject]$updateInstaller = $updateSession.CreateUpdateInstaller()

            $updateInstaller.Commit()
        } finally {
            if ($null -ne $updateSession) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateSession)
            }
            if ($null -ne $updateInstaller) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($updateInstaller)
            }
        }
    }
}

Update-TypeData -TypeName WindowsUpdate -MemberName Title -MemberType ScriptProperty -Value { return $this._updateObject.Title }
Update-TypeData -TypeName WindowsUpdate -MemberName Description -MemberType ScriptProperty -Value { return $this._updateObject.Description }
#Update-TypeData -TypeName WindowsUpdate -MemberName Category -MemberType ScriptProperty -Value { return (Select-Object -InputObject $this._updateObject -Property Categories) }
Update-TypeData -TypeName WindowsUpdate -MemberName Type -MemberType ScriptProperty -Value { return (Select-Object -InputObject $this._updateObject -Property Type) }
Update-TypeData -TypeName WindowsUpdate -MemberName IsDownloaded -MemberType ScriptProperty -Value { return $this._updateObject.IsDownloaded }
Update-TypeData -TypeName WindowsUpdate -MemberName IsInstalled -MemberType ScriptProperty -Value { return $this._updateObject.IsInstalled }
Update-TypeData -TypeName WindowsUpdate -MemberName IsPresent -MemberType ScriptProperty -Value { return $this._updateObject.IsPresent }
Update-TypeData -TypeName WindowsUpdate -MemberName IsRebootRequired -MemberType ScriptProperty -Value { return $this._updateObject.RebootRequired }
Update-TypeData -TypeName WindowsUpdate -MemberName IsHidden -MemberType ScriptProperty -Value { return $this._updateObject.IsHidden }
Update-TypeData -TypeName WindowsUpdate -MemberName IsMandatory -MemberType ScriptProperty -Value { return $this._updateObject.IsMandatory }
Update-TypeData -TypeName WindowsUpdate -MemberName IsEulaAccepted -MemberType ScriptProperty -Value { return $this._updateObject.EulaAccepted }
Update-TypeData -TypeName WindowsUpdate -MemberName EulaText -MemberType ScriptProperty -Value { return $this._updateObject.EulaText }
Update-TypeData -TypeName WindowsUpdate -MemberName UpdateID -MemberType ScriptProperty -Value { return $this._updateObject.Identity.UpdateID }

Update-TypeData -TypeName WindowsUpdate -DefaultDisplayPropertySet @(
    'Title',
    'Description'
    'Category',
    'Type',
    'IsDownloaded',
    'IsInstalled',
    'IsPresent',
    'IsRebootRequired',
    'IsHidden',
    'IsMandatory',
    'IsEulaAccepted',
    'EulaText'
)