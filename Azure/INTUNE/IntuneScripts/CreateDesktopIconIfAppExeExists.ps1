#
# Update the folderPath and the executableName for the program you want to create an icon on the desktop for.
#

param (
    [string]$folderPath = "C:\\Program Files (x86)\\Nuance\\Dragon Medical One 2024",
    [string]$executableName = "sod.exe"
)

function Find-ExecutableAndCreateShortcut {
    try {
        # Combine the folder path and executable name
        $executablePath = Join-Path -Path $folderPath -ChildPath $executableName

        if (Test-Path $executablePath) {
            Write-Output "Found $executableName in $folderPath"

            # Get the Desktop path
            $desktop = [Environment]::GetFolderPath("Desktop")
            $shortcutPath = Join-Path -Path $desktop -ChildPath "$($executableName -replace '\.exe$', '').lnk"

            # Create a shortcut
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $executablePath
            $shortcut.WorkingDirectory = $folderPath
            $shortcut.IconLocation = $executablePath
            $shortcut.Save()

            Write-Output "Shortcut created on Desktop: $shortcutPath"
        } else {
            Write-Output "$executableName not found in $folderPath"
        }
    } catch {
        Write-Output "An error occurred: $_"
    }
}

# Execute the function
Find-ExecutableAndCreateShortcut
