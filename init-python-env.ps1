# ----------------------------------------------------------------------------
# -------------------------------- set up ------------------------------------
# ----------------------------------------------------------------------------
# 1. allow script execution
# 2. install chocolatey package manager
# 3. install git
# 4. create projects directory 
# 5. clone python environment set up repository
# ---------------------------------------------------------------------------- 
Set-ExecutionPolicy Bypass -Scope Process -Force

Write-Host -ForegroundColor Green "Installing $PROJECTS_PATH chocolatey package manager ..."
Unblock-File -Path 'C:\Users\WDAGUtilityAccount\analytics-env-set-up\init-python-env.ps1'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
Write-Host -ForegroundColor Green "chocolatey installed created!"

Write-Host -ForegroundColor Yellow "Installing git via  chocolatey..."
choco install git -y
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
Write-Host -ForegroundColor Green "git installed!"

$PROJECTS_PATH = "$HOME\projects"
$PYTHON_ENV_SET_UP_REPO = "analytics-env-set-up"
$JUPYTER_STARTER_KIT_REPO = "jupyter-starter-kit"

if (-Not (Test-Path -Path $PROJECTS_PATH)) {
    Write-Host -ForegroundColor Green "Creating $PROJECTS_PATH directory ..."
    New-Item -Path $PROJECTS_PATH -ItemType Directory
    Write-Host -ForegroundColor Green "Directory $PROJECTS_PATH created!"
} else {
    Write-Host -ForegroundColor Yellow "Directory $PROJECTS_PATH already exists!"
}

if (-Not (Test-Path -Path "$PROJECTS_PATH\$PYTHON_ENV_SET_UP_REPO")) {
    Write-Host -ForegroundColor Yellow "Cloning $PYTHON_ENV_SET_UP_REPO ..."
    Set-Location $PROJECTS_PATH
    git clone https://github.com/specialkapa/$PYTHON_ENV_SET_UP_REPO.git
    Write-Host -ForegroundColor Green "Cloned successfully!"
} else {
    Write-Host -ForegroundColor Green "Local $PYTHON_ENV_SET_UP_REPO repo already exists!"
}

# ----------------------------------------------------------------------------
# ---------------------------- install python --------------------------------
# ----------------------------------------------------------------------------
# 1. wipe existing python if exists
# 2. install python from scratch 
# 3. add it to PATH 
# ---------------------------------------------------------------------------- 

Set-Location $HOME

try {
    Get-Command python -ErrorAction Stop
    Write-Host -ForegroundColor Yellow "Uninstalling existing Python..."
    pipx uninstall ipython
    pipx uninstall poetry 
    python -m pip freeze | ForEach-Object { python -m pip uninstall -y $_.Split('==')[0] }
	Write-Host -ForegroundColor Green "existing python uninstalled!"
}
catch {
    Write-Host -ForegroundColor Red "Python command does not exist!"
}

Write-Host -ForegroundColor Yellow "Downloading and installing Python 3.11.2 ..."
Invoke-WebRequest https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe -OutFile python-3.11.2-amd64.exe
Start-Process -Wait -FilePath "python-3.11.2-amd64.exe" -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1"
Remove-Item "python-3.11.2-amd64.exe"
Write-Host -ForegroundColor Green "python 3.11.2 installed!"
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

$pythonExecutable = (Get-Command python | Select-Object -Property Source).Source
$pythonPath = (Get-Item $pythonExecutable).DirectoryName

Write-Host "Setting Python 3.11.2 as the global Python version..."
$env:Path = "$pythonPath;$pythonPath\Scripts;" + $env:Path
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

# ----------------------------------------------------------------------------
# ----------------------- install python packages ----------------------------
# ----------------------------------------------------------------------------
# 1. install pipx
# 2. install poetry via pipx
# 3. install ipython via pipx
# 4. install python packages from requirements.txt 
# 5. install chromium via playwright
# ---------------------------------------------------------------------------- 
Write-Host -ForegroundColor Yellow "Installing pipx ..."
python -m pip install pipx
python -m pipx ensurepath
Write-Host -ForegroundColor Green "pipx installed!"

Write-Host -ForegroundColor Yellow "Installing IPython and Poetry with pipx ..."
pipx install ipython
pipx install poetry
Write-Host -ForegroundColor Green "poetry and ipython have been installed!"

Write-Host -ForegroundColor Yellow "Installing python dependencies"
pip install -r "$PROJECTS_PATH\$PYTHON_ENV_SET_UP_REPO\requirements.txt"
Write-Host -ForegroundColor Green "python dependencies are installed!"

Write-Host -ForegroundColor Yellow "Installing chormium via playwright ..."
playwright install chromium
Write-Host -ForegroundColor Green "chromium is installed!"
Remove-Item -Path "$PROJECTS_PATH\$PYTHON_ENV_SET_UP_REPO" -Recurse -Force

# ----------------------------------------------------------------------------
# ----------------------- clone jupyter start kit ----------------------------
# ----------------------------------------------------------------------------
# Clone jupyter-starter-kit
# ---------------------------------------------------------------------------- 
Set-Location $PROJECTS_PATH
if ( (Test-Path -Path $PROJECTS_PA"$PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO")) {
    Remove-Item -Path "$PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO" -Recurse -Force
}
Write-Host -ForegroundColor Yellow "Cloning $JUPYTER_STARTER_KIT_REPO ..."
git clone https://github.com/specialkapa/$JUPYTER_STARTER_KIT_REPO.git
Write-Host -ForegroundColor Green "Cloned successfully!"

# ----------------------------------------------------------------------------
# ----------------------- create jupyter shortcuts ---------------------------
# ----------------------------------------------------------------------------
# Create jupyter notebook/lab shortcuts
# ---------------------------------------------------------------------------- 
function Generate-DesktopShortcut {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetPath,
        [Parameter(Mandatory=$true)]
        [string]$ShortcutName,
        [string]$Description = "My Shortcut Description",
        [string]$IconLocation = "",
        [string]$WorkingDirectory
    
    )
    
    $ShortcutPath = [System.Environment]::GetFolderPath('Desktop') + "\$ShortcutName.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.WindowStyle = 7
    $Shortcut.Description = $Description
    $Shortcut.IconLocation = if ($IconLocation -ne "") { $IconLocation } else { "$TargetPath, 0" }
    $Shortcut.WorkingDirectory = $WorkingDirectory
    $Shortcut.Save()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WScriptShell) | Out-Null
    Write-Output "Shortcut created at $ShortcutPath"
}

Write-Host -ForegroundColor Yellow "Creating jupyter shorcuts in Desktop ..."
$JupyterLabPath = (Get-Command jupyter-lab | Select-Object -Property Source).Source
$JupyterNotebookPath = (Get-Command jupyter-notebook | Select-Object -Property Source).Source

Generate-DesktopShortcut `
    -TargetPath $JupyterLabPath `
    -ShortcutName "jupyter-lab" `
    -Description "launches jupyter lab" `
    -IconLocation "$PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO\img\jupyter-logo.ico" `
    -WorkingDirectory $PROJECTS_PATH

Generate-DesktopShortcut `
    -TargetPath $JupyterNotebookPath `
    -ShortcutName "jupyter-notebook" `
    -Description "launches jupyter notebook" `
    -IconLocation "$PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO\img\jupyter-logo.ico" `
    -WorkingDirectory $PROJECTS_PATH

Generate-DesktopShortcut `
    -TargetPath $JupyterLabPath `
    -ShortcutName "jupyter-starter-kit-demo" `
    -Description "launches jupyter lab on the jupyter-starter-kit directory" `
    -IconLocation "$PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO\img\jupyter-logo.ico" `
    -WorkingDirectory $PROJECTS_PATH\$JUPYTER_STARTER_KIT_REPO

Write-Host -ForegroundColor Green "Jupyter shortcuts ready!"

# ----------------------------------------------------------------------------
# ----------------------- create jupyter shortcuts ---------------------------
# ----------------------------------------------------------------------------
# Display dialog box for new starters
# ---------------------------------------------------------------------------- 
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Info'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'WindowsDefaultLocation'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280, 40) 
$label.Text = "Look for the jupyter shortcuts on your Desktop. If you are new I suggest launching the jupyter-starter-kit-demo shortcut!"

$buttonHeight = 23
$buttonWidth = 100
$buttonX = ($form.ClientSize.Width - $buttonWidth) / 2
$buttonY = $form.ClientSize.Height - $buttonHeight - 10

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point($buttonX, $buttonY)
$button.Size = New-Object System.Drawing.Size($buttonWidth,$buttonHeight)
$button.Text = 'Launch Shortcut'

$button.Add_Click({
    $shortcutPath = "$env:USERPROFILE\Desktop\jupyter-starter-kit-demo.lnk" 
    if (Test-Path $shortcutPath) {
        Start-Process $shortcutPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Shortcut not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Controls.Add($label)
$form.Controls.Add($button)
$form.ShowDialog()

exit