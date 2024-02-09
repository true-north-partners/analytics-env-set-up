Write-Host -ForegroundColor Blue `
"
                   __       __   __
.---.-.-----.---.-|  .--.--|  |_|__.----.-----.
|  _  |     |  _  |  |  |  |   _|  |  __|__ --|
|___._|__|__|___._|__|___  |____|__|____|_____|
                     |_____|
"

Write-Host -ForegroundColor Green `
"
                    __                                         __   
.-----.-----.--.--.|__|.----.-----.-----.--------.-----.-----.|  |_ 
|  -__|     |  |  ||  ||   _|  _  |     |        |  -__|     ||   _|
|_____|__|__|\___/ |__||__| |_____|__|__|__|__|__|_____|__|__||____|
                                                                    
"

Write-Host -ForegroundColor Blue `
"
        
.-----.-----.|  |_.--.--.-----.
|__ --|  -__||   _|  |  |  _  |
|_____|_____||____|_____|   __|
                        |__|                                                           
"


$WELCOME = "This repository is dedicated to hosting a PowerShell script designed 
to be used on a Windows sandbox during startup. The purpose of the script is to 
automate the setup of a Python environment tailored for analytics work. This is 
particularly useful for projects requiring a consistent and automated setup 
process for analytics environments on a sandbox.

This script is designed to be used in conjunction with the Windows Sandbox for 
now given it is still experimental. Eventually this project will be modifed to 
work on a regular Windows environment.

For further details visit https://github.com/true-north-partners/analytics-env-set-up. 
"

$MESSAGE = "Sit back and look at the magic happening! I all I need to do is 
double click the spin-up-sandbox.bat batch file, sit back, grab a beer and let 
the magic happen. Specifically, the script will automate the following tasks:
   - Creation of a Python virtual environment.
   - Installation of Python packages specified in requirements.txt.
   - Cloning of https://bitbucket.org/ghb-credit-risk/jupyter-starter-kit 
     repository.
   - Creation of jupyter shortcuts
"
./write-typewritter.ps1 $WELCOME -speed 10
./write-typewritter.ps1 $MESSAGE -speed 10

 